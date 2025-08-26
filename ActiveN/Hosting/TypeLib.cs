namespace ActiveN.Hosting;

public sealed class TypeLib
{
    private TypeLib(string filePath, Guid typeLibId)
    {
        FilePath = filePath ?? throw new ArgumentNullException(nameof(filePath));
        TypeLibId = typeLibId;
    }

    public string FilePath { get; }
    public Guid TypeLibId { get; }
    public uint Lcid { get; private set; }
    public SYSKIND SysKind { get; private set; }
    public ushort MajorVersion { get; private set; }
    public ushort MinorVersion { get; private set; }
    public ushort LibFlags { get; private set; }
    public string? Name { get; private set; }
    public string? Documentation { get; private set; }

    public override string ToString() => $"{Name} ({TypeLibId}, v{MajorVersion}.{MinorVersion}, lcid={Lcid}, syskind={SysKind}, flags=0x{LibFlags:X4})";

    private static ComObject<ITypeLib> LoadTypeLib(string filePath, bool throwOnError = true)
    {
        ArgumentNullException.ThrowIfNull(filePath);
        Functions.LoadTypeLibEx(PWSTR.From(filePath), REGKIND.REGKIND_NONE, out var obj).ThrowOnError(throwOnError);
        return new ComObject<ITypeLib>(obj);
    }

    public void RegisterForCoClass(Guid clsid, RegistryKey registryRoot)
    {
        ArgumentNullException.ThrowIfNull(registryRoot);
        using var key = ComRegistration.EnsureWritableSubKey(registryRoot, Path.Combine(ComRegistration.ClsidRegistryKey, clsid.ToString("B")));
        using var typeLib = key.CreateSubKey("TypeLib", true);
        typeLib.SetValue(null, TypeLibId.ToString("B"));
    }

#pragma warning disable CA1822 // Mark members as static
#pragma warning disable IDE0060 // Remove unused parameter
    public void UnregisterForCoClass(Guid clsid, RegistryKey registryRoot)
#pragma warning restore IDE0060 // Remove unused parameter
#pragma warning restore CA1822 // Mark members as static
    {
        ArgumentNullException.ThrowIfNull(registryRoot);
        // nothing to do here, the whole CLSID key is removed by ComRegistration
    }

    public bool Register(bool perUser, bool throwOnError = true)
    {
        using var typeLib = LoadTypeLib(FilePath, throwOnError);
        if (typeLib == null)
            return false;

        if (perUser)
        {
            Functions.RegisterTypeLibForUser(typeLib.Object, PWSTR.From(FilePath), PWSTR.Null).ThrowOnError(throwOnError);
        }
        else
        {
            Functions.RegisterTypeLib(typeLib.Object, PWSTR.From(FilePath), PWSTR.Null).ThrowOnError(throwOnError);
        }
        return true;
    }

    public void UnregisterTypeLib(bool perUser, bool throwOnError = true)
    {
        if (perUser)
        {
            Functions.UnRegisterTypeLibForUser(TypeLibId, MajorVersion, MinorVersion, Lcid, SysKind).ThrowOnError(throwOnError);
        }
        else
        {
            Functions.UnRegisterTypeLib(TypeLibId, MajorVersion, MinorVersion, Lcid, SysKind).ThrowOnError(throwOnError);
        }
    }

    public static bool HasEmbeddedTypeLib(string filePath)
    {
        ArgumentNullException.ThrowIfNull(filePath);
        var module = Functions.LoadLibraryExW(PWSTR.From(filePath), 0, LOAD_LIBRARY_FLAGS.LOAD_LIBRARY_AS_DATAFILE | LOAD_LIBRARY_FLAGS.LOAD_LIBRARY_AS_IMAGE_RESOURCE);
        if (module == 0)
            return false;

        try
        {
            // 1 here matches csproj's <TypeLibRc> number before TYPELIB
            return Functions.FindResourceW(module, new PWSTR(1), PWSTR.From("TYPELIB")) != 0;
        }
        finally
        {
            Functions.FreeLibrary(module);
        }
    }

    public static unsafe TypeLib? Load(string filePath, bool throwOnError = true)
    {
        using var typeLib = LoadTypeLib(filePath, throwOnError);
        if (typeLib == null)
            return null;

        typeLib.Object.GetLibAttr(out var ptr).ThrowOnError(throwOnError);
        if (ptr == 0)
            return null;

        var attr = *(TLIBATTR*)ptr;
        var info = new TypeLib(filePath, attr.guid)
        {
            Lcid = attr.lcid,
            SysKind = attr.syskind,
            MajorVersion = attr.wMajorVerNum,
            MinorVersion = attr.wMinorVerNum,
            LibFlags = attr.wLibFlags
        };
        typeLib.Object.ReleaseTLibAttr(ptr);

        BSTR name;
        BSTR doc;
        typeLib.Object.GetDocumentation(-1, (nint)(&name), (nint)(&doc), out _, 0);
        info.Name = name.ToString();
        info.Documentation = doc.ToString();
        if (name.Value != 0) BSTR.Dispose(ref name);
        if (doc.Value != 0) BSTR.Dispose(ref doc);
        return info;
    }
}
