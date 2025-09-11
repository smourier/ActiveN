namespace ActiveN.Utilities;

// Winforms .NET Core is not yet AOT-capable, we cannot use its PropertyGrid yet, so we use .NET Framework's PropertyGrid for now
// since it's always installed and doesn't require any additional dependencies.
// Note that if the process is a .NET Framework process, it's possible that the .NET Core Windows Forms will be loaded automatically instead of .NET Framework's.
// In this case, custom attributes like [Category("blah")] will not work as expected because .NET Core's property grid will mix things up.
public class PropertyGrid
{
    public const string AssemblyName = "ActiveN.PropertyGrid";
    public const string TypeName = AssemblyName + ".PropertiesForm";
    public const string MethodName = "Show";

    private static bool _netFxPropertyGridFilesEnsured;
    public static string EnsureNetFxPropertyGridFiles(bool force = false)
    {
        var fileVersion = Assembly.GetExecutingAssembly().GetFileVersion();
        var path = Path.Combine(Path.GetTempPath(), $"ActiveN{fileVersion}", $"{AssemblyName}.dll");
        if (_netFxPropertyGridFilesEnsured && !force)
            return path;

        using var stream = Assembly.GetExecutingAssembly().GetManifestResourceStream($"{typeof(PropertyGrid).Namespace}.{AssemblyName}.dll") ?? throw new InvalidOperationException();
        TracingUtilities.Trace($"Ensuring {path}");

        var needsCopy = !File.Exists(path) || new FileInfo(path).Length != stream.Length;
#if DEBUG
        needsCopy = true;
#endif

        if (needsCopy)
        {
            var dir = Path.GetDirectoryName(path);
            if (!Directory.Exists(dir))
            {
                Directory.CreateDirectory(Path.GetDirectoryName(path) ?? throw new InvalidOperationException());
            }

            FileSystemUtilities.WrapSharingViolations(() =>
            {
                using var fileStream = File.Create(path);
                stream.CopyTo(fileStream);
            }, (ex, i) =>
            {
                if (i == FileSystemUtilities.DefaultWrapSharingViolationsRetryCount - 1)
                {
                    TracingUtilities.Trace($"Error copying {path}: {ex}");
                    return false;
                }

                Thread.Sleep(FileSystemUtilities.DefaultWrapSharingViolationsWaitTime);
                return true;
            });
        }

        using var pdbStream = Assembly.GetExecutingAssembly().GetManifestResourceStream($"{typeof(PropertyGrid).Namespace}.{AssemblyName}.pdb");
        if (pdbStream != null)
        {
            var pdbPath = Path.Combine(Path.GetTempPath(), $"ActiveN{fileVersion}", $"{AssemblyName}.pdb");
            TracingUtilities.Trace($"Ensuring {pdbPath}");

            needsCopy = !File.Exists(pdbPath) || new FileInfo(pdbPath).Length != pdbStream.Length;
#if DEBUG
            needsCopy = true;
#endif

            if (needsCopy)
            {
                // same directory as .dll
                FileSystemUtilities.WrapSharingViolations(() =>
                {
                    using var fileStream = File.Create(pdbPath);
                    stream.CopyTo(fileStream);
                },
                (ex, i) =>
                {
                    if (i == FileSystemUtilities.DefaultWrapSharingViolationsRetryCount - 1)
                    {
                        TracingUtilities.Trace($"Error copying {pdbPath}: {ex}");
                        return false;
                    }

                    Thread.Sleep(FileSystemUtilities.DefaultWrapSharingViolationsWaitTime);
                    return true;
                });
            }
        }

        _netFxPropertyGridFilesEnsured = File.Exists(path) && new FileInfo(path).Length == stream.Length;
        return path;
    }

    private static ClrRuntimeHost? GetHost(bool useLoaded = true)
    {
        using var host = new ClrHost();
        if (useLoaded)
        {
            foreach (var rt in host.EnumerateLoadedRuntimes(Process.GetCurrentProcess()))
            {
                TracingUtilities.Trace($"Found loaded runtime: {rt.Version} started: {rt.IsStarted}");
                if (rt.Version?.StartsWith("v4") == true)
                {
                    var rth = rt.GetHost();
                    TracingUtilities.Trace($" host status: {rt.IsStarted}");
                    if (rt.IsStarted == null)
                    {
                        rth.Start();
                    }

                    rt.Dispose();
                    return rth;
                }
                rt.Dispose();
            }
        }

        foreach (var rt in host.EnumerateInstalledRuntimes())
        {
            TracingUtilities.Trace($"Found installed runtime: {rt.Version}");
            if (rt.Version?.StartsWith("v4") == true)
            {
                var rth = rt.GetHost();
                TracingUtilities.Trace($" host status: {rt.IsStarted}");
                if (rt.IsStarted == null)
                {
                    rth.Start();
                }
                rt.Dispose();
                return rth;
            }
            rt.Dispose();
        }
        return null;
    }

    public static HRESULT Show(HWND parent, RECT rect, object obj, bool throwOnError = true)
    {
        ArgumentNullException.ThrowIfNull(obj);
        var path = EnsureNetFxPropertyGridFiles();
        using var host = GetHost() ?? throw new InvalidOperationException("Cannot find any .NET Framework runtime");
        TracingUtilities.Trace($"path: {path} host: {host} typeName: {TypeName} methodName: {MethodName}");

        // add object to SPM so the NET Framework PropertyGrid can find it
        using var spm = ComObject<ISharedPropertyGroupManager>.CoCreate(Constants.CLSID_SharedPropertyGroupManager) ?? throw new InvalidOperationException("Cannot create SharedPropertyGroupManager");
        using var groupName = new Bstr("ActiveN");
        var isoMode = (int)LockModes.LockSetGet;
        var releaseModes = (int)ReleaseModes.Standard;

        spm.Object.CreatePropertyGroup(groupName, ref isoMode, ref releaseModes, out var gexists, out var spgObj).ThrowOnError(throwOnError);
        using var spg = new ComObject<ISharedPropertyGroup>(spgObj);

        using var propertyName = new Bstr("SelectedObject");
        spg.Object.CreateProperty(propertyName, out var pexists, out var propObj).ThrowOnError(throwOnError);
        using var prop = new ComObject<ISharedProperty>(propObj);

        return DirectN.Extensions.Com.ComObject.WithComInstance(obj, unk =>
        {
            using var v = new Variant(unk, VARENUM.VT_UNKNOWN);
            prop.Object.put_Value(v.Detached).ThrowOnError(throwOnError);

            HRESULT hr = host.ExecuteInDefaultAppDomain(path, TypeName, MethodName, $"parent:{parent.Value}|rect:{rect}");
            hr.ThrowOnError(throwOnError);
            return hr;
        });
    }
}
