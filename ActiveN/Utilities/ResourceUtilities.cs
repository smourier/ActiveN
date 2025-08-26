namespace ActiveN.Utilities;

public static class ResourceUtilities
{
    public static bool HasEmbeddedTypeLib(string filePath, int id = 1) => HasEmbeddedResource(filePath, "TYPELIB", id);
    public static bool HasEmbeddedResource(string filePath, string resourceType, int id = 1)
    {
        ArgumentNullException.ThrowIfNull(filePath);
        ArgumentNullException.ThrowIfNull(resourceType);
        var module = Functions.LoadLibraryExW(PWSTR.From(filePath), 0, LOAD_LIBRARY_FLAGS.LOAD_LIBRARY_AS_DATAFILE | LOAD_LIBRARY_FLAGS.LOAD_LIBRARY_AS_IMAGE_RESOURCE);
        if (module == 0)
            return false;

        try
        {
            // 1 here matches csproj's <TypeLibRc> number before TYPELIB
            return Functions.FindResourceW(module, new PWSTR(id), PWSTR.From(resourceType)) != 0;
        }
        finally
        {
            Functions.FreeLibrary(module);
        }
    }

    public static bool HasEmbeddedManifest(string filePath, int id = 1) => HasEmbeddedResource(filePath, Constants.RT_MANIFEST.Value, id);
    public static bool HasEmbeddedBitmap(string filePath, int id = 1) => HasEmbeddedResource(filePath, Constants.RT_BITMAP.Value, id);
    public static bool HasEmbeddedResource(string filePath, nint resourceType, int id = 1)
    {
        ArgumentNullException.ThrowIfNull(filePath);
        var module = Functions.LoadLibraryExW(PWSTR.From(filePath), 0, LOAD_LIBRARY_FLAGS.LOAD_LIBRARY_AS_DATAFILE | LOAD_LIBRARY_FLAGS.LOAD_LIBRARY_AS_IMAGE_RESOURCE);
        if (module == 0)
            return false;

        try
        {
            // 1 here matches csproj's <TypeLibRc> number before TYPELIB
            return Functions.FindResourceW(module, new PWSTR(id), new PWSTR(resourceType)) != 0;
        }
        finally
        {
            Functions.FreeLibrary(module);
        }
    }
}
