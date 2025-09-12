// Copyright (c) Aelyo Softworks S.A.S.. All rights reserved.
// See LICENSE in the project root for license information.

namespace ActiveN.Hosting;

public abstract partial class ComRegistration
{
    private readonly Lazy<string> _dllPath;
    private readonly ConcurrentDictionary<Guid, ClassFactory> _classFactories = new();

    [UnconditionalSuppressMessage("Trimming", "IL2026:Members annotated with 'RequiresUnreferencedCodeAttribute' require dynamic access otherwise can break functionality when trimming application code", Justification = "Will mostly work in debug only")]
    protected ComRegistration(IReadOnlyList<ComRegistrationType> comTypes)
    {
        ArgumentNullException.ThrowIfNull(comTypes);
        if (comTypes.Count == 0)
            throw new ArgumentException("At least one COM type must be specified.", nameof(comTypes));

        ComTypes = comTypes;
        Application.CanShowFatalError = false;
        Application.ExitAllOnLastWindowRemoved = false;

        // add all known assemblies that might have GUIDs to trace
        var assemblies = new HashSet<Assembly>
        {
            typeof(IUnknown).Assembly, // DirectN
            typeof(Window).Assembly, // DirectN.Extensions
            Assembly.GetExecutingAssembly(),
        };

        var entry = Assembly.GetEntryAssembly();
        if (entry != null)
        {
            assemblies.Add(entry);
        }

        foreach (var type in comTypes)
        {
            assemblies.Add(type.Type.Assembly);
        }

        foreach (var asm in assemblies)
        {
            GuidNames.AddClassesGuids(asm);
        }

        _dllPath = new Lazy<string>(GetDllPath);

        //TracingUtilities.TraceToFile = true;
        var process = SystemUtilities.CurrentProcess;
        if (process.ProcessName.EqualsIgnoreCase("devenv"))
        {
            TracingUtilities.TraceToFile = true; // enable tracing when running from Visual Studio
        }

        TracingUtilities.Trace($"Path: {DllPath} bitness: {nint.Size} types: {string.Join(", ", ComTypes.Select(t => t.Type.FullName))} entry asm: '{DllPath}' process: {process.MainModule?.FileName ?? process.ProcessName} lowbox: {SystemUtilities.IsAppContainer()}");
        RegisterEmbeddedTypeLib = ResourceUtilities.HasEmbeddedTypeLib(DllPath);
        TracingUtilities.Trace($"RegisterEmbeddedTypeLib: '{RegisterEmbeddedTypeLib}'");
    }

    public IReadOnlyList<ComRegistrationType> ComTypes { get; }
    public IDictionary<Guid, ClassFactory> ClassFactories => _classFactories;
    public string DllPath => _dllPath.Value;
    public virtual bool RegisterEmbeddedTypeLib { get; protected set; }
    public virtual ComRegistrationTarget Target { get; protected set; }
    public virtual bool CanUnload { get; protected set; }
    public string? ThunkDllPath { get; protected set; }
    public string RegistrationDllPath => ThunkDllPath ?? DllPath;

#if DIRECTN_DEBUG
    public bool EnableComObjectTraces
    {
        get => DirectN.Extensions.Com.ComObject.EnableTraces;
        set => DirectN.Extensions.Com.ComObject.EnableTraces = value;
    }
#endif

    protected virtual internal HRESULT CreateInstance(ClassFactory classFactory, out object? instance)
    {
        ArgumentNullException.ThrowIfNull(classFactory);
        instance = null;
        foreach (var comType in ComTypes)
        {
            if (classFactory.Clsid == comType.Type.GUID)
            {
                var ctor = comType.Type.GetConstructor(Type.EmptyTypes);
                if (ctor != null)
                {
                    instance = ctor.Invoke(null);
                    return Constants.S_OK;
                }
            }
        }

        return Constants.E_NOINTERFACE;
    }

    protected virtual ComRegistrationContext CreateRegistrationContext(RegistryKey root, ComRegistrationType type) => new(this, root, type);
    protected virtual ClassFactory CreateClassFactory(Guid clsid) => new(clsid, this);

    protected virtual HRESULT RegisterServer() => TracingUtilities.WrapErrors(() =>
    {
        TracingUtilities.Trace($"Path: {DllPath}");

        TypeLib? typeLib = null;
        if (RegisterEmbeddedTypeLib)
        {
            typeLib = TypeLib.Load(DllPath);
            if (typeLib != null)
            {
                typeLib.Register(Target);
                TracingUtilities.Trace($"Registered type library {typeLib.TypeLibId} '{typeLib.Name}' ('{typeLib.Documentation}') version {typeLib.MajorVersion}.{typeLib.MinorVersion}");
            }
        }

        var root = GetRegistryRoot(Target);
        foreach (var type in ComTypes)
        {
            var ctx = CreateRegistrationContext(GetRegistryRoot(Target), type);
            ctx.TypeLib = typeLib;
            RegisterType(ctx);
        }
        return Constants.S_OK;
    }, () => UnregisterServer());

    protected virtual void RegisterType(ComRegistrationContext context)
    {
        ArgumentNullException.ThrowIfNull(context);
        TracingUtilities.Trace($"Type: {context.Type.Type.FullName} guid: {context.Clsid}");
        RegisterInProcessComObject(context);

        if (RegisterEmbeddedTypeLib)
        {
            context.TypeLib?.RegisterForCoClass(context);
        }

        var miscStatus = context.MiscStatus;
        var method = context.Type.Type.GetMethod(nameof(RegisterType), BindingFlags.Public | BindingFlags.Static);
        if (method != null)
        {
            var misc = context.Type.Type.GetCustomAttribute<MiscStatusAttribute>();
            if (misc != null)
            {
                context.MiscStatus = misc.Value;
            }

            method.Invoke(null, [context]);
        }
        context.MiscStatus = miscStatus;
    }

    public static RegistryKey GetRegistryRoot(ComRegistrationTarget target)
    {
        if (target == ComRegistrationTarget.HKEY_CURRENT_USER)
            return Registry.CurrentUser;

        if (target == ComRegistrationTarget.HKEY_LOCAL_MACHINE)
            return Registry.LocalMachine;

        return SystemUtilities.GetTokenElevationType() == TokenElevationType.Limited ? Registry.CurrentUser : Registry.LocalMachine;
    }

    protected virtual HRESULT UnregisterServer() => TracingUtilities.WrapErrors(() =>
    {
        TracingUtilities.Trace($"Path: {DllPath}");
        var root = GetRegistryRoot(Target);

        TypeLib? typeLib = null;
        if (RegisterEmbeddedTypeLib)
        {
            typeLib = TypeLib.Load(DllPath);
            if (typeLib != null)
            {
                // don't throw on error here, we want to continue unregistering other stuff
                var hr = typeLib.Unregister(Target, false);
                TracingUtilities.Trace($"Unregistered type library {typeLib.TypeLibId} '{typeLib.Name}' ('{typeLib.Documentation}') version {typeLib.MajorVersion}.{typeLib.MinorVersion} hr: {hr}");
            }
        }

        foreach (var type in ComTypes)
        {
            var ctx = CreateRegistrationContext(GetRegistryRoot(Target), type);
            ctx.TypeLib = typeLib;
            UnregisterType(ctx);
        }
        return Constants.S_OK;
    });

    // unregistering should not throw but try to continue
    protected virtual void UnregisterType(ComRegistrationContext context)
    {
        ArgumentNullException.ThrowIfNull(context);
        UnregisterComObject(context);

        if (RegisterEmbeddedTypeLib)
        {
            context.TypeLib?.UnregisterForCoClass(context.Clsid, context.RegistryRoot);
        }

        var method = context.Type.Type.GetMethod(nameof(UnregisterType), BindingFlags.Public | BindingFlags.Static);
        method?.Invoke(null, [context]);
    }

    protected virtual object? GetClassObject(Guid clsid, Guid iid)
    {
        TracingUtilities.Trace($"Path: {DllPath} clsid: {clsid} iid: {iid.GetName()}");
        foreach (var type in ComTypes)
        {
            TracingUtilities.Trace($"Type: {type.Type.FullName} guid: {type.Type.GUID}");
            if (clsid == type.Type.GUID &&
                (iid == typeof(IClassFactory).GUID || iid == typeof(IUnknown).GUID))
            {
                if (!_classFactories.TryGetValue(clsid, out var classFactory))
                {
                    classFactory = CreateClassFactory(clsid);
                    _classFactories[clsid] = classFactory;
                }
                TracingUtilities.Trace($"ClassFactory: {classFactory}");
                return classFactory;
            }
        }
        return null;
    }

    protected unsafe HRESULT GetClassObject(nint rclsid, nint riid, nint ppv) => TracingUtilities.WrapErrors(() =>
    {
        if (rclsid == 0 || riid == 0 || ppv == 0)
            return Constants.E_POINTER;

        var clsid = *(Guid*)rclsid;
        var iid = *(Guid*)riid;
        var obj = GetClassObject(clsid, iid);
        var unk = DirectN.Extensions.Com.ComObject.GetOrCreateComInstance(obj, iid);
        *(nint*)ppv = unk;
        TracingUtilities.Trace($"unk: {obj}");
        return unk == 0 ? Constants.E_NOINTERFACE : Constants.S_OK;
    });

    protected virtual HRESULT CanUnloadNow() => TracingUtilities.WrapErrors(() =>
    {
        TracingUtilities.Trace($"Path: {DllPath} CanUnload: {CanUnload}");
        TracingUtilities.Flush(); // host asking to quit, it's a good time to flush traces
        return CanUnload ? (uint)Constants.S_OK : (uint)Constants.S_FALSE;
    });

    // call for example
    // "regsvr32 /i:user /n" to register in HKCU without calling DllRegisterServer
    // "regsvr32 /i:user /n /u" to unregister from HKCU
    protected virtual HRESULT Install(bool install, nint cmdLinePtr) => TracingUtilities.WrapErrors(() =>
    {
        TracingUtilities.Trace($"Path: {DllPath}");
        if (cmdLinePtr != 0)
        {
            var cmdLine = Marshal.PtrToStringUni(cmdLinePtr);
            TracingUtilities.Trace($"CmdLine: {cmdLine}");
            if (cmdLine.EqualsIgnoreCase("user"))
            {
                Target = ComRegistrationTarget.HKEY_CURRENT_USER;
            }
            else if (cmdLine.EqualsIgnoreCase("machine") || cmdLine.EqualsIgnoreCase("admin"))
            {
                Target = ComRegistrationTarget.HKEY_LOCAL_MACHINE;
            }
        }

        return install ? RegisterServer() : UnregisterServer();
    }, () =>
    {
        if (install)
        {
            UnregisterServer();
        }
    });

    protected HRESULT ThunkInit(nint thunkDllPathPtr) => TracingUtilities.WrapErrors(() =>
    {
        ThunkDllPath = Marshal.PtrToStringUni(thunkDllPathPtr);
        TracingUtilities.Trace($"Path: {DllPath} ThunkDllPathPtr: {ThunkDllPath}");
        return 0;
    });

    public virtual void RegisterInProcessComObject(ComRegistrationContext context)
    {
        ArgumentNullException.ThrowIfNull(context);

        var threadingModel = context.ThreadingModel?.Nullify() ?? "Both";
        TracingUtilities.Trace($"Registering {context.FullName} from {RegistrationDllPath} with threading model '{threadingModel}'...");
        using var serverKey = EnsureWritableSubKey(context.RegistryRoot, Path.Combine(context.ClsidRegistryKey, context.Clsid.ToString("B"), "InprocServer32"));
        serverKey.SetValue(null, RegistrationDllPath);
        serverKey.SetValue("ThreadingModel", threadingModel);

        // ProgId is optional
        var progid = context.Type.Type.GetCustomAttribute<ProgIdAttribute>()?.Value.Nullify();
        if (progid != null)
        {
            using var key = EnsureWritableSubKey(context.RegistryRoot, Path.Combine(context.ClsidRegistryKey, context.Clsid.ToString("B")));
            using var progIdKey = EnsureWritableSubKey(key, "ProgId");
            progIdKey.SetValue(null, progid);

            using var viProgIdKey = EnsureWritableSubKey(key, "VersionIndependentProgID");
            viProgIdKey.SetValue(null, progid);

            using var ckey = EnsureWritableSubKey(context.RegistryRoot, Path.Combine(context.ClassesRegistryKey, progid, "CLSID"));
            ckey.SetValue(null, context.Clsid.ToString("B"));
        }

        var name = context.Type.Type.GetCustomAttribute<DisplayNameAttribute>()?.DisplayName.Nullify();
        if (name != null)
        {
            using var key = EnsureWritableSubKey(context.RegistryRoot, Path.Combine(context.ClsidRegistryKey, context.Clsid.ToString("B")));
            key.SetValue(null, name);
        }
        TracingUtilities.Trace($"Registered {context.FullName}.");
    }

    public virtual void UnregisterComObject(ComRegistrationContext context)
    {
        ArgumentNullException.ThrowIfNull(context);

        TracingUtilities.Trace($"Unregistering {context.FullName}...");
        using var key = context.RegistryRoot.OpenSubKey(context.ClsidRegistryKey, true);
        key?.DeleteSubKeyTree(context.Clsid.ToString("B"), false);

        // ProgId is optional
        var att = context.Type.Type.GetCustomAttribute<ProgIdAttribute>();
        if (att != null && !string.IsNullOrWhiteSpace(att.Value))
        {
            var progid = att.Value.Trim();
            using var ckey = context.RegistryRoot.OpenSubKey(context.ClassesRegistryKey, true);
            ckey?.DeleteSubKeyTree(progid, false);
        }

        TracingUtilities.Trace($"Unregistered {context.FullName}.");
    }

    public static RegistryKey EnsureWritableSubKey(RegistryKey root, string name)
    {
        var key = root.OpenSubKey(name, true);
        if (key != null)
            return key;

        var parentName = Path.GetDirectoryName(name);
        if (string.IsNullOrEmpty(parentName))
            return root.CreateSubKey(name);

        using var parentKey = EnsureWritableSubKey(root, parentName);
        return parentKey.CreateSubKey(Path.GetFileName(name));
    }

#if DEBUG
    // this only works if *not* published as AOT

    [UnconditionalSuppressMessage("SingleFile", "IL3000:Avoid accessing Assembly file path when publishing as a single file", Justification = "Only used when not published as AOT")]
    protected virtual string GetDllPath() => GetType().Assembly.Location;

#else
    // this only works if published as AOT

    private static unsafe delegate* unmanaged<uint> GetModuleFunctionPointer() => &DoNothing; // any func pointer will do

    // this won't really be exported, just used to fool AOT compiler to get a valid module function pointer
    [UnmanagedCallersOnly(EntryPoint = nameof(DoNothing))]
    public static uint DoNothing() => 0;

    protected virtual unsafe string GetDllPath()
    {
        var ptr = GetModuleFunctionPointer();
        const int GET_MODULE_HANDLE_EX_FLAG_FROM_ADDRESS = 4;
        const int GET_MODULE_HANDLE_EX_FLAG_UNCHANGED_REFCOUNT = 2;
        if (!Functions.GetModuleHandleExW(GET_MODULE_HANDLE_EX_FLAG_FROM_ADDRESS | GET_MODULE_HANDLE_EX_FLAG_UNCHANGED_REFCOUNT, new PWSTR((nint)ptr), out var module))
            throw new Win32Exception(Marshal.GetLastPInvokeError());

        var size = 256; // 128 wchars
        do
        {
            using var pwstr = new AllocPwstr(size);
            var msize = Functions.GetModuleFileNameW(module, pwstr, pwstr.SizeInChars);
            // note we could check for ERROR_INSUFFICIENT_BUFFER but it's not always set for some reason (for example, under Excel, very weird)
            if (msize != 0 && msize < pwstr.SizeInChars) // if msize == insize, it means the buffer was too small
                return pwstr.ToString()!;

            size *= 2;
        }
        while (true);
    }

#endif
}
