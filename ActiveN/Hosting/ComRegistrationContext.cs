// Copyright (c) Aelyo Softworks S.A.S.. All rights reserved.
// See LICENSE in the project root for license information.

namespace ActiveN.Hosting;

public class ComRegistrationContext(
    ComRegistration registration,
    RegistryKey registryRoot,
    ComRegistrationType type)
{
    public ComRegistration Registration { get; } = registration;
    public RegistryKey RegistryRoot { get; } = registryRoot;
    public ComRegistrationType Type { get; } = type;
    public bool PerUser => RegistryRoot == Registry.CurrentUser;
    public Guid Clsid => Type.Type.GUID;
    public string? FullName => Type.Type.FullName;
    public string ClassesRegistryKey => @"Software\Classes";
    public string ClsidRegistryKey => ClassesRegistryKey + @"\CLSID";

    // not sure what a good default is here
    public virtual OLEMISC MiscStatus { get; set; } = OLEMISC.OLEMISC_RECOMPOSEONRESIZE | OLEMISC.OLEMISC_CANTLINKINSIDE | OLEMISC.OLEMISC_INSIDEOUT | OLEMISC.OLEMISC_ACTIVATEWHENVISIBLE | OLEMISC.OLEMISC_SETCLIENTSITEFIRST;
    public virtual TypeLib? TypeLib { get; set; }

    // default is "Both" note: not all of the code currently supports full multithreading
    // also note aggregation doesn't work if server & client threading models are not compatible (MTA vs STA and the reverse)
    public virtual string? ThreadingModel { get; protected set; }

    public virtual IList<Guid> ImplementedCategories { get; } =
    [
        ControlCategories.CATID_ActiveN,
        ControlCategories.CATID_Programmable,
        ControlCategories.CATID_Insertable,
        ControlCategories.CATID_SafeForScripting,
        ControlCategories.CATID_SafeForInitializing,
        ControlCategories.CATID_ActiveXControls,
        ControlCategories.CATID_Control,
    ];
}
