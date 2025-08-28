namespace ActiveN.Hosting;

public class ComRegistrationContext(
    ComRegistration registration,
    RegistryKey registryRoot,
    ComRegistrationType type)
{
    public ComRegistration Registration { get; } = registration;
    public RegistryKey RegistryRoot { get; } = registryRoot;
    public ComRegistrationType Type { get; } = type;
    public virtual Guid GUID { get; set; } = type.Type.GUID;

    // not sure what a good default is here
    public virtual OLEMISC MiscStatus { get; set; } = OLEMISC.OLEMISC_RECOMPOSEONRESIZE | OLEMISC.OLEMISC_CANTLINKINSIDE | OLEMISC.OLEMISC_INSIDEOUT | OLEMISC.OLEMISC_ACTIVATEWHENVISIBLE | OLEMISC.OLEMISC_SETCLIENTSITEFIRST;
    public virtual TypeLib? TypeLib { get; set; }

    public virtual IList<Guid> ImplementedCategories { get; } =
    [
        ControlCategories.CATID_ActiveN,
        ControlCategories.CATID_Programmable,
        ControlCategories.CATID_Insertable,
        ControlCategories.CATID_SafeForScripting,
        ControlCategories.CATID_SafeForInitializing,
        ControlCategories.CATID_Control,
    ];
}
