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
}
