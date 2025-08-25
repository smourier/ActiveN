namespace ActiveN.Hosting;

public class ComRegistrationContext(BaseComRegistration registration, RegistryKey registryRoot)
{
    public BaseComRegistration Registration { get; } = registration;
    public RegistryKey RegistryRoot { get; } = registryRoot;
}
