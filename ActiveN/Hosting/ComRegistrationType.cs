namespace ActiveN.Hosting;

// this class is just a type holder with the required annotations
public class ComRegistrationType([DynamicallyAccessedMembers(
    DynamicallyAccessedMemberTypes.PublicMethods |
    DynamicallyAccessedMemberTypes.PublicProperties |
    DynamicallyAccessedMemberTypes.PublicConstructors)] Type type)
{
    [DynamicallyAccessedMembers(
        DynamicallyAccessedMemberTypes.PublicMethods |
        DynamicallyAccessedMemberTypes.PublicProperties |
        DynamicallyAccessedMemberTypes.PublicConstructors)]
    public Type Type { get; } = type;

    public override string ToString() => Type.ToString();
}
