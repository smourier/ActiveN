namespace ActiveN.Hosting;

// this class is just a type holder with the required annotations
public class ComRegistrationType([DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.PublicMethods | DynamicallyAccessedMemberTypes.PublicProperties)] Type type)
{
    [DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.PublicMethods | DynamicallyAccessedMemberTypes.PublicProperties)]
    public Type Type { get; } = type;

    public override string ToString() => Type.ToString();
}
