namespace ActiveN.Samples.SimpleComObject;

[Guid("b87ad785-2cd4-4e7c-a5c0-82e11e54639c")]
[GeneratedComInterface]
public partial interface ISimple
{
    [PreserveSig]
    [return: MarshalAs(UnmanagedType.Error)]
    HRESULT Add(double left, double right, out double ret);
}
