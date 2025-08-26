namespace ActiveN.Samples.HelloWorld;

[Guid("00000002-2126-40c8-a2f3-8fa83f8cc1f6")]
[GeneratedComInterface]
public partial interface IHelloWorldControl : IDispatch
{
    HRESULT ComputePi(out double ret);
}
