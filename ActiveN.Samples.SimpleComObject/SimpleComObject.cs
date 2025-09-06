namespace ActiveN.Samples.SimpleComObject;

[Guid("a2a2d7bf-773d-49ca-bf99-c3d17ab4225b")]
[ProgId("ActiveN.Samples.SimpleComObject.SimpleComObject")]
[DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.PublicProperties | DynamicallyAccessedMemberTypes.PublicMethods)]
[GeneratedComClass]
public partial class SimpleComObject : BaseDispatch, ISimple, ISimpleDual
{
#pragma warning disable CA1822 // Mark members as static; no since we're dealing with COM instance methods & properties

    public string Name { get; set; } = "SimpleComObject";
    public double Add(double left, double right) => left + right;
    public double Multiply(double left, double right) => left * right;

#pragma warning restore CA1822 // Mark members as static

    #region interfaces implementation

    // note we implement both ISimple and ISimpleDual here explicitly but this is not mandatory
    // I prefer to do this so we expose public methods and properties directly on the class with a nicer .NET style
    HRESULT ISimpleDual.get_Name(out BSTR value) { value = new BSTR(Marshal.StringToBSTR(Name)); return Constants.S_OK; }
    HRESULT ISimpleDual.set_Name(BSTR value) { Name = value.ToString() ?? string.Empty; return Constants.S_OK; }
    HRESULT ISimpleDual.Multiply(double left, double right, out double ret) { ret = Multiply(left, right); return Constants.S_OK; }
    HRESULT ISimple.Add(double left, double right, out double ret) { ret = Add(left, right); return Constants.S_OK; }

    #endregion

    #region Mandatory overrides
    protected override ComRegistration ComRegistration => ComHosting.Instance;
    #endregion

    #region IDispatch support

    HRESULT IDispatch.GetIDsOfNames(in Guid riid, PWSTR[] rgszNames, uint cNames, uint lcid, int[] rgDispId) =>
        GetIDsOfNames(in riid, rgszNames, cNames, lcid, rgDispId);

    HRESULT IDispatch.Invoke(int dispIdMember, in Guid riid, uint lcid, DISPATCH_FLAGS wFlags, in DISPPARAMS pDispParams, nint pVarResult, nint pExcepInfo, nint puArgErr) =>
        Invoke(dispIdMember, riid, lcid, wFlags, pDispParams, pVarResult, pExcepInfo, puArgErr);

    HRESULT IDispatch.GetTypeInfo(uint iTInfo, uint lcid, out ITypeInfo ppTInfo) =>
        GetTypeInfo(iTInfo, lcid, out ppTInfo);

    HRESULT IDispatch.GetTypeInfoCount(out uint pctinfo) =>
        GetTypeInfoCount(out pctinfo);

    #endregion
}
