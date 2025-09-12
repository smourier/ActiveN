// Copyright (c) Aelyo Softworks S.A.S.. All rights reserved.
// See LICENSE in the project root for license information.

namespace ActiveN.Samples.PdfView;

[DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.PublicProperties | DynamicallyAccessedMemberTypes.PublicMethods)]
[GeneratedComClass]
public partial class PdfViewPage : BaseDispatch, IPdfViewPage
{
    private readonly PdfViewControl _control;

    public PdfViewPage(PdfViewControl control, PdfPage page)
    {
        ArgumentNullException.ThrowIfNull(control);
        ArgumentNullException.ThrowIfNull(page);
        _control = control;
        Page = page;
    }

    public PdfPage Page { get; }
    public int Index => (int)Page.Index;
    public double Width => Page.Size.Width;
    public double Height => Page.Size.Height;
    public float PreferredZoom => Page.PreferredZoom;
    public PdfPageRotation Rotation => Page.Rotation;

    public void ExtractTo(VARIANT output)
    {
        var window = _control.Window ?? throw new Exception("No file was opened.");

        using var pv = Variant.Attach(ref output, false);
        if (pv.Value is string path)
        {
            window.ExtractPage(Page, path).Wait();
            return;
        }

        if (pv.Value is IStream istream)
        {
            window.ExtractPage(Page, new StreamOnIStream(istream)).Wait();
            return;
        }

        throw new NotSupportedException($"{nameof(output)} must be a file path or an stream.");
    }

    HRESULT IPdfViewPage.get_Index(out int value) { value = Index; return Constants.S_OK; }
    HRESULT IPdfViewPage.get_Width(out double value) { value = Width; return Constants.S_OK; }
    HRESULT IPdfViewPage.get_Height(out double value) { value = Height; return Constants.S_OK; }
    HRESULT IPdfViewPage.get_PreferredZoom(out float value) { value = PreferredZoom; return Constants.S_OK; }
    HRESULT IPdfViewPage.get_Rotation(out PdfPageRotation value) { value = Rotation; return Constants.S_OK; }
    HRESULT IPdfViewPage.ExtractTo(VARIANT output) { ExtractTo(output); return Constants.S_OK; }

    protected override void Dispose(bool disposing)
    {
        base.Dispose(disposing);
        Page.Dispose();
    }

    #region Mandatory overrides
    protected override ComRegistration ComRegistration => ComHosting.Instance;
    protected override Guid DispatchInterfaceId => typeof(IPdfViewPage).GUID;
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
