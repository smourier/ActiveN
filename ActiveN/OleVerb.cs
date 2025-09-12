// Copyright (c) Aelyo Softworks S.A.S.. All rights reserved.
// See LICENSE in the project root for license information.

namespace ActiveN;

public class OleVerb(OLEVERB verb)
{
    public event EventHandler<HRESULTValueEventArgs<(MSG? msg, IOleClientSite activeSite, HWND hwndParent)>>? Invoking;

    public OLEVERB Verb => verb;

    public override string ToString() => $"{Verb.lVerb}: {Verb.lpszVerbName}";

    protected virtual void OnInvoking(object sender, HRESULTValueEventArgs<(MSG? msg, IOleClientSite activeSite, HWND hwndParent)> e) => Invoking?.Invoke(this, e);
    protected virtual internal HRESULT Invoke(MSG? msg, IOleClientSite activeSite, HWND hwndParent)
    {
        var e = new HRESULTValueEventArgs<(MSG? msg, IOleClientSite activeSite, HWND hwndParent)>((msg, activeSite, hwndParent));
        OnInvoking(this, e);
        return e.Result;
    }
}
