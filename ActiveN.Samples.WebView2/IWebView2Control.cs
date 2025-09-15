// Copyright (c) Aelyo Softworks S.A.S.. All rights reserved.
// See LICENSE in the project root for license information.

namespace ActiveN.Samples.WebView2;

[Guid("9d51ea97-30dd-4e70-99ae-a5085fbf8095")]
[GeneratedComInterface]
public partial interface IWebView2Control : IDispatch
{
#pragma warning disable IDE1006 // Naming Styles
    [PreserveSig]
    [return: MarshalAs(UnmanagedType.Error)]
    HRESULT set_Source(BSTR value);

    [PreserveSig]
    [return: MarshalAs(UnmanagedType.Error)]
    HRESULT get_Source(out BSTR value);

    [PreserveSig]
    [return: MarshalAs(UnmanagedType.Error)]
    HRESULT get_StatusBarText(out BSTR value);

    [PreserveSig]
    [return: MarshalAs(UnmanagedType.Error)]
    HRESULT get_DocumentTitle(out BSTR value);

    [PreserveSig]
    [return: MarshalAs(UnmanagedType.Error)]
    HRESULT get_BrowserVersion(out BSTR value);
#pragma warning restore IDE1006 // Naming Styles

    [PreserveSig]
    [return: MarshalAs(UnmanagedType.Error)]
    HRESULT Navigate(BSTR uri);

    [PreserveSig]
    [return: MarshalAs(UnmanagedType.Error)]
    HRESULT NavigateToString(BSTR htmlContent);

    [PreserveSig]
    [return: MarshalAs(UnmanagedType.Error)]
    HRESULT GoBack();

    [PreserveSig]
    [return: MarshalAs(UnmanagedType.Error)]
    HRESULT GoForward();

    [PreserveSig]
    [return: MarshalAs(UnmanagedType.Error)]
    HRESULT Reload();
}
