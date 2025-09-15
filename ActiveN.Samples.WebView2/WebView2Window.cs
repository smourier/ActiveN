// Copyright (c) Aelyo Softworks S.A.S.. All rights reserved.
// See LICENSE in the project root for license information.

namespace ActiveN.Samples.WebView2;

public class WebView2Window : Window
{
    private ComObject<ICoreWebView2Controller>? _controller;
    private ComObject<ICoreWebView2>? _webView2;
    private EventRegistrationToken _navigationCompleted;
    private EventRegistrationToken _documentTitleChanged;
    private EventRegistrationToken _newWindowRequested;
    private EventRegistrationToken _frameNavigationCompleted;

    public event EventHandler<ValueEventArgs<ICoreWebView2NavigationCompletedEventArgs>>? NavigationCompleted;
    public event EventHandler<ValueEventArgs<ICoreWebView2NavigationCompletedEventArgs>>? FrameNavigationCompleted;
    public event EventHandler<ValueEventArgs<ICoreWebView2NewWindowRequestedEventArgs>>? NewWindowRequested;
    public event EventHandler<ValueEventArgs<string?>>? DocumentTitleChanged;

    public WebView2Window(HWND parentHandle, WINDOW_STYLE style, RECT rect, string? source)
        : base(title: nameof(WebView2Window), parentHandle: parentHandle, style: style, rect: rect)
    {
        // this checks WebView2Loader.dll is present somewhere (file path or embedded resource)
        WebView2Utilities.Initialize(Assembly.GetExecutingAssembly());

        // this checks WebView2 itself is installed
        BrowserVersion = WebView2Utilities.GetAvailableCoreWebView2BrowserVersionString() ?? "<not installed>";

        // use webview2 user data folder under local app data to ensure it'll work wherever the app is installed
        var path = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), GetType().Namespace!);

        var hr = global::WebView2.Functions.CreateCoreWebView2EnvironmentWithOptions(PWSTR.Null, PWSTR.From(path), null!,
            new CoreWebView2CreateCoreWebView2EnvironmentCompletedHandler((result, env) =>
            {
                if (result.IsError)
                {
                    TracingUtilities.Trace($"WebView controller cannot be initialized: {result}.");
                    return;
                }

                env.CreateCoreWebView2Controller(Handle, new CoreWebView2CreateCoreWebView2ControllerCompletedHandler((result, controller) =>
                {
                    _controller = new ComObject<ICoreWebView2Controller>(controller);
                    _controller.Object.get_CoreWebView2(out var webView2).ThrowOnError();

                    _webView2 = new ComObject<ICoreWebView2>(webView2);
                    _webView2.Object.add_FrameNavigationCompleted(new CoreWebView2NavigationCompletedEventHandler((sender, args) =>
                    {
                        OnFrameNavigationCompleted(this, args);
                    }), ref _frameNavigationCompleted);

                    _webView2.Object.add_NavigationCompleted(new CoreWebView2NavigationCompletedEventHandler((sender, args) =>
                    {
                        OnNavigationCompleted(this, args);
                    }), ref _navigationCompleted);

                    _webView2.Object.add_DocumentTitleChanged(new CoreWebView2DocumentTitleChangedEventHandler((sender, args) =>
                    {
                        sender.get_DocumentTitle(out var p);
                        using var pwstr = new Pwstr(p.Value);
                        OnDocumentTitleChanged(this, pwstr.ToString());
                    }), ref _documentTitleChanged);

                    _webView2.Object.add_NewWindowRequested(new CoreWebView2NewWindowRequestedEventHandler((sender, args) =>
                    {
                        OnNewWindowRequested(this, args);
                    }), ref _newWindowRequested);

                    controller.put_Bounds(ClientRect).ThrowOnError();

                    if (string.IsNullOrWhiteSpace(source))
                    {
                        var text = $"WebView2 V{BrowserVersion} - {RuntimeInformation.ProcessArchitecture} - .NET V{Environment.Version}";
                        var html = $"<body style='margin:0;padding:0'><p style='height:100vh;background-image:linear-gradient(90deg,#e3ffe7 0%,#d9e7ff 100%);font-family:consolas;display:flex;justify-content: center;align-items:center'>{text}</p>";
                        webView2.NavigateToString(PWSTR.From(html));
                    }
                    else
                    {
                        webView2.Navigate(PWSTR.From(source));
                    }
                }));
            }));
        if (hr.IsError)
        {
            TracingUtilities.Trace($"CreateCoreWebView2EnvironmentWithOptions failed: {hr}");
            _controller = null;
        }
    }

    public string Source
    {
        get
        {
            if (_webView2 == null)
                return string.Empty;

            _webView2.Object.get_Source(out var p).ThrowOnError();
            using var pwstr = new Pwstr(p.Value);
            return pwstr.ToString() ?? string.Empty;
        }
        set
        {
            ArgumentNullException.ThrowIfNull(value);
            Navigate(value);
        }
    }

    public string DocumentTitle
    {
        get
        {
            if (_webView2 == null)
                return string.Empty;

            _webView2.Object.get_DocumentTitle(out var p).ThrowOnError();
            using var pwstr = new Pwstr(p.Value);
            return pwstr.ToString() ?? string.Empty;
        }
    }

    public string StatusBarText
    {
        get
        {
            if (_webView2 == null)
                return string.Empty;

            var webView2_12 = _webView2.As<ICoreWebView2_12>();
            if (webView2_12 == null)
                return string.Empty;

            webView2_12.Object.get_StatusBarText(out var p).ThrowOnError();
            using var pwstr = new Pwstr(p.Value);
            return pwstr.ToString() ?? string.Empty;
        }
    }

    public void Reload() => _webView2?.Object.Reload().ThrowOnError();
    public void GoBack() => _webView2?.Object.GoBack().ThrowOnError();
    public void GoForward() => _webView2?.Object.GoForward().ThrowOnError();

    public void Navigate(string uri)
    {
        ArgumentNullException.ThrowIfNull(uri);
        TracingUtilities.Trace($"Navigating to {uri}...");
        _webView2?.Object.Navigate(PWSTR.From(uri)).ThrowOnError();
    }

    public void NavigateToString(string htmlContent)
    {
        ArgumentNullException.ThrowIfNull(htmlContent);
        TracingUtilities.Trace($"Navigating to HTML content: `{htmlContent}`");
        _webView2?.Object.NavigateToString(PWSTR.From(htmlContent)).ThrowOnError();
    }

    protected virtual void OnFrameNavigationCompleted(object? sender, ICoreWebView2NavigationCompletedEventArgs args)
        => FrameNavigationCompleted?.Invoke(this, new ValueEventArgs<ICoreWebView2NavigationCompletedEventArgs>(args));

    protected virtual void OnNavigationCompleted(object? sender, ICoreWebView2NavigationCompletedEventArgs args)
        => NavigationCompleted?.Invoke(this, new ValueEventArgs<ICoreWebView2NavigationCompletedEventArgs>(args));

    protected virtual void OnDocumentTitleChanged(object? sender, string? title)
        => DocumentTitleChanged?.Invoke(this, new ValueEventArgs<string?>(title));

    protected virtual void OnNewWindowRequested(object? sender, ICoreWebView2NewWindowRequestedEventArgs args)
        => NewWindowRequested?.Invoke(this, new ValueEventArgs<ICoreWebView2NewWindowRequestedEventArgs>(args));

    protected override bool OnResized(WindowResizedType type, SIZE size)
    {
        _controller?.Object.put_Bounds(ClientRect).ThrowOnError();
        return base.OnResized(type, size);
    }

    protected override void Dispose(bool disposing)
    {
        if (disposing)
        {
            if (_navigationCompleted.value != 0)
            {
                _webView2?.Object.remove_FrameNavigationCompleted(_navigationCompleted);
                _navigationCompleted.value = 0;
            }

            _webView2?.Dispose();
            _webView2 = null;
            _controller?.Dispose();
            _controller = null;
        }
        base.Dispose(disposing);
    }

    public static string BrowserVersion { get; private set; } = "<not initialized>";
}
