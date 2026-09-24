// Copyright (c) Aelyo Softworks S.A.S.. All rights reserved.
// See LICENSE in the project root for license information.

namespace ActiveN.Samples.WebView2;

public class WebView2Window : Window
{
    private IComObject<ICoreWebView2Controller>? _controller;
    private IComObject<ICoreWebView2>? _webView2;
    private CoreWebView2Events? _webView2Events;

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

        _ = InitializeAsync(path, source);
    }

    private async Task InitializeAsync(string userDataFolder, string? source)
    {
        try
        {
            using var environment = await global::WebView2.Functions.CreateCoreWebView2EnvironmentWithOptionsAsync(null, userDataFolder, null) ?? throw new InvalidOperationException("The WebView2 environment could not be created.");
            _controller = await environment.CreateCoreWebView2ControllerAsync(Handle) ?? throw new InvalidOperationException("The WebView2 controller could not be created.");
            _webView2 = _controller.CoreWebView2 ?? throw new InvalidOperationException("The WebView2 controller has no WebView.");

            _webView2Events = new CoreWebView2Events(_webView2);
            _webView2Events.FrameNavigationCompleted += (sender, args) => OnFrameNavigationCompleted(this, args);
            _webView2Events.NavigationCompleted += (sender, args) => OnNavigationCompleted(this, args);
            _webView2Events.DocumentTitleChanged += (sender, args) => OnDocumentTitleChanged(this, DocumentTitle);
            _webView2Events.NewWindowRequested += (sender, args) => OnNewWindowRequested(this, args);

            _controller.Bounds = ClientRect;

            if (string.IsNullOrWhiteSpace(source))
            {
                var text = $"WebView2 V{BrowserVersion} - {RuntimeInformation.ProcessArchitecture} - .NET V{Environment.Version}";
                var html = $"<body style='margin:0;padding:0'><p style='height:100vh;background-image:linear-gradient(90deg,#e3ffe7 0%,#d9e7ff 100%);font-family:consolas;display:flex;justify-content: center;align-items:center'>{text}</p>";
                _webView2.NavigateToString(html);
            }
            else
            {
                _webView2.Navigate(source);
            }
        }
        catch (Exception ex)
        {
            TracingUtilities.Trace($"WebView2 cannot be initialized: {ex.Message}");
        }
    }

    public string Source
    {
        get
        {
            if (_webView2 == null)
                return string.Empty;

            return _webView2.Source ?? string.Empty;
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

            return _webView2.DocumentTitle ?? string.Empty;
        }
    }

    public string StatusBarText
    {
        get
        {
            if (_webView2?.Object is not ICoreWebView2_12)
                return string.Empty;

            return _webView2.StatusBarText ?? string.Empty;
        }
    }

    public void Reload() => _webView2?.Reload();
    public void GoBack() => _webView2?.GoBack();
    public void GoForward() => _webView2?.GoForward();

    public void Navigate(string uri)
    {
        ArgumentNullException.ThrowIfNull(uri);
        TracingUtilities.Trace($"Navigating to {uri}...");
        _webView2?.Navigate(uri);
    }

    public void NavigateToString(string htmlContent)
    {
        ArgumentNullException.ThrowIfNull(htmlContent);
        TracingUtilities.Trace($"Navigating to HTML content: `{htmlContent}`");
        _webView2?.NavigateToString(htmlContent);
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
        _controller?.Bounds = ClientRect;
        return base.OnResized(type, size);
    }

    protected override void Dispose(bool disposing)
    {
        if (disposing)
        {
            _webView2Events?.Dispose();
            _webView2Events = null;
            _webView2?.Dispose();
            _webView2 = null;
            _controller?.Object.Close();
            _controller?.Dispose();
            _controller = null;
        }
        base.Dispose(disposing);
    }

    public static string BrowserVersion { get; private set; } = "<not initialized>";
}
