namespace ActiveN.Samples.WebView2;

public class WebView2Window(HWND parentHandle, WINDOW_STYLE style, RECT rect)
    : Window(title: nameof(WebView2Window), parentHandle: parentHandle, style: style, rect: rect)
{
    protected override bool OnPaint(HDC hdc, PAINTSTRUCT ps) { Paint(hdc, ClientRect); return true; }
    internal static void Paint(HDC hdc, RECT rc)
    {
        TracingUtilities.Trace($"Paint hdc: {hdc}");
        var text = $"Hello from ActiveN .NET {Environment.Version} ({DateTime.Now})";
        _ = Functions.TextOutW(hdc, (rc.left + rc.right) / 2, (rc.top + rc.bottom) / 2, PWSTR.From(text), text.Length);
    }
}
