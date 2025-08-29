namespace ActiveN.Samples.HelloWorld;

public class HelloWorldWindow(HWND parentHandle, WINDOW_STYLE style, RECT rect)
    : Window(parentHandle: parentHandle, style: style, rect: rect)
{
    protected override bool OnPaint(HDC hdc, PAINTSTRUCT ps)
    {
        TracingUtilities.Trace($"hdc: {hdc} erase: {ps.fErase} restore: {ps.fRestore} incUpdate: {ps.fIncUpdate} rcPaint: {ps.rcPaint}");
        var text = "Hello from ActiveN HelloWorldWindow!";
        _ = Functions.TextOutW(hdc, 10, 10, PWSTR.From(text), text.Length);
        return true;
    }
}
