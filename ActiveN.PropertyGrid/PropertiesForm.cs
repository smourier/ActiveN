namespace ActiveN.PropertyGrid;

public class PropertiesForm : Form
{
    public PropertiesForm()
    {
        FormBorderStyle = FormBorderStyle.None;
        Text = "ActiveN Property Grid";
        Grid = new System.Windows.Forms.PropertyGrid
        {
            Dock = DockStyle.Fill,
            HelpVisible = false,
            CommandsVisibleIfAvailable = false,
            ToolbarVisible = false
        };
        Controls.Add(Grid);
    }

    protected override bool ShowWithoutActivation => true;
    public nint ParentHandle { get; set; }
    public System.Windows.Forms.PropertyGrid Grid { get; }

    protected override void WndProc(ref Message m)
    {
        Trace("Msg: " + MessageDecoder.Decode(m));
        if (m.Msg == WM_PARENTNOTIFY)
        {
            var eventMsg = (int)(m.WParam.ToInt64() & 0xFFFF);
            if (eventMsg == WM_LBUTTONDOWN ||
                eventMsg == WM_MBUTTONDOWN ||
                eventMsg == WM_RBUTTONDOWN ||
                eventMsg == WM_XBUTTONDOWN ||
                eventMsg == WM_POINTERDOWN)
            {
                SetFocus(ParentHandle);
            }
        }
        base.WndProc(ref m);
    }

    private void SizeToParent()
    {
        GetClientRect(ParentHandle, out var rect);
        Location = new Point(rect.left, rect.top);
        Size = new Size(rect.right - rect.left, rect.bottom - rect.top);
    }

#pragma warning disable IDE0060 // Remove unused parameter
    public static int Show(string argument) // imposed by ClrRuntimeHost
#pragma warning restore IDE0060 // Remove unused parameter
    {
        try
        {
            AppDomain.CurrentDomain.FirstChanceException += (s, e) =>
            {
                Trace($"FirstChanceException: {e.Exception}");
            };

            AppDomain.CurrentDomain.UnhandledException += (s, e) =>
            {
                Trace($"UnhandledException: {e.ExceptionObject}");
            };

            var form = new PropertiesForm();

            var spmType = Type.GetTypeFromProgID("MTxSpm.SharedPropertyGroupManager");
            dynamic spm = Activator.CreateInstance(spmType!)!;
            var group = spm.CreatePropertyGroup("ActiveN", 0, 0, false);
            var property = group.CreateProperty("SelectedObject", false);
            var value = property.Value;

            var args = ParseArguments(argument);
            Trace($"value: {value} argument: {argument} parsed args: {string.Join(",", args.Select(kv => kv.Key + "=" + kv.Value))}");
            if (!args.TryGetValue("parent", out var parent) || !ulong.TryParse(parent, out var parentHandle))
                throw new InvalidOperationException("No valid parent HWND specified.");

            if (args.ContainsKey("toolbarvisible"))
            {
                form.Grid.ToolbarVisible = true;
            }

            var parentHwnd = (nint)parentHandle;
            form.ParentHandle = parentHwnd;
            form.Grid.SelectedObject = value;
            SetParent(form.Handle, parentHwnd);

            // parent's size should be fixed (OLE property page), no need to subclass/monitor size changes
            form.SizeToParent();
            form.Show();
            Trace($"return");
            return 0;
        }
        catch (Exception ex)
        {
            Trace($"Error showing property grid: {ex}");
            return ex.HResult;
        }
    }

    private static Dictionary<string, string> ParseArguments(string argument)
    {
        var result = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        var split = argument.Split(['|'], StringSplitOptions.RemoveEmptyEntries);
        foreach (var s in split)
        {
            var kv = s.Split([':'], 2);
            var key = kv[0].Trim();
            if (!string.IsNullOrWhiteSpace(key))
            {
                if (kv.Length == 1)
                {
                    result[key] = string.Empty;
                }
                else
                {
                    result[key] = kv[1].Trim();
                }
            }
        }
        return result;
    }

    [DllImport("user32")]
    private static extern nint SetParent(nint hWndChild, nint hWndNewParent);

    [DllImport("user32")]
    private static extern nint SetFocus(nint hWnd);

    [DllImport("user32")]
    private static extern bool GetClientRect(nint hWndChild, out RECT hWndNewParent);

#pragma warning disable IDE1006 // Naming Styles
    private const int WM_PARENTNOTIFY = 0x0210;
    private const int WM_LBUTTONDOWN = 0x0201;
    private const int WM_MBUTTONDOWN = 0x0207;
    private const int WM_RBUTTONDOWN = 0x0204;
    private const int WM_XBUTTONDOWN = 0x020B;
    private const int WM_POINTERDOWN = 0x0246;
#pragma warning restore IDE1006 // Naming Styles

    private partial struct RECT
    {
        public int left;
        public int top;
        public int right;
        public int bottom;
    }

    public static void Trace(string? text = null, [CallerMemberName] string? methodName = null, [CallerFilePath] string? filePath = null)
    {
        var name = filePath != null ? Path.GetFileNameWithoutExtension(filePath) : null;
        EventProvider.Default.WriteMessageEvent($"[{Environment.CurrentManagedThreadId}]{name}:: {methodName}: {text}");
    }
}
