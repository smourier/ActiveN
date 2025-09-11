using System.Collections.Concurrent;
using System.Drawing;
using System.Linq;

namespace ActiveN.PropertyGrid;

public class PropertiesForm : Form
{
    private static readonly ConcurrentDictionary<int, PropertiesForm> _instances = new();

    protected override bool ShowWithoutActivation => true;

    public PropertiesForm()
    {
        Text = "ActiveN Property Grid";
        StartPosition = FormStartPosition.Manual;
        Location = new Point(100, 100);
        Size = new Size(400, 600);
        Grid = new System.Windows.Forms.PropertyGrid
        {
            Dock = DockStyle.Fill
        };
        Controls.Add(Grid);
    }

    public System.Windows.Forms.PropertyGrid Grid { get; }

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

            if (!_instances.TryGetValue(Environment.CurrentManagedThreadId, out var form))
            {
                form = new PropertiesForm();
                Trace($"Already shown");
                _instances[Environment.CurrentManagedThreadId] = form;
            }

            var spmType = Type.GetTypeFromProgID("MTxSpm.SharedPropertyGroupManager");
            dynamic spm = Activator.CreateInstance(spmType!)!;
            var group = spm.CreatePropertyGroup("ActiveN", 0, 0, false);
            var property = group.CreateProperty("SelectedObject", false);
            var value = property.Value;

            Win32Window? parent = null;
            var args = ParseArguments(argument);
            Trace($"value: {value} argument: {argument} args: {string.Join(",", args.Select(kv => kv.Key + "=" + kv.Value))}");
            if (args.TryGetValue("parent", out var parentHwnd))
            {
                parent = new Win32Window((nint)ulong.Parse(parentHwnd));
            }

            var rect = new RECT { left = 0, top = 0, right = 200, bottom = 200 };
            if (parent != null && parent.Handle != 0)
            {
                GetClientRect(parent.Handle, out var rc);
                rect = rc;
            }
            else if (args.TryGetValue("rect", out var rects))
            {
                var rc = RECT.Parse(rects);
                if (rc != null)
                {
                    rect = rc.Value;
                }
            }

            form.Grid.SelectedObject = value;
            //form.Show(parent);
            Trace($"shown: 0x{parent?.Handle:X} rect: {rect}");
            if (parent != null)
            {
                SetParent(form.Handle, parent.Handle);
                //MoveWindow(form.Handle, rect.left, rect.top, rect.right - rect.left, rect.bottom - rect.top, true);
            }
            form.Show();
            //form.ShowDialog(parent);
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
    private static extern bool GetClientRect(nint hWndChild, out RECT hWndNewParent);

    [DllImport("user32")]
    private static extern bool MoveWindow(IntPtr hWnd, int X, int Y, int nWidth, int nHeight, bool bRepaint);

    private partial struct RECT
    {
        public int left;
        public int top;
        public int right;
        public int bottom;

        public static RECT? Parse(string s)
        {
            if (string.IsNullOrWhiteSpace(s))
                return null;

            var parts = s.Split([','], StringSplitOptions.RemoveEmptyEntries);
            if (parts.Length == 4
                && int.TryParse(parts[0], out var left)
                && int.TryParse(parts[1], out var top)
                && int.TryParse(parts[2], out var right)
                && int.TryParse(parts[3], out var bottom))
                return new RECT
                {
                    left = left,
                    top = top,
                    right = right,
                    bottom = bottom
                };
            return null;
        }
        public override readonly string ToString() => $"({left},{top})-({right},{bottom})";
    }

    private sealed class Win32Window(nint handle) : IWin32Window
    {
        public nint Handle { get; } = handle;
    }

    public static void Trace(string? text = null, [CallerMemberName] string? methodName = null, [CallerFilePath] string? filePath = null)
    {
        var name = filePath != null ? Path.GetFileNameWithoutExtension(filePath) : null;
        EventProvider.Default.WriteMessageEvent($"[{Environment.CurrentManagedThreadId}]{name}:: {methodName}: {text}");
    }
}
