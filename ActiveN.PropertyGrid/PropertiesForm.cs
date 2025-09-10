namespace ActiveN.PropertyGrid;

public class PropertiesForm : Form
{
    public PropertiesForm()
    {
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
            var form = new PropertiesForm();
            var spmType = Type.GetTypeFromProgID("MTxSpm.SharedPropertyGroupManager");
            dynamic spm = Activator.CreateInstance(spmType!)!;
            var group = spm.CreatePropertyGroup("ActiveN", 0, 0, false);
            var property = group.CreateProperty("SelectedObject", false);
            var value = property.Value;
            Trace($"value: {value}");

            form.Grid.SelectedObject = value;

            var ret = form.ShowDialog();

            Trace($"ret: {ret}");
            return 0;
        }
        catch (Exception ex)
        {
            Trace($"Error showing property grid: {ex}");
            return ex.HResult;
        }
    }

    public static void Trace(string? text = null, [CallerMemberName] string? methodName = null, [CallerFilePath] string? filePath = null)
    {
        var name = filePath != null ? Path.GetFileNameWithoutExtension(filePath) : null;
        EventProvider.Default.WriteMessageEvent($"[{Environment.CurrentManagedThreadId}]{name}:: {methodName}: {text}");
    }
}
