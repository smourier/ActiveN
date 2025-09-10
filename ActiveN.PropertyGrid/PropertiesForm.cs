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
            var property = group.CreateProperty(typeof(IPropertyGridObject).GUID.ToString(), false);
            var value = property.Value;
            Trace($"value: {value}");
            var obj = value as IPropertyGridObject;
            Trace($"obj: {obj}");

            var desc = PropertyGridTypeDescriptor.BuildCustomTypeDescriptor(obj);
            form.Grid.SelectedObject = obj;

            form.ShowDialog();

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

    private sealed class PropertyGridTypeDescriptor(PropertyDescriptor[] descriptors) : CustomTypeDescriptor
    {
        public override PropertyDescriptorCollection GetProperties() => new(descriptors);

        public static ICustomTypeDescriptor BuildCustomTypeDescriptor(IPropertyGridObject? obj)
        {
            var props = new List<PropertyDescriptor>();
            if (obj != null)
            {
                obj.GetProperties(out var variant);
                Trace($"variant: {variant}");
            }
            return new PropertyGridTypeDescriptor([.. props]);
        }
    }
}
