namespace ActiveN;

[GeneratedComClass]
public partial class PropertyNotifySinkConnectionPoint : BaseConnectionPoint, IDisposable
{
    protected override IComObject GetFromPointer(nint ptr) => DirectN.Extensions.Com.ComObject.FromPointer<IPropertyNotifySink>(ptr) ?? throw new InvalidOperationException();
    public override Guid InterfaceId => typeof(IPropertyNotifySink).GUID;

    public virtual void OnChanged(int dispId)
    {
        foreach (var kv in Sinks)
        {
            var sink = kv.Value.As<IPropertyNotifySink>();
            sink?.Object.OnChanged(dispId);
        }
    }

    public virtual void OnRequestEdit(int dispId)
    {
        foreach (var kv in Sinks)
        {
            var sink = kv.Value.As<IPropertyNotifySink>();
            sink?.Object.OnRequestEdit(dispId);
        }
    }
}

