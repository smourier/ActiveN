namespace ActiveN;

[GeneratedComClass]
public partial class DispatchConnectionPoint(Guid sourceInterfaceId) : BaseConnectionPoint, IDisposable
{
    protected override IComObject GetFromPointer(nint ptr) => DirectN.Extensions.Com.ComObject.FromPointer<IDispatch>(ptr) ?? throw new InvalidOperationException();
    public override Guid InterfaceId { get; } = sourceInterfaceId;

    public virtual unsafe void InvokeMember(int dispId, params object?[]? parameters) => TracingUtilities.WrapErrors(() =>
    {
        TracingUtilities.Trace($"dispid {dispId} with {parameters?.Length ?? 0} parameters. Sinks: {Sinks.Count}");
        if (Sinks.Count == 0)
            return;

        Variant[]? variants = null;
        VARIANT[]? vars;
        var dp = new DISPPARAMS();
        if (parameters?.Length > 0)
        {
            variants = new Variant[parameters.Length];
            vars = new VARIANT[parameters.Length];
            for (var i = 0; i < parameters.Length; i++)
            {
                variants[i] = new Variant(parameters[i]);
                vars[i] = variants[i].Detached;
            }

            dp.cArgs = (uint)parameters.Length;
            dp.rgvarg = vars.AsPointer();
        }

        try
        {
            foreach (var kv in Sinks)
            {
                var disp = kv.Value.As<IDispatch>();
                disp?.Object.Invoke(dispId, Guid.Empty, 0, DISPATCH_FLAGS.DISPATCH_METHOD, dp, 0, 0, 0);
            }
        }
        finally
        {
            variants.Dispose();
        }
    });
}

