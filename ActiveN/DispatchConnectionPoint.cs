// Copyright (c) Aelyo Softworks S.A.S.. All rights reserved.
// See LICENSE in the project root for license information.

namespace ActiveN;

[GeneratedComClass]
public partial class DispatchConnectionPoint(Guid sourceInterfaceId) : BaseConnectionPoint
{
    protected override IComObject GetFromPointer(nint ptr) => DirectN.Extensions.Com.ComObject.FromPointer<IDispatch>(ptr) ?? throw new InvalidOperationException();
    public override Guid InterfaceId { get; } = sourceInterfaceId;
    public override bool IsIDispatch => true;

    public virtual unsafe void InvokeMember(int dispId, params object?[]? parameters) => TracingUtilities.WrapErrors(() =>
    {
        TracingUtilities.Trace($"dispid {dispId} with {parameters?.Length ?? 0} parameters. Sinks: {Sinks.Count}");
        if (Sinks.Count > 0)
        {
            var count = parameters?.Length ?? 0;
            Variant[]? variants = null;
            var vars = new VARIANT[count];
            if (count > 0)
            {
                variants = new Variant[count];
                for (var i = 0; i < count; i++)
                {
                    variants[i] = new Variant(parameters![i]);
                    vars[i] = variants[i].Detached;
                }
            }

            try
            {
                fixed (VARIANT* pVars = vars)
                {
                    var dp = new DISPPARAMS
                    {
                        cArgs = (uint)count,
                        rgvarg = (nint)pVars,
                    };

                    foreach (var kv in Sinks)
                    {
                        var disp = kv.Value.As<IDispatch>();
                        disp?.Object.Invoke(dispId, Guid.Empty, 0, DISPATCH_FLAGS.DISPATCH_METHOD, dp, 0, 0, 0);
                    }
                }
            }
            finally
            {
                variants.Dispose();
            }
        }
        return Constants.S_OK;
    });
}

