// Copyright (c) Aelyo Softworks S.A.S.. All rights reserved.
// See LICENSE in the project root for license information.

namespace ActiveN;

[AttributeUsage(AttributeTargets.Method | AttributeTargets.Property, Inherited = true)]
public sealed class PropertyPageAttribute(string clsid) : Attribute
{
    public string Clsid { get; } = clsid;
    public Guid? Guid
    {
        get
        {
            if (System.Guid.TryParse(Clsid, out var guid))
                return guid;

            return null;
        }
    }

    public string? DefaultString { get; set; }
}
