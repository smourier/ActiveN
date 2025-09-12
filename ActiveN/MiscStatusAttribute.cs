// Copyright (c) Aelyo Softworks S.A.S.. All rights reserved.
// See LICENSE in the project root for license information.

namespace ActiveN;

[AttributeUsage(AttributeTargets.Class, Inherited = true, AllowMultiple = false)]
public sealed class MiscStatusAttribute(OLEMISC value) : Attribute
{
    public OLEMISC Value { get; } = value;
}
