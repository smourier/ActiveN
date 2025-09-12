// Copyright (c) Aelyo Softworks S.A.S.. All rights reserved.
// See LICENSE in the project root for license information.

namespace ActiveN;

public class DispatchCategory(PROPCAT category, string categoryName)
{
    public static DispatchCategory Misc { get; } = new DispatchCategory(PROPCAT.PROPCAT_Misc, "Misc");

    public PROPCAT Category { get; } = category;
    public string Name { get; } = categoryName;

    public virtual string GetLocalizedName(uint lcid) => Name;

    public override string ToString() => $"{Category} ({Name})";
}
