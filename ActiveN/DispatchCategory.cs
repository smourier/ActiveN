namespace ActiveN;

public class DispatchCategory(PROPCAT category, string categoryName)
{
    public static DispatchCategory Misc { get; } = new DispatchCategory(PROPCAT.PROPCAT_Misc, "Misc");

    public PROPCAT Category { get; } = category;
    public string Name { get; } = categoryName;

    public virtual string GetLocalizedName(uint lcid) => Name;

    public override string ToString() => $"{Category} ({Name})";
}
