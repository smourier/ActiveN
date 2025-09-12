// Copyright (c) Aelyo Softworks S.A.S.. All rights reserved.
// See LICENSE in the project root for license information.

namespace ActiveN;

public class DispatchType([DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.PublicProperties | DynamicallyAccessedMemberTypes.PublicMethods)] Type type)
{
    private readonly Dictionary<string, DispatchMember> _membersByName = new(StringComparer.OrdinalIgnoreCase);
    private readonly Dictionary<int, DispatchMember> _memberByDispIds = [];
    private readonly HashSet<string> _restrictedNames = new(StringComparer.OrdinalIgnoreCase);
    private readonly Dictionary<PROPCAT, DispatchCategory> _categories = [];

    [DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.PublicProperties | DynamicallyAccessedMemberTypes.PublicMethods)]
    public Type Type { get; } = type ?? throw new ArgumentNullException(nameof(type));

    public override string ToString() => Type.FullName ?? Type.Name;

    protected virtual DispatchMember CreateMember(int dispId, DispatchCategory category, MemberInfo? info) => new(dispId, category, info);
    protected virtual DispatchCategory CreateCategory(PROPCAT category, string categoryName) => new(category, categoryName);

    protected virtual DispatchCategory GetCategory(MemberInfo? info)
    {
        var name = info?.GetCustomAttribute<CategoryAttribute>()?.Category.Nullify();
        if (name == null)
            return DispatchCategory.Misc;

        foreach (var field in typeof(PROPCAT).GetFields(BindingFlags.Public | BindingFlags.Static))
        {
            if (field.GetValue(null) is not PROPCAT value)
                continue;

            const string prefix = "PROPCAT_";
            if (field.Name.Length > prefix.Length && field.Name.StartsWith(prefix))
            {
                var propcatName = field.Name[prefix.Length..];
                if (propcatName.EqualsIgnoreCase(name))
                {
                    if (!_categories.TryGetValue(value, out var category))
                    {
                        category = CreateCategory(value, name);
                        _categories[value] = category;
                    }
                    return category;
                }
            }
        }

        // custom categories
        var cat = _categories.Values.FirstOrDefault(c => c.Name.EqualsIgnoreCase(name)); // we use values but it's ok since we don't expect many categories
        if (cat == null)
        {
            cat = CreateCategory((PROPCAT)_categories.Count + 1, name);
            _categories[cat.Category] = cat;
        }
        return cat;
    }

    public virtual unsafe void AddTypeInfoDispids(ITypeInfo typeInfo)
    {
        ArgumentNullException.ThrowIfNull(typeInfo);
        var typeAttr = TypeLib.GetAttributes(typeInfo);
        if (!typeAttr.HasValue)
            return;

        for (uint i = 0; i < typeAttr.Value.cImplTypes; i++)
        {
            if (typeInfo.GetRefTypeOfImplType(i, out var href).IsError)
                continue;

            if (typeInfo.GetRefTypeInfo(href, out var obj).IsError || obj == null)
                continue;

            var typeName = TypeLib.GetName(obj, -1);
            TracingUtilities.Trace($"type: '{typeName}'");

            using var refTypeInfo = new ComObject<ITypeInfo>(obj);
            var refTypeAttr = TypeLib.GetAttributes(refTypeInfo.Object);
            if (!refTypeAttr.HasValue)
                continue;

            for (uint f = 0; f < refTypeAttr.Value.cFuncs; f++)
            {
                var funcDesc = TypeLib.GetFuncDesc(refTypeInfo.Object, f);
                if (!funcDesc.HasValue)
                    continue;

                // skip restricted (QueryInterface, AddRef, Invoke, etc.)
                var name = TypeLib.GetName(refTypeInfo.Object, funcDesc.Value.memid);
                TracingUtilities.Trace($"funcDesc: id: {funcDesc.Value.memid} name:'{name}' kind: {funcDesc.Value.funckind} invkind: {funcDesc.Value.invkind} params: {funcDesc.Value.cParams} paramsOpt: {funcDesc.Value.cParamsOpt} flags: {funcDesc.Value.wFuncFlags}");
                if (name == null)
                    continue;

                if (funcDesc.Value.wFuncFlags.HasFlag(FUNCFLAGS.FUNCFLAG_FRESTRICTED))
                {
                    TracingUtilities.Trace($"restricted: {name}");
                    _restrictedNames.Add(name);
                    continue;
                }

                // if null, means the method/property is in the TLB but not in the actual type
                var memberInfo = (MemberInfo?)Type.GetProperty(name, BindingFlags.Public | BindingFlags.Instance)
                    ?? Type.GetMethod(name, BindingFlags.Public | BindingFlags.Instance);

                var category = GetCategory(memberInfo);
                var member = CreateMember(funcDesc.Value.memid, category, memberInfo) ?? throw new InvalidOperationException();

                var page = memberInfo?.GetCustomAttribute<PropertyPageAttribute>();
                if (page != null)
                {
                    member.PropertyPageId = page.Guid;
                    member.DefaultString = page.DefaultString;
                }

                _membersByName[memberInfo?.Name ?? name] = member;
                _memberByDispIds[member.DispId] = member;
            }
        }
    }

    public virtual void AddReflectionDispids(int autoDispidsBase)
    {
        // note we don't support overloaded methods & properties
        // add only members not already added by type info
        var methods = Type.GetMethods(BindingFlags.Public | BindingFlags.Instance);
        for (var i = 0; i < methods.Length; i++)
        {
            var method = methods[i];
            if (method.Attributes.HasFlag(MethodAttributes.SpecialName)) // like get_ / set_ / add_ / remove_ / etc.
                continue;

            if (_restrictedNames.Contains(method.Name))
                continue;

            if (_membersByName.ContainsKey(method.Name))
                continue;

            var browsable = method.GetCustomAttribute<BrowsableAttribute>()?.Browsable;
            if (browsable.HasValue && !browsable.Value)
                continue;

            // allow developer to customize name & dispid using attributes
            var name = method.GetCustomAttribute<ComAliasNameAttribute>()?.Value ?? method.Name;
            var dispid = method.GetCustomAttribute<DispIdAttribute>()?.Value ?? autoDispidsBase + _membersByName.Count;

            var category = GetCategory(method);
            var member = CreateMember(dispid, category, method) ?? throw new InvalidOperationException();

            var page = method?.GetCustomAttribute<PropertyPageAttribute>();
            if (page != null)
            {
                member.PropertyPageId = page.Guid;
                member.DefaultString = page.DefaultString;
            }

            member.PropertyPageId = method?.GetCustomAttribute<PropertyPageAttribute>()?.Guid;
            _membersByName[name] = member;
            _memberByDispIds[member.DispId] = member;
        }

        var properties = Type.GetProperties(BindingFlags.Public | BindingFlags.Instance);
        for (var i = 0; i < properties.Length; i++)
        {
            var property = properties[i];
            if (_restrictedNames.Contains(property.Name))
                continue;

            if (_membersByName.ContainsKey(property.Name))
                continue;

            var browsable = property.GetCustomAttribute<BrowsableAttribute>()?.Browsable;
            if (browsable.HasValue && !browsable.Value)
                continue;

            // allow developer to customize name & dispid using attributes
            var name = property.GetCustomAttribute<ComAliasNameAttribute>()?.Value ?? property.Name;
            var dispid = property.GetCustomAttribute<DispIdAttribute>()?.Value ?? autoDispidsBase + _membersByName.Count;

            var category = GetCategory(property);
            var member = CreateMember(dispid, category, property) ?? throw new InvalidOperationException();
            member.IsReadOnly = !property.CanWrite || property.GetCustomAttribute<ReadOnlyAttribute>()?.IsReadOnly == true;

            var page = property?.GetCustomAttribute<PropertyPageAttribute>();
            if (page != null)
            {
                member.PropertyPageId = page.Guid;
                member.DefaultString = page.DefaultString;
            }

            member.PropertyPageId = property?.GetCustomAttribute<PropertyPageAttribute>()?.Guid;
            _membersByName[name] = member;
            _memberByDispIds[member.DispId] = member;
        }

#if DEBUG
        foreach (var name in _restrictedNames)
        {
            TracingUtilities.Trace($"type: {Type.Name} restricted: {name}");
        }

        foreach (var kv in _memberByDispIds)
        {
            TracingUtilities.Trace($"type: {Type.Name} dispid: {kv.Key} (0x{kv.Key:X}) => {kv.Value}");
        }

        foreach (var kv in _membersByName)
        {
            TracingUtilities.Trace($"type: {Type.Name} name: '{kv.Key}' => {kv.Value}");
        }
#endif
    }

    public virtual DispatchCategory GetCategory(PROPCAT category)
    {
        if (_categories.TryGetValue(category, out var cat))
            return cat;

        return DispatchCategory.Misc;
    }

    public virtual DispatchMember? GetMember(int dispId)
    {
        if (!_memberByDispIds.TryGetValue(dispId, out var member))
            return null;

        return member;
    }

    public MemberInfo? GetMemberInfo(int dispId) => GetMember(dispId)?.Info;
    public virtual bool TryGetDispId(string name, out int dispId)
    {
        dispId = 0;
        if (!_membersByName.TryGetValue(name, out var member))
            return false;

        dispId = member.DispId;
        return true;
    }
}
