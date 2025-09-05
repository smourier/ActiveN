namespace ActiveN.Utilities;

public class HRESULTValueEventArgs<T>(T value, bool isValueReadOnly = true, bool isCancellable = false) : ValueEventArgs<T>(value, isValueReadOnly, isCancellable)
{
    public virtual HRESULT Result { get; set; } = Constants.E_NOTIMPL;
}
