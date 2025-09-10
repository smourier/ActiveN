namespace ActiveN.PropertyGrid;

public sealed partial class EventProvider : IDisposable
{
    public static Guid DefaultGuid { get; set; } = new("964d4572-adb9-4f3a-8170-fcbecec27467");
    public static readonly EventProvider Default = new(DefaultGuid);

    private long _handle;
    public Guid Id { get; }

    public EventProvider(Guid id)
    {
        Id = id;
        var hr = EventRegister(id, 0, 0, out _handle);
        if (hr != 0)
            throw new Win32Exception(hr);
    }

    public bool WriteMessageEvent(string text, byte level = 0, long keywords = 0) => EventWriteString(_handle, level, keywords, text) == 0;

    public void Dispose()
    {
        var handle = Interlocked.Exchange(ref _handle, 0);
        if (handle != 0)
        {
            _ = EventUnregister(handle);
        }
    }

    [DllImport("advapi32")]
    private static extern int EventRegister(in Guid ProviderId, nint EnableCallback, nint CallbackContext, out long RegHandle);

    [DllImport("advapi32")]
    private static extern int EventUnregister(long RegHandle);

    [DllImport("advapi32", CharSet = CharSet.Unicode)]
    private static extern int EventWriteString(long RegHandle, byte Level, long Keyword, string String);
}
