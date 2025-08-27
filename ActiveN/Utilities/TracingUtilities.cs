namespace ActiveN.Utilities;

public static class TracingUtilities
{
    // write all traces on a single thread to avoid garbling the output
    private static int _threadId;
    private static TextWriter? _traceWriter;
    private static readonly SingleThreadTaskScheduler _writerScheduler = new SingleThreadTaskScheduler(thread =>
    {
        _threadId = thread.ManagedThreadId;
        thread.Name = "ActiveN Trace Writer";

        var process = Process.GetCurrentProcess();
        var dir = Path.Combine(Path.GetTempPath(), "ActiveN", process.ProcessName);
        if (!Directory.Exists(dir))
        {
            try
            {
                Directory.CreateDirectory(dir);
            }
            catch
            {
                // ignore
            }
        }
        Directory.SetCurrentDirectory(dir);
        _traceWriter = new StreamWriter(Path.Combine(dir, $"{Guid.NewGuid():N}.ActiveN.log")) { AutoFlush = true };
        _traceWriter?.WriteLine($"[{DateTime.Now:HH:mm:ss.fff}][{_threadId}]Log starting, process: {process.ProcessName} ({process.Id})");
        return true;
    });

    private static void TraceToWriter(string? text = null, string? methodName = null, string? filePath = null)
    {
        if (Environment.CurrentManagedThreadId == _threadId)
        {
            trace();
            return;
        }

        _ = Task.Factory.StartNew(trace, CancellationToken.None, TaskCreationOptions.None, _writerScheduler);

        void trace()
        {
            _traceWriter?.WriteLine($"[{DateTime.Now:HH:mm:ss.fff}][{Environment.CurrentManagedThreadId}]{(filePath != null ? Path.GetFileNameWithoutExtension(filePath) : null)}:: {methodName}: {text}");
        }
    }

    public static void FlushTextWriter()
    {
        // don't use scheduler thread for this
        _traceWriter?.Flush();
    }

    public static void Trace(string? text = null, [CallerMemberName] string? methodName = null, [CallerFilePath] string? filePath = null)
    {
        TraceToWriter(text, methodName, filePath);
        var name = filePath != null ? Path.GetFileNameWithoutExtension(filePath) : null;
        EventProvider.Default.WriteMessageEvent($"[{Environment.CurrentManagedThreadId}]{name}:: {methodName}: {text}");
    }
}
