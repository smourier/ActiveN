namespace ActiveN.Utilities;

public static class TracingUtilities
{
    // write all traces on a single thread to avoid garbling the output
    private static int _threadId;
    private static TextWriter? _traceWriter;
    private static readonly SingleThreadTaskScheduler _writerScheduler = new(thread =>
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

    public static T? WrapErrors<T>(this Func<T> action, Action? actionOnError = null)
    {
        ArgumentNullException.ThrowIfNull(action);
        try
        {
            return action();
        }
        catch (SecurityException se)
        {
            // transform this one as a well-known access denied
            Trace($"Ex: {se}");
            if (actionOnError != null)
            {
                try
                {
                    actionOnError();
                }
                catch (Exception ex2)
                {
                    Trace($"Ex2: {ex2}");
                    // continue;
                }
            }
            if (typeof(T) == typeof(uint))
                return (T)(object)(uint)Constants.E_ACCESSDENIED;

            if (typeof(T) == typeof(int))
                return (T)(object)(int)Constants.E_ACCESSDENIED;

            if (typeof(T) == typeof(HRESULT))
                return (T)(object)Constants.E_ACCESSDENIED;

            return default;
        }
        catch (Exception ex)
        {
            Trace($"Ex: {ex}");
            if (actionOnError != null)
            {
                try
                {
                    actionOnError();
                }
                catch (Exception ex2)
                {
                    Trace($"Ex2: {ex2}");
                    // continue;
                }
            }
            if (typeof(T) == typeof(uint))
                return (T)(object)(uint)ex.HResult;

            if (typeof(T) == typeof(int))
                return (T)(object)ex.HResult;

            if (typeof(T) == typeof(HRESULT))
                return (T)(object)(HRESULT)ex.HResult;

            return default;
        }
    }

    public static uint WrapErrors(Action action, Action? actionOnError = null)
    {
        ArgumentNullException.ThrowIfNull(action);
        try
        {
            action();
            return 0;
        }
        catch (SecurityException se)
        {
            // transform this one as a well-known access denied
            Trace($"Ex: {se}");
            if (actionOnError != null)
            {
                try
                {
                    actionOnError();
                }
                catch (Exception ex2)
                {
                    Trace($"Ex2: {ex2}");
                    // continue;
                }
            }
            return (uint)Constants.E_ACCESSDENIED;
        }
        catch (Exception ex)
        {
            Trace($"Ex: {ex}");
            if (actionOnError != null)
            {
                try
                {
                    actionOnError();
                }
                catch (Exception ex2)
                {
                    Trace($"Ex2: {ex2}");
                    // continue;
                }
            }
            return (uint)ex.HResult;
        }
    }
}
