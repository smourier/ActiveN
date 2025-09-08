﻿namespace ActiveN.Utilities;

public static class TracingUtilities
{
    // for tracing only
    [Guid("9a604ee7-e630-3ded-9444-baae247075ab")]
    private partial interface IFormattable { }

    // office's VBA asks for this
    [Guid("049948d1-5686-11cf-8e0d-00aa00a74c5c")]
    private partial interface IPersistHTML { }

    [Guid("e19c7100-9709-4db7-9373-e7b518b47086")]
    private partial interface IOfficeSafetyRelated { }

    // office's VBA asks for this
    [Guid("C62A69F0-16DC-11CE-9E98-00AA00574A4F")]
#pragma warning disable IDE1006 // Naming Styles
    private partial interface IForms20Form { }
#pragma warning restore IDE1006 // Naming Styles

    // write all traces on a single thread to avoid garbling the output
    private static int _threadId;
    private static TextWriter? _traceWriter;
    private static SingleThreadTaskScheduler? _writerScheduler;

    public static bool TraceToFile
    {
        get => _traceWriter != null;
        set
        {
            if (value && _traceWriter == null)
            {
                _writerScheduler = new(thread =>
                {
                    _threadId = thread.ManagedThreadId;
                    thread.Name = "ActiveN Trace Writer";

                    var dir = Path.Combine(Path.GetTempPath(), "ActiveN", SystemUtilities.CurrentProcess.ProcessName);
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
                    _traceWriter?.WriteLine($"[{DateTime.Now:HH:mm:ss.fff}][{_threadId}]Log starting, bitness: {nint.Size} process: {SystemUtilities.CurrentProcess.MainModule?.FileName ?? SystemUtilities.CurrentProcess.ProcessName} lowbox: {SystemUtilities.IsAppContainer()}");
                    return true;
                });
            }
            else if (!value && _traceWriter != null)
            {
                var tw = Interlocked.Exchange(ref _traceWriter, null);
                if (tw != null)
                {
                    tw.Flush();
                    tw.Dispose();
                }

                Interlocked.Exchange(ref _writerScheduler, null)?.Dispose();
            }
        }
    }

    private static void TraceToWriter(string? text = null, string? methodName = null, string? filePath = null)
    {
        if (Environment.CurrentManagedThreadId == _threadId)
        {
            trace();
            return;
        }

        var scheduler = _writerScheduler;
        if (scheduler != null)
        {
            _ = Task.Factory.StartNew(trace, CancellationToken.None, TaskCreationOptions.None, scheduler);
        }

        void trace()
        {
            _traceWriter?.WriteLine($"[{DateTime.Now:HH:mm:ss.fff}][{Environment.CurrentManagedThreadId}]{(filePath != null ? Path.GetFileNameWithoutExtension(filePath) : null)}:: {methodName}: {text}");
        }
    }

    public static void Flush()
    {
        // don't use scheduler thread for this
        _traceWriter?.Flush();
        _writerScheduler?.TriggerDequeue();
    }

    public static void Trace(string? text = null, [CallerMemberName] string? methodName = null, [CallerFilePath] string? filePath = null)
    {
        TraceToWriter(text, methodName, filePath);
        var name = filePath != null ? Path.GetFileNameWithoutExtension(filePath) : null;
        EventProvider.Default.WriteMessageEvent($"[{Environment.CurrentManagedThreadId}]{name}:: {methodName}: {text}");
    }

    public static T? WrapErrors<T>(this Func<T> action, Action? actionOnError = null, [CallerMemberName] string? methodName = null, [CallerFilePath] string? filePath = null)
    {
        ArgumentNullException.ThrowIfNull(action);
        try
        {
            return action();
        }
        catch (SecurityException se)
        {
            // transform this one as a well-known access denied
            Trace($"Security Error: {se}", methodName, filePath);
            SetError(se.GetInterestingExceptionMessage(), methodName, filePath);
            if (actionOnError != null)
            {
                try
                {
                    actionOnError();
                }
                catch (Exception ex2)
                {
                    Trace($"Security Error2: {ex2}", methodName, filePath);
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
            Trace($"Error: {ex}", methodName, filePath);
            SetError(ex.GetInterestingExceptionMessage(), methodName, filePath);
            if (actionOnError != null)
            {
                try
                {
                    actionOnError();
                }
                catch (Exception ex2)
                {
                    Trace($"Error2: {ex2}", methodName, filePath);
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

    public static HRESULT WrapErrors(this Func<HRESULT> action, Action? actionOnError = null, [CallerMemberName] string? methodName = null, [CallerFilePath] string? filePath = null)
    {
        ArgumentNullException.ThrowIfNull(action);
        try
        {
            var hr = action();
            if (hr != 0)
            {
                Trace($"COM Error: {hr}", methodName, filePath);
                if (actionOnError != null)
                {
                    try
                    {
                        actionOnError();
                    }
                    catch (Exception ex2)
                    {
                        Trace($"COM Error2: {ex2}", methodName, filePath);
                        // continue;
                    }
                }
            }
            return hr;
        }
        catch (SecurityException se)
        {
            // transform this one as a well-known access denied
            Trace($"Security Error: {se}", methodName, filePath);
            SetError(se.GetInterestingExceptionMessage(), methodName, filePath);
            if (actionOnError != null)
            {
                try
                {
                    actionOnError();
                }
                catch (Exception ex2)
                {
                    Trace($"Security Error2: {ex2}", methodName, filePath);
                    // continue;
                }
            }
            return Constants.E_ACCESSDENIED;
        }
        catch (Exception ex)
        {
            Trace($"Error: {ex}", methodName, filePath);
            SetError(ex.GetInterestingExceptionMessage(), methodName, filePath);
            if (actionOnError != null)
            {
                try
                {
                    actionOnError();
                }
                catch (Exception ex2)
                {
                    Trace($"Error2: {ex2}", methodName, filePath);
                    // continue;
                }
            }

            return ex.HResult;
        }
    }

    private static string? GetSource(string? methodName, string? filePath)
    {
        var source = methodName.Nullify();
        filePath = filePath.Nullify();
        if (source == null)
            return filePath;

        if (filePath == null)
            return source;

        return $"{filePath}:{source}";
    }

    public static void SetError(string? description, [CallerMemberName] string? methodName = null, [CallerFilePath] string? filePath = null)
    {
        if (string.IsNullOrWhiteSpace(description))
            return;

        ComError.SetError(description, GetSource(methodName, filePath));
    }
}
