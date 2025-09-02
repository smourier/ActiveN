namespace ActiveN.Samples.PdfView;

public class PdfViewWindow : Window
{
    private const int _buttonsHeight = 24;
    private const int _buttonsPadding = 10;
    private const int _buttonsWidth = 100;
    private const string _title = "ActiveN Pdf View";

    private Font? _font;
    private PdfWindow? _pdfWindow;
    private PdfDocument? _pdfDocument;
    private PdfPage? _pdfPage;
    private HWND _previousButton;
    private HWND _nextButton;
    private bool _showControls = true;
    private D3DCOLORVALUE _backgroundColor = D3DCOLORVALUE.White;

    public event EventHandler? FileOpened;
    public event EventHandler? FileClosed;
    public event EventHandler? PageChanged;

    public PdfViewWindow(HWND parentHandle, WINDOW_STYLE style, RECT rect)
        : base(_title, parentHandle: parentHandle, style: style, rect: rect)
    {
        // we need another (child) window for rendering PDF
        // as DirectX eats everything in a window and we wouldn't be able to see the buttons
        _pdfWindow = new PdfWindow(this);

        // add buttons to the window
        CreateButton("Open", _buttonsPadding, _buttonsPadding, _buttonsWidth, _buttonsHeight, (int)ButtonId.Open);
        _previousButton = CreateButton("Previous Page", _buttonsPadding + _buttonsWidth + _buttonsPadding, _buttonsPadding, _buttonsWidth, _buttonsHeight, (int)ButtonId.Previous);
        _nextButton = CreateButton("Next Page", _buttonsPadding + (_buttonsWidth + _buttonsPadding) * 2, _buttonsPadding, _buttonsWidth, _buttonsHeight, (int)ButtonId.Next);

        // set standard font for all buttons
        _font = GetMessageBoxFont();
        foreach (var win in AllChildWindows)
        {
            _font?.Set(win.Handle);
        }

        LoadPage(0);
    }

    protected PdfDocument? PdfDocument => _pdfDocument;
    protected PdfPage? PdfPage => _pdfPage;
    public string? FilePath { get; private set; }
    public bool IsPasswordProtected => _pdfDocument?.IsPasswordProtected ?? false;
    public int PageCount => (int?)_pdfDocument?.PageCount ?? -1;
    public virtual bool ShowControls
    {
        get => _showControls;
        set
        {
            if (_showControls == value)
                return;

            _showControls = value;
            RunTaskOnUIThread(UpdatePdfWindowSize, true);
        }
    }

    public virtual D3DCOLORVALUE BackgroundColor
    {
        get => _backgroundColor;
        set
        {
            if (_backgroundColor.Equals(value))
                return;

            _backgroundColor = value;
            _pdfWindow?.Invalidate();
        }
    }

    public virtual async Task OpenFile(string filePath)
    {
        ArgumentNullException.ThrowIfNull(filePath);
        var file = await StorageFile.GetFileFromPathAsync(filePath);
        if (file != null)
        {
            await OpenFile(file);
        }
    }

    public virtual async Task OpenFile(StorageFile file)
    {
        ArgumentNullException.ThrowIfNull(file);
        try
        {
            _pdfDocument = await PdfDocument.LoadFromFileAsync(file);
            FilePath = file.Path;
            _pdfPage?.Dispose();

            // this is not currently handled
            if (_pdfDocument.IsPasswordProtected)
                throw new Exception("The PDF file is password protected.");

            if (_pdfDocument.PageCount == 0)
                throw new Exception("The PDF file has no page to display.");

            _ = RunTaskOnUIThread(OnFileOpened);
            LoadPage(0);
        }
        catch (Exception ex)
        {
            throw new Exception("Failed to open PDF file.", ex);
        }
    }

    public virtual void MovePage(int delta)
    {
        if (_pdfPage == null || _pdfDocument == null)
            return;

        var page = _pdfPage.Index + delta;
        if (page < 0 || page >= _pdfDocument.PageCount)
            return;

        if (page == _pdfPage.Index)
            return;

        LoadPage((uint)page);
    }

    public virtual void CloseFile()
    {
        if (_pdfDocument == null)
            return;

        _pdfDocument = null;
        LoadPage(0);
        _ = RunTaskOnUIThread(OnFileClosed);
    }

    protected virtual void OnFileOpened() => FileOpened?.Invoke(this, EventArgs.Empty);
    protected virtual void OnFileClosed() => FileClosed?.Invoke(this, EventArgs.Empty);
    protected virtual void OnPageChanged() => PageChanged?.Invoke(this, EventArgs.Empty);

    private async Task OpenFile()
    {
        var picker = new FileOpenPicker();
        picker.FileTypeFilter.Add(".pdf");

        InitializeWithWindow.Initialize(picker, Handle);
        var file = await picker.PickSingleFileAsync();
        if (file != null)
        {
            await OpenFile(file);
        }
    }

    private void LoadPage(uint index)
    {
        _pdfPage?.Dispose();
        _pdfPage = _pdfDocument?.GetPage(index);
        _pdfWindow?.Invalidate();
        UpdatePdfWindowSize();
        if (_pdfPage != null)
        {
            Text = $"{_title} - {FilePath} - Page {_pdfPage.Index + 1} / {_pdfDocument!.PageCount}";
            Functions.EnableWindow(_previousButton, _pdfPage.Index > 0);
            Functions.EnableWindow(_nextButton, _pdfPage.Index < _pdfDocument.PageCount - 1);
        }
        else
        {
            Text = _title;
            Functions.EnableWindow(_previousButton, false);
            Functions.EnableWindow(_nextButton, false);
        }

        _ = RunTaskOnUIThread(OnPageChanged);
    }

    protected override LRESULT? WindowProc(HWND hwnd, uint msg, WPARAM wParam, LPARAM lParam)
    {
        // handle buttons clicks
        switch (msg)
        {
            case MessageDecoder.WM_COMMAND:
                var id = wParam.Value.LOWORD();
                switch ((ButtonId)id)
                {
                    case ButtonId.Open:
                        _ = OpenFile();
                        break;

                    case ButtonId.Next:
                        MovePage(1);
                        break;

                    case ButtonId.Previous:
                        MovePage(-1);
                        break;
                }
                return 0; // handled
        }
        return base.WindowProc(hwnd, msg, wParam, lParam);
    }

    // resize pdf window to fit the client area of this window w/o buttons
    private void UpdatePdfWindowSize()
    {
        var win = _pdfWindow;
        if (win != null)
        {
            RECT rc;
            int offset;
            if (ShowControls)
            {
                rc = ClientRect;
                offset = _buttonsHeight + 2 * _buttonsPadding;
            }
            else
            {
                rc = ClientRect;
                offset = 0;
            }
            win.ResizeAndMove(0, offset, rc.Width, rc.Height - offset);
        }
    }

    protected override bool OnResized(WindowResizedType type, SIZE size)
    {
        if (type != WindowResizedType.Minimized)
        {
            UpdatePdfWindowSize();
        }
        return base.OnResized(type, size);
    }

    protected override DiagnosticsInformation CreateDiagnosticsInformation() => new(Assembly.GetExecutingAssembly(), this, ", ");
    protected override void Dispose(bool disposing)
    {
        TracingUtilities.Trace($"disposing: {disposing}");
        Interlocked.Exchange(ref _pdfWindow, null)?.Dispose();
        Interlocked.Exchange(ref _pdfPage, null)?.Dispose();
        Interlocked.Exchange(ref _font, null)?.Dispose();
        base.Dispose(disposing);
    }

    private enum ButtonId
    {
        Open = 1,
        Next = 2,
        Previous = 3,
    }

    private sealed class PdfWindow(PdfViewWindow parent)
        : CompositionWindow(_title, WINDOW_STYLE.WS_VISIBLE | WINDOW_STYLE.WS_CHILD, parentHandle: parent.Handle)
    {
        private ComObject<IPdfRendererNative>? _pdfRendererNative;

        protected override void CreateDeviceResources()
        {
            base.CreateDeviceResources();

            using var device = Device.As<IDXGIDevice>()!;
            Functions.PdfCreateRenderer(device.Object, out var obj).ThrowOnError();
            _pdfRendererNative = new ComObject<IPdfRendererNative>(obj);
        }

        private static Size GetScaleFactor(Size availableSize, Size contentSize)
        {
            var scaleX = contentSize.Width == 0 ? 0 : availableSize.Width / contentSize.Width;
            var scaleY = contentSize.Height == 0 ? 0 : availableSize.Height / contentSize.Height;
            var minscale = scaleX < scaleY ? scaleX : scaleY;
            scaleX = scaleY = minscale;
            return new Size(scaleX, scaleY);
        }

        protected unsafe override void Render(IComObject<ID3D11DeviceContext> deviceContext, IComObject<IDXGISwapChain1> swapChain)
        {
            if (GraphicsDevice == null)
                return;

            var surface = GraphicsDevice.CreateDrawingSurface(ClientRect.ToSize(), DirectXPixelFormat.B8G8R8A8UIntNormalized, DirectXAlphaMode.Premultiplied);
            RootVisual.Brush = Compositor.CreateSurfaceBrush(surface);
            using var interop = surface.AsComObject<ICompositionDrawingSurfaceInterop>();
            using var dc = interop.BeginDraw(null);
            dc.Clear(parent.BackgroundColor);

            if (parent._pdfPage != null && _pdfRendererNative != null)
            {
                var pageUnk = ((IWinRTObject)parent._pdfPage).NativeObject.ThisPtr; // no AddRef needed

                // resize to fit window's height or width
                var rc = ClientRect;
                var size = parent._pdfPage.Size;

                var factor = GetScaleFactor(rc.ToSize(), size);
                var width = size.Width * factor.Width;
                var height = size.Height * factor.Height;

                // center if needed
                // there may be an existing transform, so we need to combine it and restore it
                D2D_MATRIX_3X2_F? xf = null;
                if (width < rc.Width || height < rc.Height)
                {
                    xf = dc.GetTransform();
                    dc.SetTransform(xf.Value * D2D_MATRIX_3X2_F.Translation(
                        Math.Max(0, (float)((rc.Width - width) / 2)),
                        Math.Max(0, (float)((rc.Height - height) / 2)))
                        );
                }

                // note sure why but PDF_RENDER_PARAMS's BackgroundColor seems to not be used
                // so we do it ourselves
                using var backgroundBrush = dc.CreateSolidColorBrush(D3DCOLORVALUE.White);
                dc.FillRectangle(new D2D_RECT_F(0, 0, width, height), backgroundBrush);

                var renderParams = new PDF_RENDER_PARAMS
                {
                    //BackgroundColor = D3DCOLORVALUE.White, // doesn't seem to be honored
                    DestinationWidth = (uint)width,
                    DestinationHeight = (uint)height,
                };

                // render
                _pdfRendererNative.Object.RenderPageToDeviceContext(pageUnk, dc.Object, (nint)(&renderParams)).ThrowOnError();

                // restore previous transform
                if (xf.HasValue)
                {
                    dc.SetTransform(xf.Value);
                }
            }

            interop.EndDraw();
        }

        protected override void Dispose(bool disposing)
        {
            TracingUtilities.Trace($"disposing: {disposing}");
            Interlocked.Exchange(ref _pdfRendererNative, null)?.Dispose();
            base.Dispose(disposing);
        }
    }
}
