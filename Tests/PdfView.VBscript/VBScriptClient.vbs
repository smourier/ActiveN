Set server = CreateObject("ActiveN.Samples.PdfView.PdfViewControl")

server.OpenFile "sample.pdf"
WScript.Echo "Number of pages: " & server.PageCount

For i = 0 to 9
    Set page = server.GetPage(i)
    WScript.Echo "Extracting page " & page.Index & ": " & page.Width & " x " & page.Height & " (rotation: " & page.Rotation & ")"
    page.ExtractTo "bin\sample_page_" & page.Index & ".png"
Next

