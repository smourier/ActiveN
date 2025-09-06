' VBScript is an IDispatch-only client, so we can only call methods and properties of ISimpleDual interface
Set server = CreateObject("ActiveN.Samples.SimpleComObject.SimpleComObject")
WScript.Echo server.Name
WScript.Echo "6 x 7 = " & server.Multiply(6, 7)
WScript.Echo "6 + 7 = " & server.Add(6, 7)
