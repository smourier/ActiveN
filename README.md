# ActiveN

ActiveN is a lightweight framework for building classic COM components and OLE/ActiveX controls in modern fully AOT-compatible .NET, with registration-ready deployment.

It lets you author controls and automation objects that run inside legacy or current COM hosts (VBA, VB6, scripting engines, test containers, etc.) without relying on WinForms or WPF (since they are currently *not* AOT-compatible).

## Why ActiveN?

Traditional .NET COM interop (RegAsm, COM-visible assemblies, RCWs) is not compatible with Native AOT and often pulls large runtime dependencies. ActiveN focuses on:
- Relatively small, self-contained, AOT-publishable binaries
- Explicit control over COM identity (GUIDs, interfaces, type libraries)
- Support for aggregation (critical for host compatibility like Excel)
- Control custom implementation (in-place activation, UI, events, persistence)
- Hosting DirectX (Direct2D, DirectComposition, etc.) content in legacy hosts

If you need to modernize or replace aging C++ / ATL / VB6 controls with .NET code while keeping host compatibility, this framework can help.

## Supported Hosts

Most hosts that load OLE/ActiveX controls (or late-bound automation objects) should work. Tested hosts include:

- Excel VBA
- Word VBA
- VBScript (cscript.exe)
- VB6 (yes, the "legacy" 32-bit one)
- Visual FoxPro 9 (yes the also "legacy" 32-bit one)
- TstCon – official [ActiveX Control Test Container](https://github.com/Microsoft/VCSamples/tree/master/VC2010Samples/MFC/ole/TstCon)
- TstCon64 – 64-bit fork: [TstCon64](https://github.com/smourier/TstCon64)

> Tip: Always test both 32-bit and 64-bit where applicable.

---

## Highlights

- COM hosting and registration support (full support for `regsvr32`)
- COM class, interface, and events authoring (IDispatch, dual & custom interfaces, TypeLib support)
- COM aggregation support (required for some hosts such as Excel VBA)
- ActiveX / OLE control support (IOleObject, IOleControl, IOleInPlaceActiveObject, IPersistStreamInit, etc.)
- IConnectionPointContainer / IConnectionPoint event sink infrastructure
- OLE control persistence support (state save/restore)
- Type library (TLB) generation from `.IDL` files (you own the IDL surface and its .NET counterpart)
- Native AOT publish path for minimal footprint and simplified deployment
- Facilitated .NET Core AOT development & debugging with [AotNetComHost](https://github.com/smourier/AotNetComHost)
- No dependency on WinForms or WPF (you will have to to windowing / rendering yourself by using 3rd parties or DirectX / Direct2D / DirectComposition code)

## Repository Structure

- **ActiveN** – Core library providing the COM/ActiveX infrastructure. Depends only on .NET and [DirectNAot](https://github.com/smourier/DirectNAot) for AOT-friendly Windows interop.
- **ActiveN.PropertyGrid** – Optional runtime helper enabling a PropertyGrid-backed "custom" OLE property page (see remarks below).

### ActiveX Sample Controls

- **ActiveN.Samples.SimpleComObject** – Simple automation object (no UI)
- **ActiveN.Samples.HelloWorld** – Minimal ActiveX control
- **ActiveN.Samples.Pdfview** – PDF viewer hosting sample
- **ActiveN.Samples.WebView2** – WebView2 browser hosting sample (uses [WebView2Aot](https://github.com/smourier/WebView2Aot))

### Tests

- **SimpleComObject.VBscript** – VBScript exercising IDispatch
- **PdfView.VBscript** – Extracts images from a PDF via the PdfView control
- **WebView2.Excel** – Excel VBA integration demo
- **WebView2.VB6** – VB6 host project

---
## Building and Publishing

Ensure the Windows SDK and MIDL toolchain are installed (Visual Studio workload: "Desktop development with C++") and run normal __Build__ in Visual Studio.

__Debug__ and __Release__ builds are handled differently. They use a custom import and target to generate the TLB and other Win32 resources for ActiveX support.
This custom step is included in the sample project files and invokes Windows SDK tools like MIDL and RC:

```xml
...
<Import Project="..\ActiveN\ComObjects.targets" />
...
``` 

### Debug
Debug builds are standard .NET builds (no AOT publishing, but AOT compatibility is checked) for easier debugging and faster iteration.

In this mode, the .csproj will create the regular .NET binaries, for example *myCustom.dll** and will add a *myCustom.comthunk.dll* file that is the native .dll that must be registered with regsvr32.
```xml
...
<Target Name="AotNetComHost" AfterTargets="Build" Condition="'$(Configuration)' == 'Debug'">
	<ItemGroup>
		<AotNetComHost Include="$(ProjectDir)..\ActiveN\External\AotNetComHost.$(Platform).dll" />
	</ItemGroup>
	<Copy SourceFiles="@(AotNetComHost)" DestinationFolder="$(TargetDir)" SkipUnchangedFiles="true" />
	<Move ContinueOnError="false" SourceFiles="$(OutDir)AotNetComHost.$(Platform).dll" DestinationFiles="$(OutDir)$(AssemblyName).comthunk.dll" />
</Target>
...
``` 

More info on this development/debugging technique is available in [AotNetComHost](https://github.com/smourier/AotNetComHost)

Here is an example of a Debug build output:

![image][Debug Build](assets/debug_build.png)

### Release
__Release__ builds are not supposed to be used directly: they are only intermediate steps before Native AOT publishing.
 
For Native AOT publishing, you can use Visual Studio __Publish__ command: right-click project, select Publish..., then use provided sample .pubxml files or create one.
For example, you can use the provided *FolderProfile.x64.pubxml* to publish a self-contained 64-bit binary.

Here is an example of a publish profile for a self-contained 32-bit binary:

![image][Publish Profile](assets/publish_x86.png)]

And here is an example of a Release build output, which includes the one and only dll and extra .PDB files to ease debugging (that you don't need to distribute):

![image][Release Build](assets/release_build.png)

## Registration
Since we're building true COM components, registration is done via `regsvr32` (no need for RegAsm or similar tools).
In Debug mode, you register the *myCustom.comthunk.dll* file, while in Release/Publish mode, you register the single *myCustom.dll* file.


## Testing

## Quick Start (Authoring a Custom COM Component / ActiveX Control)

1. Copy one of the sample projects as a template. The .csproj from the sample projects includes all necessary build steps and references.
2. Rename the project, namespaces, and classes.
3. Replace every GUID (in both `.IDL` and `.cs`) with newly generated ones.
4. Keep `.IDL` and C# interface definitions synchronized (the framework does not auto-generate .NET code from IDL code).
5. Implement your logic (methods, properties, events, windowing, rendering).
7. Register (via provided registration helper or `regsvr32` path).
8. Load in your target host and test.

> Recommendation: Maintain a short checklist script for re-generating TLB and re-registering after interface changes.

---

## Minimal Conceptual Flow


