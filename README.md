# ActiveN

ActiveN is a lightweight framework for building classic COM components and OLE/ActiveX controls in modern fully AOT-compatible .NET, with registration-ready deployment.

It lets you author controls and automation objects that run inside legacy or current COM hosts (VBA, VB6, scripting engines, test containers, etc.) without relying on WinForms or WPF (since they are currently *not* AOT-compatible).

## Why ActiveN?

.NET Core doesn't provide any facility for building OLE/ActiveX controls and lacks some important COM features. ActiveN focuses on:
- Relatively small, self-contained, AOT-publishable binaries
- Explicit control over COM identity (GUIDs, interfaces, type libraries)
- Support for aggregation (critical for host compatibility like Excel)
- Control custom implementation (in-place activation, UI, events, persistence)
- x86, x64, and ARM64 support
- Hosting DirectX (Direct3D11, Direct2D, DirectComposition, etc.) content in legacy hosts

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
- HKCU and HKLM registration support
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
> ActiveN doesn't use the `EnableComHosting` MSBuild property since it doesn't provide (as of today) the necessary AOT compatibility.

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

 ![Debug Build](/assets/debug_build.png?raw=true)

### Release
__Release__ builds are not supposed to be used directly: they are only intermediate steps before Native AOT publishing.
 
For Native AOT publishing, you can use Visual Studio __Publish__ command: right-click project, select Publish..., then use provided sample .pubxml files or create one.
For example, you can use the provided *FolderProfile.x64.pubxml* to publish a self-contained 64-bit binary.

Here is an example of a publish profile for a self-contained 32-bit binary:

 ![Publish Profile](/assets/publish_x86.png?raw=true)]

And here is an example of a Release build output, which includes the one and only dll and extra .PDB files to ease debugging (that you don't need to distribute):

 ![Release Build](/assets/release_build.png?raw=true)

## Registration
Since we're building true COM components, registration is done via `regsvr32` (no need for RegAsm or similar tools).
In Debug mode, you register the *myCustom.comthunk.dll* file, while in Release/Publish mode, you register the single *myCustom.dll* file.

Contrary to .NET Core built-in registration (see this https://github.com/dotnet/runtime/issues/45750), ActiveN supports HKCU and HKLM registration

By default, DLL registration depends on the privileges of the current user:

* If running elevated (as Administrator): registration is performed under HKLM (machine-wide).

* If running without elevation: registration is performed under HKCU (per-user).

You can explicitly override this default behavior when calling regsvr32:

    regsvr32 myCustom.dll
=> Uses the default behavior (HKLM if elevated, otherwise HKCU).

    regsvr32 /n /i:user myCustom.dll
=> Forces registration under HKCU (per-user).

    regsvr32 /n /i:admin myCustom.dll
=> Forces registration under HKLM (machine-wide). Requires administrative privileges.


> Tip: Since Windows Vista or so, regsvr32 automatically detects the bitness of the target DLL. If needed, it will restart itself with the correct architecture (32-bit or 64-bit), so you don’t have to worry about manually choosing the right executable.

---
## Quick Start (Authoring a Custom COM Component / ActiveX Control)

1. Copy one of the sample projects as a template. The .csproj from the sample projects includes all necessary build steps and references.
2. Rename the project, namespaces, and classes.
3. Replace every GUID (in both `.IDL` and `.cs`) with newly generated ones.
4. Keep `.IDL` and C# interface definitions synchronized (the framework does not auto-generate .NET code from IDL code).
5. Implement your logic (methods, properties, events, windowing, rendering).
7. Register (via provided registration helper or `regsvr32` path).
8. Load in your target host and test.

> Recommendation: Make sure you always keep the .IDL and interfaces .cs in sync.

## Design Mode Property Pages (PropertyGrid Helper)

In traditional ActiveX controls, property pages are implemented as separate COM objects that the host can instantiate and display when the user wants to edit control properties.
The problem is they are based on Win32 dialog boxes, which are not easy to implement in .NET, especially in an AOT-compatible way.

To simplify this, ActiveN provides an optional helper library, `ActiveN.PropertyGrid`, that implements a generic property page based on the standard **.NET Framework 4** `PropertyGrid` control.

Yes, you read correctly: the .NET Framework 4 Winforms' `PropertyGrid` control is used in a .NET Core AOT-compatible component. The reasons are:
- The `PropertyGrid` control is a standard Winforms control that is part of .NET Framework.
- The .NET framework is itself part of the OS now (no need to deploy .NET Framework 4 with your component).
- .NET Core's Winforms and WPF are not AOT-compatible (yet?), so we can't use them directly.

So in the end, the .NET Framework is very AOT-compatible and what happens is ActiveN hosts the .NET Framework CLR to display the property page, while your control and its logic (property get/set) remain in .NET Core.

> This is optional, ActiveN doesn't require this. You can still implement your own property pages using Win32 if you want.

Here is a screenshot of the PropertyGrid-based property page in action (from the `ActiveN.Samples.HelloWorld` sample) in an Excel Spreadsheet in design mode:

 ![Property Page](/assets/custom_property_page.png?raw=true)

 ## Some screenshots:

 Microsoft VB6 Hosting WebView2 (design mode)
 
 <img width="1361" height="784" alt="VB6 ActiveX Control .NET Core Design" src="https://github.com/user-attachments/assets/08b4eee8-b895-4fb8-b963-a7fcdbec7105" />

 User mode:

 <img width="836" height="547" alt="VB6 ActiveX Control .NET Core" src="https://github.com/user-attachments/assets/8461f3c9-bf58-4bf9-9b7f-060629095375" />

 VB6 Code:
 
<img width="383" height="218" alt="B6 ActiveX Control .NET Core Code" src="https://github.com/user-attachments/assets/ff67df06-4a02-4fe5-ba3e-b86063002797" />

 Microsoft Word (Office 365 x64) hosting PDF View (design mode)

 <img width="925" height="553" alt="Word ActiveX Control .NET Core PDF Design" src="https://github.com/user-attachments/assets/522fb215-59ac-4a5f-8caf-1681c27c4a03" />

 User mode:

<img width="989" height="772" alt="Word ActiveX Control .NET Core PDF" src="https://github.com/user-attachments/assets/d07b6da8-1307-4800-a779-4148b3ff1612" />

 Microsoft Excel (Office 365 x64) hosting WebView2 (user mode)

<img width="1197" height="767" alt="Excel ActiveX Control .NET Core WebView2" src="https://github.com/user-attachments/assets/c30bad01-f5e2-445c-94e9-9f955930cf21" />

Microsoft Visual FoxPro 9 hosting WebView2 (design mode):

<img width="1140" height="718" alt="Visual Foxpro 9 ActiveX Control .NET Core PDF" src="https://github.com/user-attachments/assets/8ae064ff-0892-41a8-ad6f-685b4400ab99" />
