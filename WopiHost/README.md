Introduction
==========
[![Build status](https://ci.appveyor.com/api/projects/status/l7jn00f4fxydpbed?svg=true)](https://ci.appveyor.com/project/petrsvihlik/wopihost) 
[![codecov](https://codecov.io/gh/petrsvihlik/WopiHost/branch/master/graph/badge.svg)](https://codecov.io/gh/petrsvihlik/WopiHost) 
[![Gitter](https://img.shields.io/gitter/room/nwjs/nw.js.svg)](https://gitter.im/ms-wopi/)
[![.NET Core](https://img.shields.io/badge/netcore-2.0-692079.svg)](https://www.microsoft.com/net/learn/get-started/windows)


This project is a sample implementation of a [WOPI host](http://blogs.msdn.com/b/officedevdocs/archive/2013/03/20/introducing-wopi.aspx). Basically, it allows developers to integrate custom datasources with Office Online Server (formerly Office Web Apps) or any other WOPI client by implementing a bunch of interfaces.

Features / improvements compared to existing samples on the web
-----------------------
 - clean WebAPI built with ASP.NET Core MVC, no references to System.Web
 - uses new ASP.NET Core features (configuration, etc.)
 - can be self-hosted or run under IIS
 - file manipulation is extracted to own layer of abstraction (there is no dependency on System.IO)
   - example implementation included (provider for Windows file system)
   - file identifiers can be anything (doesn't have to correspond with the file's name in the file system)
 - custom token authentication middleware
 - DI used everywhere
 - URL generator
   - based on a WOPI discovery module
 - all references are NuGets
 
Usage
=====

Prerequisites
-------------
 - [Visual Studio 2017 with the .NET Core workload](https://www.microsoft.com/net/core)

Building the app
----------------
The WopiHost app targets `net46` and `netcoreapp2.0`. You can choose which one you want to use. If you get errors that Microsoft.CobaltCore.15.0.0.0.nupkg can't be found then just remove the reference or see the chapter "Cobalt" below.
 
Configuration
-----------
WopiHost.Web\Properties\launchSettings.json
- `WopiHostUrl` - used by URL generator
- `WopiClientUrl` - used by the discovery module to load WOPI client URL templates

WopiHost\Properties\launchSettings.json
- `WopiFileProviderAssemblyName` - name of assembly containing implementation of WopiHost.Abstractions interfaces
- `WopiRootPath` - provider-specific setting used by WopiFileSystemProvider (which is an implementation of IWopiFileProvider working with System.IO)
- `server.urls` - hosting URL(s) used by Kestrel. [Read more...](http://andrewlock.net/configuring-urls-with-kestrel-iis-and-iis-express-with-asp-net-core/)

Running the application
-----------------------
Once you've successfully built the app you can:

- run it directly from the Visual Studio using [IIS Express or selfhosted](/img/debug.png?raw=true).
  - make sure you run both `WopiHost` and `WopiHost.Web`. You can set them both as [startup projects](/img/multiple_projects.png?raw=true)
- run it from the `cmd`
  - navigate to the WopiHost folder and run `dotnet run`
- run it in IIS (tested in IIS 8.5)
  - navigate to the WopiHost folder and run `dnu publish --runtime active`
  - copy the files from WopiHost\bin\output to your desired web application directory
  - run the web.cmd file as administrator, wait for it to finish and close it (Ctrl+C and y)
  - create a new application in IIS and set the physical path to the wwwroot in the web application directory
  - make sure the site you're adding it to has a binding with port 5000
  - go to the application settings and change the value of `dnx clr` to `clr` and the value of `dnx-version` to `1.0.0-rc1`
  - in the same window, add all the configuration settings

Compatible WOPI Clients
-------
Running the application only makes sense with a WOPI client as its counterpart. WopiHost is compatible with the following clients:

 - Office Online Server 2016 ([deployment guidelines](https://technet.microsoft.com/en-us/library/jj219455(v=office.16).aspx))
 - Office Online https://wopi.readthedocs.io/en/latest/

Note that WopiHost will always be compatible only with the latest version of OOS because Microsoft also [supports only the latest version](https://blogs.office.com/2016/11/18/office-online-server-november-release/).

The deployment of OOS/OWA requires the server to be part of a domain. If your server is not part of any domain (e.g. you're running it in a VM sandbox) it can be overcame e.g. by installing [DC role](http://social.technet.microsoft.com/wiki/contents/articles/12370.windows-server-2012-set-up-your-first-domain-controller-step-by-step.aspx). After it's deployed you can safely remove the role and the OWA server will remain functional.
To test your OWA server [follow the instructions here](https://blogs.technet.microsoft.com/office_web_apps_server_2013_support_blog/2013/12/27/how-to-test-viewing-office-documents-using-the-office-web-apps-2013-viewer/).
To remove the OWA instance use [`Remove-OfficeWebAppsMachine`](http://sharepointjack.com/2014/fun-configuring-office-web-apps-2013-owa/).

Cobalt
------
In the past (in Office Web Apps 2013), some actions required support of MS-FSSHTTP protocol (also known as "cobalt"). This is no longer true with Office Online Server 2016.
However, if the WOPI client discovers (via [SupportsCobalt](http://wopi.readthedocs.io/projects/wopirest/en/latest/files/CheckFileInfo.html#term-supportscobalt) property) that the WOPI host supports cobalt, it'll use it as it's more efficient.

If you want to make the project work with Office Web Apps 2013 SP1 ([deployment guidelines](https://technet.microsoft.com/en-us/library/jj219455(v=office.15).aspx)), you'll need to create a NuGet package called Microsoft.CobaltCore.15.0.0.0.nupkg containing Microsoft.CobaltCore.dll. This DLL is part of Office Web Apps 2013 / Office Online Server 2016 and its license doesn't allow public distribution and therefore it's not part of this repository. Please make sure your OWA/OOS server and user connecting to it have valid licenses before you start using it.

 1. Locate Microsoft.CobaltCore.dll (you can find it in the GAC of the OWA server): `C:\Windows\Microsoft.NET\assembly\GAC_MSIL\Microsoft.CobaltCore\`
 2. Install [NuGet Package Explorer](https://npe.codeplex.com/)
 3. Use .nuspec located in the [Microsoft.CobaltCore](https://github.com/petrsvihlik/WopiHost/tree/master/Microsoft.CobaltCore) folder to create new package
 4. Put the .nupkg to your local NuGet feed
 5. Configure Visual Studio to use your local NuGet feed

Note: the Microsoft.CobaltCore.dll targets the full .NET Framework, it's not possible to use it in an application that targets .NET Core.


