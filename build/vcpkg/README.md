# Using wxAutoExcel with vcpkg

wxAutoExcel, being a superniche library, does not have an official
vcpkg port. However, a custom overlay is provided for convenience here.
Tested only with x64-windows and x64-windows-static.

## Installing
1. Optionally, run *update-ref.bat* to make the portfile use the latest
commit in the wxAutoExcel master branch (Requires GIT in path); or 
manually update REF and SHA512 there as needed. As the portfile is bundled
with the repo, its REF cannot port to itself (unknown at commit time).
2. Copy folder *WXAUTOEXCEL-SRCDIR/build/vcpkg/wxautoexcel*
to the folder with your custom vcpgk ports, e.g. *c:/dev/libs/my-vcpkg-ports*.
3. Install the port with `vcpkg install wxautoexcel --overlay-ports=c:/dev/libs/my-vcpkg-ports/wxautoexcel`
(using the actual overlay port path).

## Using
The only difference compared to using an official port is that you must
specify the path to the overlay port in vcpkg.

You can add the path either to the `VCPKG_OVERLAY_PORTS` environment variable or
to the `overlay-ports` field in your project's *vcpkg-configuration.json*.