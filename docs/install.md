

# Installing, Building, and Using wxAutoExcel

## 1 Installation

The first step, is to download the source archive or clone the code
from GitHub and put it in any directory. If at all possible, please
avoid a directory name with spaces in its name, as this can cause problems.
In the following text this directory will be referred to as `WXAUTOEXCEL-SRCDIR`.

## 2 Building wxAutoExcel

### 2.1 General information
wxAutoExcel is built in Debug and/or Release configurations with CMake (v3.14+ required),
`find_package()` is employed to find wxWidgets build to use. Microsoft Visual C++ or MinGW
(GCC or clang) compilers are supported.

Before building wxAutoExcel, you can change some of its compile-time
settings in *WXAUTOEXCEL-SRCDIR/include/wx/wxAutoExcel_setup.h*. These settings
are mostly useful if you don't plan to use some of wxAutoExcel features
such as Shapes or Charts and wish to minimize the size of the wxAutoExcel
(dynamic) library.

### 2.2 Build options

- `wxAutoExcel_BUILD_SHARED`: whether to build static (lib) or shared (dll) wxAutoExcel library.
  `BOOL`, defaults to the value of `BUILD_SHARED_LIBS`.
- `wxAutoExcel_BUILD_LINK_WX_SHARED`: whether to link to wxWidgets dynamically.
  When linking  dynamically, the `wxWidgets_LIB_DIR` CMake variable should be set to wxWidgets
  dll folder and not the lib folder.
  `BOOL`, defaults to the value of `BUILD_SHARED_LIBS`.
- `wxAutoExcel_BUILD_USE_STATIC_RUNTIME`: whether to link to the static CRT (and other compiler libraries).
  This setting must match the setting used when building wxWidgets (`wxBUILD_USE_STATIC_RUNTIME`) and cannot
  be used when building wxAutoExcel in the shared mode or linking to wxWidgets dynamically.
  `BOOL`, defaults to `OFF`.
- `wxAutoExcel_BUILD_USE_PRECOMPILED`: whether to use precompiled headers (requires CMake v3.16+).
  `BOOL`, defaults to `ON`.
- `wxAutoExcel_BUILD_VENDOR`: Short string identifying the library builder (used in the DLL name).
  `STRING`, defaults to `custom`.
- `wxAutoExcel_BUILD_INSTALL`: whether to install wxAutoExcel.
  `BOOL`, defaults to `OFF`.
- `wxAutoExcel_BUILD_SAMPLES`: whether to build the bundled samples.
  `BOOL`, defaults to the value of `PROJECT_IS_TOP_LEVEL`.

Static or import libraries will be built in the *WXAUTOEXCEL-BUILDDIR/lib*
folder. Dynamic libraries and samples' executables will be built in the
*WXAUTOEXCEL-BUILDDIR/bin* folder. When using a multi-config generator, the bin
folder contains subfolders for the individual configurations (e.g., Debug).
The same hierarchy is followed for the folder with installed wxAutoExcel
(*WXAUTOEXCEL-INSTALLDIR*) but subfolders for configurations are not used.

### 2.3 An Example
This example shows how to build wxAutoExcel, using MSVS 2022 CMake generator in Debug and Release shared configurations.
Let's say wxAutoExcel source directory (*WXAUTOEXCEL-SRCDIR*) is *c:/dev/libs/wxAutoExcel*, the build
directory (*WXAUTOEXCEL-BUILDDIR*) is *c:/dev/libs/wxAutoExcel-build-vc17-x64-DLL*, and the installed
directory (*WXAUTOEXCEL-INSTALLDIR*) is *c:/dev/libs/wxAutoExcel-build-vc17-x64-DLL-installed*.
All following commands are to be run from *WXAUTOEXCEL-BUILDDIR*.

#### Configure

    cmake -G "Visual Studio 17 2022" -DBUILD_SHARED_LIBS=ON -DwxAutoExcel_BUILD_SAMPLES=OFF -DwxAutoExcel_BUILD_INSTALL=ON -S ../wxAutoExcel -B .

#### Build

    cmake --build . --config Debug
    cmake --build . --config Release

#### Install (optional)
We can install Debug and Release build in the same folder:

    cmake --install . --config Debug --prefix ../wxAutoExcel-build-vc17-x64-DLL-installed
    cmake --install . --config Release --prefix ../wxAutoExcel-build-vc17-x64-DLL-installed

### 2.4 Library naming scheme

The names of static and import libraries include version and indication whether
this is a Debug configuration (those have *d* after the version number).
The names of shared (DLL) libraries include more detailed version and provide also
compiler, architecture, and Vendor identification. For example this is how the file
names will look for Debug configuration of version 2.0.0 with Vendor set to default
*custom*.

#### For MSVC 2022 x64
*Import library:* wxAutoExcel20d.lib
*DLL:* wxAutoExcel200d_vc143_x64_custom.dll

#### For GCC 13.2.0 x86
*Import library:* libwxAutoExcel20d.a
*DLL:* wxAutoExcel200d_gcc1320_x32_custom.dll

#### For clang 17.0.0 x64
*Import library:* libwxAutoExcel20d.a
*DLL:* wxAutoExcel200d_clang1700_x64_custom.dll

## 3 Building applications using wxAutoExcel

### 3.1 CMake with `find_package()`

The first step is calling `find_package(wxAutoExcel REQUIRED CONFIG)` 
and the second step is calling `target_link_libraries()`, using `wxAutoExcel`
as the library name. For example, your application CMakeLists.txt may look like this

    # .....
    find_package(wxWidgets 3.2.0 REQUIRED COMPONENTS core base)
    include(${wxWidgets_USE_FILE})
    
    find_package(wxAutoExcel REQUIRED CONFIG)
    # .....
    target_link_libraries(${PROJECT_NAME} PRIVATE wxAutoExcel ${wxWidgets_LIBRARIES})

If you did not install wxAutoExcel to a location known to CMake, you may need to tell CMake where to find it.
The easiest way may be setting CMake variable `wxAutoExcel_ROOT` to *WXAUTOEXCEL-BUILDDIR* or *WXAUTOEXCEL-INSTALLDIR*.
For example, you may build the application from its build folder, setting `wxAutoExcel_ROOT`, like this

    cmake -G "Visual Studio 17 2022" -DwxAutoExcel_ROOT=c:/dev/libs/wxAutoExcel-build-vc17-x64-DLL-installed -S ../MyApp -B .

### 3.2 CMake with `add_subdirectory()`

The first step is calling `add_subdirectory()` and the second step is calling
`target_link_libraries()`, using `wxAutoExcel` as the library name.

For example, if you have wxAutoExcel source code in your application's folder *3rdparty/wxAutoExcel*
(e.g., *c:/dev/apps/MyApp/3rdparty/wxAutoExcel*), your application CMakeLists.txt may look like this

    # ..... make sure wxWidgets is available
    # maybe set some of wxAutoExcel_BUILD options here....
    add_subdirectory(3rdparty/wxAutoExcel)
    # ..... 
    target_link_libraries(${PROJECT_NAME} PRIVATE wxAutoExcel wx::core wx::base)

### 3.3 Manual setup

If you do not wish to build your application with CMake, then in addition to the standard wxWidgets-related
and other settings, you need to: 
- Add `WXAUTOEXCEL-SRCDIR/include` or `WXAUTOEXCEL-INSTALLDIR/include` to the compiler include path.
- Add `WXAUTOEXCEL-BUILDDIR/lib/[Debug or Release]` or `WXAUTOEXCEL-INSTALLDIR/lib` to the libraries path.
- If you are using the DLL build of wxAutoExcel, add `WXUSINGDLL_WXAUTOEXCEL`
  to the preprocessor definitions.
- Finally, add the appropriate wxAutoExcel (import) library (see the Library Naming Scheme chapter) to the
  list of libraries to link with, for Debug and Release targets separately.
