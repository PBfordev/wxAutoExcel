# Installing, Building, and Using wxAutoExcel

## 1 Installation

The first step, is to download the source archive or clone the code
from GitHub and put it in any directory. If at all possible, please
avoid a directory name with spaces in its name, as this can cause problems.
In the following text this directory will be referred to as `WXAUTOEXCEL-SRCDIR`.

## 2 Building wxAutoExcel

### 2.1 General information
wxAutoExcel is built in Debug and/or Release configurations with CMake (v3.16+ required),
`find_package()` is employed to find wxWidgets build to use. Microsoft Visual C++ and MinGW
(GCC or clang) compilers are supported.

Before building wxAutoExcel, you can change some of its compile-time
settings in *WXAUTOEXCEL-SRCDIR/include/wx/wxAutoExcel_setup.h*. These settings
are mostly useful if you don't plan to use some of wxAutoExcel features
such as Shapes or Charts and wish to minimize the size of the wxAutoExcel
(dynamic) library.

### 2.2 CMake build options

- `wxAutoExcel_BUILD_SHARED`: whether to build static or shared (dll) wxAutoExcel library.
  `BOOL`, defaults to the value of `BUILD_SHARED_LIBS`.
- `wxAutoExcel_BUILD_LINK_WX_SHARED`: whether to link to wxWidgets dynamically.
  When linking  dynamically, the `wxWidgets_LIB_DIR` CMake variable should be set to wxWidgets
  dll folder and not the lib folder.
  `BOOL`, defaults to the value of `BUILD_SHARED_LIBS`.
- `wxAutoExcel_BUILD_USE_STATIC_RUNTIME`: whether to link to the static CRT (and other compiler libraries).
  This setting must match the setting used when building wxWidgets (`wxBUILD_USE_STATIC_RUNTIME`) and cannot
  be used when building wxAutoExcel as shared or linking to wxWidgets dynamically.
  `BOOL`, defaults to `OFF`, unless using an MSVC generator with `CMAKE_MSVC_RUNTIME_LIBRARY` set to use
  a statically-linked CRT library.
- `wxAutoExcel_BUILD_USE_PRECOMPILED`: whether to use precompiled headers.
  `BOOL`, defaults to `ON`.
- `wxAutoExcel_BUILD_VENDOR`: Short string identifying the library builder (used in the DLL name).
  `STRING`, defaults to `custom`.
- `wxAutoExcel_BUILD_INSTALL`: whether to install wxAutoExcel.
  `BOOL`, defaults to `OFF`.
- `wxAutoExcel_BUILD_SAMPLES`: whether to build the bundled samples.
  `BOOL`, defaults to the value of `PROJECT_IS_TOP_LEVEL`.

### 2.3 Example
This example shows how to build wxAutoExcel, using MSVS 2022 CMake generator in Debug and Release shared configurations.
Let's say wxAutoExcel source directory (*WXAUTOEXCEL-SRCDIR*) is *c:/dev/libs/wxAutoExcel*, the build
directory (*WXAUTOEXCEL-BUILDDIR*) is *c:/dev/libs/wxAutoExcel-build-vc17-x64-DLL*, and the installed
directory (*WXAUTOEXCEL-INSTALLDIR*) is *c:/dev/libs/wxAutoExcel-vc17-x64-DLL-installed*.
All following commands are to be run from *WXAUTOEXCEL-BUILDDIR*.

#### Configure

    cmake -G "Visual Studio 17 2022" -DBUILD_SHARED_LIBS=ON -DwxAutoExcel_BUILD_INSTALL=ON -S ../wxAutoExcel -B .

#### Build

    cmake --build . --config Debug
    cmake --build . --config Release

#### Install (optional)
We can install Debug and Release builds in the same folder:

    cmake --install . --config Debug --prefix ../wxAutoExcel-vc17-x64-DLL-installed
    cmake --install . --config Release --prefix ../wxAutoExcel-vc17-x64-DLL-installed

### 2.4 Library naming scheme

The names of static and import libraries include version and indication whether
this is a Debug configuration (those have *d* after the version number).
The names of shared (DLL) libraries include more detailed version and provide also
compiler, architecture, and Vendor identification.

## 3 Building applications using wxAutoExcel

### 3.1 CMake with `find_package()`

The first step is calling `find_package(wxAutoExcel REQUIRED CONFIG)` 
and the second step is calling `target_link_libraries()`, using `wxAutoExcel::wxAutoExcel`
as the library target name. For example, your application CMakeLists.txt may look like this

    # .....
    find_package(wxWidgets 3.2 REQUIRED COMPONENTS core base)
    if(wxWidgets_USE_FILE)
      include(${wxWidgets_USE_FILE})
    endif()
    
    find_package(wxAutoExcel 2.0 REQUIRED CONFIG)
    # .....
    target_link_libraries(${PROJECT_NAME} PRIVATE wxAutoExcel::wxAutoExcel ${wxWidgets_LIBRARIES})

If you did not install wxAutoExcel to a location known to CMake, you may need to tell CMake where to find it.
There are many ways to do that, perhaps the easiest one may be setting CMake variable `wxAutoExcel_ROOT`
(requires declaring minimum CMake version as 3.12 or newer) to *WXAUTOEXCEL-BUILDDIR* or *WXAUTOEXCEL-INSTALLDIR*.
For example, when building your application from its build folder, `wxAutoExcel_ROOT` may be set like this

    cmake -G "Visual Studio 17 2022" -DwxAutoExcel_ROOT=c:/dev/libs/wxAutoExcel-vc17-x64-DLL-installed -S ../MyApp -B .

### 3.2 CMake with `add_subdirectory()`

The first step is calling `add_subdirectory()` and the second step is calling
`target_link_libraries()`, using `wxAutoExcel::wxAutoExcel` as the library target name.

For example, if you have wxAutoExcel source code in your application's folder *3rdparty/wxAutoExcel*
(e.g., *c:/dev/apps/MyApp/3rdparty/wxAutoExcel*), your application CMakeLists.txt may look like this

    # ..... make sure wxWidgets is available .....
    # maybe set some of wxAutoExcel_BUILD options here....
    add_subdirectory(3rdparty/wxAutoExcel)
    # .....
    target_link_libraries(${PROJECT_NAME} PRIVATE wxAutoExcel::wxAutoExcel ${wxWidgets_LIBRARIES})

### 3.3 CMake with `FetchContent`

The first step is calling `include(FetchContent)`, `FetchContent_Declare()`,
and `FetchContent_MakeAvailable()`; the second step is calling `target_link_libraries()`,
using `wxAutoExcel::wxAutoExcel` as the library target name.

For example, your application CMakeLists.txt may look like this

    # ..... make sure wxWidgets is available .....
    include(FetchContent)
    FetchContent_Declare(
      wxAutoExcel
      GIT_REPOSITORY https://github.com/PBfordev/wxAutoExcel
      GIT_TAG        4396ab0d5ec75e4ebaedf97f832213d28ba457aa # use the actual desired tag instead
    )
    # maybe set some of wxAutoExcel_BUILD options here....
    FetchContent_MakeAvailable(wxAutoExcel)
    # .....
    target_link_libraries(${PROJECT_NAME} PRIVATE wxAutoExcel::wxAutoExcel ${wxWidgets_LIBRARIES})


### 3.4 Manual setup (not recommended)

If you do not wish to build your application with CMake, then in addition to the standard wxWidgets-related
and other settings, you need to: 
- Add `WXAUTOEXCEL-SRCDIR/include` or `WXAUTOEXCEL-INSTALLDIR/include` to the compiler include path.
- Add `WXAUTOEXCEL-BUILDDIR/lib` or `WXAUTOEXCEL-INSTALLDIR/lib` to the linker libraries path.
- If you are using the DLL build of wxAutoExcel, add `WXUSINGDLL_WXAUTOEXCEL`
  to the preprocessor definitions.
- Finally, add the appropriate wxAutoExcel (import) library to the
  list of libraries to link with, for Debug and Release targets separately.
