md %build_dir% 
cd %build_dir%

echo wxWidgets_CONFIGURATION:STRING=mswu > CMakeCache.txt
echo wxWidgets_LIB_DIR:PATH=%wxWidgets_LIBRARIES% >> CMakeCache.txt
echo wxWidgets_ROOT_DIR:PATH=%wxWidgets_ROOT_DIR% >> CMakeCache.txt
echo WX_LIB_DIR:INTERNAL=%wxWidgets_LIBRARIES% >> CMakeCache.txt
echo WX_ROOT_DIR:INTERNAL=%wxWidgets_ROOT_DIR% >> CMakeCache.txt
echo wxAutoExcel_BUILD_SHARED:BOOL=OFF >> CMakeCache.txt 
echo wxAutoExcel_BUILD_LINK_WX_SHARED:BOOL=ON >> CMakeCache.txt    

goto %toolset%

:msbuild
rem "C:\Program Files\Microsoft SDKs\Windows\v7.1\Bin\SetEnv.cmd" /x64
CALL "C:\Program Files (x86)\Microsoft Visual Studio %VisualStudioVersion%\VC\vcvarsall.bat" %platform%
cmake -Wno-dev -G "Visual Studio 14 2015 Win64" %project_dir%
msbuild "ALL_BUILD.vcxproj" /consoleloggerparameters:Verbosity=minimal /target:Build  /p:Configuration=%configuration% /p:Platform=%platform% /logger:"C:\Program Files\AppVeyor\BuildAgent\Appveyor.MSBuildLogger.dll"
goto :eof

:nmake
echo CMAKE_BUILD_TYPE:STRING=%configuration% >> CMakeCache.txt
CALL "C:\Program Files (x86)\Microsoft Visual Studio %VisualStudioVersion%\VC\vcvarsall.bat" %platform%
cmake -Wno-dev -G "NMake Makefiles" %project_dir%
nmake -f makefile
goto :eof

:mingw
echo CMAKE_BUILD_TYPE:STRING=%configuration% >> CMakeCache.txt
echo wxAutoExcel_BUILD_GCC_STD_OPTION:STRING=gnu++ >> CMakeCache.txt
echo wxAutoExcel_BUILD_CXX_STANDARD:STRING=11 >> CMakeCache.txt
set path=C:\MinGW\bin;"C:\Program Files (x86)\CMake\bin"
cmake -Wno-dev -G "MinGW Makefiles" %project_dir%
mingw32-make -j2 -f makefile
goto :eof

:eof