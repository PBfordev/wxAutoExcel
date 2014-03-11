Introduction
*************

wxAutoExcel is a wxWidgets (www.wxwidgets.org) library for automating Microsoft Excel.


Building wxAutoExcel
*********************

wxAutoExcel comes with Microsoft Visual C++ 2008 and Code::Blocks (using GCC) project files. Go to the folder wxAutoExcel\build and you will find wxAutoExcel_vc9.sln for MSVC and wxAutoExcel_gcc.cbp for Code::Blocks there.
Open the project file of your choice and build one or more of its configurations. There are four of them: Debug, Release, DLL Debug and DLL Release.
Verify the build succeeded and the libraries were produced in the wxAutoExcel\lib\ folder. In order to successfully compile wxAutoExcel with provided project files, it is expected that you have set a system environment variable WXWIN, pointing to the folder where you have installed wxWidxets, e.g. WXWIN=c:\wxWidgets.


Adding wxAutoExcel to your wxWidgets project
*********************************************

This applies to MSVC, but the procedure should be similar with any other IDE. You need to do the following for all your configurations (e.g. Debug, Release...):
1. Go to the project configuration properties.
2. In "C/C++ / General" add wxAutoExcel's include directory into "Additional include directories" (e.g. "c:\wxAutoExcel\include").
3. In "Linker / General" add wxAutoExcel's library directory into "Additional library directories" (e.g. "c:\wxAutoExcel\lib\vc_lib").
4. In "Linker / Input" add wxAutoExcel's library into "Additional dependencies" (wxAutoExcel100ud.lib for the Debug configuration and wxAutoExcel100u.lib for the Release one).
5. Optionally add <wx/wxAutoExcel.h> to your precompiled header file to speed up compilation.

Include <wx/wxAutoExcel.h> in files referring to any wxAutoExcel class and don't forget that all those classes reside in wxAutoExcel namespace.