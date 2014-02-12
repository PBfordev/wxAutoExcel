/*! @mainpage %wxAutoExcel
 *
 * @section intro_sec Introduction
 *
 * This is the introduction.
 *
 * @section install_sec Installation
  * @subsection step1 Download and decompress
 Decompress the .zip file into any directory while preserving the folder hierarchy. You should avoid using directories with spaces in their names.

 * @subsection step2 Building wxAutoExcel
  %wxAutoExcel comes with MSVC 2008 and Code::Blocks (using GCC) project files. Go to the folder "wxAutoExcel\build" and you will find wxAutoExcel_vc9.sln for 
 MSVC and wxAutoExcel_gcc.cbp for Code::Blocks there.
 Open the project file of your choice and build one of more of its configurations. There are four of them: Debug, Release, DLL Debug and DLL Release. 
 Verify the build succeeded and the libraries were produced in the %wxAutoExcel\\lib\\ folder. In order to successfully compile %wxAutoExcel with provided project files,
 it is assumed that you have set a system environment variable WXWIN, pointing to the folder where you have installed wxWidxets, e.g. WXWIN=c:\\wxWidgets.

 * @subsection step3 Adding wxAutoExcel to your project

This applies to Microsoft Visual C++, but the procedure should be similar in any other  IDE. You need to do the following for all your configurations (e.g. Debug, Release...):
@li Go to  your project's configuration properties.
@li In "C/C++ / General" add %wxAutoExcel's include directory into "Additional include directories" (e.g. "c:\wxAutoExcel\include").
@li In "Linker / General" add %wxAutoExcel's library directory into "Additional library directories" (e.g. "c:\wxAutoExcel\\lib\vc_lib").
@li In "Linker / Input" add %wxAutoExcel's library into "Additional dependencies" (%wxAutoExcel100ud.lib for the Debug configuration and %wxAutoExcel100u.lib for the Release one).
@li Optionally add "wx/wxAutoExcel.h" to your precompiled header file to speed up compilation.

Include <wx/%wxAutoExcel.h> in files referring to any %wxAutoExcel class, don't forget that all these classes are in wxAutoExcel namespace.

 *  
 * 
 */