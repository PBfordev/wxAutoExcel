/*! @mainpage %wxAutoExcel
 *
 * @section intro_sec Introduction
 *
 * This is the introduction.
 *
 * @section install_sec Installation
  * @subsection step1 Download and decompress
 Just decompress the .zip file manually into any directory (you should avoid using directories with spaces in their names).

 * @subsection step2 Using Microsoft Visual C++ IDE
 
 (a) Building %wxAutoExcel

 wxAutoExcel comes with MSVC 2008 project files. Go to the folder "%wxAutoExcel\build" and you will find file %wxAutoExcel_vc9.sln there.
 Open it and build the solution, which has two targets: DEBUG and RELEASE. Verify the build succeeded and files %wxAutoExcel100u.lib
 and %wxAutoExcel100ud.lib were produced in the %wxAutoExcel\\lib\\vc_lib folder. In order to successfully compile %wxAutoExcel with provided project files,
 it is assumed that you have set a system environment variable WXWIN, pointing to the folder where you have installed wxWidxets, e.g. WXWIN=c:\\wxWidgets-2.9.5.

 (b) Adding %wxAutoExcel to your MSVC project

You need to do the following for all your configurations (e.g. Debug, Release):
@li Go to  your project's configuration properties.
@li In "C/C++ / General" add %wxAutoExcel's include directory into "Additional include directories" (e.g. c:\\%wxAutoExcel\\include").
@li In "Linker / General" add %wxAutoExcel's library directory into "Additional library directories" (e.g. c:\\wxAutoExcel\\lib\\vc_lib").
@li In "Linker / Input" add %wxAutoExcel's library into "Additional dependencies" (%wxAutoExcel100ud.lib for the Debug configuration and %wxAutoExcel100u.lib for the release one).
@li Optionally add "wx/%wxAutoExcel.h" to your precompiled header file to speed up compilation.

Include <wx/%wxAutoExcel.h> in the file referring to any %wxAutoExcel class, don't forget that all the classes are in wxAutoExcel namespace.

 *  
 * 
 */