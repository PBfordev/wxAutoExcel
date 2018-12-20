# wxAutoExcel  [![Build status](https://ci.appveyor.com/api/projects/status/q9r7w07abnhwno78/branch/master?svg=true)](https://ci.appveyor.com/project/pbfordev/wxautoexcel/branch/master)

Introduction
---------
wxAutoExcel is a [wxWidgets](http://www.wxwidgets.org) (requires v3.1 or newer) 
C++ library attempting to make automating Microsoft Excel easier.

Platforms
---------
 
Microsoft Windows, requires Microsoft Excel to be installed.

Installing and Using wxAutoExcel
---------
See [docs/install.txt] (https://github.com/PBfordev/wxAutoExcel/blob/master/docs/install.txt) for instructions how to set-up and build wxAutoExcel and how to
use it in your projects. I strongly suggest checking the tutorial and bundled samples 
to see the basics of wxAutoExcel in action.
Documentation is available at http://pbfordev.github.io/wxAutoExcel/

Example of Code Using wxAutoExcel
---------
The simple code below: (1) starts a new Microsoft Excel instance, adds a new workbook,
(2) writes a string into the A1 cell of the first worksheet of the newly added workbook,
(3) sets the text color of the A1 cell to blue, and (4) displays the window of the new
Microsoft Excel instance.

```cpp
#include <wx/wxAutoExcel.h>

using namespace wxAutoExcel;

// the following code is assumed to be inside 
// a function returning a bool

wxExcelApplication app = wxExcelApplication::CreateInstance();
if ( !app )
{
    wxLogError(_("Could not launch Microsoft Excel. Please check that it is properly installed."));
    return false;
}

wxExcelWorkbook workbook = app.GetWorkbooks().Add();
if ( !workbook )
{
    wxLogError(_("Failed to create a new workbook."));
    return false;
}

wxExcelRange range = workbook.GetWorksheets()[1].GetRange("A1");
range = "Hello, World!";
range.GetFont().SetColor(*wxBLUE);

app.SetVisible(true);


Licence
---------
[wxWidgets licence](https://github.com/wxWidgets/wxWidgets/blob/master/docs/licence.txt) 