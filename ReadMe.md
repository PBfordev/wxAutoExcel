# wxAutoExcel
[![GitHub MSVS 2022 wxWidgets 3.2](https://github.com/PBfordev/wxAutoExcel/actions/workflows/build-msvs2022.yml/badge.svg)](https://github.com/PBfordev/wxAutoExcel/actions/workflows/build-msvs2022.yml) 
[![GitHub MSVS 2022 wxWidgets master](https://github.com/PBfordev/wxAutoExcel/actions/workflows/build-msvs2022.yml-wxmaster/badge.svg)](https://github.com/PBfordev/wxAutoExcel/actions/workflows/build-msvs2022-wxmaster.yml) [![GitHub MSYS2](https://github.com/PBfordev/wxAutoExcel/actions/workflows/build-msys2.yml/badge.svg)](https://github.com/PBfordev/wxAutoExcel/actions/workflows/build-msys2.yml) 
[![GitHub Build docs](https://github.com/PBfordev/wxAutoExcel/actions/workflows/build-online-docs.yml/badge.svg)](https://github.com/PBfordev/wxAutoExcel/actions/workflows/build-online-docs.yml)

Introduction
---------
wxAutoExcel is a [wxWidgets](http://www.wxwidgets.org) (requires v3.1 or newer) 
C++ library attempting to make automating Microsoft Excel easier.

Platforms
---------
 
Microsoft Windows, requires Microsoft Excel to be installed.

Installing and Using wxAutoExcel
---------
See [docs/install.md](https://github.com/PBfordev/wxAutoExcel/blob/master/docs/install.md) for instructions how to set-up and build wxAutoExcel and how to
use it in your projects. I strongly suggest checking the tutorial and bundled samples 
to see the basics of wxAutoExcel in action.
Documentation is available at https://pbfordev.github.io/wxAutoExcel/

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

app.SetVisible(true);

wxExcelWorkbook workbook = app.GetWorkbooks().Add();
if ( !workbook )
{
    wxLogError(_("Failed to create a new workbook."));
    return false;
}

wxExcelRange range = workbook.GetWorksheets()[1].GetRange("A1");
range = "Hello, World!";
range.GetFont().SetColor(*wxBLUE);
```

Licence
---------
[wxWidgets licence](https://github.com/wxWidgets/wxWidgets/blob/master/docs/licence.txt) 