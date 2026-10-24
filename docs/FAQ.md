# wxAutoExcel FAQ

## Table of Contents

- [How do I start with wxAutoExcel?](#how-do-i-start-with-wxautoexcel)
- [Which licence is wxAutoExcel released under?](#which-licence-is-wxautoexcel-released-under)
- [Can I use wxAutoExcel without wxWidgets?](#can-i-use-wxautoexcel-without-wxwidgets)
- [Can I use wxAutoExcel to manipulate Microsoft Excel files on a computer without Microsoft Excel installed?](#can-i-use-wxautoexcel-to-manipulate-microsoft-excel-files-on-a-computer-without-microsoft-excel-installed)
- [wxWidgets is a multiplatform toolkit - can I use wxAutoExcel to automate Microsoft Excel on OSX?](#wxwidgets-is-a-multiplatform-toolkit---can-i-use-wxautoexcel-to-automate-microsoft-excel-on-osx)
- [Which versions of Microsoft Excel are supported?](#which-versions-of-microsoft-excel-are-supported)
- [Does wxAutoExcel support all the features Microsoft Excel exposes?](#does-wxautoexcel-support-all-the-features-microsoft-excel-exposes)
- [I have multiple versions of Microsoft Excel installed - can I choose which one to use for automation?](#i-have-multiple-versions-of-microsoft-excel-installed---can-i-choose-which-one-to-use-for-automation)
- [Is it possible to read from / write to a two-dimensional Range?](#is-it-possible-to-read-from--write-to-a-two-dimensional-range)
- [How can I use wxAutoExcel with Microsoft Excel localized into languages other than English?](#how-can-i-use-wxautoexcel-with-microsoft-excel-localized-into-languages-other-than-english)
- [How can I make my wxAutoExcel code faster?](#how-can-i-make-my-wxautoexcel-code-faster)
- [Can I use wxAutoExcel from a worker thread?](#can-i-use-wxautoexcel-from-a-worker-thread)
- [Why does an Excel process keep running after my application quits?](#why-does-an-excel-process-keep-running-after-my-application-quits)
- [How can I work with Excel objects not implemented in wxAutoExcel?](#how-can-i-work-with-excel-objects-not-implemented-in-wxautoexcel)
- [It does not work?!](#it-does-not-work)


## How do I start with wxAutoExcel?

1. Follow the install instructions in docs\install.md.
2. Read the [tutorial](https://pbfordev.github.io/wxAutoExcel/page_tutorial.html) and then take a look at the bundled  samples to familiarize yourself with wxAutoExcel, start with *Minimal* sample.

It is assumed you are familiar with VBA and Microsoft Excel object model,
as wxAutoExcel is just its C++ wrapper. If you are not very familiar with 
those, I believe that Microsoft Excel's Macro Recorder can be a very useful tool. 
When I want to do something in the code and I am not sure how, I usually 
just record a macro and then "translate" it from VBA to wxAutoExcel.


## Which licence is wxAutoExcel released under?

wxAutoExcel uses the MIT license.


## Can I use wxAutoExcel without wxWidgets?

No, wxAutoExcel requires the rudimentary OLE Automation support 
as well as some generic classes provided by wxWidgets. However,
it may be possible to use wxAutoExcel in an application which
does not use wxWidgets as its GUI toolkit, take a look at 
*PureWin32* sample for an inspiration.

## Can I use wxAutoExcel to manipulate Microsoft Excel files on a computer without Microsoft Excel installed?

No, wxAutoExcel uses OLE Automation which requires Microsoft Excel
to be installed and working properly.


## wxWidgets is a multiplatform toolkit - can I use wxAutoExcel to automate Microsoft Excel on OSX?

No, OLE Automation is available only on Microsoft Windows.


## Which versions of Microsoft Excel are supported?

wxAutoExcel was initially developed along Microsoft Excel 2007. Generally speaking,
wxAutoExcel can be used at least with Excel 2003-2021 / Office 365. Obviously, one 
needs to be careful to use only those objects, methods, and properties that are available for the installed Excel version. `wxExcelApplication` has a couple of version methods, such as `GetVersionAsDouble()`  or `IsVersionAtLeast_()`, it is also possible to obtain the list of properties and methods for a given object, see 
`wxExcelObject::GetPropertyAndMethodNames_()`.


## Does wxAutoExcel support all the features Microsoft Excel exposes?

No, it does not come even close, see the class list in the HTML 
documentation for what is available. Even when a class is listed there,
some of its methods or properties may not be implemented.
Events are not supported at all.

## I have multiple versions of Microsoft Excel installed - can I choose which one to use for automation?

Unfortunately, this is not possible (without modifying Registry).


## Is it possible to read from / write to a two-dimensional Range?

Yes. The simplest way is `wxExcelRangeData`, a block of cells indexed by
row and column, which `wxExcelRange::GetValues_()` and
`wxExcelRange::SetValues_()` transfer at once, as fast as possible:

        // range is a valid wxExcelRange of 100 rows and 3 columns
        wxExcelRangeData data(100, 3);

        for ( size_t row = 0; row < data.GetRowCount(); ++row )
            for ( size_t column = 0; column < data.GetColumnCount(); ++column )
                data(row, column) = static_cast<double>(row * column);

        range.SetValues_(data);

`wxExcelRangeData::GetCell()` reads a cell only when its value is of the
requested type, e.g. a number as a `double` or `long`, a string as
a `wxString`, or a date as a `wxDateTime`, and `wxExcelRangeData::IsCellError()`
tells whether the cell contains an error value, such as #N/A.

The data must have the same number of rows and columns as the range,
otherwise nothing is set and an error is reported. The formulas of all cells
are transferred in the same way with `wxExcelRange::GetFormulas_()` and
`wxExcelRange::SetFormulas_()`.

Without it, `wxExcelRange::GetValue()` returns the values of all cells of
a range in a list ordered by columns, `wxExcelRange::SetValue()` accepts
a list for a range with a single row, and both can use a two-dimensional
SAFEARRAY (`wxVariantDataSafeArray`), please see *BulkData* sample.


## How can I use wxAutoExcel with Microsoft Excel localized into languages other than English?

I suggest calling `SetAutomationLCID_(wxExcelObject::lcidEnglishUS)` on wxAutoExcel objects. The LCID is propagated towards the child wxAutoExcel objects, e.g., when you call it on `wxExcelApplication`, all of its properties like workbooks
will have the LCID set, all worksheets of the workbooks etc.
This way you can use English names of formulas, styles etc. regardless of
the language Microsoft Excel is localized into. This unfortunately does not
work in several cases. Excel interprets the formulas passed to
`wxExcelFormatConditions::Add()`, `wxExcelValidation::Add()`, and
`wxExcelNames::Add()` with the function names and the separators of the system
settings, e.g. a decimal comma, and the items of a validation list are
separated with the list separator of the system, see
`wxExcelApplication::GetInternational()` with `xlListSeparator`.
A name can be added with a simple formula and then given the English one
with `wxExcelName::SetRefersTo()`, which uses the automation locale.
Microsoft Excel also seems getting worse at it, e.g., since version 2013 you cannot
use English `NumberFormat` string for `TickLabels` in charts. 
Be aware that the automation locale affects how the values (e.g., those
including decimal or thousand separators, formatted as dates etc.) are 
interpreted in Microsoft Excel. E.g., string "1,234" can be considered an integer 
if the locale has "," as the thousand separator but also as "1.234" if "," is the
decimal separator.
On the other hand, the automation locale does not affect the list separator
in the addresses of ranges with several areas, which is always the one of
Windows regional settings, see `wxExcelApplication::GetListSeparator_()`.
To use the names of functions and the separators of the language
Microsoft Excel is localized into, use the "Local" properties, such as
`wxExcelRange::SetFormulaLocal()`.


## How can I make my wxAutoExcel code faster?

Every call of a wxAutoExcel method is a call to the Excel process, which takes
much longer than a call within your application. Most time is therefore saved
by making fewer calls:

- Read and write the values of many cells at once, using `wxExcelRangeData`
  with `wxExcelRange::GetValues_()` and `wxExcelRange::SetValues_()`, instead
  of cell by cell. They use `Value2` by default, which is also a bit faster
  than `Value`, as Excel does not convert the dates and currency values.
- Keep the objects you use repeatedly, e.g. a worksheet or a range, instead of
  obtaining them again for each call.
- Iterating over the cells of a range, e.g. with a range-based for loop,
  takes several calls for each cell. Use it for formatting individual cells,
  not for reading or writing the values.
- To find the cells with a value, use `wxExcelRange::Find()` and
  `wxExcelRange::FindNext()`, or `wxExcelRange::FindAll_()`, instead of
  reading the values of all cells.
- Format a whole range at once instead of its cells one by one.

Excel also does less work when it does not redraw its window or recalculate
the formulas after each change. Remember to restore the settings when done,
also when an error occurs:

        // app is a valid instance of wxExcelApplication
        const bool screenUpdating = app.GetScreenUpdating();
        const XlCalculation calculation = app.GetCalculation();

        app.SetScreenUpdating(false);
        app.SetCalculation(xlCalculationManual);

        // ... many changes ...

        app.SetCalculation(calculation);
        app.SetScreenUpdating(screenUpdating);

Setting the calculation requires an open workbook. Showing alerts can be turned
off with `wxExcelApplication::SetDisplayAlerts()`, which also prevents Excel
from waiting for a user who cannot see its hidden window.


## Can I use wxAutoExcel from a worker thread?

It is best to use wxAutoExcel only from the main thread of your application.
wxWidgets initializes OLE for the main thread, wxAutoExcel objects wrap COM
interfaces which belong to the thread which obtained them and must not be
used from another thread. A worker thread would need to initialize OLE itself
(e.g. with a `wxOleInitializer` from *wx/msw/ole/oleutils.h*), obtain its own
wxExcelApplication and all the other objects, and destroy them before OLE is
uninitialized. The error mode set with `wxExcelObject::SetErrorMode_()` is
shared by all threads.

This would not make the automation faster either, as Excel processes the
calls one by one in its main thread, and it rejects them when it is busy,
e.g. when the user is editing a cell in a visible Excel.


## Why does an Excel process keep running after my application quits?

An Excel instance started with `wxExcelApplication::CreateInstance()` runs
hidden until `wxExcelApplication::Quit()` is called, even after your
application ends. Excel also keeps running as long as any wxAutoExcel object
(e.g. a workbook or range) of the instance exists, so make sure all of them
are destroyed after calling `Quit()`, e.g. by keeping them in a scope which
ends before that of the application object. If your application crashes or
ends without calling `Quit()`, the hidden EXCEL.EXE process must be ended in
Task Manager.

Before quitting, close the workbooks or mark them as saved with
`wxExcelWorkbook::SetSaved(true)`, otherwise Excel may ask whether to save
them and wait for an answer in its hidden window. When you attach to an
instance the user works with, e.g. with `wxExcelApplication::GetInstance()`,
do not call `Quit()`, as it would close the user's workbooks too.


## How can I work with Excel objects not implemented in wxAutoExcel?

Unless you are willing to add their support to wxAutoExcel or can persuade me
to do that myself, you will have to resort to pure `wxAutomationObject` calls. 
The following example demonstrates a hypothetical situation where 
Worksheet and Worksheets objects were not implemented by wxAutoExcel 
and you wanted to output the name of each worksheet in a workbook:

       // workbook must be a valid wxExcelWorkbook object
       wxAutomationObject worksheets;
    
        if ( !workbook.GetUnimplementedObject_("Worksheets", worksheets) )
            return;
        
        const long count = worksheets.GetProperty("Count");
        for ( long i = 1; i <= count; i++ )
        {
            wxAutomationObject worksheet;
    
            if ( wxExcelObject::GetUnimplementedCollectionItem_(worksheets, i, worksheet, true) )
            {            
                wxLogMessage(_("Name of the worksheet with index %ld is \"%s\"."), 
                    i, (wxString)worksheet.GetProperty("Name"));        
            }    
        }   


## It does not work?!

Please be always as specific as possible when reporting an issue.
If it is build-related, make sure you followed the install instructions.
Provide all information required to successfully resolve the issue, 
including the error messages.
If your wxAutoExcel code does not work as expected, please:
1. Make the shortest possible self-contained compilable example demonstrating
   the issue, preferably based on the most appropriate bundled sample.
2. Create and test the VBA equivalent of the code from (1) to make sure the issue
   is within wxAutoExcel and not your code or Microsoft Excel itself.
3. Post the issue in wxCode section of wxWidgets forums (https://forums.wxwidgets.org/viewforum.php?f=30) or report it on GitHub (https://github.com/pbfordev/wxAutoExcel/issues).
