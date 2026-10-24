/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Copyright:   (c) 2016 PB <pbfordev@gmail.com>
// License:     MIT license
//////////////////////////////////////////////////////////////////////////// 

/**

@page page_tutorial Tutorial 

@tableofcontents

In this tutorial it is assumed that you have built wxAutoExcel library
in required configurations and added it to your project, 
as described in <a href='https://github.com/pbfordev/wxAutoExcel/blob/master/docs/install.md'>docs/install.md</a>.
It is also assumed you <tt>\#include <wx/wxAutoExcel.h></tt> and are <tt>using namespace %wxAutoExcel</tt>.

The tutorial is very brief and it is recommended to check out the bundled
samples, starting with the Minimal sample.

@section page_tutorial_introduction Introduction
<b>Classes and methods</b>

wxAutoExcel classes are named the same as MS Excel VBA classes, except their names 
start with <i>wxExcel</i>, e.g., <i>Range</i> is @c wxExcelRange. 
Method names are the same as those of underlying MS Excel class, e.g.,
<i>Range.Activate</i> is @c wxExcelRange::Activate().
Properties are implemented as methods, prefixed with @c Get and/or @c Set, 
e.g. <i>Range.Value</i> is @c wxExcelRange::GetValue() and @c wxExcelRange::SetValue(). 
All wxAutoExcel classes are derived from @c wxExcelObject. This class has several
utility methods, their names end with an underscore so they can be easily distinguished
from the methods that wrap the underlying MS Excel class methods and properties. 
Notable exceptions to this rule are @c wxExcelApplication::CreateInstance() and
@c wxExcelApplication::GetInstance().

@section page_tutorial_application Obtaining and closing wxExcelApplication instance

<b>Starting a new instance</b>
@code
    wxExcelApplication app = wxExcelApplication::CreateInstance();
    if ( !app ) 
    {
        wxLogError(_("Failed to create an instance of MS Excel application."));        
    }
@endcode

<b>Attaching to a running instance (if any exists) </b>

This works similarly to @c wxAutomationObject::GetInstance(), 
i.e., you can pass flags which affect its behaviour, such as 
the (default) @c wxAutomationInstance_CreateIfNeeded.
If you attempt attaching to a running instance this way, and there
is more than one instance running, you cannot affect which one you get. 
@code
    wxExcelApplication app = wxExcelApplication::GetInstance();
    if ( !app ) 
    {
        wxLogError(_("Failed to obtain an instance of MS Excel application."));
    }
@endcode

<b>Attaching to an instance with a specified workbook open</b>
@code
    wxExcelApplication app = wxExcelApplication::GetInstance("c:\\budget.xlsx");
    if ( !app ) 
    {
        wxLogError(_("Failed to obtain an instance of MS Excel application."));
    }
@endcode

<b>Showing Excel window</b>

MS Excel launched in automation mode has its window hidden, so you need
to tell it to show itself if needed.
@code
    // app is a valid instance of wxExcelApplication
    app.SetVisible(true);
@endcode

<b>Closing the application</b>

Once you are done with the application, do not want to use any
of its objects and wish to close it.
@code
    // app is a valid instance of wxExcelApplication
    app.Quit();
@endcode


@section page_tutorial_localised Working with localised MS Excel
If you want to use wxAutoExcel with Excel localised into languages other 
than English, I recommend setting its automation LCID to US English, e.g. 
@code
    app.SetAutomationLCID_(wxExcelObject::lcidEnglishUS);
@endcode
and you should be able to use English names of formulas and styles etc. 
regardless of the Excel user language. The LCID will be inherited by all 
"children" of the object. This unfortunately does not work in all cases 
and can have some side effects, see the matching entry in <a href='https://github.com/pbfordev/wxAutoExcel/blob/master/docs/FAQ.md'>docs/FAQ.md</a>.

@section page_tutorial_workbook Working with workbooks
<b>Adding a new workbook</b>
@code
    // app is a valid instance of wxExcelApplication
    wxExcelWorkbook workbook = app.GetWorkbooks().Add();
    if ( !workbook ) 
    {
        wxLogError(_("Failed to create a new workbook."));
    }
@endcode

<b>Opening an existing workbook</b>
@code
    // app is a valid instance of wxExcelApplication
    // fileName is the full path of an existing file
    wxExcelWorkbook workbook = app.GetWorkbooks().Open(fileName);
    if ( !workbook )
    {
        wxLogError(_("Failed to open a workbook %s."), fileName);
    }
@endcode

<b>Obtaining ActiveWorkbook</b>
@code
    // app is a valid instance of wxExcelApplication

    wxExcelWorkbook workbook;
    // There may not be an active workbook in the application
    // so make sure to suppress the possible error message
    {
        wxAutoExcelObjectErrorModeOverrider emo(wxExcelObject::Err_DoNothing, true);
        workbook = app.GetActiveWorkbook(); 
    }
    if ( !workbook ) 
    {
        wxLogError(_("Failed to obtain ActiveWorkbook."));
    }
@endcode


@section page_tutorial_worksheet Working with worksheets
<b>Obtaining Worksheets collection and enumerating worksheets</b>
@code
    // workbook is a valid wxExcelWorkbook instance
    wxExcelWorksheets worksheets = workbook.GetWorksheets();
    if ( !worksheets ) 
    {
        wxLogError(_("Failed to obtain Worksheets."));
        return;
    }
    
    wxExcelWorksheet worksheet;
    long count = worksheets.GetCount();
    for ( long i = 1; i <= count; i++ )
    {
        worksheet = worksheets[i];
    }
@endcode

With C++11 or newer, the items of collections can also be iterated over
with a range-based for loop, see wxAutoExcel::wxExcelCollectionIterator.
The loop variable should be a value, not a reference.
@code
    // workbook is a valid wxExcelWorkbook instance
    for ( wxExcelWorksheet worksheet : workbook.GetWorksheets() )
    {
        wxLogMessage("Worksheet '%s'", worksheet.GetName());
    }
@endcode

<b>Adding a worksheet</b>

Simply adding a worksheet (it will be placed before the active worksheet)
@code
    // worksheets is a valid wxExcelWorksheets instance
    wxExcelWorksheet worksheet = worksheets.Add();
    // now add three more worksheets
    worksheets.Add(3); 
@endcode

<b>Adding a worksheet in the front of all others</b>
@code
    // worksheets is a valid wxExcelWorksheets instance
    wxExcelWorksheet worksheet = worksheets.AddAfterOrBefore(worksheets[1], false);
@endcode

@section page_tutorial_range Working with ranges
<b>Obtaining a range</b>

See the bundled samples for more complex examples of obtaining ranges.
@code
    // sheet is a valid wxExcelWorksheet instance
    wxExcelRange range = sheet.GetRange("B2:C10");
    
    // range2 will be a cell with an absolute address of "B2"
    wxExcelRange range2 = range.GetRange("A1");
    
    wxExcelRange usedRange = sheet.GetUsedRange(); 
@endcode

<b>Reading and writing range values</b>

See the bundled samples for more complex examples.
@code
    // worksheet is a valid wxExcelWorksheet instance
    wxExcelRange range = worksheet.GetRange("A1");
    wxVariant value;
    
    range.SetValue(12.3);
    value = range.GetValue(); // 12.3
@endcode

<b>Reading and writing the values of many cells at once</b>

Transferring all values at once with @c wxExcelRangeData is much faster
than cell by cell. The data must have the same number of rows and columns
as the range.
@code
    // worksheet is a valid wxExcelWorksheet instance
    wxExcelRange range = worksheet.GetRange("A1:C100");
    wxExcelRangeData data(100, 3); // rows and columns are numbered from 0

    for ( size_t row = 0; row < data.GetRowCount(); ++row )
    {
        data(row, 0) = wxString::Format("Item %d", static_cast<int>(row) + 1);
        data(row, 1) = 2.5 * row;
        data(row, 2) = row % 2 == 0;
    }
    range.SetValues_(data);

    wxExcelRangeData values;
    if ( range.GetValues_(values) )
    {
        // values(99, 1) is 247.5
    }
@endcode

<b>Reading the values of cells of the expected types</b>

wxExcelRangeData::GetCell() obtains the value only when the cell contains
a value of the requested type, so the type of the variant need not be checked.
@code
    // values were obtained with wxExcelRange::GetValues_()
    for ( size_t row = 0; row < values.GetRowCount(); ++row )
    {
        wxString name;
        double price;

        if ( values.GetCell(row, 0, name) && values.GetCell(row, 1, price) )
            wxLogMessage("%s costs %.2f", name, price);
        else if ( values.IsCellError(row, 1) )
            wxLogMessage("The price in row %zu is an error", row + 1);
    }
@endcode

<b>Iterating over the cells, rows, or columns of a range</b>

With C++11 or newer, a range can be iterated over with a range-based for loop.
It is much slower than transferring the values at once with
wxExcelRangeData, as each cell takes several calls to Excel.
@code
    // range is a valid wxExcelRange instance with a single area
    for ( wxExcelRange cell : range )
    {
        // the cells, across the rows
    }

    for ( wxExcelRange row : range.GetRows() )
    {
        row.SetRowHeight(20);
        for ( wxExcelRange cell : row.GetCells() )
        {
            // the cells of the row
        }
    }
@endcode

<b>Formatting a range</b>
@code
    // range is a valid wxExcelRange instance
    wxExcelFont font = range.GetFont();
    font.SetBold(true);
    font.SetColor(*wxBLUE);
    font.SetSize(font.GetSize() * 1.5); // 150% of default size
    
    range.SetHorizontalAlignment(xlCenter);
    range.GetInterior().SetColor(*wxLIGHT_GREY);
    range.BorderAround(WXAEEP(xlDouble), WXAEEP(xlThick), NULL, wxBLUE);
    range.GetBorders()[xlInsideHorizontal].SetWeight(xlThin);
@endcode

@section page_tutorial_tasks Common tasks
<b>Opening a workbook read-only</b>
@code
    // app is a valid instance of wxExcelApplication
    // fileName is the full path of an existing file
    wxExcelWorkbook workbook = app.GetWorkbooks().Open(fileName, NULL, true);
@endcode

<b>Finding the last used row in a column</b>

The same as <tt>Cells(Rows.Count, 1).End(xlUp).Row</tt> in VBA. When the
column is empty, the result is 1.
@code
    // worksheet is a valid wxExcelWorksheet instance
    wxExcelRange column = worksheet.GetColumns(1);
    wxExcelRange lastCell = column.GetItem(column.GetRows().GetCount()).GetEnd(xlUp);
    long lastRow = lastCell.GetRow();
@endcode

<b>Copying a range to another workbook</b>
@code
    // source is a valid wxExcelRange instance, target a valid wxExcelWorkbook instance
    wxExcelRange destination = target.GetWorksheets()[1].GetRange("C3");
    source.Copy(&destination);
@endcode

<b>Copying a worksheet to another workbook</b>
@code
    // worksheet is a valid wxExcelWorksheet instance, target a valid wxExcelWorkbook instance
    // place the copy before the first worksheet of target
    worksheet.CopyAfterOrBefore(target.GetWorksheets()[1], false);
@endcode

<b>Saving a workbook in the xlsx format and closing it</b>
@code
    // workbook is a valid wxExcelWorkbook instance
    // fileName is the full path of the file to save the workbook as
    XlFileFormat format = xlOpenXMLWorkbook;
    if ( !workbook.SaveAs(fileName, &format) )
    {
        wxLogError(_("Failed to save the workbook as %s."), fileName);
    }
    workbook.Close(false);
@endcode

<b>Creating a table</b>
@code
    // worksheet is a valid wxExcelWorksheet instance, A1:C10 contain
    // the data with the column names in the first row
    wxExcelRange source = worksheet.GetRange("A1:C10");
    XlYesNoGuess hasHeaders = xlYes;
    wxExcelListObject table = worksheet.GetListObjects().Add(&source, &hasHeaders);
    table.SetName("Sales");
@endcode

<b>Sorting a range</b>
@code
    // worksheet is a valid wxExcelWorksheet instance, A1:C10 contain
    // the data with the column names in the first row
    wxExcelSort sort = worksheet.GetSort();
    XlSortOrder descending = xlDescending;

    sort.GetSortFields().Clear();
    sort.GetSortFields().Add(worksheet.GetRange("C2:C10"), xlSortOnValues, &descending);
    sort.SetRange(worksheet.GetRange("A1:C10"));
    sort.SetHeader(xlYes);
    sort.Apply();
@endcode

<b>Highlighting cells with conditional formatting</b>

Excel interprets the formulas of conditional formats with the separators
and function names of the system settings, regardless of the automation LCID.
@code
    // range is a valid wxExcelRange instance
    XlFormatConditionOperator greater = xlGreater;
    wxExcelFormatCondition condition = range.GetFormatConditions().Add(xlCellValue, &greater, "=100");
    condition.GetInterior().SetColor(*wxRED);

    // the colour of a cell including the conditional formatting
    wxColour colour = range.GetItem(1).GetDisplayFormat().GetInterior().GetColor();
@endcode

<b>Finding all cells with a value</b>

wxExcelRange::FindAll_() returns the cells in the order of the search,
starting with the upper-left cell of the range. Excel remembers the search
settings, so it is better to set them all.
@code
    // worksheet is a valid wxExcelWorksheet instance
    wxExcelRange range = worksheet.GetUsedRange();
    wxExcelRangeVector found;
    XlFindLookIn lookIn = xlValues;
    XlLookAt lookAt = xlWhole;
    XlSearchOrder byRows = xlByRows;

    if ( range.FindAll_("Total", found, &lookIn, &lookAt, &byRows, false) )
    {
        for ( size_t i = 0; i < found.size(); ++i )
            found[i].GetFont().SetBold(true);
    }
@endcode

<b>Replacing text in a range</b>

When the range is a single cell, Excel replaces in the whole worksheet.
@code
    // range is a valid wxExcelRange instance
    XlLookAt lookAt = xlPart;
    range.Replace("colour", "color", &lookAt, NULL, false);
@endcode

@section page_tutorial_chart Working with charts

Check the bundled charts sample for more examples.

<b>Adding an embedded chart</b>
@code
    wxExcelChart chart;
    
    chart = worksheet.GetShapes().AddChart(xlXYScatterLinesNoMarkers, 1, 1, 250, 250).GetChart();
    chart.SetHasTitle(true);
    chart.GetChartTitle().SetText(_("My chart"));
    chart.SetPlotBy(xlColumns);
    chart.SetSourceData(sourceRange);
@endcode

<b>Adding a chart sheet </b>
@code
    wxExcelChart chart = workbook.GetCharts().Add();
    chart.SetChartType(xl3DColumnClustered);
    chart.SetName(_("My 3D clustered chart"));
@endcode

<b>Customising chart elements</b>
@code
    // chart is of xlXYScatterLinesNoMarkers type
     
    // customise the y axis
    wxExcelAxis axis = chart.Axes(xlValue);
    
    axis.SetHasTitle(true);
    axis.GetAxisTitle().SetCaption(_("Some values"));
    axis.SetMajorTickMark(xlTickMarkCross);
    axis.GetTickLabels().GetFont().SetSize(8);
    axis.SetMinimumScale(0);
    axis.SetMaximumScale(200);
    
    // customise series
    wxExcelSeries series = chart.SeriesCollection()[1];
    
    series.SetName(_("My series 1"));
    series.GetFormat().GetLine().SetWeight(1.5);
    series.GetFormat().GetLine().GetForeColor().SetRGB(*wxBLUE);
    
    // customise individual point appearance
    wxExcelPoints points = series.Points();
    wxExcelPoint point = points[1]; // first point

    point.SetMarkerForegroundColor(*wxRED);
    point.SetMarkerSize(7);
    point.SetMarkerStyle(xlMarkerStylePlus);
@endcode


@section page_tutorial_errors Handling errors
<b>wxWidgets error handling in @c wxAutomationObject</b>

When an OLE automation call performed by @c wxAutomationObject fails, wxWidgets
tells the user with @c wxLogError(), see @c ShowException() in <tt>WXWIN/src/msw/ole/automatn.cpp</tt>).
This behaviour may not be most desirable, as the user probably does not understand the
error message and the error information is not propagated to the calling code created 
by the application programmer.
Unfortunately, the only way to prevent displaying the error to the user is to suppress
logging with @c wxLogNull. Creating a @c wxLogNull instance suppresses all logging for the
calling thread, so it has to be used with caution. 

<b>%wxAutoExcel error handling</b>

By default, wxAutoExcel in the release mode (i.e., when @c NDEBUG is defined) prevents errors 
described above to be shown to user by wxWidgets, the setting is controlled by @c WXAUTOEXCEL_SHOW_WXAUTOMATION_ERROR 
defined in @c wxAutoExcel_setup.h. How wxAutoExcel itself behaves when an error is encountered
is controlled by a wxAutoExcel-wide setting, see @c wxExcelObject::SetErrorMode_() and
@c wxExcelObject::GetErrorMode_(). It is highly recommended to use @c wxAutoExcelObjectErrorModeOverrider
class to control how the errors are dealt with instead of calling the two methods.
For example, one can have this code before and at the same scope with the part with wxAutoExcel calls,
as the programmer may be interested in seeing the error message while the user is probably not.

@code
#ifdef NDEBUG
    wxAutoExcelObjectErrorModeOverrider emo(wxExcelObject::Err_DoNothing, false);
#else    
    wxLog::SetLogLevel(wxLOG_Trace);
    wxLog::AddTraceMask(wxTRACE_AutoExcel);
#endif 
@endcode

When calling a method or property of a wxAutoExcel class wrapping a MS Excel object,
the application programmer can learn whether the call succeeded by calling its <tt>operator bool()</tt>.
This operator returns true if the object has a valid automation interface and the last call
(method or property) on it succeeded, false otherwise.

@code        
    wxExcelWorkbook workbook = app.GetActiveWorkbook();
    if ( !workbook )
    {
        wxLogError(_("Could not obtain ActiveWorkbook."));
        return false;
    }

    workbook.ExportAsFixedFormat(xlTypePDF, "This is Invalid File Name *:?");
    if ( !workbook )
    {
        wxLogError(_("Could not export ActiveWorkbook to PDF."));
        return false;
    }
@endcode

*/