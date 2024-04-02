/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Copyright:   (c) 2016 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
//////////////////////////////////////////////////////////////////////////// 

/**

@page page_tutorial Tutorial 

@tableofcontents

In this tutorial it is assumed that you have built wxAutoExcel library
in required configurations and added it to your project, 
as described in <a href='https://github.com/pbfordev/wxAutoExcel/blob/master/docs/install.md'>docs/install.md</a>.
It is also assumed you <tt>\#include <wx/wxAutoExcel.h></tt> and are <tt>\#using namespace %wxAutoExcel</tt>.

The tutorial is very brief and it is recommended to check out the bundled
samples, starting with the Minimal sample.

@section page_tutorial_introduction Introduction
<b>Classes and methods</b>

wxAutoExcel classes are named the same as MS Excel VBA classes, except their names 
start with <i>wxExcel</i>, e.g., <i>Range</i> is @c wxExcelRange. 
Method names are the same as those of underlying MS Excel class, e.g.,
<i>Range.Activate</i> is @c wxExcelRange::Activate().
Properties are are implemented as methods, prefixed with @c Get and/or @c Set, 
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

<b>Attaching to any running instance </b>

This works analogically to @c wxAutomationObject::GetInstance(), 
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
    wxExcelWorksheets wsheets = workbook.GetWorksheets();
    if ( !wsheets ) 
    {
        wxLogError(_("Failed to obtain Worksheets."));
        return;
    }
    
    wxExcelWorksheet wsheet;
    long count = wsheets.GetCount();
    for ( long i = 1; i <= count; i++ )
    {
        wsheet = wsheets[i];
    }    
@endcode

<b>Adding a worksheet</b>

Simply adding a worksheet, it will be placed after the last existing worksheet
@code
    // wsheets is a valid wxExcelWorksheets instance
    wxExcelWorksheet wsheet = wsheets.Add();
    // now add three more worksheets
    wsheets.Add(3); 
@endcode

<b>Adding a worksheet in the front of all others</b>
@code
    // wsheets is a valid wxExcelWorksheets instance
    wxExcelWorksheet wsheet = wsheets.AddAfterOrBefore(wsheets[1], false);
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

<b>Reading and writing to/fro a range</b>

See the bundled samples for more complex examples.
See the bulkdata sample for an example on how to efficiently
transfer a large number of values, using @c wxSafeArray.
@code
    // sheet is a valid wxExcelWorksheet instance
    wxExcelRange range = sheet.GetRange("A1");
    
    range.SetValue(12.3);
    value = range.GetValue(); // 12.3
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
#ifdef _DEBUG
    wxLog::SetLogLevel(wxLOG_Trace);
    wxLog::AddTraceMask(wxTRACE_AutoExcel);                                  
#else    
    wxAutoExcelObjectErrorModeOverrider emo(wxExcelObject::Err_DoNothing, false);
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