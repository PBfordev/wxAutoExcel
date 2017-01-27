/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Copyright:   (c) 2016 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
//////////////////////////////////////////////////////////////////////////// 

/**

@page page_tutorial Tutorial 

@tableofcontents

In this tutorial it is  assumed that you have built wxAutoExcel library
in required configurations, added it to your project, and set up all
the paths as described in <a href='https://github.com/pbfordev/wxAutoExcel/blob/master/docs/install.txt'>docs/install.txt</a>.
It is also assumed you \#included <wx/wxAutoExcel.h> and are \#using %wxAutoExcel namespace.

The tutorial is very brief and it is recommended to check out the bundled
samples, starting with the Minimal sample.

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

This works analogically to wxAutomationObject::GetInstance(), 
i.e., you can pass flags which affect its behaviour, such as 
the (default) wxAutomationInstance_CreateIfNeeded.
If you attempt attach to a running instance this way, and there
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


@section page_tutorial_localised Working with localised Microsoft Excel
If you want to use wxAutoExcel with Excel localised into languages other 
than English, I recommend setting its automation LCID to US English, e.g. 
@code 
 	   app.SetAutomationLCID_(wxExcelObject::lcidEnglishUS);
@endcode
and you should be able to use English names of formulas and styles etc. 
regardless of the Excel user language. The LCID will be inherited by all 
"children" of the object. This unfortunately does not work in all cases 
and can have some side effects, see the matching entry in docs\FAQ.txt.

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
        wxAutoExcelObjectErrorModeOverrider emo(0, true);
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

@section page_tutorial_range Working with Ranges
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
See the bulkadata sample for an example on how to efficiently
transfer a large number of values, using wxSafeArray.
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
    chart.GetChartTitle().SetText("My chart");        
    chart.SetPlotBy(xlColumns);
    chart.SetSourceData(sourceRange);                      
@endcode

<b>Adding a chart sheet </b>
@code
    wxExcelChart chart = workbook.GetCharts().Add();    
    chart.SetChartType(xl3DColumnClustered);
    chart.SetName("My 3D clustered chart");
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
    wxExcelPoint point = series[1]; // first point

    point.SetMarkerForegroundColor(*wxRED);
    point.SetMarkerSize(7);
    point.SetMarkerStyle(xlMarkerStylePlus);                  
@endcode

 */