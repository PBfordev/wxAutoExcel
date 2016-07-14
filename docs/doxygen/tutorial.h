/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Copyright:   (c) 2016 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
//////////////////////////////////////////////////////////////////////////// 

/**

@page page_tutorial Tutorial 

@tableofcontents

In this tutorial it assumed that you have built wxAutoExcel library
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
    // display MS Excel window if required
	// MS Excel launched in automation mode has its window hidden
	app.SetVisible(true); 
@endcode

<b>Attaching to any running instance </b>

This works analogically to wxAutomationObject::GetInstance(), 
i.e. you can pass flags which affect its behaviour, such as 
the (default) wxAutomationInstance_CreateIfNeeded.
If you attempt attach to a running instance this way, and there
are more than one instance running, you cannot affect
which one you get. 
@code 
    wxExcelApplication app = wxExcelApplication::GetInstance();
    if ( !app ) 
    {
        wxLogError(_("Failed to obtain an instance of MS Excel application."));
    }    
@endcode

<b>Attaching to an instance with specified workbook open</b>
@code 
    wxExcelApplication app = wxExcelApplication::GetInstance("c:\\budget.xlsx");
    if ( !app ) 
    {
        wxLogError(_("Failed to obtain an instance of MS Excel application."));
    }    
@endcode

<b>Closing the application</b>

Once you are done with the application, do not want to use any
of its objects and wish to close it call
@code 
    wxExcelApplication app;
	// create and use the application 	
	app.Quit();	
@endcode


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

<b>Opening existing workbook</b>
@code 
    // app is a valid instance of wxExcelApplication    
	// fileName is the full path of an existing file	
	wxExcelWorkbook workbook = app.GetWorkbooks().Open(fileName);
    if ( !workbook ) 
    {
        wxLogError(_("Failed to open a workbook."));        
    }
@endcode

 */