#include "getdata.h"
#include "enum2string.h"


// few select Application properties
void ExcelSpy::GetApplicationData(wxExcelApplication& app, wxStringPairVector& data)
{
    data.push_back(std::make_pair("Name", app.GetName()));
    data.push_back( std::make_pair("Version", app.GetVersion()) );
    data.push_back( std::make_pair("Build", wxString::Format("%g", app.GetBuild())) );
    data.push_back( std::make_pair("Product code", app.GetProductCode()) );
    data.push_back( std::make_pair("Operating system", app.GetOperatingSystem()) );
    data.push_back( std::make_pair("Organization name", app.GetOrganizationName()) );
    data.push_back( std::make_pair("User name", app.GetUserName()) );
    data.push_back( std::make_pair("Path", app.GetPath()) );
    data.push_back( std::make_pair("Startup path", app.GetStartupPath()) );
    data.push_back( std::make_pair("Alternate startup path", app.GetAltStartupPath()) );
    data.push_back( std::make_pair("Default file path", app.GetDefaultFilePath()) );
    data.push_back( std::make_pair("Templates path", app.GetTemplatesPath()) );
    data.push_back( std::make_pair("Library path", app.GetLibraryPath()) );
    data.push_back( std::make_pair("User library path", app.GetUserLibraryPath()) );
    data.push_back( std::make_pair("Active printer", app.GetActivePrinter()) );

    data.push_back( std::make_pair("Standard font", app.GetStandardFont()) );
    data.push_back( std::make_pair("Standard font size", wxString::Format("%g", app.GetStandardFontSize())) );    
    data.push_back( std::make_pair("Default save format", XlFileFormat_ToStr(app.GetDefaultSaveFormat())) );
    data.push_back( std::make_pair("Sheets in new workbook", wxString::Format("%ld", app.GetSheetsInNewWorkbook())) );    


    data.push_back( std::make_pair("Use system separators", app.GetUseSystemSeparators() ? "True" : "False") );
    data.push_back( std::make_pair("Decimal separator", app.GetDecimalSeparator()) );
    data.push_back( std::make_pair("Thousands separator", app.GetThousandsSeparator()) );
    data.push_back( std::make_pair("Measurement unit", XlMeasurementUnits_ToStr(app.GetMeasurementUnit())) );
    
    data.push_back( std::make_pair("Language settingsSheets: User interface language", 
        wxString::Format("%ld", app.GetLanguageSettings().GetLanguageID(msoLanguageIDUI))) );
    data.push_back( std::make_pair("Language settingsSheets: Execution mode language", 
        wxString::Format("%ld", app.GetLanguageSettings().GetLanguageID(msoLanguageIDExeMode))) );
    
    data.push_back( std::make_pair("Application security", MsoAutomationSecurity_ToStr(app.GetAutomationSecurity())) );
    data.push_back( std::make_pair("Always use ClearType", app.GetAlwaysUseClearType() ? "True" : "False") );
    data.push_back( std::make_pair("Show windows in Taskbar", app.GetShowWindowsInTaskbar() ? "True" : "False") );
    data.push_back( std::make_pair("Ask to update links", app.GetAskToUpdateLinks() ? "True" : "False") );
}

// Application.International properties
void ExcelSpy::GetInternationalData(wxExcelApplication& app, wxStringPairVector& data)
{
    wxVariant v = app.GetInternational();
    
    data.push_back( std::make_pair("xlCountryCode", v[xlCountryCode-1].MakeString()) );
    data.push_back( std::make_pair("xlCountrySetting", v[xlCountrySetting-1].MakeString()) );
    data.push_back( std::make_pair("xlDecimalSeparator", v[xlDecimalSeparator-1].MakeString()) );
    data.push_back( std::make_pair("xlThousandsSeparator", v[xlThousandsSeparator-1].MakeString()) );
    data.push_back( std::make_pair("xlListSeparator", v[xlListSeparator-1].MakeString()) );
    data.push_back( std::make_pair("xlUpperCaseRowLetter", v[xlUpperCaseRowLetter-1].MakeString()) );
    data.push_back( std::make_pair("xlUpperCaseColumnLetter", v[xlUpperCaseColumnLetter-1].MakeString()) );
    data.push_back( std::make_pair("xlLowerCaseRowLetter", v[xlLowerCaseRowLetter-1].MakeString()) );
    data.push_back( std::make_pair("xlLowerCaseColumnLetter", v[xlLowerCaseColumnLetter-1].MakeString()) );
    data.push_back( std::make_pair("xlLeftBracket", v[xlLeftBracket-1].MakeString()) );
    data.push_back( std::make_pair("xlRightBracket", v[xlRightBracket-1].MakeString()) );
    data.push_back( std::make_pair("xlLeftBrace", v[xlLeftBrace-1].MakeString()) );
    data.push_back( std::make_pair("xlRightBrace", v[xlRightBrace-1].MakeString()) );
    data.push_back( std::make_pair("xlColumnSeparator", v[xlColumnSeparator-1].MakeString()) );
    data.push_back( std::make_pair("xlRowSeparator", v[xlRowSeparator-1].MakeString()) );
    data.push_back( std::make_pair("xlAlternateArraySeparator", v[xlAlternateArraySeparator-1].MakeString()) );
    data.push_back( std::make_pair("xlDateSeparator", v[xlDateSeparator-1].MakeString()) );
    data.push_back( std::make_pair("xlTimeSeparator", v[xlTimeSeparator-1].MakeString()) );
    data.push_back( std::make_pair("xlYearCode", v[xlYearCode-1].MakeString()) );
    data.push_back( std::make_pair("xlMonthCode", v[xlMonthCode-1].MakeString()) );
    data.push_back( std::make_pair("xlDayCode", v[xlDayCode-1].MakeString()) );
    data.push_back( std::make_pair("xlHourCode", v[xlHourCode-1].MakeString()) );
    data.push_back( std::make_pair("xlMinuteCode", v[xlMinuteCode-1].MakeString()) );
    data.push_back( std::make_pair("xlSecondCode", v[xlSecondCode-1].MakeString()) );
    data.push_back( std::make_pair("xlCurrencyCode", v[xlCurrencyCode-1].MakeString()) );
    data.push_back( std::make_pair("xlGeneralFormatName", v[xlGeneralFormatName-1].MakeString()) );
    data.push_back( std::make_pair("xlCurrencyDigits", v[xlCurrencyDigits-1].MakeString()) );
    data.push_back( std::make_pair("xlCurrencyNegative", v[xlCurrencyNegative-1].MakeString()) );
    data.push_back( std::make_pair("xlNoncurrencyDigits", v[xlNoncurrencyDigits-1].MakeString()) );
    data.push_back( std::make_pair("xlMonthNameChars", v[xlMonthNameChars-1].MakeString()) );
    data.push_back( std::make_pair("xlWeekdayNameChars", v[xlWeekdayNameChars-1].MakeString()) );
    data.push_back( std::make_pair("xlDateOrder", v[xlDateOrder-1].MakeString()) );
    data.push_back( std::make_pair("xl24HourClock", v[xl24HourClock-1].MakeString()) );
    data.push_back( std::make_pair("xlNonEnglishFunctions", v[xlNonEnglishFunctions-1].MakeString()) );
    data.push_back( std::make_pair("xlMetric", v[xlMetric-1].MakeString()) );
    data.push_back( std::make_pair("xlCurrencySpaceBefore", v[xlCurrencySpaceBefore-1].MakeString()) );
    data.push_back( std::make_pair("xlCurrencyBefore", v[xlCurrencyBefore-1].MakeString()) );
    data.push_back( std::make_pair("xlCurrencyMinusSign", v[xlCurrencyMinusSign-1].MakeString()) );
    data.push_back( std::make_pair("xlCurrencyTrailingZeros", v[xlCurrencyTrailingZeros-1].MakeString()) );
    data.push_back( std::make_pair("xlCurrencyLeadingZeros", v[xlCurrencyLeadingZeros-1].MakeString()) );
    data.push_back( std::make_pair("xlMonthLeadingZero", v[xlMonthLeadingZero-1].MakeString()) );
    data.push_back( std::make_pair("xlDayLeadingZero", v[xlDayLeadingZero-1].MakeString()) );
    data.push_back( std::make_pair("xl4DigitYears", v[xl4DigitYears-1].MakeString()) );
    data.push_back( std::make_pair("xlMDY", v[xlMDY-1].MakeString()) );
    data.push_back( std::make_pair("xlTimeLeadingZero", v[xlTimeLeadingZero-1].MakeString()) );
}

// RecentFiles data
void ExcelSpy::GetRecentFilesData(wxExcelApplication& app, wxStringPairVector& data)
{
    wxExcelRecentFiles files = app.GetRecentFiles();    
    long count = files.GetCount();
        
    data.push_back( std::make_pair("Maximum", wxString::Format("%ld", files.GetMaximum())) );
    data.push_back( std::make_pair("Count", wxString::Format("%ld", files.GetCount())) );
    for ( long l = 1; l <= count; l++ )
    {
        data.push_back( std::make_pair(wxString::Format("File %ld", l), files[l].GetPath()) );
    }
}

void FillAddInsOrAddIns2Data(wxExcelAddInsBase* addInsBase, wxStringPairVector& data)
{
    if ( !addInsBase->IsOk_() )
        return;
    
    long count = addInsBase->GetCount();
    for ( long l = 1; l <= count; l++ )
    {
        wxExcelAddIn addIn = addInsBase->GetItem(l);
        wxString info;
        
        info << "Installed=";
        info << (addIn.GetInstalled() ? "Yes" : "No"); info << "; ";
        info << "IsOpen=";
        info << (addIn.GetIsOpen() ? "Yes" : "No"); info << "; ";
        info << "Path=" << addIn.GetPath();

        data.push_back(std::make_pair(addIn.GetName(), info));
    }

}

// AddIns data
void ExcelSpy::GetAddInsData(wxExcelApplication& app, wxStringPairVector& data)
{
    wxExcelAddIns addIns = app.GetAddIns();    
    FillAddInsOrAddIns2Data(&addIns, data);    
}

// AddIns2 data
void ExcelSpy::GetAddIns2Data(wxExcelApplication& app, wxStringPairVector& data)
{
    wxExcelAddIns2 addIns2 = app.GetAddIns2();
    FillAddInsOrAddIns2Data(&addIns2, data);
}


// few select Workbook properties
void ExcelSpy::GetWorkbookData(wxExcelApplication& app, wxExcelWorkbook& workbook, wxStringPairVector& data)
{    
    data.push_back( std::make_pair("Full name", workbook.GetFullName()) );    
    data.push_back( std::make_pair("File format", XlFileFormat_ToStr(workbook.GetFileFormat())) );
    if ( app.IsVersionAtLeast_(wxExcelApplication::evExcel2007) )
        data.push_back( std::make_pair("Excel8CompatibilityMode", workbook.GetExcel8CompatibilityMode() ? "True" : "False") );
    
    data.push_back( std::make_pair("Sheets.Count", wxString::Format("%ld", workbook.GetSheets().GetCount())) );
    data.push_back( std::make_pair("WorkSheets.Count", wxString::Format("%ld", workbook.GetWorksheets().GetCount())) );

#if WXAUTOEXCEL_USE_CHARTS        
        wxExcelCharts charts = workbook.GetCharts();
        
        if ( charts )
        {
            data.push_back( std::make_pair("Charts.Count", wxString::Format("%ld", charts.GetCount())) );
        }
#endif // #if WXAUTOEXCEL_USE_CHARTS    
        
    data.push_back( std::make_pair("Create backup", workbook.GetCreateBackup() ? "True" : "False") );
    data.push_back( std::make_pair("Has password", workbook.GetHasPassword() ? "True" : "False") );
    if ( app.IsVersionAtLeast_(wxExcelApplication::evExcel2007) )
        data.push_back( std::make_pair("Has VB project", workbook.GetHasVBProject() ? "True" : "False") );
}

// document properties
void ExcelSpy::GetDocumentPropertiesData(wxExcelDocumentProperties props, wxStringPairVector& data)
{
    long count = props.GetCount();
    wxExcelDocumentProperty p;
    wxVariant val;    
    
    for ( long l = 1; l <= count; l++ )
    {                
        p = props[l];                            
        
        // Excel can fail with an error if a property has not been assigned a value
        // so we need to override its default error handling here        
        wxAutoExcelObjectErrorModeOverrider emo(wxExcelObject::Err_DoNothing, true);
        val = p.GetValue();                                
        data.push_back( std::make_pair(p.GetName(), p ? val.MakeString() : "<Not set>") );        
    }    
}

// Styles
void ExcelSpy::GetStylesData(wxExcelWorkbook& workbook, wxStringPairVector& data)
{
    wxExcelStyles styles = workbook.GetStyles();

    if ( !styles )
        return;
    
    long count = styles.GetCount();
    wxExcelStyle style;
    wxString name, nameLocal;

    for ( long l = 1; l <= count; l++ )
    {
        style = styles[l];                

        name = style.GetName();
        if ( style.GetBuiltIn() )
            name += " (Built-in)";
        nameLocal = style.GetNameLocal();        
        data.push_back( std::make_pair(name, nameLocal) );    
    }    
}

// Names
void ExcelSpy::GetNamesData(wxExcelWorkbook& workbook, wxStringPairVector& data)
{
    wxExcelNames names = workbook.GetNames();

    if ( !names )
        return;

    long count = names.GetCount();
    wxExcelName name;

    for ( long l = 1; l <= count; l++ )
    {
        name = names[l];

        data.push_back(std::make_pair(name.GetName(), name.GetValue()));
    }
}

// few select Sheets properties
void ExcelSpy::GetSheetsData(wxExcelSheets& sheets, wxStringPairVector& data)
{
    data.push_back( std::make_pair("Count", wxString::Format("%ld", sheets.GetCount())) );
}

// few select Sheet properties
void ExcelSpy::GetSheetData(wxExcelSheet& sheet, wxStringPairVector& data)
{
    data.push_back( std::make_pair("Name", sheet.GetName()) );
    
    wxString s;
    XlSheetType xlType = sheet.GetType();
    
    switch (xlType)
    {
        case xlChart:
            s = "Chart";
            break;
        case xlDialogSheet:
            s = "Dialog sheet";
            break;
        case xlExcel4IntlMacroSheet:
            s = "Excel version 4 international macro sheet";
            break;
        case xlExcel4MacroSheet:
            s = "Excel version 4 macro sheet";
            break;
        case xlWorksheet:
            s = "Worksheet";
            break;
        default:
            s = "Uknown sheet type";
    }    

    // work around Sheet.Type returning wrong values for Charts
    if ( sheet.IsChart() )
         s = "Chart";    
    data.push_back( std::make_pair("Type", s) );        
}

// few select Worksheets properties
void ExcelSpy::GetWorksheetsData(wxExcelWorksheets& sheets, wxStringPairVector& data)
{
    data.push_back( std::make_pair("Count", wxString::Format("%ld", sheets.GetCount())) );
}

// few select Worksheet properties
void ExcelSpy::GetWorksheetData(wxExcelWorksheet& sheet, wxStringPairVector& data)
{
    data.push_back( std::make_pair("Name", sheet.GetName()) );
    data.push_back( std::make_pair("Index", wxString::Format("%ld", sheet.GetIndex())) );    

    data.push_back( std::make_pair("Visible", XlSheetVisibility_ToStr(sheet.GetVisible())) );
    data.push_back( std::make_pair("Standard height", wxString::Format("%g", sheet.GetStandardHeight())) );    
    data.push_back( std::make_pair("Standard width", wxString::Format("%g", sheet.GetStandardWidth())) );    
    data.push_back( std::make_pair("FilterMode", sheet.GetFilterMode() ? "True" : "False") );
    data.push_back( std::make_pair("AutoFilterMode", sheet.GetAutoFilterMode() ? "True" : "False") );
    data.push_back( std::make_pair("DisplayPageBreaks", sheet.GetDisplayPageBreaks() ? "True" : "False") );
    data.push_back( std::make_pair("ProtectContents", sheet.GetProtectContents() ? "True" : "False") );
    
}

// few select PageSetup properties
void ExcelSpy::GetPageSetupData(wxExcelPageSetup& pageSetup, wxStringPairVector& data)
{    
    if ( !pageSetup )
        return;

    wxString s;
                
    {
        // page setup for chart sheets does not use the four following properties
        // so let's suprress eventual error reports
        wxAutoExcelObjectErrorModeOverrider emo(wxExcelObject::Err_DoNothing, true);        
        
        pageSetup.GetPrintArea();
        if ( pageSetup )
        {
            data.push_back( std::make_pair("Print area", pageSetup.GetPrintArea()) );
            data.push_back( std::make_pair("Print headings", pageSetup.GetPrintHeadings() ? "True" : "False") );
            data.push_back( std::make_pair("Print gridlines", pageSetup.GetPrintGridlines() ? "True" : "False") );
            
            XlOrder order = pageSetup.GetOrder();
            switch ( order )
            {
                case xlDownThenOver:
                    s = "DownThenOver";
                    break;
                case xlOverThenDown:
                    s = "OverThenDown";
                    break;
                default:
                    s = wxString::Format("Unknown (%ld)", (long)order);
                    break;
            }
            data.push_back( std::make_pair("Order", s) );
        } 
        else
        {
            data.push_back( std::make_pair("Print area", "not supported for charts") );
            data.push_back( std::make_pair("Print headings", "not supported for charts") );
            data.push_back( std::make_pair("Print gridlines", "not supported for charts") );
            data.push_back( std::make_pair("Order", "not supported for charts") );
        }

    }

    data.push_back( std::make_pair("Paper size", XlPaperSize_ToStr(pageSetup.GetPaperSize())) );
    
    XlPageOrientation orientation = pageSetup.GetOrientation();
    
    switch ( orientation )
    {
        case xlLandscape:
            s = "Landscape";
            break;
        case xlPortrait:
            s = "Portrait";
            break;
        default:
            s = wxString::Format("Unknown (%ld)", (long)orientation);
            break;
    }
    data.push_back( std::make_pair("Orientation", s) );   

    data.push_back( std::make_pair("Top margin", wxString::Format("%g pt", pageSetup.GetTopMargin())) );
    data.push_back( std::make_pair("Left margin", wxString::Format("%g pt", pageSetup.GetLeftMargin())) );
    data.push_back( std::make_pair("Right margin", wxString::Format("%g pt", pageSetup.GetRightMargin())) );
    data.push_back( std::make_pair("Bottom margin", wxString::Format("%g pt", pageSetup.GetBottomMargin())) );

    data.push_back( std::make_pair("Header margin", wxString::Format("%g pt", pageSetup.GetHeaderMargin())) );
    data.push_back( std::make_pair("Footer margin", wxString::Format("%g pt", pageSetup.GetFooterMargin())) );

    data.push_back( std::make_pair("Left header", pageSetup.GetLeftHeader()) );
    data.push_back( std::make_pair("Center header", pageSetup.GetCenterHeader()) );
    data.push_back( std::make_pair("Right header", pageSetup.GetRightHeader()) );
    data.push_back( std::make_pair("Left footer", pageSetup.GetLeftFooter()) );
    data.push_back( std::make_pair("Center footer", pageSetup.GetCenterFooter()) );
    data.push_back( std::make_pair("Right footer", pageSetup.GetRightFooter()) );
    
}

wxString wxXlTriboolToString(const wxXlTribool& tb)
{
    if ( tb.IsTrue() )
        return "True";
    if ( tb.IsFalse() )
        return "False";

    return "Default / Undetermined";
}

// few select Range properties
void ExcelSpy::GetRangeData(wxExcelRange& range, wxStringPairVector& data)
{
    data.push_back( std::make_pair("Address", range.GetAddress()) );    
    data.push_back( std::make_pair("Count", wxString::Format("%ld", range.GetCount())) );
    data.push_back( std::make_pair("CountLarge", range.GetCountLarge().ToString()) );
    
    data.push_back( std::make_pair("Column", wxString::Format("%ld", range.GetColumn())) );
    data.push_back( std::make_pair("Row", wxString::Format("%ld", range.GetRow())) );    
    data.push_back( std::make_pair("Columns.Count", wxString::Format("%ld", range.GetColumns().GetCount())) );
    data.push_back( std::make_pair("Rows.Count", wxString::Format("%ld", range.GetRows().GetCount())) );

    data.push_back( std::make_pair("Width", wxString::Format("%g pt", range.GetWidth())) );
    data.push_back( std::make_pair("Height", wxString::Format("%g pt", range.GetHeight())) );

    data.push_back( std::make_pair("UseStandardWidth", wxXlTriboolToString(range.GetUseStandardWidth())) );       
    data.push_back( std::make_pair("UseStandardHeight", wxXlTriboolToString(range.GetUseStandardHeight())) );       

    data.push_back( std::make_pair("A1.Text", range.GetRange("A1").GetText()) );       

}


void ExcelSpy::GetOLEObjectsData(wxExcelOLEObjects& objects, wxStringPairVector& data)
{
    data.push_back( std::make_pair("Count", wxString::Format("%ld", objects.GetCount())) );
}

void ExcelSpy::GetOLEObjectData(wxExcelOLEObject& object, wxStringPairVector& data)
{    
    data.push_back( std::make_pair("Name", object.GetName()) );    

    wxString s;
    XlOLEType oType = object.GetOLEType();
    
    switch ( oType )
    {
        case xlOLEControl:
            s = "ActiveX control";
            break;
        case xlOLEEmbed:
            s = "Embedded OLE object";
            break;
        case xlOLELink:
            s = "OLE link";
            break;
        default:
            s = wxString::Format("Unknown (%ld)", (long)oType);
            break;
    }
    data.push_back( std::make_pair("Type", s) );
    if ( oType == xlOLELink )
    {
        data.push_back( std::make_pair("SourceName", object.GetSourceName()) );
    }
    data.push_back( std::make_pair("progID", object.GetprogID()) );    

    data.push_back( std::make_pair("Index", wxString::Format("%ld", object.GetIndex())) );
    data.push_back( std::make_pair("AutoLoad", object.GetAutoLoad() ? "True" : "False") );
    if ( oType == xlOLELink )
    {
        data.push_back( std::make_pair("AutoUpdate", object.GetAutoUpdate() ? "True" : "False") );
    }
    data.push_back( std::make_pair("Enabled", object.GetEnabled() ? "True" : "False") );     
    data.push_back( std::make_pair("Visible", object.GetVisible() ? "True" : "False") );     

    data.push_back( std::make_pair("Top", wxString::Format("%g pt", object.GetTop())) );
    data.push_back( std::make_pair("Left", wxString::Format("%g pt", object.GetLeft())) );
    data.push_back( std::make_pair("Width", wxString::Format("%g pt", object.GetWidth())) );
    data.push_back( std::make_pair("Height", wxString::Format("%g pt", object.GetHeight())) );
}

void ExcelSpy::GetHyperlinksData(wxExcelHyperlinks& links, wxStringPairVector& data)
{
    data.push_back( std::make_pair("Count", wxString::Format("%ld", links.GetCount())) );
}

void ExcelSpy::GetHyperlinkData(wxExcelHyperlink& link, wxStringPairVector& data)
{    
    data.push_back( std::make_pair("Name", link.GetName()) );    
    data.push_back( std::make_pair("Address", link.GetAddress()) );    
    data.push_back( std::make_pair("Subaddress", link.GetSubAddress()) );        
    wxString s;
    MsoHyperlinkType lType = link.GetType();
    
    switch ( lType )
    {
        case msoHyperlinkRange  :
            s = "Range";
            break;
        case msoHyperlinkShape:
            s = "Shape";
            break;
        case msoHyperlinkInlineShape:
            s = "Inline shape";
            break;
        default:
            s = wxString::Format("Unknown (%ld)", (long)lType);
            break;
    }
    data.push_back( std::make_pair("Type", s) );    

    data.push_back( std::make_pair("ScreenTip", link.GetScreenTip()) );    
    data.push_back( std::make_pair("Text to display", lType == msoHyperlinkRange ? link.GetTextToDisplay() : "n/a") );        
}


#if WXAUTOEXCEL_USE_CHARTS

void ExcelSpy::GetChartsData(wxExcelCharts& charts, wxStringPairVector& data)
{
    data.push_back( std::make_pair("Count", wxString::Format("%ld", charts.GetCount())) );
}


void ExcelSpy::GetChartObjectsData(wxExcelChartObjects& chartObjects, wxStringPairVector& data)
{
    data.push_back( std::make_pair("Count", wxString::Format("%ld", chartObjects.GetCount())) );
}

void ExcelSpy::GetChartObjectData(wxExcelChartObject& chartObject, wxStringPairVector& data)
{
    data.push_back( std::make_pair("Name", chartObject.GetName()) );    
    
    data.push_back( std::make_pair("Enabled", chartObject.GetEnabled() ? "True" : "False") );     
    data.push_back( std::make_pair("Visible", chartObject.GetVisible() ? "True" : "False") );     
    data.push_back( std::make_pair("Locked", chartObject.GetLocked() ? "True" : "False") );     

    data.push_back( std::make_pair("Top", wxString::Format("%g pt", chartObject.GetTop())) );
    data.push_back( std::make_pair("Left", wxString::Format("%g pt", chartObject.GetLeft())) );
    data.push_back( std::make_pair("Width", wxString::Format("%g pt", chartObject.GetWidth())) );
    data.push_back( std::make_pair("Height", wxString::Format("%g pt", chartObject.GetHeight())) );
    data.push_back( std::make_pair("Rounded corners", chartObject.GetRoundedCorners() ? "True" : "False") );  

    wxExcelChart chart = chartObject.GetChart();
    GetChartData(chart, data);
}

void ExcelSpy::GetChartData(wxExcelChart& chart, wxStringPairVector& data)
{
    data.push_back( std::make_pair("Chart.Name", chart.GetName()) );
    data.push_back( std::make_pair("Chart.Type", XlChartType_ToStr(chart.GetChartType())) );
    
    {
        wxAutoExcelObjectErrorModeOverrider emo(wxExcelObject::Err_DoNothing, true);
        long index = chart.GetIndex();
        if ( index > 0 )
            data.push_back( std::make_pair("Chart.Index", wxString::Format("%ld", index)) );
        else
            data.push_back( std::make_pair("Chart.Index", "n/a (embedded chart)") );
    }
    wxString s;

    if ( chart.GetHasTitle() )
        s = chart.GetChartTitle().GetCaption();
    else
        s = "<empty>";
    data.push_back( std::make_pair("Chart.Title", s) );
}

#endif // #if WXAUTOEXCEL_USE_CHARTS

#if WXAUTOEXCEL_USE_SHAPES

void ExcelSpy::GetShapesData(wxExcelShapes& shapes, wxStringPairVector& data)
{
    data.push_back( std::make_pair("Count", wxString::Format("%ld", shapes.GetCount())) );
}

void ExcelSpy::GetShapeData(wxExcelShape& shape, wxStringPairVector& data)
{
    data.push_back( std::make_pair("Name", shape.GetName()) );            
    data.push_back( std::make_pair("Type", MsoShapeType_ToStr(shape.GetType())) );
    data.push_back( std::make_pair("Visible", shape.GetVisible() ? "True" : "False") );     
    data.push_back( std::make_pair("Top", wxString::Format("%g pt", shape.GetTop())) );
    data.push_back( std::make_pair("Left", wxString::Format("%g pt", shape.GetLeft())) );
    data.push_back( std::make_pair("Width", wxString::Format("%g pt", shape.GetWidth())) );
    data.push_back( std::make_pair("Height", wxString::Format("%g pt", shape.GetHeight())) );   
}

#endif // #if WXAUTOEXCEL_USE_SHAPES

// Comments
void ExcelSpy::GetCommentsData(wxExcelWorksheet& sheet, wxStringPairVector& data)
{
    wxExcelComments comments = sheet.GetComments();

    if ( !comments )
        return;
    
    long count = comments.GetCount();
    wxExcelComment comment;

    for ( long l = 1; l <= count; l++ )
    {
        comment = comments[l];

        data.push_back( std::make_pair(comment.GetAuthor(), comment.Text()) );    
    }    
}
