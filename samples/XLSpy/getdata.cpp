#include "getdata.h"
#include "enum2string.h"

// Adds the name and the value to data, or <error> instead of the value
// when the call of the object which obtained the value failed
void AddValue(wxStringPairVector& data, const wxString& name,
              const wxExcelObject& object, const wxString& value)
{
    data.push_back( std::make_pair(name, object ? value : wxString("<error>")) );
}


// few select Application properties
void ExcelSpy::GetApplicationData(wxExcelApplication& app, wxStringPairVector& data)
{
    AddValue(data, "Name", app, app.GetName());
    AddValue(data, "Version", app, app.GetVersion());
    AddValue(data, "Build", app, wxString::Format("%g", app.GetBuild()));
    AddValue(data, "Full build", app, app.GetFullBuild_());
    AddValue(data, "Product code", app, app.GetProductCode());
    AddValue(data, "Operating system", app, app.GetOperatingSystem());
    AddValue(data, "Organization name", app, app.GetOrganizationName());
    AddValue(data, "User name", app, app.GetUsername());
    AddValue(data, "Path", app, app.GetPath());
    AddValue(data, "Startup path", app, app.GetStartupPath());
    AddValue(data, "Alternate startup path", app, app.GetAltStartupPath());
    AddValue(data, "Default file path", app, app.GetDefaultFilePath());
    AddValue(data, "Templates path", app, app.GetTemplatesPath());
    AddValue(data, "Library path", app, app.GetLibraryPath());
    AddValue(data, "User library path", app, app.GetUserLibraryPath());
    AddValue(data, "Active printer", app, app.GetActivePrinter());

    AddValue(data, "Standard font", app, app.GetStandardFont());
    AddValue(data, "Standard font size", app, wxString::Format("%g", app.GetStandardFontSize()));
    AddValue(data, "Default save format", app, XlFileFormat_ToStr(app.GetDefaultSaveFormat()));
    AddValue(data, "Sheets in new workbook", app, wxString::Format("%ld", app.GetSheetsInNewWorkbook()));


    AddValue(data, "Use system separators", app, app.GetUseSystemSeparators() ? "True" : "False");
    AddValue(data, "Decimal separator", app, app.GetDecimalSeparator());
    AddValue(data, "Thousands separator", app, app.GetThousandsSeparator());
    AddValue(data, "Measurement unit", app, XlMeasurementUnits_ToStr(app.GetMeasurementUnit()));

    data.push_back( std::make_pair("Language settings: User interface language",
        wxString::Format("%ld", app.GetLanguageSettings().GetLanguageID(msoLanguageIDUI))) );
    data.push_back( std::make_pair("Language settings: Execution mode language",
        wxString::Format("%ld", app.GetLanguageSettings().GetLanguageID(msoLanguageIDExeMode))) );

    AddValue(data, "Application security", app, MsoAutomationSecurity_ToStr(app.GetAutomationSecurity()));
    AddValue(data, "Always use ClearType", app, app.GetAlwaysUseClearType() ? "True" : "False");
    AddValue(data, "Show windows in Taskbar", app, app.GetShowWindowsInTaskbar() ? "True" : "False");
    AddValue(data, "Ask to update links", app, app.GetAskToUpdateLinks() ? "True" : "False");
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

    AddValue(data, "Maximum", files, wxString::Format("%ld", files.GetMaximum()));
    AddValue(data, "Count", files, wxString::Format("%ld", files.GetCount()));
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


void ExcelSpy::GetFileExportConvertersData(wxExcelApplication& app, wxStringPairVector& data)
{
    wxExcelFileExportConverters converters = app.GetFileExportConverters();
    const long count = converters.GetCount();

    for ( long l = 1; l <= count; l++ )
    {
        wxExcelFileExportConverter converter = converters[l];
        const long fileFormat = converter.GetFileFormat();
        const wxString description = converter.GetDescription();
        const wxString extensions = converter.GetExtensions();

        data.push_back( std::make_pair(wxString::Format("%s (%ld)", description, fileFormat), extensions) );
    }

}

// few select Workbook properties
void ExcelSpy::GetWorkbookData(wxExcelApplication& app, wxExcelWorkbook& workbook, wxStringPairVector& data)
{
    AddValue(data, "Full name", workbook, workbook.GetFullName());
    AddValue(data, "File format", workbook, XlFileFormat_ToStr(workbook.GetFileFormat()));
    if ( app.IsVersionAtLeast_(wxExcelApplication::evExcel2007) )
        AddValue(data, "Excel8CompatibilityMode", workbook, workbook.GetExcel8CompatibilityMode() ? "True" : "False");

    data.push_back( std::make_pair("Sheets.Count", wxString::Format("%ld", workbook.GetSheets().GetCount())) );
    data.push_back( std::make_pair("WorkSheets.Count", wxString::Format("%ld", workbook.GetWorksheets().GetCount())) );

#if WXAUTOEXCEL_USE_CHARTS
        wxExcelCharts charts = workbook.GetCharts();

        if ( charts )
        {
            AddValue(data, "Charts.Count", charts, wxString::Format("%ld", charts.GetCount()));
        }
#endif // #if WXAUTOEXCEL_USE_CHARTS

    AddValue(data, "Create backup", workbook, workbook.GetCreateBackup() ? "True" : "False");
    AddValue(data, "Has password", workbook, workbook.GetHasPassword() ? "True" : "False");
    if ( app.IsVersionAtLeast_(wxExcelApplication::evExcel2007) )
        AddValue(data, "Has VB project", workbook, workbook.GetHasVBProject() ? "True" : "False");
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

    const long count = styles.GetCount();

    for ( long l = 1; l <= count; l++ )
    {
        wxExcelStyle style = styles[l];
        wxString name, nameLocal;

        name = style.GetName();
        if ( style.GetBuiltIn() )
            name += " (Built-in)";
        nameLocal = style.GetNameLocal();
        data.push_back( std::make_pair(name, nameLocal) );
    }
}

// TableStyles
void ExcelSpy::GetTableStylesData(wxExcelWorkbook& workbook, wxStringPairVector& data)
{
    wxExcelTableStyles styles = workbook.GetTableStyles();

    if ( !styles )
        return;

    const long count = styles.GetCount();

    for ( long l = 1; l <= count; l++ )
    {
        wxExcelTableStyle style = styles[l];
        wxString name, nameLocal;

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

// Colors
void ExcelSpy::GetColorsData(wxExcelWorkbook& workbook, wxStringPairVector& data)
{
    const wxVector<wxColour> colors = workbook.GetColors();

    for ( size_t i = 0; i < colors.size(); ++i )
    {
        data.push_back(std::make_pair(wxString::Format("Color %zu", i + 1), colors[i].GetAsString(wxC2S_HTML_SYNTAX)));
    }
}


// few select Sheets properties
void ExcelSpy::GetSheetsData(wxExcelSheets& sheets, wxStringPairVector& data)
{
    AddValue(data, "Count", sheets, wxString::Format("%ld", sheets.GetCount()));
}

// few select Sheet properties
void ExcelSpy::GetSheetData(wxExcelSheet& sheet, wxStringPairVector& data)
{
    AddValue(data, "Name", sheet, sheet.GetName());

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
    AddValue(data, "Count", sheets, wxString::Format("%ld", sheets.GetCount()));
}

// few select Worksheet properties
void ExcelSpy::GetWorksheetData(wxExcelWorksheet& sheet, wxStringPairVector& data)
{
    AddValue(data, "Name", sheet, sheet.GetName());
    AddValue(data, "Index", sheet, wxString::Format("%ld", sheet.GetIndex()));

    AddValue(data, "Visible", sheet, XlSheetVisibility_ToStr(sheet.GetVisible()));
    AddValue(data, "Standard height", sheet, wxString::Format("%g", sheet.GetStandardHeight()));
    AddValue(data, "Standard width", sheet, wxString::Format("%g", sheet.GetStandardWidth()));
    AddValue(data, "FilterMode", sheet, sheet.GetFilterMode() ? "True" : "False");
    AddValue(data, "AutoFilterMode", sheet, sheet.GetAutoFilterMode() ? "True" : "False");
    AddValue(data, "DisplayPageBreaks", sheet, sheet.GetDisplayPageBreaks() ? "True" : "False");
    AddValue(data, "ProtectContents", sheet, sheet.GetProtectContents() ? "True" : "False");

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
            AddValue(data, "Print area", pageSetup, pageSetup.GetPrintArea());
            AddValue(data, "Print headings", pageSetup, pageSetup.GetPrintHeadings() ? "True" : "False");
            AddValue(data, "Print gridlines", pageSetup, pageSetup.GetPrintGridlines() ? "True" : "False");

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
            AddValue(data, "Order", pageSetup, s);
        }
        else
        {
            data.push_back( std::make_pair("Print area", "not supported for charts") );
            data.push_back( std::make_pair("Print headings", "not supported for charts") );
            data.push_back( std::make_pair("Print gridlines", "not supported for charts") );
            data.push_back( std::make_pair("Order", "not supported for charts") );
        }

    }

    AddValue(data, "Paper size", pageSetup, XlPaperSize_ToStr(pageSetup.GetPaperSize()));

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
    AddValue(data, "Orientation", pageSetup, s);

    AddValue(data, "Top margin", pageSetup, wxString::Format("%g pt", pageSetup.GetTopMargin()));
    AddValue(data, "Left margin", pageSetup, wxString::Format("%g pt", pageSetup.GetLeftMargin()));
    AddValue(data, "Right margin", pageSetup, wxString::Format("%g pt", pageSetup.GetRightMargin()));
    AddValue(data, "Bottom margin", pageSetup, wxString::Format("%g pt", pageSetup.GetBottomMargin()));

    AddValue(data, "Header margin", pageSetup, wxString::Format("%g pt", pageSetup.GetHeaderMargin()));
    AddValue(data, "Footer margin", pageSetup, wxString::Format("%g pt", pageSetup.GetFooterMargin()));

    AddValue(data, "Left header", pageSetup, pageSetup.GetLeftHeader());
    AddValue(data, "Center header", pageSetup, pageSetup.GetCenterHeader());
    AddValue(data, "Right header", pageSetup, pageSetup.GetRightHeader());
    AddValue(data, "Left footer", pageSetup, pageSetup.GetLeftFooter());
    AddValue(data, "Center footer", pageSetup, pageSetup.GetCenterFooter());
    AddValue(data, "Right footer", pageSetup, pageSetup.GetRightFooter());

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
    AddValue(data, "Address", range, range.GetAddress());
    AddValue(data, "Count", range, wxString::Format("%ld", range.GetCount()));
    AddValue(data, "CountLarge", range, range.GetCountLarge().ToString());

    AddValue(data, "Column", range, wxString::Format("%ld", range.GetColumn()));
    AddValue(data, "Row", range, wxString::Format("%ld", range.GetRow()));
    data.push_back( std::make_pair("Columns.Count", wxString::Format("%ld", range.GetColumns().GetCount())) );
    data.push_back( std::make_pair("Rows.Count", wxString::Format("%ld", range.GetRows().GetCount())) );

    AddValue(data, "Width", range, wxString::Format("%g pt", range.GetWidth()));
    AddValue(data, "Height", range, wxString::Format("%g pt", range.GetHeight()));

    AddValue(data, "UseStandardWidth", range, wxXlTriboolToString(range.GetUseStandardWidth()));
    AddValue(data, "UseStandardHeight", range, wxXlTriboolToString(range.GetUseStandardHeight()));

    // the address is relative to the range, so A1 is its top-left cell
    wxExcelRange topLeftCell = range.GetRange("A1");
    const wxString topLeftCellAddress = topLeftCell.GetAddress(false, false);
    AddValue(data, wxString::Format("Top-left cell (%s).Text", topLeftCellAddress), topLeftCell, topLeftCell.GetText());

}


void ExcelSpy::GetOLEObjectsData(wxExcelOLEObjects& objects, wxStringPairVector& data)
{
    AddValue(data, "Count", objects, wxString::Format("%ld", objects.GetCount()));
}

void ExcelSpy::GetOLEObjectData(wxExcelOLEObject& object, wxStringPairVector& data)
{
    AddValue(data, "Name", object, object.GetName());

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
    AddValue(data, "Type", object, s);
    if ( oType == xlOLELink )
    {
        AddValue(data, "SourceName", object, object.GetSourceName());
    }
    AddValue(data, "progID", object, object.GetprogID());

    AddValue(data, "Index", object, wxString::Format("%ld", object.GetIndex()));
    AddValue(data, "AutoLoad", object, object.GetAutoLoad() ? "True" : "False");
    if ( oType == xlOLELink )
    {
        AddValue(data, "AutoUpdate", object, object.GetAutoUpdate() ? "True" : "False");
    }
    AddValue(data, "Enabled", object, object.GetEnabled() ? "True" : "False");
    AddValue(data, "Visible", object, object.GetVisible() ? "True" : "False");

    AddValue(data, "Top", object, wxString::Format("%g pt", object.GetTop()));
    AddValue(data, "Left", object, wxString::Format("%g pt", object.GetLeft()));
    AddValue(data, "Width", object, wxString::Format("%g pt", object.GetWidth()));
    AddValue(data, "Height", object, wxString::Format("%g pt", object.GetHeight()));
}

void ExcelSpy::GetHyperlinksData(wxExcelHyperlinks& links, wxStringPairVector& data)
{
    AddValue(data, "Count", links, wxString::Format("%ld", links.GetCount()));
}

void ExcelSpy::GetHyperlinkData(wxExcelHyperlink& link, wxStringPairVector& data)
{
    AddValue(data, "Name", link, link.GetName());
    AddValue(data, "Address", link, link.GetAddress());
    AddValue(data, "Subaddress", link, link.GetSubAddress());
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
    AddValue(data, "Type", link, s);

    AddValue(data, "ScreenTip", link, link.GetScreenTip());
    AddValue(data, "Text to display", link, lType == msoHyperlinkRange ? link.GetTextToDisplay() : "n/a");
}


void ExcelSpy::GetListObjectsData(wxExcelListObjects& objects, wxStringPairVector& data)
{
    AddValue(data, "Count", objects, wxString::Format("%ld", objects.GetCount()));
}

void ExcelSpy::GetListObjectData(wxExcelListObject& object, wxStringPairVector& data)
{
    const XlListObjectSourceType sType = object.GetSourceType();
    wxExcelRange range = object.GetRange();
    wxExcelRange dataBodyRange = object.GetDataBodyRange();
    wxExcelTableStyle style = object.GetTableStyle();

    AddValue(data, "Name", object, object.GetName());
    AddValue(data, "DisplayName", object, object.GetDisplayName());
    data.push_back( std::make_pair("SourceType", XlListObjectSourceType_ToStr(sType)) );
    AddValue(data, "Range.Address", range, range.GetAddress());
    AddValue(data, "DataBodyRange.Address", dataBodyRange, dataBodyRange.GetAddress());
    AddValue(data, "Style.Name", style, style.GetName());
    AddValue(data, "Style.NameLocal", style, style.GetNameLocal());
    AddValue(data, "Style.BuiltIn", style, style.GetBuiltIn() ? "Yes" : "No");
}


#if WXAUTOEXCEL_USE_CHARTS

void ExcelSpy::GetChartsData(wxExcelCharts& charts, wxStringPairVector& data)
{
    AddValue(data, "Count", charts, wxString::Format("%ld", charts.GetCount()));
}


void ExcelSpy::GetChartObjectsData(wxExcelChartObjects& chartObjects, wxStringPairVector& data)
{
    AddValue(data, "Count", chartObjects, wxString::Format("%ld", chartObjects.GetCount()));
}

void ExcelSpy::GetChartObjectData(wxExcelChartObject& chartObject, wxStringPairVector& data)
{
    AddValue(data, "Name", chartObject, chartObject.GetName());

    AddValue(data, "Enabled", chartObject, chartObject.GetEnabled() ? "True" : "False");
    AddValue(data, "Visible", chartObject, chartObject.GetVisible() ? "True" : "False");
    AddValue(data, "Locked", chartObject, chartObject.GetLocked() ? "True" : "False");

    AddValue(data, "Top", chartObject, wxString::Format("%g pt", chartObject.GetTop()));
    AddValue(data, "Left", chartObject, wxString::Format("%g pt", chartObject.GetLeft()));
    AddValue(data, "Width", chartObject, wxString::Format("%g pt", chartObject.GetWidth()));
    AddValue(data, "Height", chartObject, wxString::Format("%g pt", chartObject.GetHeight()));
    AddValue(data, "Rounded corners", chartObject, chartObject.GetRoundedCorners() ? "True" : "False");

    wxExcelChart chart = chartObject.GetChart();
    GetChartData(chart, data);
}

void ExcelSpy::GetChartData(wxExcelChart& chart, wxStringPairVector& data)
{
    AddValue(data, "Chart.Name", chart, chart.GetName());
    AddValue(data, "Chart.Type", chart, XlChartType_ToStr(chart.GetChartType()));

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
    AddValue(data, "Count", shapes, wxString::Format("%ld", shapes.GetCount()));
}

void ExcelSpy::GetShapeData(wxExcelShape& shape, wxStringPairVector& data)
{
    AddValue(data, "Name", shape, shape.GetName());
    AddValue(data, "Type", shape, MsoShapeType_ToStr(shape.GetType()));
    AddValue(data, "Visible", shape, shape.GetVisible() ? "True" : "False");
    AddValue(data, "Top", shape, wxString::Format("%g pt", shape.GetTop()));
    AddValue(data, "Left", shape, wxString::Format("%g pt", shape.GetLeft()));
    AddValue(data, "Width", shape, wxString::Format("%g pt", shape.GetWidth()));
    AddValue(data, "Height", shape, wxString::Format("%g pt", shape.GetHeight()));
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

// Threaded Comments
void ExcelSpy::GetCommentsThreadedData(wxExcelWorksheet& sheet, wxStringPairVector& data)
{
    wxAutoExcelObjectErrorModeOverrider emo(wxExcelObject::Err_DoNothing, true);

    wxExcelCommentsThreaded comments = sheet.GetCommentsThreaded();

    if ( !comments )
        return;

    const long count = comments.GetCount();

    for ( long l = 1; l <= count; l++ )
    {
        wxExcelCommentThreaded comment = comments[l];

        data.push_back( std::make_pair(wxString::Format("%s (%s)",
            comment.GetAuthor().GetName(), comment.GetDate().MakeString()),
            comment.Text()) );

        wxExcelCommentsThreaded replies = comment.GetReplies();
        const long replyCount = replies.GetCount();

        for ( long r = 1; r <= replyCount; r++ )
        {
            wxExcelCommentThreaded reply = replies[r];

            data.push_back( std::make_pair(wxString::Format("  %s (%s)",
            reply.GetAuthor().GetName(), reply.GetDate().MakeString()),
            reply.Text()) );
        }
    }
}
