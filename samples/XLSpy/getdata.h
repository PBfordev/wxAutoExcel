#ifndef _GETDATA_H
#define _GETDATA_H

#include <vector>
#include <utility>

#include <wx/wx.h>

#include <wx/wxAutoExcel.h>


typedef std::pair<wxString, wxString> wxStringPair;
typedef std::vector<wxStringPair> wxStringPairVector;


using namespace wxAutoExcel;

class ExcelSpy
{
public:
    static void GetApplicationData(wxExcelApplication& app, wxStringPairVector& data);
    static void GetInternationalData(wxExcelApplication& app, wxStringPairVector& data);
    static void GetRecentFilesData(wxExcelApplication& app, wxStringPairVector& data);
    static void GetAddInsData(wxExcelApplication& app, wxStringPairVector& data);
    static void GetAddIns2Data(wxExcelApplication& app, wxStringPairVector& data);
    
    static void GetWorkbookData(wxExcelApplication& app, wxExcelWorkbook& workbook, wxStringPairVector& data);
    static void GetDocumentPropertiesData(wxExcelDocumentProperties props, wxStringPairVector& data);
    static void GetStylesData(wxExcelWorkbook& workbook, wxStringPairVector& data);    
    static void GetNamesData(wxExcelWorkbook& workbook, wxStringPairVector& data);
    
    static void GetSheetsData(wxExcelSheets& sheets, wxStringPairVector& data);
    static void GetSheetData(wxExcelSheet& sheet, wxStringPairVector& data);
    
    static void GetWorksheetsData(wxExcelWorksheets& sheets, wxStringPairVector& data);
    static void GetWorksheetData(wxExcelWorksheet& sheet, wxStringPairVector& data);

    static void GetPageSetupData(wxExcelPageSetup& pageSetup, wxStringPairVector& data);    
    static void GetCommentsData(wxExcelWorksheet& sheet, wxStringPairVector& data);        
    
    static void GetRangeData(wxExcelRange& range, wxStringPairVector& data);
        
    static void GetOLEObjectsData(wxExcelOLEObjects& objects, wxStringPairVector& data);
    static void GetOLEObjectData(wxExcelOLEObject& object, wxStringPairVector& data);

    static void GetHyperlinksData(wxExcelHyperlinks& links, wxStringPairVector& data);
    static void GetHyperlinkData(wxExcelHyperlink& link, wxStringPairVector& data);

#if WXAUTOEXCEL_USE_CHARTS
    static void GetChartsData(wxExcelCharts& charts, wxStringPairVector& data);
    static void GetChartObjectsData(wxExcelChartObjects& chartObjects, wxStringPairVector& data);
    static void GetChartObjectData(wxExcelChartObject& chartObject, wxStringPairVector& data);
    static void GetChartData(wxExcelChart& chart, wxStringPairVector& data);
#endif // #if WXAUTOEXCEL_USE_CHARTS

#if WXAUTOEXCEL_USE_SHAPES
    static void GetShapesData(wxExcelShapes& shapes, wxStringPairVector& data);
    static void GetShapeData(wxExcelShape& shape, wxStringPairVector& data);
#endif // #if WXAUTOEXCEL_USE_SHAPES
};








#endif // #define _GETDATA_H