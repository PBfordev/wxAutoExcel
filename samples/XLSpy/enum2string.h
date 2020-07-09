#ifndef _ENUM2STRING_H
#define _ENUM2STRING_H

#include <wx/wx.h>

#include <wx/wxAutoExcel.h>


using namespace wxAutoExcel;

// contains function converting enum values
// to human readable strings


wxString XlMeasurementUnits_ToStr(XlMeasurementUnits val);
wxString XlFileFormat_ToStr(XlFileFormat val);
wxString XlChartType_ToStr(XlChartType val);
wxString XlSheetVisibility_ToStr(XlSheetVisibility val);
wxString XlPaperSize_ToStr(XlPaperSize val);

wxString MsoAutomationSecurity_ToStr(MsoAutomationSecurity val);
wxString MsoShapeType_ToStr(MsoShapeType val);

wxString XlListObjectSourceType_ToStr(XlListObjectSourceType val);


#endif // _ENUM2STRING_H