/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Modified by:
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
/////////////////////////////////////////////////////////////////////////////


#include "wx/wxAutoExcel_prec.h"

#include "wx/wxAutoExcelChart.h"

#if WXAUTOEXCEL_USE_CHARTS

#include <wx/dynarray.h>

#include "wx/wxAutoExcelAxes.h"
#include "wx/wxAutoExcelChartArea.h"
#include "wx/wxAutoExcelChartGroups.h"
#include "wx/wxAutoExcelChartObjects.h"
#include "wx/wxAutoExcelChartTitle.h"
#include "wx/wxAutoExcelDataTable.h"
#include "wx/wxAutoExcelFloor.h"
#include "wx/wxAutoExcelHyperlinks.h"
#include "wx/wxAutoExcelLegend.h"
#include "wx/wxAutoExcelOLEObjects.h"
#include "wx/wxAutoExcelPageSetup.h"
#include "wx/wxAutoExcelRange.h"
#include "wx/wxAutoExcelPlotArea.h"
#include "wx/wxAutoExcelSeries.h"
#include "wx/wxAutoExcelSeriesCollection.h"
#include "wx/wxAutoExcelShapes.h"
#include "wx/wxAutoExcelSheet.h"
#include "wx/wxAutoExcelTab.h"
#include "wx/wxAutoExcelWalls.h"
#include "wx/wxAutoExcelWorksheet.h"

#include "wx/wxAutoExcel_private.h"

namespace wxAutoExcel {

// ***** class wxExcelChart METHODS *****

void wxExcelChart::Activate()
{
    WXAUTOEXCEL_CALL_METHOD0_RET("Activate", "null");
}

void wxExcelChart::ApplyChartTemplate(const wxString& fileName)
{
    WXAUTOEXCEL_CALL_METHOD1_RET("ApplyChartTemplate", fileName, "null");
}

void wxExcelChart::ApplyDataLabels(XlDataLabelsType* type, wxXlTribool legendKey,
                                   wxXlTribool autoText, wxXlTribool hasLeaderLines,
                                   wxXlTribool showSeriesName, wxXlTribool showCategoryName,
                                   wxXlTribool showValue, wxXlTribool showPercentage,
                                   wxXlTribool showBubbleSize, const wxString& separator)
{
    wxVariantVector args;

    WXAUTOEXCEL_OPTIONALCPP_TO_OPTIONALVARIANT_NAME_VECTOR(Type, ((long*)type), args);
    WXAUTOEXCEL_OPTIONALCPPTBOOL_TO_OPTIONALVARIANT_NAME_VECTOR(LegendKey, legendKey, args);
    WXAUTOEXCEL_OPTIONALCPPTBOOL_TO_OPTIONALVARIANT_NAME_VECTOR(AutoText, autoText, args);
    WXAUTOEXCEL_OPTIONALCPPTBOOL_TO_OPTIONALVARIANT_NAME_VECTOR(HasLeaderLines, hasLeaderLines, args);
    WXAUTOEXCEL_OPTIONALCPPTBOOL_TO_OPTIONALVARIANT_NAME_VECTOR(ShowSeriesName, showSeriesName, args);
    WXAUTOEXCEL_OPTIONALCPPTBOOL_TO_OPTIONALVARIANT_NAME_VECTOR(ShowCategoryName, showCategoryName, args);
    WXAUTOEXCEL_OPTIONALCPPTBOOL_TO_OPTIONALVARIANT_NAME_VECTOR(ShowValue, showValue, args);
    WXAUTOEXCEL_OPTIONALCPPTBOOL_TO_OPTIONALVARIANT_NAME_VECTOR(ShowPercentage, showPercentage, args);
    WXAUTOEXCEL_OPTIONALCPPTBOOL_TO_OPTIONALVARIANT_NAME_VECTOR(ShowBubbleSize, showBubbleSize, args);
    WXAUTOEXCEL_OPTIONALCPPSTR_TO_OPTIONALVARIANT_NAME_VECTOR(Separator, separator, args);

    WXAUTOEXCEL_CALL_METHODARR_RET("ApplyDataLabels", args, "null");
}


void wxExcelChart::ApplyLayout(long layout, XlChartType* chartType)
{
    wxVariant vLayout(layout, wxS("Layout"));
    WXAUTOEXCEL_OPTIONALCPP_TO_OPTIONALVARIANT_NAME(ChartType, ((long*)chartType));

    WXAUTOEXCEL_CALL_METHOD2_RET("ApplyLayout", vLayout, vChartType, "null");
}

wxExcelAxis wxExcelChart::Axes(XlAxisType type,  XlAxisGroup* axisGroup)
{    
    wxVariant vType((long)type, wxS("Type"));
    WXAUTOEXCEL_OPTIONALCPP_TO_OPTIONALVARIANT_NAME(AxisGroup, ((long*)axisGroup));

    wxExcelAxis axis;
    WXAUTOEXCEL_CALL_METHOD2_OBJECT("Axes", vType, vAxisGroup, axis);
}

wxExcelChartGroups wxExcelChart::ChartGroups()
{
    wxExcelChartGroups groups;
    WXAUTOEXCEL_CALL_METHOD0_OBJECT("ChartGroups", groups);
}

wxExcelChartGroup wxExcelChart::ChartGroups(long index)
{
    wxExcelChartGroup group; 
    WXAUTOEXCEL_CALL_METHOD1_OBJECT("ChartGroups", index, group);
}

wxExcelChartObjects wxExcelChart::ChartObjects()
{
    wxExcelChartObjects objects;

    WXAUTOEXCEL_CALL_METHOD0_OBJECT("ChartObjects", objects);
}

wxExcelChartObjects wxExcelChart::ChartObjects(const wxVector<long>& indices)
{
    wxExcelChartObjects objects;

    wxCHECK(indices.size() > 0, objects);

    wxVariant vIndices;

    vIndices.NullList();
    for (size_t i = 0; i < indices.size(); i++)
        vIndices.Append(indices[i]);

    WXAUTOEXCEL_CALL_METHOD1_OBJECT("ChartObjects", vIndices, objects);
}

wxExcelChartObjects wxExcelChart::ChartObjects(const wxArrayString& names)
{
    wxExcelChartObjects objects;

    wxCHECK(names.size() > 0, objects);

    wxVariant vNames;

    vNames.NullList();
    for (size_t i = 0; i < names.size(); i++)
        vNames.Append(names[i]);

    WXAUTOEXCEL_CALL_METHOD1_OBJECT("ChartObjects", vNames, objects);
}


void wxExcelChart::ChartWizard(wxExcelRange* source, XlChartType* gallery, long* format,
                               XlRowCol* plotBy, long* categoryLabels, long* seriesLabels, 
                               wxXlTribool hasLegend, const wxString& title, const wxString& categoryTitle, 
                               const wxString& valueTitle, const wxString& extraTitle)
{    
    wxVariant vSource;
    
    if ( source )
    {
        if ( !ObjectToVariant(source, vSource, wxS("Source")) )
            return;
    }
    
    wxVariantVector args;
    
    args.push_back(vSource);
    WXAUTOEXCEL_OPTIONALCPP_TO_OPTIONALVARIANT_NAME_VECTOR(Gallery, ((long*)gallery), args);
    WXAUTOEXCEL_OPTIONALCPP_TO_OPTIONALVARIANT_NAME_VECTOR(Format, format, args);
    WXAUTOEXCEL_OPTIONALCPP_TO_OPTIONALVARIANT_NAME_VECTOR(PlotBy, ((long*)plotBy), args);
    WXAUTOEXCEL_OPTIONALCPP_TO_OPTIONALVARIANT_NAME_VECTOR(CategoryLabels, categoryLabels, args);
    WXAUTOEXCEL_OPTIONALCPP_TO_OPTIONALVARIANT_NAME_VECTOR(SeriesLabels, seriesLabels, args);
    WXAUTOEXCEL_OPTIONALCPPTBOOL_TO_OPTIONALVARIANT_NAME_VECTOR(HasLegend, hasLegend, args);
    WXAUTOEXCEL_OPTIONALCPPSTR_TO_OPTIONALVARIANT_NAME_VECTOR(Title, title, args);
    WXAUTOEXCEL_OPTIONALCPPSTR_TO_OPTIONALVARIANT_NAME_VECTOR(CategoryTitle, categoryTitle, args);
    WXAUTOEXCEL_OPTIONALCPPSTR_TO_OPTIONALVARIANT_NAME_VECTOR(ValueTitle, valueTitle, args);
    WXAUTOEXCEL_OPTIONALCPPSTR_TO_OPTIONALVARIANT_NAME_VECTOR(ExtraTitle, extraTitle, args);
    
    WXAUTOEXCEL_CALL_METHODARR_RET("ChartWizard", args, "null");
}

void wxExcelChart::CheckSpelling(const wxString& customDictionary, wxXlTribool ignoreUpperCase,
                                 wxXlTribool alwaysSuggest, MsoLanguageID* spellLang)
{
    WXAUTOEXCEL_OPTIONALCPPSTR_TO_OPTIONALVARIANT_NAME(CustomDictionary, customDictionary);
    WXAUTOEXCEL_OPTIONALCPPTBOOL_TO_OPTIONALVARIANT_NAME(IgnoreUpperCase, ignoreUpperCase);
    WXAUTOEXCEL_OPTIONALCPPTBOOL_TO_OPTIONALVARIANT_NAME(AlwaysSuggest, alwaysSuggest);
    WXAUTOEXCEL_OPTIONALCPP_TO_OPTIONALVARIANT_NAME(SpellLang, ((long*)spellLang));

    WXAUTOEXCEL_CALL_METHOD4_RET("CheckSpelling", vCustomDictionary, vIgnoreUpperCase, vAlwaysSuggest, vSpellLang, "null");
}


void wxExcelChart::ClearToMatchStyle()
{
    WXAUTOEXCEL_CALL_METHOD0_RET("ClearToMatchStyle", "null");
}

void wxExcelChart::CopyPicture(XlPictureAppearance* appearance, XlCopyPictureFormat* format)
{
    WXAUTOEXCEL_OPTIONALCPP_TO_OPTIONALVARIANT_NAME(Appearance, ((long*)appearance));
    WXAUTOEXCEL_OPTIONALCPP_TO_OPTIONALVARIANT_NAME(Format, ((long*)format));

    WXAUTOEXCEL_CALL_METHOD2_RET("CopyPicture", vAppearance, vFormat, "null");    
}


void wxExcelChart::Delete()
{
    WXAUTOEXCEL_CALL_METHOD0_RET("Delete", "null");
}

void wxExcelChart::Deselect()
{
    WXAUTOEXCEL_CALL_METHOD0_RET("Deselect", "null");
}


bool wxExcelChart::Export(const wxString& fileName, const wxString& filterName, wxXlTribool interactive)
{
    wxASSERT ( !fileName.empty() );
    
    WXAUTOEXCEL_OPTIONALCPPSTR_TO_OPTIONALVARIANT_NAME(FileName, fileName);
    WXAUTOEXCEL_OPTIONALCPPSTR_TO_OPTIONALVARIANT_NAME(FilterName, filterName);
    WXAUTOEXCEL_OPTIONALCPPTBOOL_TO_OPTIONALVARIANT_NAME(Interactive, interactive);

    WXAUTOEXCEL_CALL_METHOD3("Export", vFileName, vFilterName, vInteractive, "bool", false);
    return vResult.GetBool();
}


wxExcelChart wxExcelChart::Location(XlChartLocation where, const wxString& name)
{
    wxVariant vWhere((long)where, wxS("Where"));
    WXAUTOEXCEL_OPTIONALCPPSTR_TO_OPTIONALVARIANT_NAME(Name, name);

    wxExcelChart chart;
    WXAUTOEXCEL_CALL_METHOD2_OBJECT("Location", vWhere, vName, chart);
}

bool wxExcelChart::Move()
{
    WXAUTOEXCEL_CALL_METHOD0_BOOL("Move");
}

bool wxExcelChart::MoveAfterOrBefore(wxExcelSheet sheetAfterOrBefore, bool after)
{
    return DoOrderedCopyOrMove(false, sheetAfterOrBefore, after);
}

bool wxExcelChart::DoOrderedCopyOrMove(bool copy, wxExcelSheet sheetAfterOrBefore, bool after)
{
    wxVariant vAfterOrBefore;

    if ( ObjectToVariant(&sheetAfterOrBefore, vAfterOrBefore) )
    {
        vAfterOrBefore.SetName(after? wxS("After") : wxS("Before"));       
        WXAUTOEXCEL_CALL_METHOD1_BOOL(copy ? "Copy" : "Move", vAfterOrBefore);
    }
    return false;
}

wxExcelOLEObjects wxExcelChart::OLEObjects()
{
    wxExcelOLEObjects objects;
    WXAUTOEXCEL_CALL_METHOD0_OBJECT("OLEObjects", objects);
}

void wxExcelChart::Paste(XlPasteType* type)
{
    WXAUTOEXCEL_OPTIONALCPP_TO_OPTIONALVARIANT_NAME(Type, ((long*)type));

    WXAUTOEXCEL_CALL_METHOD1_RET("Paste", vType, "null");    
}

bool wxExcelChart::PrintOut(long* from, long* to, long* copies, wxXlTribool preview, const wxString& activePrinter,
                            wxXlTribool printToFile, wxXlTribool collate, const wxString& prToFileName)
{
    wxVariantVector args;

    WXAUTOEXCEL_OPTIONALCPP_TO_OPTIONALVARIANT_NAME_VECTOR(From, from, args);
    WXAUTOEXCEL_OPTIONALCPP_TO_OPTIONALVARIANT_NAME_VECTOR(To, to, args);
    WXAUTOEXCEL_OPTIONALCPP_TO_OPTIONALVARIANT_NAME_VECTOR(Copies, copies, args);
    WXAUTOEXCEL_OPTIONALCPPTBOOL_TO_OPTIONALVARIANT_NAME_VECTOR(Preview, preview, args);
    WXAUTOEXCEL_OPTIONALCPPSTR_TO_OPTIONALVARIANT_NAME_VECTOR(ActivePrinter, activePrinter, args);
    WXAUTOEXCEL_OPTIONALCPPTBOOL_TO_OPTIONALVARIANT_NAME_VECTOR(PrintToFile, printToFile, args);
    WXAUTOEXCEL_OPTIONALCPPTBOOL_TO_OPTIONALVARIANT_NAME_VECTOR(Collate, collate, args);
    WXAUTOEXCEL_OPTIONALCPPSTR_TO_OPTIONALVARIANT_NAME_VECTOR(PrToFileName, prToFileName, args);

    return PrintOut(args);
}

bool wxExcelChart::PrintOut(const wxVariantVector& args)
{
    WXAUTOEXCEL_CALL_METHODARR("PrintOut", args, "bool", false);
    return vResult.GetBool();
}

bool wxExcelChart::PrintPreview(wxXlTribool enableChanges)
{
    WXAUTOEXCEL_OPTIONALCPPTBOOL_TO_OPTIONALVARIANT_NAME(EnableChanges, enableChanges);

    WXAUTOEXCEL_CALL_METHOD1_BOOL("PrintPreview", vEnableChanges);
}

void wxExcelChart::Protect(const wxString& password, wxXlTribool shapes, 
                           wxXlTribool contents, wxXlTribool userInterfaceOnly)
{
    WXAUTOEXCEL_OPTIONALCPPSTR_TO_OPTIONALVARIANT_NAME(Password, password);
    WXAUTOEXCEL_OPTIONALCPPTBOOL_TO_OPTIONALVARIANT_NAME(Shapes, shapes);
    WXAUTOEXCEL_OPTIONALCPPTBOOL_TO_OPTIONALVARIANT_NAME(Contents, contents);
    WXAUTOEXCEL_OPTIONALCPPTBOOL_TO_OPTIONALVARIANT_NAME(UserInterfaceOnly, userInterfaceOnly);   
}

void wxExcelChart::Refresh()
{
    WXAUTOEXCEL_CALL_METHOD0_RET("Refresh", "null");
}

void wxExcelChart::SaveAs(const wxString& fileName, XlFileFormat* fileFormat,
                const wxString& password, const wxString& writeResPassword,
                wxXlTribool readOnlyRecommended, wxXlTribool createBackup,                
                wxXlTribool addToMru, wxXlTribool local)
{
    wxVariantVector args;

    WXAUTOEXCEL_OPTIONALCPPSTR_TO_OPTIONALVARIANT_NAME_VECTOR(FileName, fileName, args);
    WXAUTOEXCEL_OPTIONALCPP_TO_OPTIONALVARIANT_NAME_VECTOR(FileFormat, ((long*)fileFormat), args);
    WXAUTOEXCEL_OPTIONALCPPSTR_TO_OPTIONALVARIANT_NAME_VECTOR(Password, password, args);
    WXAUTOEXCEL_OPTIONALCPPSTR_TO_OPTIONALVARIANT_NAME_VECTOR(WriteResPassword, writeResPassword, args);
    WXAUTOEXCEL_OPTIONALCPPTBOOL_TO_OPTIONALVARIANT_NAME_VECTOR(ReadOnlyRecommended, readOnlyRecommended, args);
    WXAUTOEXCEL_OPTIONALCPPTBOOL_TO_OPTIONALVARIANT_NAME_VECTOR(CreateBackup, createBackup, args);    
    WXAUTOEXCEL_OPTIONALCPPTBOOL_TO_OPTIONALVARIANT_NAME_VECTOR(AddToMru, addToMru, args);
    WXAUTOEXCEL_OPTIONALCPPTBOOL_TO_OPTIONALVARIANT_NAME_VECTOR(Local, local, args);

    SaveAs(args);
}

void wxExcelChart::SaveAs(const wxVariantVector& optionalArgs)
{
    WXAUTOEXCEL_CALL_METHODARR_RET("SaveAs", optionalArgs, "null");    
}

void wxExcelChart::SaveChartTemplate(const wxString& fileName)
{
    WXAUTOEXCEL_CALL_METHOD1_RET("SaveChartTemplate", fileName, "null");
}

void wxExcelChart::Select(wxXlTribool replace)
{
    WXAUTOEXCEL_OPTIONALCPPTBOOL_TO_OPTIONALVARIANT_NAME(Replace, replace);
    WXAUTOEXCEL_CALL_METHOD1_RET("Select", vReplace, "null");
}


wxExcelSeriesCollection wxExcelChart::SeriesCollection()
{
    wxExcelSeriesCollection series;
    WXAUTOEXCEL_CALL_METHOD0_OBJECT("SeriesCollection", series);
}

wxExcelSeries wxExcelChart::SeriesCollection(long index)
{
    wxExcelSeries series;
    WXAUTOEXCEL_CALL_METHOD1_OBJECT("SeriesCollection", index, series);
}

wxExcelSeries wxExcelChart::SeriesCollection(const wxString& name)
{
    wxExcelSeries series;
    WXAUTOEXCEL_CALL_METHOD1_OBJECT("SeriesCollection", name, series);
}


void wxExcelChart::SetBackgroundPicture(const wxString& fileName)
{
     WXAUTOEXCEL_CALL_METHOD1_RET("SetBackgroundPicture", fileName, "null");
}

void wxExcelChart::SetDefaultChart(const wxString& name)
{    
    wxVariant vName;

    if ( name.empty() )
        vName = name;
    else
        vName = (long)xlBuiltIn;    
        
    WXAUTOEXCEL_CALL_METHOD1_RET("SetDefaultChart", vName, "null");
}

void wxExcelChart::SetElement(MsoChartElementType element)
{
    WXAUTOEXCEL_CALL_METHOD1_RET("SetElement", (long)element, "null");
}

void wxExcelChart::SetSourceData(wxExcelRange source, XlRowCol* plotBy)
{
    wxVariant vSource;

    if ( ObjectToVariant(&source, vSource, wxS("Source")) )
    {     
        WXAUTOEXCEL_OPTIONALCPP_TO_OPTIONALVARIANT_NAME(PlotBy, ((long*)plotBy));
        WXAUTOEXCEL_CALL_METHOD2_RET("SetSourceData", vSource, vPlotBy, "null");
    }

}

void wxExcelChart::Unprotect(const wxString& password )
{    
    WXAUTOEXCEL_OPTIONALCPPSTR_TO_OPTIONALVARIANT_NAME(Password, password);
    WXAUTOEXCEL_CALL_METHOD1_RET("Unprotect", vPassword, "null");
}

// ***** class wxExcelChart PROPERTIES *****

bool wxExcelChart::GetAutoScaling()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("AutoScaling");
}

void wxExcelChart::SetAutoScaling(bool autoScaling)
{
    InvokePutProperty(wxS("AutoScaling"), autoScaling);
}

wxExcelWalls wxExcelChart::GetBackWall()
{
    wxExcelWalls walls;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("BackWall", walls);
}

XlBarShape wxExcelChart::GetBarShape()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("BarShape", XlBarShape, xlBox);
}

void wxExcelChart::SetBarShape(XlBarShape barShape)
{
    InvokePutProperty(wxS("BarShape"), (long)barShape);
}

wxExcelChartArea wxExcelChart::GetChartArea()
{
    wxExcelChartArea chartArea;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("ChartArea", chartArea);
}

long wxExcelChart::GetChartStyle()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("ChartStyle");
}

void wxExcelChart::SetChartStyle(long chartStyle)
{
    InvokePutProperty(wxS("ChartStyle"), chartStyle);
}

wxExcelChartTitle wxExcelChart::GetChartTitle()
{
    wxExcelChartTitle chartTitle;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("ChartTitle", chartTitle);
}

XlChartType wxExcelChart::GetChartType()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("ChartType", XlChartType, xlArea);
}

void wxExcelChart::SetChartType(XlChartType chartType)
{
    InvokePutProperty(wxS("ChartType"), (long)chartType);
}

wxString wxExcelChart::GetCodeName()
{
    WXAUTOEXCEL_PROPERTY_STRING_GET0("CodeName");
}

wxExcelDataTable wxExcelChart::GetDataTable()
{
    wxExcelDataTable table;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("DataTable", table);
}


long wxExcelChart::GetDepthPercent()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("DepthPercent");
}

void wxExcelChart::SetDepthPercent(long depthPercent)
{
    InvokePutProperty(wxS("DepthPercent"), depthPercent);
}

XlDisplayBlanksAs wxExcelChart::GetDisplayBlanksAs()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("DisplayBlanksAs", XlDisplayBlanksAs,  xlZero);
}

void wxExcelChart::SetDisplayBlanksAs(XlDisplayBlanksAs displayBlanksAs)
{
    InvokePutProperty(wxS("DisplayBlanksAs"), (long)displayBlanksAs);
}

long wxExcelChart::GetElevation()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("Elevation");
}

void wxExcelChart::SetElevation(long elevation)
{
    InvokePutProperty(wxS("Elevation"), elevation);
}

wxExcelFloor wxExcelChart::GetFloor()
{
    wxExcelFloor floor;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("Floor", floor);
}

long wxExcelChart::GetGapDepth()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("GapDepth");
}

void wxExcelChart::SetGapDepth(long gapDepth)
{
    InvokePutProperty(wxS("GapDepth"), gapDepth);
}

bool wxExcelChart::GetHasAxis(XlAxisType index1,  XlAxisGroup* index2)
{    
    wxVariant vIndex1((long)index1, wxS("Index1"));
    WXAUTOEXCEL_OPTIONALCPP_TO_OPTIONALVARIANT_NAME(Index2, ((long*)index2));

    WXAUTOEXCEL_PROPERTY_BOOL_GET2("HasAxis", vIndex1, vIndex2);
}

void wxExcelChart::SetHasAxis(bool hasAxis, XlAxisType index1,  XlAxisGroup* index2)
{
    wxVariant vIndex1((long)index1, wxS("Index1"));
    WXAUTOEXCEL_OPTIONALCPP_TO_OPTIONALVARIANT_NAME(Index2, ((long*)index2));

    InvokePutProperty(wxS("HasAxis"), hasAxis, vIndex1, vIndex2);  
}

bool wxExcelChart::GetHasDataTable()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("HasDataTable");
}

void wxExcelChart::SetHasDataTable(bool hasDataTable)
{
    InvokePutProperty(wxS("HasDataTable"), hasDataTable);
}

bool wxExcelChart::GetHasLegend()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("HasLegend");
}

void wxExcelChart::SetHasLegend(bool hasLegend)
{
    InvokePutProperty(wxS("HasLegend"), hasLegend);
}

bool wxExcelChart::GetHasTitle()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("HasTitle");
}

void wxExcelChart::SetHasTitle(bool hasTitle)
{
    InvokePutProperty(wxS("HasTitle"), hasTitle);
}

long wxExcelChart::GetHeightPercent()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("HeightPercent");
}

void wxExcelChart::SetHeightPercent(long heightPercent)
{
    InvokePutProperty(wxS("HeightPercent"), heightPercent);
}

wxExcelHyperlinks wxExcelChart::GetHyperlinks()
{
    wxExcelHyperlinks hyperlinks;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("Hyperlinks", hyperlinks);
}

long wxExcelChart::GetIndex()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("Index");
}

wxExcelLegend wxExcelChart::GetLegend()
{
    wxExcelLegend legend;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("Legend", legend);
}

wxString wxExcelChart::GetName()
{
    WXAUTOEXCEL_PROPERTY_STRING_GET0("Name");
}

void wxExcelChart::SetName(const wxString& name)
{
    InvokePutProperty(wxS("Name"), name);
}

wxExcelWorksheet wxExcelChart::GetNext()
{
    wxExcelWorksheet worksheet;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("Next", worksheet);
}

wxExcelPageSetup wxExcelChart::GetPageSetup()
{
    wxExcelPageSetup pageSetup;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("PageSetup", pageSetup);
}


long wxExcelChart::GetPerspective()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("Perspective");
}

void wxExcelChart::SetPerspective(long perspective)
{
    InvokePutProperty(wxS("Perspective"), perspective);
}

wxExcelPlotArea wxExcelChart::GetPlotArea()
{
    wxExcelPlotArea plotArea;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("PlotArea", plotArea);
}

XlRowCol wxExcelChart::GetPlotBy()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("PlotBy", XlRowCol, xlRows);
}

void wxExcelChart::SetPlotBy(XlRowCol plotBy)
{
    InvokePutProperty(wxS("PlotBy"), (long)plotBy);
}

bool wxExcelChart::GetPlotVisibleOnly()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("PlotVisibleOnly");
}

void wxExcelChart::SetPlotVisibleOnly(bool plotVisibleOnly)
{
    InvokePutProperty(wxS("PlotVisibleOnly"), plotVisibleOnly);
}

wxExcelWorksheet wxExcelChart::GetPrevious()
{
    wxExcelWorksheet worksheet;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("Previous", worksheet);
}

bool wxExcelChart::GetProtectContents()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("ProtectContents");
}

bool wxExcelChart::GetProtectData()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("ProtectData");
}

void wxExcelChart::SetProtectData(bool protectData)
{
    InvokePutProperty(wxS("ProtectData"), protectData);
}

bool wxExcelChart::GetProtectDrawingObjects()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("ProtectDrawingObjects");
}

bool wxExcelChart::GetProtectFormatting()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("ProtectFormatting");
}

void wxExcelChart::SetProtectFormatting(bool protectFormatting)
{
    InvokePutProperty(wxS("ProtectFormatting"), protectFormatting);
}

bool wxExcelChart::GetProtectionMode()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("ProtectionMode");
}

bool wxExcelChart::GetProtectSelection()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("ProtectSelection");
}

void wxExcelChart::SetProtectSelection(bool protectSelection)
{
    InvokePutProperty(wxS("ProtectSelection"), protectSelection);
}

bool wxExcelChart::GetRightAngleAxes()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("RightAngleAxes");
}

void wxExcelChart::SetRightAngleAxes(bool rightAngleAxes)
{
    InvokePutProperty(wxS("RightAngleAxes"), rightAngleAxes);
}

long wxExcelChart::GetRotation()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("Rotation");
}

void wxExcelChart::SetRotation(long rotation)
{
    InvokePutProperty(wxS("Rotation"), rotation);
}

#if WXAUTOEXCEL_USE_SHAPES

wxExcelShapes wxExcelChart::GetShapes()
{
    wxExcelShapes shapes;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("Shapes", shapes);
}

#endif // #if WXAUTOEXCEL_USE_SHAPES

bool wxExcelChart::GetShowDataLabelsOverMaximum()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("ShowDataLabelsOverMaximum");
}

void wxExcelChart::SetShowDataLabelsOverMaximum(bool showDataLabelsOverMaximum)
{
    InvokePutProperty(wxS("ShowDataLabelsOverMaximum"), showDataLabelsOverMaximum);
}

wxExcelWalls wxExcelChart::GetSideWall()
{
    wxExcelWalls walls;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("SideWall", walls);
}

wxExcelTab wxExcelChart::GetTab()
{
    wxExcelTab tab;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("Tab", tab);
}

XlSheetVisibility wxExcelChart::GetVisible()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("Visible", XlSheetVisibility, xlSheetHidden);
}

void wxExcelChart::SetVisible(XlSheetVisibility visible)
{
    InvokePutProperty(wxS("Visible"), (long)visible);
}

wxExcelWalls wxExcelChart::GetWalls()
{
    wxExcelWalls walls;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("Walls", walls);
}

} // namespace wxAutoExcel

#endif // #if WXAUTOEXCEL_USE_CHARTS
