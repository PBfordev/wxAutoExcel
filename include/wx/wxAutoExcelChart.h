/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// License:     MIT license
/////////////////////////////////////////////////////////////////////////////


#ifndef _WXAUTOEXCEL_CHART_H
#define _WXAUTOEXCEL_CHART_H

#include "wx/wxAutoExcel_defs.h"

#if WXAUTOEXCEL_USE_CHARTS

#include "wx/wxAutoExcel_object.h"
#include "wx/wxAutoExcel_enums.h"

class wxArrayString;

namespace wxAutoExcel {


    /**
    @brief Represents Microsoft Excel Chart object.
    */
   class WXDLLIMPEXP_WXAUTOEXCEL wxExcelChart : public wxExcelObject
    {
    public:
        // ***** METHODS *****

        /**
        Makes the current chart the active chart.

        [MSDN documentation for Chart.Activate](https://learn.microsoft.com/en-us/office/vba/api/excel.chart.activate%28method%29)).
        */
        void Activate();

        /**
        Applies a standard or custom chart type to a chart.

        [MSDN documentation for Chart.ApplyChartTemplate](https://learn.microsoft.com/en-us/office/vba/api/excel.chart.applycharttemplate).
        */
        void ApplyChartTemplate(const wxString& fileName);

        /**
        Applies data labels to all the series in a chart.

        [MSDN documentation for Chart.ApplyDataLabels](https://learn.microsoft.com/en-us/office/vba/api/excel.chart.applydatalabels).
        */
        void ApplyDataLabels(XlDataLabelsType* type = NULL, wxXlTribool legendKey = wxDefaultXlTribool,
                             wxXlTribool autoText = wxDefaultXlTribool, wxXlTribool hasLeaderLines = wxDefaultXlTribool,
                             wxXlTribool showSeriesName = wxDefaultXlTribool, wxXlTribool showCategoryName = wxDefaultXlTribool,
                             wxXlTribool showValue = wxDefaultXlTribool, wxXlTribool showPercentage = wxDefaultXlTribool,
                             wxXlTribool showBubbleSize = wxDefaultXlTribool, const wxString& separator = wxEmptyString);

        /**
        Applies the layouts shown in the ribbon.

        [MSDN documentation for Chart.ApplyLayout](https://learn.microsoft.com/en-us/office/vba/api/excel.chart.applylayout).
        */
        void ApplyLayout(long layout, XlChartType* chartType = NULL);

        //@{
        /**
        Returns an object that represents either a single axis or a collection of the axes on the chart.

        [MSDN documentation for Chart.Axes](https://learn.microsoft.com/en-us/office/vba/api/excel.chart.axes).
        */
        wxExcelAxes Axes();
        wxExcelAxis Axes(XlAxisType type,  XlAxisGroup* axisGroup = NULL);
        //@}

        //@{
        /**
        Returns an object that represents either a single chart group (a ChartGroup object) or a collection of all the chart groups in the chart (a ChartGroups object). The returned collection includes every type of group.

        [MSDN documentation for Chart.ChartGroups](https://learn.microsoft.com/en-us/office/vba/api/excel.chart.chartgroups).
        */
        wxExcelChartGroups ChartGroups();
        wxExcelChartGroup ChartGroups(long index);
        //@}

        //@{
        /**
        Returns an object that represents either a single embedded chart (a ChartObject object) or a collection of all the embedded charts (a ChartObjects object) on the sheet.

        [MSDN documentation for Chart.ChartObjects](https://learn.microsoft.com/en-us/office/vba/api/excel.chart.chartobjects).
        */
        wxExcelChartObjects ChartObjects();
        wxExcelChartObjects ChartObjects(const wxVector<long>& indices);
        wxExcelChartObjects ChartObjects(const wxArrayString& names);
        wxExcelChartObject ChartObjects(long index);
        wxExcelChartObject ChartObjects(const wxString& name);
        //@}

        /**
        Modifies the properties of the given chart. You can use this method to quickly format a chart without setting all the individual properties. This method is noninteractive, and it changes only the specified properties.

        [MSDN documentation for Chart.ChartWizard](https://learn.microsoft.com/en-us/office/vba/api/excel.chart.chartwizard).
        */
        void ChartWizard(wxExcelRange* source = NULL, XlChartType* gallery = NULL,
                         long* format = NULL, XlRowCol* plotBy = NULL,
                         long* categoryLabels = NULL, long* seriesLabels = NULL,
                         wxXlTribool hasLegend = wxDefaultXlTribool,
                         const wxString& title = wxEmptyString, const wxString& categoryTitle = wxEmptyString,
                         const wxString& valueTitle = wxEmptyString, const wxString& extraTitle = wxEmptyString);

        /**
        Checks the spelling of an object.

        [MSDN documentation for Chart.CheckSpelling](https://learn.microsoft.com/en-us/office/vba/api/excel.chart.checkspelling).
        */
        void CheckSpelling(const wxString& customDictionary = wxEmptyString, wxXlTribool ignoreUpperCase = wxDefaultXlTribool,
                           wxXlTribool alwaysSuggest = wxDefaultXlTribool, MsoLanguageID* spellLang = NULL);

        /**
        lears all colors on the specified chart that do not follow the color style applied to the chart.

        [MSDN documentation for Chart.ClearToMatchColorStyle](https://learn.microsoft.com/en-us/office/vba/api/excel.chart.cleartomatchcolorstyle).
        */
        void ClearToMatchColorStyle();

        /**
        Clears the chart elements formatting to automatic.

        [MSDN documentation for Chart.ClearToMatchStyle](https://learn.microsoft.com/en-us/office/vba/api/excel.chart.cleartomatchstyle).
        */
        void ClearToMatchStyle();

        /**
        Copies the selected object to the Clipboard as a picture.

        [MSDN documentation for Chart.CopyPicture](https://learn.microsoft.com/en-us/office/vba/api/excel.chart.copypicture).
        */
        void CopyPicture(XlPictureAppearance* appearance, XlCopyPictureFormat* format);

        /**
        Deletes the object.

        [MSDN documentation for Chart.Delete](https://learn.microsoft.com/en-us/office/vba/api/excel.chart.delete).
        */
        void Delete();

        /**
        Cancels the selection for the specified chart.

        [MSDN documentation for Chart.Deselect](https://learn.microsoft.com/en-us/previous-versions/office/developer/office-2007/bb223343%28v=office.12%29).
        */
        void Deselect();

        /**
        Exports the chart in a graphic format.

        [MSDN documentation for Chart.Export](https://learn.microsoft.com/en-us/office/vba/api/excel.chart.export).
        */
        bool Export(const wxString& fileName, const wxString& filterName = wxEmptyString,
                    wxXlTribool interactive = wxDefaultXlTribool);

        //@{
        /**
        Returns an object that represents either a single series (a Series object) or a collection of all the series (a SeriesCollection collection) in the chart or chart group.

        [MSDN documentation for Chart.FullSeriesCollection](https://learn.microsoft.com/en-us/office/vba/api/excel.chart.fullseriescollection).
        */
        wxExcelSeriesCollection FullSeriesCollection();
        wxExcelSeries FullSeriesCollection(long index);
        //@}


        /**
        Moves the chart to a new location.

        [MSDN documentation for Chart.Location](https://learn.microsoft.com/en-us/office/vba/api/excel.chart.location).
        */
        wxExcelChart Location(XlChartLocation where, const wxString& name = wxEmptyString);

        /**
            Moves the chart to the new workbook.
            [MSDN documentation for Chart.Move](https://learn.microsoft.com/en-us/office/vba/api/excel.chart.move).
        */
        bool Move();
        /**
            Moves the chart within the same workbook, after or before the specified sheet.
            [MSDN documentation for Chart.Move](https://learn.microsoft.com/en-us/office/vba/api/excel.chart.move).
        */
        bool MoveAfterOrBefore(wxExcelSheet sheetAfterOrBefore, bool after);

        //@{
        /**
        Returns an object that represents either a single OLE object (an OLEObject ) or a collection of all OLE objects (an OLEObjects collection) on the chart or sheet. Read-only.

        [MSDN documentation for Chart.OLEObjects](https://learn.microsoft.com/en-us/office/vba/api/excel.chart.oleobjects).
        */
        wxExcelOLEObjects OLEObjects();
        wxExcelOLEObject OLEObjects(long index);
        wxExcelOLEObject OLEObjects(const wxString& name);
        //@}

        /**
        Pastes chart data from the Clipboard into the specified chart.

        [MSDN documentation for Chart.Paste](https://learn.microsoft.com/en-us/office/vba/api/excel.chart.paste).
        */
        void Paste(XlPasteType* type = NULL);

        //@{
        /**
        Prints the object.

        [MSDN documentation for Chart.PrintOut](https://learn.microsoft.com/en-us/office/vba/api/excel.chart.printout).
        */
        bool PrintOut(long* from = NULL, long* to = NULL, long* copies = NULL, wxXlTribool preview = wxDefaultXlTribool,
                      const wxString& activePrinter = wxEmptyString, wxXlTribool printToFile = wxDefaultXlTribool,
                      wxXlTribool collate = wxDefaultXlTribool, const wxString& prToFileName= wxEmptyString);
        bool PrintOut(const wxVariantVector& args);
        //@}

        /**
        Shows a preview of the object as it would look when printed.

        [MSDN documentation for Chart.PrintPreview](https://learn.microsoft.com/en-us/office/vba/api/excel.chart.printpreview).
        */
        bool PrintPreview(wxXlTribool enableChanges = wxDefaultXlTribool);

        /**
        Protects a chart so that it cannot be modified.

        [MSDN documentation for Chart.Protect](https://learn.microsoft.com/en-us/office/vba/api/excel.chart.protect).
        */
        void Protect(const wxString& password = wxEmptyString, wxXlTribool drawingObjects = wxDefaultXlTribool, wxXlTribool contents = wxDefaultXlTribool, wxXlTribool userInterfaceOnly = wxDefaultXlTribool);

        /**
        Causes the specified chart to be redrawn immediately.

        [MSDN documentation for Chart.Refresh](https://learn.microsoft.com/en-us/office/vba/api/excel.chart.refresh).
        */
        void Refresh();

        //@{
        /**
        Saves changes to the chart in a different file.

        [MSDN documentation for Chart.SaveAs](https://learn.microsoft.com/en-us/office/vba/api/excel.chart.saveas).
        */
         void SaveAs(const wxString& fileName = wxEmptyString, XlFileFormat* fileFormat = NULL,
                     const wxString& password = wxEmptyString, const wxString& writeResPassword = wxEmptyString,
                     wxXlTribool readOnlyRecommended = wxDefaultXlTribool, wxXlTribool createBackup = wxDefaultXlTribool,
                     wxXlTribool addToMru = wxDefaultXlTribool, wxXlTribool local = wxDefaultXlTribool);

        void SaveAs(const wxVariantVector& optionalArgs);
        //@}

        /**
        Saves a custom chart template to the list of available chart templates.

        [MSDN documentation for Chart.SaveChartTemplate](https://learn.microsoft.com/en-us/office/vba/api/excel.chart.savecharttemplate).
        */
        void SaveChartTemplate(const wxString& fileName);

        /**
        Selects the object.

        [MSDN documentation for Chart.Select](https://learn.microsoft.com/en-us/office/vba/api/excel.chart.select%28method%29)).
        */
        void Select(wxXlTribool replace = wxDefaultXlTribool);

        //@{
        /**
        Returns an object that represents either a single series (a Series object) or a collection of all the series (a SeriesCollection collection) in the chart or chart group.

        [MSDN documentation for Chart.SeriesCollection](https://learn.microsoft.com/en-us/office/vba/api/excel.chart.seriescollection).
        */
        wxExcelSeriesCollection SeriesCollection();
        wxExcelSeries SeriesCollection(long index);
        wxExcelSeries SeriesCollection(const wxString& name);
        //@}

        /**
        Sets the background graphic for a chart.

        [MSDN documentation for Chart.SetBackgroundPicture](https://learn.microsoft.com/en-us/office/vba/api/excel.chart.setbackgroundpicture).
        */
        void SetBackgroundPicture(const wxString& fileName);

        /**
        Specifies the name of the chart template that Microsoft Excel uses when creating new charts.
        If name is an empty string, xlBuiltIn will be used.

        [MSDN documentation for Chart.SetDefaultChart](https://learn.microsoft.com/en-us/office/vba/api/excel.chart.setdefaultchart).
        */
        void SetDefaultChart(const wxString& name = wxEmptyString);

        /**
        Sets chart elements on a chart. Read/write MsoChartElementType.

        [MSDN documentation for Chart.SetElement](https://learn.microsoft.com/en-us/office/vba/api/excel.chart.setelement).
        */
        void SetElement(MsoChartElementType element);

        /**
        Sets the source data range for the chart.

        [MSDN documentation for Chart.SetSourceData](https://learn.microsoft.com/en-us/office/vba/api/excel.chart.setsourcedata).
        */
        void SetSourceData(wxExcelRange source, XlRowCol* plotBy = NULL);

        /**
        Removes protection from a sheet or workbook. This method has no effect if the sheet or workbook isn't protected.

        [MSDN documentation for Chart.Unprotect](https://learn.microsoft.com/en-us/office/vba/api/excel.chart.unprotect).
        */
        void Unprotect(const wxString& password = wxEmptyString);

        // ***** PROPERTIES *****

        /**
        True if Microsoft Excel scales a 3-D chart so that it's closer in size to the equivalent 2-D chart. The RightAngleAxes property must be True.

        [MSDN documentation for Chart.AutoScaling](https://learn.microsoft.com/en-us/office/vba/api/excel.chart.autoscaling).
        */
        bool GetAutoScaling();

        /**
        True if Microsoft Excel scales a 3-D chart so that it's closer in size to the equivalent 2-D chart. The RightAngleAxes property must be True.

        [MSDN documentation for Chart.AutoScaling](https://learn.microsoft.com/en-us/office/vba/api/excel.chart.autoscaling).
        */
        void SetAutoScaling(bool autoScaling);

        /**
        Returns a Walls object that allows the user to individually format the back wall of a 3-D chart. Since Excel 2007.

        [MSDN documentation for Chart.BackWall](https://learn.microsoft.com/en-us/office/vba/api/excel.chart.backwall).
        */
        wxExcelWalls GetBackWall();

        /**
        Returns the shape used with the 3-D bar or column chart. Read/write XlBarShape.

        [MSDN documentation for Chart.BarShape](https://learn.microsoft.com/en-us/office/vba/api/excel.chart.barshape).
        */
        XlBarShape GetBarShape();

        /**
        Sets the shape used with the 3-D bar or column chart. Read/write XlBarShape.

        [MSDN documentation for Chart.BarShape](https://learn.microsoft.com/en-us/office/vba/api/excel.chart.barshape).
        */
        void SetBarShape(XlBarShape barShape);

        /**
        Returns a XlCategoryLabelLevel Enumeration (Excel) constant referring to the level of where the category labels are being sourced from.

        [MSDN documentation for Chart.CategoryLabelLevel](https://learn.microsoft.com/en-us/office/vba/api/excel.chart.categorylabellevel).
        */
        long GetCategoryLabelLevel();

        /**
        Sets a XlCategoryLabelLevel Enumeration (Excel) constant referring to the level of where the category labels are being sourced from.

        [MSDN documentation for Chart.CategoryLabelLevel](https://learn.microsoft.com/en-us/office/vba/api/excel.chart.categorylabellevel).
        */
        void SetCategoryLabelLevel(long level);

        /**
        Returns a ChartArea object that represents the complete chart area for the chart.

        [MSDN documentation for Chart.ChartArea](https://learn.microsoft.com/en-us/office/vba/api/excel.chart.chartarea).
        */
        wxExcelChartArea GetChartArea();

        /**
        Returns an Integer that represents the color scheme for the chart

        [MSDN documentation for Chart.ChartColor](https://learn.microsoft.com/en-us/office/vba/api/excel.chart.chartcolor).
        */
        long GetChartColor();

        /**
        Sets an Integer that represents the color scheme for the chart

        [MSDN documentation for Chart.ChartColor](https://learn.microsoft.com/en-us/office/vba/api/excel.chart.chartcolor).
        */
        void SetChartColor(long color);

        /**
        Returns the chart style for the chart. Read/write Variant. Since Excel 2007.

        [MSDN documentation for Chart.ChartStyle](https://learn.microsoft.com/en-us/office/vba/api/excel.chart.chartstyle).
        */
        long GetChartStyle();

        /**
        Sets the chart style for the chart. Read/write Variant. Since Excel 2007.

        [MSDN documentation for Chart.ChartStyle](https://learn.microsoft.com/en-us/office/vba/api/excel.chart.chartstyle).
        */
        void SetChartStyle(long chartStyle);

        /**
        Returns a ChartTitle object that represents the title of the specified chart.

        [MSDN documentation for Chart.ChartTitle](https://learn.microsoft.com/en-us/office/vba/api/excel.chart.charttitle).
        */
        wxExcelChartTitle GetChartTitle();

        /**
        Returns the chart type. Read/write XlChartType.

        [MSDN documentation for Chart.ChartType](https://learn.microsoft.com/en-us/office/vba/api/excel.chart.charttype).
        */
        XlChartType GetChartType();

        /**
        Sets the chart type. Read/write XlChartType.

        [MSDN documentation for Chart.ChartType](https://learn.microsoft.com/en-us/office/vba/api/excel.chart.charttype).
        */
        void SetChartType(XlChartType chartType);

        /**
        Returns the code name for the object.

        [MSDN documentation for Chart.CodeName](https://learn.microsoft.com/en-us/office/vba/api/excel.chart.codename).
        */
        wxString GetCodeName();

        /**
        Returns a 32-bit integer that indicates the application in which this object was created.

        [MSDN documentation for Chart.Creator](https://learn.microsoft.com/en-us/office/vba/api/excel.chart.creator).
        */
        long GetCreator();

        /**
        Returns a DataTable object that represents the chart data table.

        [MSDN documentation for Chart.DataTable](https://learn.microsoft.com/en-us/office/vba/api/excel.chart.datatable).
        */
        wxExcelDataTable GetDataTable();

        /**
        Returns the depth of a 3-D chart as a percentage of the chart width (between 20 and 2000 percent).

        [MSDN documentation for Chart.DepthPercent](https://learn.microsoft.com/en-us/office/vba/api/excel.chart.depthpercent).
        */
        long GetDepthPercent();

        /**
        Sets the depth of a 3-D chart as a percentage of the chart width (between 20 and 2000 percent).

        [MSDN documentation for Chart.DepthPercent](https://learn.microsoft.com/en-us/office/vba/api/excel.chart.depthpercent).
        */
        void SetDepthPercent(long depthPercent);

        /**
        Returns the way that blank cells are plotted on a chart. Can be one of the XlDisplayBlanksAs constants.

        [MSDN documentation for Chart.DisplayBlanksAs](https://learn.microsoft.com/en-us/office/vba/api/excel.chart.displayblanksas).
        */
        XlDisplayBlanksAs GetDisplayBlanksAs();

        /**
        Sets the way that blank cells are plotted on a chart. Can be one of the XlDisplayBlanksAs constants.

        [MSDN documentation for Chart.DisplayBlanksAs](https://learn.microsoft.com/en-us/office/vba/api/excel.chart.displayblanksas).
        */
        void SetDisplayBlanksAs(XlDisplayBlanksAs displayBlanksAs);

        /**
        Returns the elevation of the 3-D chart view, in degrees.

        [MSDN documentation for Chart.Elevation](https://learn.microsoft.com/en-us/office/vba/api/excel.chart.elevation).
        */
        long GetElevation();

        /**
        Sets the elevation of the 3-D chart view, in degrees.

        [MSDN documentation for Chart.Elevation](https://learn.microsoft.com/en-us/office/vba/api/excel.chart.elevation).
        */
        void SetElevation(long elevation);

        /**
        Returns a Floor object that represents the floor of the 3-D chart.

        [MSDN documentation for Chart.Floor](https://learn.microsoft.com/en-us/office/vba/api/excel.chart.floor).
        */
        wxExcelFloor GetFloor();

        /**
        Returns the distance between the data series in a 3-D chart, as a percentage of the marker width. The value of this property must be between 0 and 500.

        [MSDN documentation for Chart.GapDepth](https://learn.microsoft.com/en-us/office/vba/api/excel.chart.gapdepth).
        */
        long GetGapDepth();

        /**
        Sets the distance between the data series in a 3-D chart, as a percentage of the marker width. The value of this property must be between 0 and 500.

        [MSDN documentation for Chart.GapDepth](https://learn.microsoft.com/en-us/office/vba/api/excel.chart.gapdepth).
        */
        void SetGapDepth(long gapDepth);

        /**
        Returns which axes exist on the chart. Read/write Variant.

        [MSDN documentation for Chart.HasAxis](https://learn.microsoft.com/en-us/office/vba/api/excel.chart.hasaxis).
        */
        bool GetHasAxis(XlAxisType index1,  XlAxisGroup* index2 = NULL);

        /**
        Sets which axes exist on the chart. Read/write Variant.

        [MSDN documentation for Chart.HasAxis](https://learn.microsoft.com/en-us/office/vba/api/excel.chart.hasaxis).
        */
        void SetHasAxis(bool hasAxis, XlAxisType index1,  XlAxisGroup* index2 = NULL);

        /**
        True if the chart has a data table.

        [MSDN documentation for Chart.HasDataTable](https://learn.microsoft.com/en-us/office/vba/api/excel.chart.hasdatatable).
        */
        bool GetHasDataTable();

        /**
        True if the chart has a data table.

        [MSDN documentation for Chart.HasDataTable](https://learn.microsoft.com/en-us/office/vba/api/excel.chart.hasdatatable).
        */
        void SetHasDataTable(bool hasDataTable);

        /**
        True if the chart has a legend.

        [MSDN documentation for Chart.HasLegend](https://learn.microsoft.com/en-us/office/vba/api/excel.chart.haslegend).
        */
        bool GetHasLegend();

        /**
        True if the chart has a legend.

        [MSDN documentation for Chart.HasLegend](https://learn.microsoft.com/en-us/office/vba/api/excel.chart.haslegend).
        */
        void SetHasLegend(bool hasLegend);

        /**
        True if the axis or chart has a visible title.

        [MSDN documentation for Chart.HasTitle](https://learn.microsoft.com/en-us/office/vba/api/excel.chart.hastitle).
        */
        bool GetHasTitle();

        /**
        True if the axis or chart has a visible title.

        [MSDN documentation for Chart.HasTitle](https://learn.microsoft.com/en-us/office/vba/api/excel.chart.hastitle).
        */
        void SetHasTitle(bool hasTitle);

        /**
        Returns the height of a 3-D chart as a percentage of the chart width (between 1 and 10000 percent).

        [MSDN documentation for Chart.HeightPercent](https://learn.microsoft.com/en-us/office/vba/api/excel.chart.heightpercent).
        */
        long GetHeightPercent();

        /**
        Sets the height of a 3-D chart as a percentage of the chart width (between 1 and 10000 percent).

        [MSDN documentation for Chart.HeightPercent](https://learn.microsoft.com/en-us/office/vba/api/excel.chart.heightpercent).
        */
        void SetHeightPercent(long heightPercent);

        /**
        Returns a Hyperlinks collection that represents the hyperlinks for the chart.

        [MSDN documentation for Chart.Hyperlinks](https://learn.microsoft.com/en-us/office/vba/api/excel.chart.hyperlinks).
        */
        wxExcelHyperlinks GetHyperlinks();

        /**
        Returns a Long value that represents the index number of the object within the collection of similar objects.

        [MSDN documentation for Chart.Index](https://learn.microsoft.com/en-us/office/vba/api/excel.chart.index).
        */
        long GetIndex();

        /**
        Returns a Legend object that represents the legend for the chart.

        [MSDN documentation for Chart.Legend](https://learn.microsoft.com/en-us/office/vba/api/excel.chart.legend).
        */
        wxExcelLegend GetLegend();

        /**
        Returns a String value representing the name of the object.

        [MSDN documentation for Chart.Name](https://learn.microsoft.com/en-us/office/vba/api/excel.chart.name).
        */
        wxString GetName();

        /**
        Sets a String value representing the name of the object.

        [MSDN documentation for Chart.Name](https://learn.microsoft.com/en-us/office/vba/api/excel.chart.name).
        */
        void SetName(const wxString& name);

        /**
        Returns a Worksheet object that represents the next sheet.

        [MSDN documentation for Chart.Next](https://learn.microsoft.com/en-us/office/vba/api/excel.chart.next).
        */
        wxExcelWorksheet GetNext();

        /**
        Returns a PageSetup object that contains all the page setup settings for the specified object.

        [MSDN documentation for Chart.PageSetup](https://learn.microsoft.com/en-us/office/vba/api/excel.chart.pagesetup).
        */
        wxExcelPageSetup GetPageSetup();

        /**
        Returns a Long value that represents the perspective for the 3-D chart view.

        [MSDN documentation for Chart.Perspective](https://learn.microsoft.com/en-us/office/vba/api/excel.chart.perspective).
        */
        long GetPerspective();

        /**
        Sets a Long value that represents the perspective for the 3-D chart view.

        [MSDN documentation for Chart.Perspective](https://learn.microsoft.com/en-us/office/vba/api/excel.chart.perspective).
        */
        void SetPerspective(long perspective);

        /**
        Returns a PlotArea object that represents the plot area of a chart.

        [MSDN documentation for Chart.PlotArea](https://learn.microsoft.com/en-us/office/vba/api/excel.chart.plotarea).
        */
        wxExcelPlotArea GetPlotArea();

        /**
        Returns the way columns or rows are used as data series on the chart. Can be one of the following XlRowCol constants: xlColumns or xlRows.

        [MSDN documentation for Chart.PlotBy](https://learn.microsoft.com/en-us/office/vba/api/excel.chart.plotby).
        */
        XlRowCol GetPlotBy();

        /**
        Sets the way columns or rows are used as data series on the chart. Can be one of the following XlRowCol constants: xlColumns or xlRows.

        [MSDN documentation for Chart.PlotBy](https://learn.microsoft.com/en-us/office/vba/api/excel.chart.plotby).
        */
        void SetPlotBy(XlRowCol plotBy);

        /**
        True if only visible cells are plotted. False if both visible and hidden cells are plotted.

        [MSDN documentation for Chart.PlotVisibleOnly](https://learn.microsoft.com/en-us/office/vba/api/excel.chart.plotvisibleonly).
        */
        bool GetPlotVisibleOnly();

        /**
        True if only visible cells are plotted. False if both visible and hidden cells are plotted.

        [MSDN documentation for Chart.PlotVisibleOnly](https://learn.microsoft.com/en-us/office/vba/api/excel.chart.plotvisibleonly).
        */
        void SetPlotVisibleOnly(bool plotVisibleOnly);

        /**
        Returns a Worksheet object that represents the next sheet.

        [MSDN documentation for Chart.Previous](https://learn.microsoft.com/en-us/office/vba/api/excel.chart.previous).
        */
        wxExcelWorksheet GetPrevious();

        /**
        True if the contents of the sheet are protected. For a chart, this protects the entire chart. To turn on content protection, use the Protect method with the Contents argument set to True.

        [MSDN documentation for Chart.ProtectContents](https://learn.microsoft.com/en-us/office/vba/api/excel.chart.protectcontents).
        */
        bool GetProtectContents();

        /**
        True if series formulas cannot be modified by the user.

        [MSDN documentation for Chart.ProtectData](https://learn.microsoft.com/en-us/office/vba/api/excel.chart.protectdata).
        */
        bool GetProtectData();

        /**
        True if series formulas cannot be modified by the user.

        [MSDN documentation for Chart.ProtectData](https://learn.microsoft.com/en-us/office/vba/api/excel.chart.protectdata).
        */
        void SetProtectData(bool protectData);

        /**
        True if shapes are protected. To turn on shape protection, use the Protect method with the DrawingObjects argument set to True.

        [MSDN documentation for Chart.ProtectDrawingObjects](https://learn.microsoft.com/en-us/office/vba/api/excel.chart.protectdrawingobjects).
        */
        bool GetProtectDrawingObjects();

        /**
        True if chart formatting cannot be modified by the user.

        [MSDN documentation for Chart.ProtectFormatting](https://learn.microsoft.com/en-us/office/vba/api/excel.chart.protectformatting).
        */
        bool GetProtectFormatting();

        /**
        True if chart formatting cannot be modified by the user.

        [MSDN documentation for Chart.ProtectFormatting](https://learn.microsoft.com/en-us/office/vba/api/excel.chart.protectformatting).
        */
        void SetProtectFormatting(bool protectFormatting);

        /**
        True if user-interface-only protection is turned on. To turn on user interface protection, use the Protect method with the UserInterfaceOnly argument set to True.

        [MSDN documentation for Chart.ProtectionMode](https://learn.microsoft.com/en-us/office/vba/api/excel.chart.protectionmode).
        */
        bool GetProtectionMode();

        /**
        True if chart elements cannot be selected.

        [MSDN documentation for Chart.ProtectSelection](https://learn.microsoft.com/en-us/office/vba/api/excel.chart.protectselection).
        */
        bool GetProtectSelection();

        /**
        True if chart elements cannot be selected.

        [MSDN documentation for Chart.ProtectSelection](https://learn.microsoft.com/en-us/office/vba/api/excel.chart.protectselection).
        */
        void SetProtectSelection(bool protectSelection);

        /**
        True if the chart axes are at right angles, independent of chart rotation or elevation. Applies only to 3-D line, column, and bar charts.

        [MSDN documentation for Chart.RightAngleAxes](https://learn.microsoft.com/en-us/office/vba/api/excel.chart.rightangleaxes).
        */
        bool GetRightAngleAxes();

        /**
        True if the chart axes are at right angles, independent of chart rotation or elevation. Applies only to 3-D line, column, and bar charts.

        [MSDN documentation for Chart.RightAngleAxes](https://learn.microsoft.com/en-us/office/vba/api/excel.chart.rightangleaxes).
        */
        void SetRightAngleAxes(bool rightAngleAxes);

        /**
        Returns the rotation of the 3-D chart view (the rotation of the plot area around the z-axis, in degrees). The value of this property must be from 0 to 360, except for 3-D bar charts, where the value must be from 0 to 44. The default value is 20. Applies only to 3-D charts. Read/write Variant.

        [MSDN documentation for Chart.Rotation](https://learn.microsoft.com/en-us/office/vba/api/excel.chart.rotation).
        */
        long GetRotation();

        /**
        Sets the rotation of the 3-D chart view (the rotation of the plot area around the z-axis, in degrees). The value of this property must be from 0 to 360, except for 3-D bar charts, where the value must be from 0 to 44. The default value is 20. Applies only to 3-D charts. Read/write Variant.

        [MSDN documentation for Chart.Rotation](https://learn.microsoft.com/en-us/office/vba/api/excel.chart.rotation).
        */
        void SetRotation(long rotation);

        /**
        Returns a XlSeriesNameLevel Enumeration (Excel) constant referring to the level of where the series names are being sourced from.

        [MSDN documentation for Chart.SeriesNameLevel](https://learn.microsoft.com/en-us/office/vba/api/excel.chart.seriesnamelevel).
        */
        long GetSeriesNameLevel();

        /**
        Sets a XlSeriesNameLevel Enumeration (Excel) constant referring to the level of where the series names are being sourced from.

        [MSDN documentation for Chart.SeriesNameLevel](https://learn.microsoft.com/en-us/office/vba/api/excel.chart.seriesnamelevel).
        */
        void SetSeriesNameLevel(long level);

#if WXAUTOEXCEL_USE_SHAPES
        /**
        Returns a Shapes collection that represents all the shapes on the chart sheet.

        [MSDN documentation for Chart.Shapes](https://learn.microsoft.com/en-us/office/vba/api/excel.chart.shapes).
        */
        wxExcelShapes GetShapes();
#endif // #if WXAUTOEXCEL_USE_SHAPES

        /**
        Returns whether to show the data labels when the value is greater than the maximum value on the value axis. Since Excel 2007.

        [MSDN documentation for Chart.ShowDataLabelsOverMaximum](https://learn.microsoft.com/en-us/office/vba/api/excel.chart.showdatalabelsovermaximum).
        */
        bool GetShowDataLabelsOverMaximum();

        /**
        Sets whether to show the data labels when the value is greater than the maximum value on the value axis. Since Excel 2007.

        [MSDN documentation for Chart.ShowDataLabelsOverMaximum](https://learn.microsoft.com/en-us/office/vba/api/excel.chart.showdatalabelsovermaximum).
        */
        void SetShowDataLabelsOverMaximum(bool showDataLabelsOverMaximum);

        /**
        Returns a Walls object that allows the user to individually format the side wall of a 3-D chart. Since Excel 2007.

        [MSDN documentation for Chart.SideWall](https://learn.microsoft.com/en-us/office/vba/api/excel.chart.sidewall).
        */
        wxExcelWalls GetSideWall();

        /**
        Returns a Tab object for a chart.

        [MSDN documentation for Chart.Tab](https://learn.microsoft.com/en-us/office/vba/api/excel.chart.tab).
        */
        wxExcelTab GetTab();

        /**
        Returns an XlSheetVisibility value that determines whether the object is visible.

        [MSDN documentation for Chart.Visible](https://learn.microsoft.com/en-us/office/vba/api/excel.chart.visible).
        */
        XlSheetVisibility GetVisible();

        /**
        Sets an XlSheetVisibility value that determines whether the object is visible.

        [MSDN documentation for Chart.Visible](https://learn.microsoft.com/en-us/office/vba/api/excel.chart.visible).
        */
        void SetVisible(XlSheetVisibility visible);

        /**
        Returns a Walls object that represents the walls of the 3-D chart.

        [MSDN documentation for Chart.Walls](https://learn.microsoft.com/en-us/office/vba/api/excel.chart.walls).
        */
        wxExcelWalls GetWalls();

        /**
        Returns "Chart".
        */
        virtual wxString GetAutoExcelObjectName_() const { return wxS("Chart"); }

    private:
        bool DoOrderedCopyOrMove(bool copy, wxExcelSheet sheetAfterOrBefore, bool after);
    };


} // namespace wxAutoExcel

#endif // #if WXAUTOEXCEL_USE_CHARTS

#endif // _WXAUTOEXCEL_CHART_H
