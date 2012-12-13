/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Modified by:
// RCS-ID:      $Id: $
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
/////////////////////////////////////////////////////////////////////////////


#ifndef _WXAUTOEXCEL_CHART_H
#define _WXAUTOEXCEL_CHART_H

#include "wx/wxAutoExcel_defs.h"

#if WXAUTOEXCEL_USE_CHARTS

#include "wx/wxAutoExcelObject.h"
#include "wx/wxAutoExcel_enums.h"

class wxArrayString;

namespace wxAutoExcel {


    /**
    Represents Microsoft Excel Chart object.
    */
   class WXDLLIMPEXP_WXAUTOEXCEL wxExcelChart : public wxExcelObject
    {
    public:
        // ***** METHODS *****

        /**
        Makes the current chart the active chart.

        [MSDN documentation for Chart.Activate](http://msdn.microsoft.com/en-us/library/bb211631).
        */
        void Activate();

        /**
        Applies a standard or custom chart type to a chart.

        [MSDN documentation for Chart.ApplyChartTemplate](http://msdn.microsoft.com/en-us/library/bb238815).
        */
        void ApplyChartTemplate(const wxString& fileName);

        /**
        Applies data labels to all the series in a chart.

        [MSDN documentation for Chart.ApplyDataLabels](http://msdn.microsoft.com/en-us/library/bb211636).
        */
        void ApplyDataLabels(XlDataLabelsType* type = NULL, wxXlTribool legendKey = wxDefaultXlTribool, 
                             wxXlTribool autoText = wxDefaultXlTribool, wxXlTribool hasLeaderLines = wxDefaultXlTribool, 
                             wxXlTribool showSeriesName = wxDefaultXlTribool, wxXlTribool showCategoryName = wxDefaultXlTribool, 
                             wxXlTribool showValue = wxDefaultXlTribool, wxXlTribool showPercentage = wxDefaultXlTribool, 
                             wxXlTribool showBubbleSize = wxDefaultXlTribool, const wxString& separator = wxEmptyString);

        /**
        Applies the layouts shown in the ribbon.

        [MSDN documentation for Chart.ApplyLayout](http://msdn.microsoft.com/en-us/library/bb238817).
        */
        void ApplyLayout(long layout, XlChartType* chartType = NULL);

        /**
        Returns an object that represents either a single axis or a collection of the axes on the chart.

        [MSDN documentation for Chart.Axes](http://msdn.microsoft.com/en-us/library/bb209703).
        */
        wxExcelAxis Axes(XlAxisType type,  XlAxisGroup* axisGroup = NULL);
        
        /**
        Returns an object that represents either a single chart group (a ChartGroup object) or a collection of all the chart groups in the chart (a ChartGroups object). The returned collection includes every type of group.

        [MSDN documentation for Chart.ChartGroups](http://msdn.microsoft.com/en-us/library/bb223238).
        */
        wxExcelChartGroups ChartGroups();                

        //@{
        /**
        Returns an object that represents either a single embedded chart (a ChartObject object) or a collection of all the embedded charts (a ChartObjects object) on the sheet.

        [MSDN documentation for Chart.ChartObjects](http://msdn.microsoft.com/en-us/library/bb148195).
        */
        wxExcelChartObjects ChartObjects();
        wxExcelChartObjects ChartObjects(const wxVector<long>& indices);
        wxExcelChartObjects ChartObjects(const wxArrayString& names);
        //@}

        /**
        Modifies the properties of the given chart. You can use this method to quickly format a chart without setting all the individual properties. This method is noninteractive, and it changes only the specified properties.

        [MSDN documentation for Chart.ChartWizard](http://msdn.microsoft.com/en-us/library/bb223240).
        */
        void ChartWizard(wxExcelRange* source = NULL, XlChartType* gallery = NULL, 
                         long* format = NULL, XlRowCol* plotBy = NULL, 
                         long* categoryLabels = NULL, long* seriesLabels = NULL, 
                         wxXlTribool hasLegend = wxDefaultXlTribool, 
                         const wxString& title = wxEmptyString, const wxString& categoryTitle = wxEmptyString, 
                         const wxString& valueTitle = wxEmptyString, const wxString& extraTitle = wxEmptyString);

        /**
        Checks the spelling of an object.

        [MSDN documentation for Chart.CheckSpelling](http://msdn.microsoft.com/en-us/library/bb148202).
        */
        void CheckSpelling(const wxString& customDictionary = wxEmptyString, wxXlTribool ignoreUpperCase = wxDefaultXlTribool,
                           wxXlTribool alwaysSuggest = wxDefaultXlTribool, MsoLanguageID* spellLang = NULL);

        /**
        Clears the chart elements formatting to automatic.

        [MSDN documentation for Chart.ClearToMatchStyle](http://msdn.microsoft.com/en-us/library/bb225816).
        */
        void ClearToMatchStyle();

        /**
        Copies the selected object to the Clipboard as a picture.

        [MSDN documentation for Chart.CopyPicture](http://msdn.microsoft.com/en-us/library/bb148210).
        */
        void CopyPicture(XlPictureAppearance* appearance, XlCopyPictureFormat* format);

        /**
        Deletes the object.

        [MSDN documentation for Chart.Delete](http://msdn.microsoft.com/en-us/library/bb148215).
        */
        void Delete();

        /**
        Cancels the selection for the specified chart.

        [MSDN documentation for Chart.Delete](http://msdn.microsoft.com/en-us/library/bb223343).
        */
        void Deselect();

        /**
        Exports the chart in a graphic format.

        [MSDN documentation for Chart.Export](http://msdn.microsoft.com/en-us/library/bb148223).
        */
        bool Export(const wxString& fileName, const wxString& filterName = wxEmptyString, 
                    wxXlTribool interactive = wxDefaultXlTribool);
   

        /**
        Moves the chart to a new location.

        [MSDN documentation for Chart.Location](http://msdn.microsoft.com/en-us/library/bb242020).
        */
        wxExcelChart Location(XlChartLocation where, const wxString& name = wxEmptyString);
        
        /**
            Moves the chart to the new workbook.
            [MSDN documentation for Chart.Move](http://msdn.microsoft.com/en-us/library/bb148249).
        */
        bool Move();
        /**
            Moves the chart within the same workbook, after or before the specified sheet.
            [MSDN documentation for Chart.Move](http://msdn.microsoft.com/en-us/library/bb148249).
        */
        bool MoveAfterOrBefore(wxExcelSheet sheetAfterOrBefore, bool after);
        
        /**
        Returns an object that represents either a single OLE object (an OLEObject ) or a collection of all OLE objects (an OLEObjects collection) on the chart or sheet. Read-only.

        [MSDN documentation for Chart.OLEObjects](http://msdn.microsoft.com/en-us/library/bb211658).
        */
        wxExcelOLEObjects OLEObjects();                

        /**
        Pastes chart data from the Clipboard into the specified chart.

        [MSDN documentation for Chart.Paste](http://msdn.microsoft.com/en-us/library/bb211662).
        */
        void Paste(XlPasteType* type = NULL);

        //@{
        /**
        Prints the object.

        [MSDN documentation for Chart.PrintOut](http://msdn.microsoft.com/en-us/library/bb211667).
        */
        bool PrintOut(long* from = NULL, long* to = NULL, long* copies = NULL, wxXlTribool preview = wxDefaultXlTribool,
                      const wxString& activePrinter = wxEmptyString, wxXlTribool printToFile = wxDefaultXlTribool,
                      wxXlTribool collate = wxDefaultXlTribool, const wxString& prToFileName= wxEmptyString);
        bool PrintOut(const wxVariantVector& args);
        //@}

        /**
        Shows a preview of the object as it would look when printed.

        [MSDN documentation for Chart.PrintPreview](http://msdn.microsoft.com/en-us/library/bb211671).
        */
        bool PrintPreview(wxXlTribool enableChanges = wxDefaultXlTribool);

        /**
        Protects a chart so that it cannot be modified.

        [MSDN documentation for Chart.Protect](http://msdn.microsoft.com/en-us/library/bb211674).
        */
        void Protect(const wxString& password = wxEmptyString, wxXlTribool shapes = wxDefaultXlTribool, wxXlTribool contents = wxDefaultXlTribool, wxXlTribool userInterfaceOnly = wxDefaultXlTribool);

        /**
        Causes the specified chart to be redrawn immediately.

        [MSDN documentation for Chart.Refresh](http://msdn.microsoft.com/en-us/library/bb211677).
        */
        void Refresh();

        //@{
        /**
        Saves changes to the chart in a different file.

        [MSDN documentation for Chart.SaveAs](http://msdn.microsoft.com/en-us/library/bb213939).
        */
         void SaveAs(const wxString& fileName = wxEmptyString, XlFileFormat* fileFormat = NULL,
                     const wxString& password = wxEmptyString, const wxString& writeResPassword = wxEmptyString,
                     wxXlTribool readOnlyRecommended = wxDefaultXlTribool, wxXlTribool createBackup = wxDefaultXlTribool,               
                     wxXlTribool addToMru = wxDefaultXlTribool, wxXlTribool local = wxDefaultXlTribool);

        void SaveAs(const wxVariantVector& optionalArgs);
        //@}

        /**
        Saves a custom chart template to the list of available chart templates.

        [MSDN documentation for Chart.SaveChartTemplate](http://msdn.microsoft.com/en-us/library/bb238822).
        */
        void SaveChartTemplate(const wxString& fileName);

        /**
        Selects the object.

        [MSDN documentation for Chart.Select](http://msdn.microsoft.com/en-us/library/bb213943).
        */
        void Select(wxXlTribool replace = wxDefaultXlTribool);
        
        /**
        Returns an object that represents either a single series (a Series object) or a collection of all the series (a SeriesCollection collection) in the chart or chart group.

        [MSDN documentation for Chart.SeriesCollection](http://msdn.microsoft.com/en-us/library/bb213957).
        */
        wxExcelSeriesCollection SeriesCollection();
        
        /**
        Sets the background graphic for a chart.

        [MSDN documentation for Chart.SetBackgroundPicture](http://msdn.microsoft.com/en-us/library/bb213966).
        */
        void SetBackgroundPicture(const wxString& fileName);

        /**
        Specifies the name of the chart template that Microsoft Excel uses when creating new charts.
        If name is an empty string, xlBuiltIn will be used.

        [MSDN documentation for Chart.SetDefaultChart](http://msdn.microsoft.com/en-us/library/bb237753).
        */
        void SetDefaultChart(const wxString& name = wxEmptyString);

        /**
        Sets chart elements on a chart. Read/write MsoChartElementType.

        [MSDN documentation for Chart.SetElement](http://msdn.microsoft.com/en-us/library/bb242078).
        */
        void SetElement(MsoChartElementType element);

        /**
        Sets the source data range for the chart.

        [MSDN documentation for Chart.SetSourceData](http://msdn.microsoft.com/en-us/library/bb178089).
        */
        void SetSourceData(wxExcelRange source, XlRowCol* plotBy = NULL);

        /**
        Removes protection from a sheet or workbook. This method has no effect if the sheet or workbook isn't protected.

        [MSDN documentation for Chart.Unprotect](http://msdn.microsoft.com/en-us/library/bb237794).
        */
        void Unprotect(const wxString& password = wxEmptyString);

        // ***** PROPERTIES *****

        /**
        True if Microsoft Excel scales a 3-D chart so that it's closer in size to the equivalent 2-D chart. The RightAngleAxes property must be True.

        [MSDN documentation for Chart.AutoScaling](http://msdn.microsoft.com/en-us/library/bb220861).
        */
        bool GetAutoScaling();

        /**
        True if Microsoft Excel scales a 3-D chart so that it's closer in size to the equivalent 2-D chart. The RightAngleAxes property must be True.

        [MSDN documentation for Chart.AutoScaling](http://msdn.microsoft.com/en-us/library/bb220861).
        */
        void SetAutoScaling(bool autoScaling);

        /**
        Returns a Walls object that allows the user to individually format the back wall of a 3-D chart. Since Excel 2007.

        [MSDN documentation for Chart.BackWall](http://msdn.microsoft.com/en-us/library/bb239908).
        */
        wxExcelWalls GetBackWall();

        /**
        Returns the shape used with the 3-D bar or column chart. Read/write XlBarShape.

        [MSDN documentation for Chart.BarShape](http://msdn.microsoft.com/en-us/library/bb179421).
        */
        XlBarShape GetBarShape();

        /**
        Sets the shape used with the 3-D bar or column chart. Read/write XlBarShape.

        [MSDN documentation for Chart.BarShape](http://msdn.microsoft.com/en-us/library/bb179421).
        */
        void SetBarShape(XlBarShape barShape);

        /**
        Returns a ChartArea object that represents the complete chart area for the chart.

        [MSDN documentation for Chart.ChartArea](http://msdn.microsoft.com/en-us/library/bb177343).
        */
        wxExcelChartArea GetChartArea();

        /**
        Returns the chart style for the chart. Read/write Variant. Since Excel 2007.

        [MSDN documentation for Chart.ChartStyle](http://msdn.microsoft.com/en-us/library/bb239909).
        */
        long GetChartStyle();

        /**
        Sets the chart style for the chart. Read/write Variant. Since Excel 2007.

        [MSDN documentation for Chart.ChartStyle](http://msdn.microsoft.com/en-us/library/bb239909).
        */
        void SetChartStyle(long chartStyle);

        /**
        Returns a ChartTitle object that represents the title of the specified chart.

        [MSDN documentation for Chart.ChartTitle](http://msdn.microsoft.com/en-us/library/bb177345).
        */
        wxExcelChartTitle GetChartTitle();

        /**
        Returns the chart type. Read/write XlChartType.

        [MSDN documentation for Chart.ChartType](http://msdn.microsoft.com/en-us/library/bb179424).
        */
        XlChartType GetChartType();

        /**
        Sets the chart type. Read/write XlChartType.

        [MSDN documentation for Chart.ChartType](http://msdn.microsoft.com/en-us/library/bb179424).
        */
        void SetChartType(XlChartType chartType);

        /**
        Returns the code name for the object.

        [MSDN documentation for Chart.CodeName](http://msdn.microsoft.com/en-us/library/bb179427).
        */
        wxString GetCodeName();

        /**
        Returns a 32-bit integer that indicates the application in which this object was created.

        [MSDN documentation for Chart.Creator](http://msdn.microsoft.com/en-us/library/bb256627).
        */
        long GetCreator();

        /**
        Returns the depth of a 3-D chart as a percentage of the chart width (between 20 and 2000 percent).

        [MSDN documentation for Chart.DepthPercent](http://msdn.microsoft.com/en-us/library/bb177465).
        */
        long GetDepthPercent();

        /**
        Sets the depth of a 3-D chart as a percentage of the chart width (between 20 and 2000 percent).

        [MSDN documentation for Chart.DepthPercent](http://msdn.microsoft.com/en-us/library/bb177465).
        */
        void SetDepthPercent(long depthPercent);

        /**
        Returns the way that blank cells are plotted on a chart. Can be one of the XlDisplayBlanksAs constants.

        [MSDN documentation for Chart.DisplayBlanksAs](http://msdn.microsoft.com/en-us/library/bb177483).
        */
        XlDisplayBlanksAs GetDisplayBlanksAs();

        /**
        Sets the way that blank cells are plotted on a chart. Can be one of the XlDisplayBlanksAs constants.

        [MSDN documentation for Chart.DisplayBlanksAs](http://msdn.microsoft.com/en-us/library/bb177483).
        */
        void SetDisplayBlanksAs(XlDisplayBlanksAs displayBlanksAs);

        /**
        Returns the elevation of the 3-D chart view, in degrees.

        [MSDN documentation for Chart.Elevation](http://msdn.microsoft.com/en-us/library/bb221088).
        */
        long GetElevation();

        /**
        Sets the elevation of the 3-D chart view, in degrees.

        [MSDN documentation for Chart.Elevation](http://msdn.microsoft.com/en-us/library/bb221088).
        */
        void SetElevation(long elevation);

        /**
        Returns a Floor object that represents the floor of the 3-D chart.

        [MSDN documentation for Chart.Floor](http://msdn.microsoft.com/en-us/library/bb242037).
        */
        wxExcelFloor GetFloor();

        /**
        Returns the distance between the data series in a 3-D chart, as a percentage of the marker width. The value of this property must be between 0 and 500.

        [MSDN documentation for Chart.GapDepth](http://msdn.microsoft.com/en-us/library/bb208572).
        */
        long GetGapDepth();

        /**
        Sets the distance between the data series in a 3-D chart, as a percentage of the marker width. The value of this property must be between 0 and 500.

        [MSDN documentation for Chart.GapDepth](http://msdn.microsoft.com/en-us/library/bb208572).
        */
        void SetGapDepth(long gapDepth);

        /**
        Returns which axes exist on the chart. Read/write Variant.

        [MSDN documentation for Chart.HasAxis](http://msdn.microsoft.com/en-us/library/bb208595).
        */
        bool GetHasAxis(XlAxisType index1,  XlAxisGroup* index2 = NULL);

        /**
        Sets which axes exist on the chart. Read/write Variant.

        [MSDN documentation for Chart.HasAxis](http://msdn.microsoft.com/en-us/library/bb208595).
        */
        void SetHasAxis(bool hasAxis, XlAxisType index1,  XlAxisGroup* index2 = NULL);

        /**
        True if the chart has a data table.

        [MSDN documentation for Chart.HasDataTable](http://msdn.microsoft.com/en-us/library/bb208611).
        */
        bool GetHasDataTable();

        /**
        True if the chart has a data table.

        [MSDN documentation for Chart.HasDataTable](http://msdn.microsoft.com/en-us/library/bb208611).
        */
        void SetHasDataTable(bool hasDataTable);

        /**
        True if the chart has a legend.

        [MSDN documentation for Chart.HasLegend](http://msdn.microsoft.com/en-us/library/bb208638).
        */
        bool GetHasLegend();

        /**
        True if the chart has a legend.

        [MSDN documentation for Chart.HasLegend](http://msdn.microsoft.com/en-us/library/bb208638).
        */
        void SetHasLegend(bool hasLegend);

        /**
        True if the axis or chart has a visible title.

        [MSDN documentation for Chart.HasTitle](http://msdn.microsoft.com/en-us/library/bb179455).
        */
        bool GetHasTitle();

        /**
        True if the axis or chart has a visible title.

        [MSDN documentation for Chart.HasTitle](http://msdn.microsoft.com/en-us/library/bb179455).
        */
        void SetHasTitle(bool hasTitle);

        /**
        Returns the height of a 3-D chart as a percentage of the chart width (between 1 and 10000 percent).

        [MSDN documentation for Chart.HeightPercent](http://msdn.microsoft.com/en-us/library/bb208679).
        */
        long GetHeightPercent();

        /**
        Sets the height of a 3-D chart as a percentage of the chart width (between 1 and 10000 percent).

        [MSDN documentation for Chart.HeightPercent](http://msdn.microsoft.com/en-us/library/bb208679).
        */
        void SetHeightPercent(long heightPercent);

        /**
        Returns a Hyperlinks collection that represents the hyperlinks for the chart.

        [MSDN documentation for Chart.Hyperlinks](http://msdn.microsoft.com/en-us/library/bb179456).
        */
        wxExcelHyperlinks GetHyperlinks();

        /**
        Returns a Long value that represents the index number of the object within the collection of similar objects.

        [MSDN documentation for Chart.Index](http://msdn.microsoft.com/en-us/library/bb179458).
        */
        long GetIndex();

        /**
        Returns a Legend object that represents the legend for the chart.

        [MSDN documentation for Chart.Legend](http://msdn.microsoft.com/en-us/library/bb177908).
        */
        wxExcelLegend GetLegend();

        /**
        Returns a String value representing the name of the object.

        [MSDN documentation for Chart.Name](http://msdn.microsoft.com/en-us/library/bb179461).
        */
        wxString GetName();

        /**
        Sets a String value representing the name of the object.

        [MSDN documentation for Chart.Name](http://msdn.microsoft.com/en-us/library/bb179461).
        */
        void SetName(const wxString& name);

        /**
        Returns a Worksheet object that represents the next sheet.

        [MSDN documentation for Chart.Next](http://msdn.microsoft.com/en-us/library/bb179462).
        */
        wxExcelWorksheet GetNext();

        /**
        Returns a PageSetup object that contains all the page setup settings for the specified object.

        [MSDN documentation for Chart.PageSetup](http://msdn.microsoft.com/en-us/library/bb212932).
        */
        wxExcelPageSetup GetPageSetup();

        /**
        Returns a Long value that represents the perspective for the 3-D chart view.

        [MSDN documentation for Chart.Perspective](http://msdn.microsoft.com/en-us/library/bb212934).
        */
        long GetPerspective();

        /**
        Sets a Long value that represents the perspective for the 3-D chart view.

        [MSDN documentation for Chart.Perspective](http://msdn.microsoft.com/en-us/library/bb212934).
        */
        void SetPerspective(long perspective);

        /**
        Returns a PlotArea object that represents the plot area of a chart.

        [MSDN documentation for Chart.PlotArea](http://msdn.microsoft.com/en-us/library/bb221428).
        */
        wxExcelPlotArea GetPlotArea();

        /**
        Returns the way columns or rows are used as data series on the chart. Can be one of the following XlRowCol constants: xlColumns or xlRows.

        [MSDN documentation for Chart.PlotBy](http://msdn.microsoft.com/en-us/library/bb221430).
        */
        XlRowCol GetPlotBy();

        /**
        Sets the way columns or rows are used as data series on the chart. Can be one of the following XlRowCol constants: xlColumns or xlRows.

        [MSDN documentation for Chart.PlotBy](http://msdn.microsoft.com/en-us/library/bb221430).
        */
        void SetPlotBy(XlRowCol plotBy);

        /**
        True if only visible cells are plotted. False if both visible and hidden cells are plotted.

        [MSDN documentation for Chart.PlotVisibleOnly](http://msdn.microsoft.com/en-us/library/bb221434).
        */
        bool GetPlotVisibleOnly();

        /**
        True if only visible cells are plotted. False if both visible and hidden cells are plotted.

        [MSDN documentation for Chart.PlotVisibleOnly](http://msdn.microsoft.com/en-us/library/bb221434).
        */
        void SetPlotVisibleOnly(bool plotVisibleOnly);

        /**
        Returns a Worksheet object that represents the next sheet.

        [MSDN documentation for Chart.Previous](http://msdn.microsoft.com/en-us/library/bb212937).
        */
        wxExcelWorksheet GetPrevious();

        /**
        True if the contents of the sheet are protected. For a chart, this protects the entire chart. To turn on content protection, use the Protect method with the Contents argument set to True.

        [MSDN documentation for Chart.ProtectContents](http://msdn.microsoft.com/en-us/library/bb179477).
        */
        bool GetProtectContents();

        /**
        True if series formulas cannot be modified by the user.

        [MSDN documentation for Chart.ProtectData](http://msdn.microsoft.com/en-us/library/bb209033).
        */
        bool GetProtectData();

        /**
        True if series formulas cannot be modified by the user.

        [MSDN documentation for Chart.ProtectData](http://msdn.microsoft.com/en-us/library/bb209033).
        */
        void SetProtectData(bool protectData);

        /**
        True if shapes are protected. To turn on shape protection, use the Protect method with the DrawingObjects argument set to True.

        [MSDN documentation for Chart.ProtectDrawingObjects](http://msdn.microsoft.com/en-us/library/bb179481).
        */
        bool GetProtectDrawingObjects();

        /**
        True if chart formatting cannot be modified by the user.

        [MSDN documentation for Chart.ProtectFormatting](http://msdn.microsoft.com/en-us/library/bb209035).
        */
        bool GetProtectFormatting();

        /**
        True if chart formatting cannot be modified by the user.

        [MSDN documentation for Chart.ProtectFormatting](http://msdn.microsoft.com/en-us/library/bb209035).
        */
        void SetProtectFormatting(bool protectFormatting);

        /**
        True if user-interface-only protection is turned on. To turn on user interface protection, use the Protect method with the UserInterfaceOnly argument set to True.

        [MSDN documentation for Chart.ProtectionMode](http://msdn.microsoft.com/en-us/library/bb179486).
        */
        bool GetProtectionMode();

        /**
        True if chart elements cannot be selected.

        [MSDN documentation for Chart.ProtectSelection](http://msdn.microsoft.com/en-us/library/bb209043).
        */
        bool GetProtectSelection();

        /**
        True if chart elements cannot be selected.

        [MSDN documentation for Chart.ProtectSelection](http://msdn.microsoft.com/en-us/library/bb209043).
        */
        void SetProtectSelection(bool protectSelection);

        /**
        True if the chart axes are at right angles, independent of chart rotation or elevation. Applies only to 3-D line, column, and bar charts.

        [MSDN documentation for Chart.RightAngleAxes](http://msdn.microsoft.com/en-us/library/bb209160).
        */
        bool GetRightAngleAxes();

        /**
        True if the chart axes are at right angles, independent of chart rotation or elevation. Applies only to 3-D line, column, and bar charts.

        [MSDN documentation for Chart.RightAngleAxes](http://msdn.microsoft.com/en-us/library/bb209160).
        */
        void SetRightAngleAxes(bool rightAngleAxes);

        /**
        Returns the rotation of the 3-D chart view (the rotation of the plot area around the z-axis, in degrees). The value of this property must be from 0 to 360, except for 3-D bar charts, where the value must be from 0 to 44. The default value is 20. Applies only to 3-D charts. Read/write Variant.

        [MSDN documentation for Chart.Rotation](http://msdn.microsoft.com/en-us/library/bb238472).
        */
        long GetRotation();

        /**
        Sets the rotation of the 3-D chart view (the rotation of the plot area around the z-axis, in degrees). The value of this property must be from 0 to 360, except for 3-D bar charts, where the value must be from 0 to 44. The default value is 20. Applies only to 3-D charts. Read/write Variant.

        [MSDN documentation for Chart.Rotation](http://msdn.microsoft.com/en-us/library/bb238472).
        */
        void SetRotation(long rotation);

#if WXAUTOEXCEL_USE_SHAPES
        /**
        Returns a Shapes collection that represents all the shapes on the chart sheet.

        [MSDN documentation for Chart.Shapes](http://msdn.microsoft.com/en-us/library/bb238476).
        */
        wxExcelShapes GetShapes();
#endif // #if WXAUTOEXCEL_USE_SHAPES

        /**
        Returns whether to show the data labels when the value is greater than the maximum value on the value axis. Since Excel 2007.

        [MSDN documentation for Chart.ShowDataLabelsOverMaximum](http://msdn.microsoft.com/en-us/library/bb242608).
        */
        bool GetShowDataLabelsOverMaximum();

        /**
        Sets whether to show the data labels when the value is greater than the maximum value on the value axis. Since Excel 2007.

        [MSDN documentation for Chart.ShowDataLabelsOverMaximum](http://msdn.microsoft.com/en-us/library/bb242608).
        */
        void SetShowDataLabelsOverMaximum(bool showDataLabelsOverMaximum);

        /**
        Returns a Walls object that allows the user to individually format the side wall of a 3-D chart. Since Excel 2007.

        [MSDN documentation for Chart.SideWall](http://msdn.microsoft.com/en-us/library/bb239945).
        */
        wxExcelWalls GetSideWall();

        /**
        Returns a Tab object for a chart.

        [MSDN documentation for Chart.Tab](http://msdn.microsoft.com/en-us/library/bb238491).
        */
        wxExcelTab GetTab();

        /**
        Returns an XlSheetVisibility value that determines whether the object is visible.

        [MSDN documentation for Chart.Visible](http://msdn.microsoft.com/en-us/library/bb238522).
        */
        XlSheetVisibility GetVisible();

        /**
        Sets an XlSheetVisibility value that determines whether the object is visible.

        [MSDN documentation for Chart.Visible](http://msdn.microsoft.com/en-us/library/bb238522).
        */
        void SetVisible(XlSheetVisibility visible);

        /**
        Returns a Walls object that represents the walls of the 3-D chart.

        [MSDN documentation for Chart.Walls](http://msdn.microsoft.com/en-us/library/bb223033).
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
