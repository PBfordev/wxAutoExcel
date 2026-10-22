/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// License:     MIT license
/////////////////////////////////////////////////////////////////////////////


#ifndef _WXAUTOEXCEL_CHARTOBJECTS_H
#define _WXAUTOEXCEL_CHARTOBJECTS_H

#include "wx/wxAutoExcel_defs.h"

#if WXAUTOEXCEL_USE_CHARTS

#include "wx/wxAutoExcel_object.h"
#include "wx/wxAutoExcel_enums.h"

namespace wxAutoExcel {

    /**
    @brief Represents Microsoft Excel ChartObject object.
    */
   class WXDLLIMPEXP_WXAUTOEXCEL wxExcelChartObject : public wxExcelObject
    {
    public:
        // ***** METHODS *****

        /**
        Makes the current chart the active chart.

        [MSDN documentation for ChartObject.Activate](https://learn.microsoft.com/en-us/office/vba/api/excel.chartobject.activate).
        */
        bool Activate();

        /**
        Brings the object to the front of the z-order.

        [MSDN documentation for ChartObject.BringToFront](https://learn.microsoft.com/en-us/office/vba/api/excel.chartobject.bringtofront).
        */
        bool BringToFront();

        /**
        Copies the object to the Clipboard.

        [MSDN documentation for ChartObject.Copy](https://learn.microsoft.com/en-us/office/vba/api/excel.chartobject.copy).
        */
        bool Copy();

        /**
        Copies the selected object to the Clipboard as a picture.

        [MSDN documentation for ChartObject.CopyPicture](https://learn.microsoft.com/en-us/office/vba/api/excel.chartobject.copypicture).
        */
        bool CopyPicture(XlPictureAppearance* appearance, XlCopyPictureFormat* format);

        /**
        Cuts the object to the Clipboard.

        [MSDN documentation for ChartObject.Cut](https://learn.microsoft.com/en-us/office/vba/api/excel.chartobject.cut).
        */
        bool Cut();

        /**
        Deletes the object.

        [MSDN documentation for ChartObject.Delete](https://learn.microsoft.com/en-us/office/vba/api/excel.chartobject.delete).
        */
        bool Delete();

        /**
        Duplicates the object and returns a reference to the new copy.

        [MSDN documentation for ChartObject.Duplicate](https://learn.microsoft.com/en-us/office/vba/api/excel.chartobject.duplicate).
        */
        wxExcelObject Duplicate();

        /**
        Selects the object.

        [MSDN documentation for ChartObject.Select](https://learn.microsoft.com/en-us/office/vba/api/excel.chartobject.select).
        */
        bool Select(wxXlTribool replace);

        /**
        Sends the object to the back of the z-order.

        [MSDN documentation for ChartObject.SendToBack](https://learn.microsoft.com/en-us/office/vba/api/excel.chartobject.sendtoback).
        */
        bool SendToBack();

        // ***** PROPERTIES *****

        /**
        Returns a Range object that represents the cell that lies under the lower-right corner of the object.

        [MSDN documentation for ChartObject.BottomRightCell](https://learn.microsoft.com/en-us/office/vba/api/excel.chartobject.bottomrightcell).
        */
        wxExcelRange GetBottomRightCell();

        /**
        Returns a Chart object that represents the chart contained in the object.

        [MSDN documentation for ChartObject.Chart](https://learn.microsoft.com/en-us/office/vba/api/excel.chartobject.chart).
        */
        wxExcelChart GetChart();

        /**
        True if the object is enabled.

        [MSDN documentation for ChartObject.Enabled](https://learn.microsoft.com/en-us/previous-versions/office/developer/office-2007/bb179467%28v=office.12%29).
        */
        bool GetEnabled();

        /**
        True if the object is enabled.

        [MSDN documentation for ChartObject.Enabled](https://learn.microsoft.com/en-us/previous-versions/office/developer/office-2007/bb179467%28v=office.12%29).
        */
        void SetEnabled(bool enabled);

        /**
        Returns a Double value that represents the height, in points, of the object.

        [MSDN documentation for ChartObject.Height](https://learn.microsoft.com/en-us/office/vba/api/excel.chartobject.height).
        */
        double GetHeight();

        /**
        Sets a Double value that represents the height, in points, of the object.

        [MSDN documentation for ChartObject.Height](https://learn.microsoft.com/en-us/office/vba/api/excel.chartobject.height).
        */
        void SetHeight(double height);

        /**
        Returns a Long value that represents the index number of the object within the collection of similar objects.

        [MSDN documentation for ChartObject.Index](https://learn.microsoft.com/en-us/office/vba/api/excel.chartobject.index).
        */
        long GetIndex();

        /**
        Returns a Double value that represents the distance, in points, from the left edge of the object to the left edge of column A (on a worksheet) or the left edge of the chart area (on a chart).

        [MSDN documentation for ChartObject.Left](https://learn.microsoft.com/en-us/office/vba/api/excel.chartobject.left).
        */
        double GetLeft();

        /**
        Sets a Double value that represents the distance, in points, from the left edge of the object to the left edge of column A (on a worksheet) or the left edge of the chart area (on a chart).

        [MSDN documentation for ChartObject.Left](https://learn.microsoft.com/en-us/office/vba/api/excel.chartobject.left).
        */
        void SetLeft(double left);

        /**
        Returns a Boolean value that indicates if the object is locked.

        [MSDN documentation for ChartObject.Locked](https://learn.microsoft.com/en-us/office/vba/api/excel.chartobject.locked).
        */
        bool GetLocked();

        /**
        Sets a Boolean value that indicates if the object is locked.

        [MSDN documentation for ChartObject.Locked](https://learn.microsoft.com/en-us/office/vba/api/excel.chartobject.locked).
        */
        void SetLocked(bool locked);

        /**
        Returns a String value representing the name of the object.

        [MSDN documentation for ChartObject.Name](https://learn.microsoft.com/en-us/office/vba/api/excel.chartobject.name).
        */
        wxString GetName();

        /**
        Returns a Variant value, containing an XlPlacement constant, that represents the way the object is attached to the cells below it.

        [MSDN documentation for ChartObject.Placement](https://learn.microsoft.com/en-us/office/vba/api/excel.chartobject.placement).
        */
        XlPlacement GetPlacement();

        /**
        Sets a Variant value, containing an XlPlacement constant, that represents the way the object is attached to the cells below it.

        [MSDN documentation for ChartObject.Placement](https://learn.microsoft.com/en-us/office/vba/api/excel.chartobject.placement).
        */
        void SetPlacement(XlPlacement placement);

        /**
        True if the object will be printed when the document is printed.

        [MSDN documentation for ChartObject.PrintObject](https://learn.microsoft.com/en-us/office/vba/api/excel.chartobject.printobject).
        */
        bool GetPrintObject();

        /**
        True if the object will be printed when the document is printed.

        [MSDN documentation for ChartObject.PrintObject](https://learn.microsoft.com/en-us/office/vba/api/excel.chartobject.printobject).
        */
        void SetPrintObject(bool printObject);

        /**
        True if the embedded chart frame cannot be moved, resized, or deleted through the user interface.

        [MSDN documentation for ChartObject.ProtectChartObject](https://learn.microsoft.com/en-us/office/vba/api/excel.chartobject.protectchartobject).
        */
        bool GetProtectChartObject();

        /**
        True if the embedded chart frame cannot be moved, resized, or deleted through the user interface.

        [MSDN documentation for ChartObject.ProtectChartObject](https://learn.microsoft.com/en-us/office/vba/api/excel.chartobject.protectchartobject).
        */
        void SetProtectChartObject(bool protectChartObject);

        /**
        True if the embedded chart has rounded corners.

        [MSDN documentation for ChartObject.RoundedCorners](https://learn.microsoft.com/en-us/office/vba/api/excel.chartobject.roundedcorners).
        */
        bool GetRoundedCorners();

        /**
        True if the embedded chart has rounded corners.

        [MSDN documentation for ChartObject.RoundedCorners](https://learn.microsoft.com/en-us/office/vba/api/excel.chartobject.roundedcorners).
        */
        void SetRoundedCorners(bool roundedCorners);

        /**
        Returns a Boolean value that determines if the font is a shadow font or if the object has a shadow.

        [MSDN documentation for ChartObject.Shadow](https://learn.microsoft.com/en-us/office/vba/api/excel.chartobject.shadow).
        */
        bool GetShadow();

        /**
        Sets a Boolean value that determines if the font is a shadow font or if the object has a shadow.

        [MSDN documentation for ChartObject.Shadow](https://learn.microsoft.com/en-us/office/vba/api/excel.chartobject.shadow).
        */
        void SetShadow(bool shadow);

#if WXAUTOEXCEL_USE_SHAPES
        /**
        Returns a ShapeRange object that represents the specified object or objects.

        [MSDN documentation for ChartObject.ShapeRange](https://learn.microsoft.com/en-us/office/vba/api/excel.chartobject.shaperange).
        */
        wxExcelShapeRange GetShapeRange();
#endif // #if WXAUTOEXCEL_USE_SHAPES

        /**
        Returns a Double value that represents the distance, in points, from the top edge of the object to the top of row 1 (on a worksheet) or the top of the chart area (on a chart).

        [MSDN documentation for ChartObject.Top](https://learn.microsoft.com/en-us/office/vba/api/excel.chartobject.top).
        */
        double GetTop();

        /**
        Sets a Double value that represents the distance, in points, from the top edge of the object to the top of row 1 (on a worksheet) or the top of the chart area (on a chart).

        [MSDN documentation for ChartObject.Top](https://learn.microsoft.com/en-us/office/vba/api/excel.chartobject.top).
        */
        void SetTop(double top);

        /**
        Returns a Range object that represents the cell that lies under the upper-left corner of the specified object.

        [MSDN documentation for ChartObject.TopLeftCell](https://learn.microsoft.com/en-us/office/vba/api/excel.chartobject.topleftcell).
        */
        wxExcelRange GetTopLeftCell();

        /**
        Returns a Boolean value that determines whether the object is visible.

        [MSDN documentation for ChartObject.Visible](https://learn.microsoft.com/en-us/office/vba/api/excel.chartobject.visible).
        */
        bool GetVisible();

        /**
        Sets a Boolean value that determines whether the object is visible.

        [MSDN documentation for ChartObject.Visible](https://learn.microsoft.com/en-us/office/vba/api/excel.chartobject.visible).
        */
        void SetVisible(bool visible);

        /**
        Returns a Double value that represents the width, in points, of the object.

        [MSDN documentation for ChartObject.Width](https://learn.microsoft.com/en-us/office/vba/api/excel.chartobject.width).
        */
        double GetWidth();

        /**
        Sets a Double value that represents the width, in points, of the object.

        [MSDN documentation for ChartObject.Width](https://learn.microsoft.com/en-us/office/vba/api/excel.chartobject.width).
        */
        void SetWidth(double width);

        /**
        Returns the z-order position of the object.

        [MSDN documentation for ChartObject.ZOrder](https://learn.microsoft.com/en-us/office/vba/api/excel.chartobject.zorder).
        */
        long GetZOrder();

        /**
        Returns "ChartObject".
        */
        virtual wxString GetAutoExcelObjectName_() const { return wxS("ChartObject"); }
    };

    /**
    @brief Represents Microsoft Excel ChartObjects collection.
    */
   class WXDLLIMPEXP_WXAUTOEXCEL wxExcelChartObjects : public wxExcelObject
    {
    public:
        // ***** METHODS *****

        /**
        Creates a new embedded chart.

        [MSDN documentation for ChartObjects.Add](https://learn.microsoft.com/en-us/office/vba/api/excel.chartobjects.add).
        */
        wxExcelChartObject Add(double left, double top, double width, double height);

        /**
        Copies the object to the Clipboard.

        [MSDN documentation for ChartObjects.Copy](https://learn.microsoft.com/en-us/office/vba/api/excel.chartobjects.copy).
        */
        bool Copy();

        /**
        Copies the selected object to the Clipboard as a picture. Variant.

        [MSDN documentation for ChartObjects.CopyPicture](https://learn.microsoft.com/en-us/office/vba/api/excel.chartobjects.copypicture).
        */
        bool CopyPicture(XlPictureAppearance* appearance, XlCopyPictureFormat* format);

        /**
        Cuts the object to the Clipboard.

        [MSDN documentation for ChartObjects.Cut](https://learn.microsoft.com/en-us/office/vba/api/excel.chartobjects.cut).
        */
        bool Cut();

        /**
        Deletes the object.

        [MSDN documentation for ChartObjects.Delete](https://learn.microsoft.com/en-us/office/vba/api/excel.chartobjects.delete).
        */
        bool Delete();

        /**
        Duplicates the object and returns a reference to the new copy.

        [MSDN documentation for ChartObjects.Duplicate](https://learn.microsoft.com/en-us/office/vba/api/excel.chartobjects.duplicate).
        */
        wxExcelChartObjects Duplicate();

        //@{
        /**
        Returns a single object from a collection.

        [MSDN documentation for ChartObjects.Item](https://learn.microsoft.com/en-us/office/vba/api/excel.chartobjects.item).
        */
        wxExcelChartObject Item(long index);
        wxExcelChartObject operator[](long index);
        wxExcelChartObject Item(const wxString& name);
        wxExcelChartObject operator[](const wxString& name);
        //@}

        /**
        Selects the object.

        [MSDN documentation for ChartObjects.Select](https://learn.microsoft.com/en-us/office/vba/api/excel.chartobjects.select).
        */
        bool Select(wxXlTribool replace);

        // ***** PROPERTIES *****

        /**
        Returns a Long value that represents the number of objects in the collection.

        [MSDN documentation for ChartObjects.Count](https://learn.microsoft.com/en-us/office/vba/api/excel.chartobjects.count).
        */
        long GetCount();

        /**
        True if the object is enabled.

        [MSDN documentation for ChartObjects.Enabled](https://learn.microsoft.com/en-us/previous-versions/office/developer/office-2007/bb212697%28v=office.12%29).
        */
        bool GetEnabled();

        /**
        True if the object is enabled.

        [MSDN documentation for ChartObjects.Enabled](https://learn.microsoft.com/en-us/previous-versions/office/developer/office-2007/bb212697%28v=office.12%29).
        */
        void SetEnabled(bool enabled);

        /**
        Returns a Double value that represents the height, in points, of the object.

        [MSDN documentation for ChartObjects.Height](https://learn.microsoft.com/en-us/office/vba/api/excel.chartobjects.height).
        */
        double GetHeight();

        /**
        Sets a Double value that represents the height, in points, of the object.

        [MSDN documentation for ChartObjects.Height](https://learn.microsoft.com/en-us/office/vba/api/excel.chartobjects.height).
        */
        void SetHeight(double height);

        /**
        Returns a Double value that represents the distance, in points, from the left edge of the object to the left edge of column A (on a worksheet) or the left edge of the chart area (on a chart).

        [MSDN documentation for ChartObjects.Left](https://learn.microsoft.com/en-us/office/vba/api/excel.chartobjects.left).
        */
        double GetLeft();

        /**
        Sets a Double value that represents the distance, in points, from the left edge of the object to the left edge of column A (on a worksheet) or the left edge of the chart area (on a chart).

        [MSDN documentation for ChartObjects.Left](https://learn.microsoft.com/en-us/office/vba/api/excel.chartobjects.left).
        */
        void SetLeft(double left);

        /**
        True if the embedded chart frame cannot be moved, resized, or deleted through the user interface. Since Excel 2007.

        [MSDN documentation for ChartObjects.ProtectChartObject](https://learn.microsoft.com/en-us/office/vba/api/excel.chartobjects.protectchartobject).
        */
        bool GetProtectChartObject();

        /**
        True if the embedded chart frame cannot be moved, resized, or deleted through the user interface. Since Excel 2007.

        [MSDN documentation for ChartObjects.ProtectChartObject](https://learn.microsoft.com/en-us/office/vba/api/excel.chartobjects.protectchartobject).
        */
        void SetProtectChartObject(bool protectChartObject);

#if WXAUTOEXCEL_USE_SHAPES
        /**
        Returns a ShapeRange object that represents the specified object or objects.

        [MSDN documentation for ChartObjects.ShapeRange](https://learn.microsoft.com/en-us/office/vba/api/excel.chartobjects.shaperange).
        */
        wxExcelShapeRange GetShapeRange();
#endif // #if WXAUTOEXCEL_USE_SHAPES

        /**
        Returns a Double value that represents the distance, in points, from the top edge of the object to the top of row 1 (on a worksheet) or the top of the chart area (on a chart).

        [MSDN documentation for ChartObjects.Top](https://learn.microsoft.com/en-us/office/vba/api/excel.chartobjects.top).
        */
        double GetTop();

        /**
        Sets a Double value that represents the distance, in points, from the top edge of the object to the top of row 1 (on a worksheet) or the top of the chart area (on a chart).

        [MSDN documentation for ChartObjects.Top](https://learn.microsoft.com/en-us/office/vba/api/excel.chartobjects.top).
        */
        void SetTop(double top);

        /**
        Returns a Boolean value that determines whether the object is visible.

        [MSDN documentation for ChartObjects.Visible](https://learn.microsoft.com/en-us/office/vba/api/excel.chartobjects.visible).
        */
        bool GetVisible();

        /**
        Sets a Boolean value that determines whether the object is visible.

        [MSDN documentation for ChartObjects.Visible](https://learn.microsoft.com/en-us/office/vba/api/excel.chartobjects.visible).
        */
        void SetVisible(bool visible);

        /**
        Returns a Double value that represents the width, in points, of the object.

        [MSDN documentation for ChartObjects.Width](https://learn.microsoft.com/en-us/office/vba/api/excel.chartobjects.width).
        */
        double GetWidth();

        /**
        Sets a Double value that represents the width, in points, of the object.

        [MSDN documentation for ChartObjects.Width](https://learn.microsoft.com/en-us/office/vba/api/excel.chartobjects.width).
        */
        void SetWidth(double width);

        /**
        Returns "ChartObjects".
        */
        virtual wxString GetAutoExcelObjectName_() const { return wxS("ChartObjects"); }
    };


} // namespace wxAutoExcel

#endif // #if WXAUTOEXCEL_USE_CHARTS

#endif // _WXAUTOEXCEL_CHARTOBJECTS_H
