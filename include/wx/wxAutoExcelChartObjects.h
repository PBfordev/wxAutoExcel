/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Modified by:
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
/////////////////////////////////////////////////////////////////////////////


#ifndef _WXAUTOEXCEL_CHARTOBJECTS_H
#define _WXAUTOEXCEL_CHARTOBJECTS_H

#include "wx/wxAutoExcel_defs.h"

#if WXAUTOEXCEL_USE_CHARTS

#include "wx/wxAutoExcel_object.h"
#include "wx/wxAutoExcel_enums.h"

namespace wxAutoExcel {

    /**
    Represents Microsoft Excel ChartObject object.
    */
   class WXDLLIMPEXP_WXAUTOEXCEL wxExcelChartObject : public wxExcelObject
    {
    public:
        // ***** METHODS *****

        /**
        Makes the current chart the active chart.

        [MSDN documentation for ChartObject.Activate](http://msdn.microsoft.com/en-us/library/bb148254).
        */
        bool Activate();

        /**
        Brings the object to the front of the z-order.

        [MSDN documentation for ChartObject.BringToFront](http://msdn.microsoft.com/en-us/library/bb148260).
        */
        bool BringToFront();

        /**
        Copies the object to the Clipboard.

        [MSDN documentation for ChartObject.Copy](http://msdn.microsoft.com/en-us/library/bb148262).
        */
        bool Copy();

        /**
        Copies the selected object to the Clipboard as a picture.

        [MSDN documentation for ChartObject.CopyPicture](http://msdn.microsoft.com/en-us/library/bb148266).
        */
        bool CopyPicture(XlPictureAppearance* appearance, XlCopyPictureFormat* format);

        /**
        Cuts the object to the Clipboard.

        [MSDN documentation for ChartObject.Cut](http://msdn.microsoft.com/en-us/library/bb148268).
        */
        bool Cut();

        /**
        Deletes the object.

        [MSDN documentation for ChartObject.Delete](http://msdn.microsoft.com/en-us/library/bb148271).
        */
        bool Delete();

        /**
        Duplicates the object and returns a reference to the new copy.

        [MSDN documentation for ChartObject.Duplicate](http://msdn.microsoft.com/en-us/library/bb148275).
        */
        wxExcelObject Duplicate();

        /**
        Selects the object.

        [MSDN documentation for ChartObject.Select](http://msdn.microsoft.com/en-us/library/bb213913).
        */
        bool Select(wxXlTribool replace);

        /**
        Sends the object to the back of the z-order.

        [MSDN documentation for ChartObject.SendToBack](http://msdn.microsoft.com/en-us/library/bb213917).
        */
        bool SendToBack();

        // ***** PROPERTIES *****

        /**
        Returns a Range object that represents the cell that lies under the lower-right corner of the object.

        [MSDN documentation for ChartObject.BottomRightCell](http://msdn.microsoft.com/en-us/library/bb179464).
        */
        wxExcelRange GetBottomRightCell();

        /**
        Returns a Chart object that represents the chart contained in the object. 

        [MSDN documentation for ChartObject.Chart](http://msdn.microsoft.com/en-us/library/bb179465).
        */
        wxExcelChart GetChart();

        /**
        True if the object is enabled.

        [MSDN documentation for ChartObject.Enabled](http://msdn.microsoft.com/en-us/library/bb179467).
        */
        bool GetEnabled();

        /**
        True if the object is enabled.

        [MSDN documentation for ChartObject.Enabled](http://msdn.microsoft.com/en-us/library/bb179467).
        */
        void SetEnabled(bool enabled);

        /**
        Returns a Double value that represents the height, in points, of the object.

        [MSDN documentation for ChartObject.Height](http://msdn.microsoft.com/en-us/library/bb179469).
        */
        double GetHeight();

        /**
        Sets a Double value that represents the height, in points, of the object.

        [MSDN documentation for ChartObject.Height](http://msdn.microsoft.com/en-us/library/bb179469).
        */
        void SetHeight(double height);

        /**
        Returns a Long value that represents the index number of the object within the collection of similar objects.

        [MSDN documentation for ChartObject.Index](http://msdn.microsoft.com/en-us/library/bb179471).
        */
        long GetIndex();

        /**
        Returns a Double value that represents the distance, in points, from the left edge of the object to the left edge of column A (on a worksheet) or the left edge of the chart area (on a chart).

        [MSDN documentation for ChartObject.Left](http://msdn.microsoft.com/en-us/library/bb179473).
        */
        double GetLeft();

        /**
        Sets a Double value that represents the distance, in points, from the left edge of the object to the left edge of column A (on a worksheet) or the left edge of the chart area (on a chart).

        [MSDN documentation for ChartObject.Left](http://msdn.microsoft.com/en-us/library/bb179473).
        */
        void SetLeft(double left);

        /**
        Returns a Boolean value that indicates if the object is locked.

        [MSDN documentation for ChartObject.Locked](http://msdn.microsoft.com/en-us/library/bb212683).
        */
        bool GetLocked();

        /**
        Sets a Boolean value that indicates if the object is locked.

        [MSDN documentation for ChartObject.Locked](http://msdn.microsoft.com/en-us/library/bb212683).
        */
        void SetLocked(bool locked);

        /**
        Returns a String value representing the name of the object.

        [MSDN documentation for ChartObject.Name](http://msdn.microsoft.com/en-us/library/bb212685).
        */
        wxString GetName();

        /**
        Returns a Variant value, containing an XlPlacement constant, that represents the way the object is attached to the cells below it.

        [MSDN documentation for ChartObject.Placement](http://msdn.microsoft.com/en-us/library/bb212688).
        */
        XlPlacement GetPlacement();

        /**
        Sets a Variant value, containing an XlPlacement constant, that represents the way the object is attached to the cells below it.

        [MSDN documentation for ChartObject.Placement](http://msdn.microsoft.com/en-us/library/bb212688).
        */
        void SetPlacement(XlPlacement placement);

        /**
        True if the object will be printed when the document is printed.

        [MSDN documentation for ChartObject.PrintObject](http://msdn.microsoft.com/en-us/library/bb212690).
        */
        bool GetPrintObject();

        /**
        True if the object will be printed when the document is printed.

        [MSDN documentation for ChartObject.PrintObject](http://msdn.microsoft.com/en-us/library/bb212690).
        */
        void SetPrintObject(bool printObject);

        /**
        True if the embedded chart frame cannot be moved, resized, or deleted through the user interface.

        [MSDN documentation for ChartObject.ProtectChartObject](http://msdn.microsoft.com/en-us/library/bb209031).
        */
        bool GetProtectChartObject();

        /**
        True if the embedded chart frame cannot be moved, resized, or deleted through the user interface.

        [MSDN documentation for ChartObject.ProtectChartObject](http://msdn.microsoft.com/en-us/library/bb209031).
        */
        void SetProtectChartObject(bool protectChartObject);

        /**
        True if the embedded chart has rounded corners.

        [MSDN documentation for ChartObject.RoundedCorners](http://msdn.microsoft.com/en-us/library/bb148872).
        */
        bool GetRoundedCorners();

        /**
        True if the embedded chart has rounded corners.

        [MSDN documentation for ChartObject.RoundedCorners](http://msdn.microsoft.com/en-us/library/bb148872).
        */
        void SetRoundedCorners(bool roundedCorners);

        /**
        Returns a Boolean value that determines if the font is a shadow font or if the object has a shadow.

        [MSDN documentation for ChartObject.Shadow](http://msdn.microsoft.com/en-us/library/bb148876).
        */
        bool GetShadow();

        /**
        Sets a Boolean value that determines if the font is a shadow font or if the object has a shadow.

        [MSDN documentation for ChartObject.Shadow](http://msdn.microsoft.com/en-us/library/bb148876).
        */
        void SetShadow(bool shadow);

#if WXAUTOEXCEL_USE_SHAPES
        /**
        Returns a ShapeRange object that represents the specified object or objects.

        [MSDN documentation for ChartObject.ShapeRange](http://msdn.microsoft.com/en-us/library/bb148879).
        */
        wxExcelShapeRange GetShapeRange();
#endif // #if WXAUTOEXCEL_USE_SHAPES

        /**
        Returns a Double value that represents the distance, in points, from the top edge of the object to the top of row 1 (on a worksheet) or the top of the chart area (on a chart).

        [MSDN documentation for ChartObject.Top](http://msdn.microsoft.com/en-us/library/bb238436).
        */
        double GetTop();

        /**
        Sets a Double value that represents the distance, in points, from the top edge of the object to the top of row 1 (on a worksheet) or the top of the chart area (on a chart).

        [MSDN documentation for ChartObject.Top](http://msdn.microsoft.com/en-us/library/bb238436).
        */
        void SetTop(double top);

        /**
        Returns a Range object that represents the cell that lies under the upper-left corner of the specified object.

        [MSDN documentation for ChartObject.TopLeftCell](http://msdn.microsoft.com/en-us/library/bb238442).
        */
        wxExcelRange GetTopLeftCell();

        /**
        Returns a Boolean value that determines whether the object is visible.

        [MSDN documentation for ChartObject.Visible](http://msdn.microsoft.com/en-us/library/bb238452).
        */
        bool GetVisible();

        /**
        Sets a Boolean value that determines whether the object is visible.

        [MSDN documentation for ChartObject.Visible](http://msdn.microsoft.com/en-us/library/bb238452).
        */
        void SetVisible(bool visible);

        /**
        Returns a Double value that represents the width, in points, of the object.

        [MSDN documentation for ChartObject.Width](http://msdn.microsoft.com/en-us/library/bb238458).
        */
        double GetWidth();

        /**
        Sets a Double value that represents the width, in points, of the object.

        [MSDN documentation for ChartObject.Width](http://msdn.microsoft.com/en-us/library/bb238458).
        */
        void SetWidth(double width);

        /**
        Returns the z-order position of the object.

        [MSDN documentation for ChartObject.ZOrder](http://msdn.microsoft.com/en-us/library/bb238465).
        */
        long GetZOrder();

        /**
        Returns "ChartObject".
        */
        virtual wxString GetAutoExcelObjectName_() const { return wxS("ChartObject"); }
    };

    /**
    Represents Microsoft Excel ChartObjects collection.
    */
   class WXDLLIMPEXP_WXAUTOEXCEL wxExcelChartObjects : public wxExcelObject
    {
    public:
        // ***** METHODS *****

        /**
        Creates a new embedded chart.

        [MSDN documentation for ChartObjects.Add](http://msdn.microsoft.com/en-us/library/bb148425).
        */
        wxExcelChartObject Add(double left, double top, double width, double height);

        /**
        Copies the object to the Clipboard.

        [MSDN documentation for ChartObjects.Copy](http://msdn.microsoft.com/en-us/library/bb148431).
        */
        bool Copy();

        /**
        Copies the selected object to the Clipboard as a picture. Variant.

        [MSDN documentation for ChartObjects.CopyPicture](http://msdn.microsoft.com/en-us/library/bb148434).
        */
        bool CopyPicture(XlPictureAppearance* appearance, XlCopyPictureFormat* format);

        /**
        Cuts the object to the Clipboard.

        [MSDN documentation for ChartObjects.Cut](http://msdn.microsoft.com/en-us/library/bb211647).
        */
        bool Cut();

        /**
        Deletes the object.

        [MSDN documentation for ChartObject.Delete](http://msdn.microsoft.com/en-us/library/bb148271).
        */
        bool Delete();

        /**
        Duplicates the object and returns a reference to the new copy.

        [MSDN documentation for ChartObjects.Duplicate](http://msdn.microsoft.com/en-us/library/bb211654).
        */
        wxExcelChartObjects Duplicate();

        //@{
        /**
        Returns a single object from a collection.

        [MSDN documentation for ChartObjects.Item](http://msdn.microsoft.com/en-us/library/bb211657).
        */
        wxExcelChartObject Item(long index);
        wxExcelChartObject operator[](long index);
        wxExcelChartObject Item(const wxString& name);
        wxExcelChartObject operator[](const wxString& name);
        //@}

        /**
        Selects the object.

        [MSDN documentation for ChartObjects.Select](http://msdn.microsoft.com/en-us/library/bb213926).
        */
        bool Select(wxXlTribool replace);

        // ***** PROPERTIES *****

        /**
        Returns a Long value that represents the number of objects in the collection.

        [MSDN documentation for ChartObjects.Count](http://msdn.microsoft.com/en-us/library/bb212693).
        */
        long GetCount();

        /**
        True if the object is enabled.

        [MSDN documentation for ChartObjects.Enabled](http://msdn.microsoft.com/en-us/library/bb212697).
        */
        bool GetEnabled();

        /**
        True if the object is enabled.

        [MSDN documentation for ChartObjects.Enabled](http://msdn.microsoft.com/en-us/library/bb212697).
        */
        void SetEnabled(bool enabled);

        /**
        Returns a Double value that represents the height, in points, of the object.

        [MSDN documentation for ChartObjects.Height](http://msdn.microsoft.com/en-us/library/bb212699).
        */
        double GetHeight();

        /**
        Sets a Double value that represents the height, in points, of the object.

        [MSDN documentation for ChartObjects.Height](http://msdn.microsoft.com/en-us/library/bb212699).
        */
        void SetHeight(double height);

        /**
        Returns a Double value that represents the distance, in points, from the left edge of the object to the left edge of column A (on a worksheet) or the left edge of the chart area (on a chart).

        [MSDN documentation for ChartObjects.Left](http://msdn.microsoft.com/en-us/library/bb212700).
        */
        double GetLeft();

        /**
        Sets a Double value that represents the distance, in points, from the left edge of the object to the left edge of column A (on a worksheet) or the left edge of the chart area (on a chart).

        [MSDN documentation for ChartObjects.Left](http://msdn.microsoft.com/en-us/library/bb212700).
        */
        void SetLeft(double left);        

        /**
        True if the embedded chart frame cannot be moved, resized, or deleted through the user interface. Since Excel 2007.

        [MSDN documentation for ChartObjects.ProtectChartObject](http://msdn.microsoft.com/en-us/library/bb242557).
        */
        bool GetProtectChartObject();

        /**
        True if the embedded chart frame cannot be moved, resized, or deleted through the user interface. Since Excel 2007.

        [MSDN documentation for ChartObjects.ProtectChartObject](http://msdn.microsoft.com/en-us/library/bb242557).
        */
        void SetProtectChartObject(bool protectChartObject);

#if WXAUTOEXCEL_USE_SHAPES
        /**
        Returns a ShapeRange object that represents the specified object or objects.

        [MSDN documentation for ChartObjects.ShapeRange](http://msdn.microsoft.com/en-us/library/bb238402).
        */
        wxExcelShapeRange GetShapeRange();
#endif // #if WXAUTOEXCEL_USE_SHAPES

        /**
        Returns a Double value that represents the distance, in points, from the top edge of the object to the top of row 1 (on a worksheet) or the top of the chart area (on a chart).

        [MSDN documentation for ChartObjects.Top](http://msdn.microsoft.com/en-us/library/bb238410).
        */
        double GetTop();

        /**
        Sets a Double value that represents the distance, in points, from the top edge of the object to the top of row 1 (on a worksheet) or the top of the chart area (on a chart).

        [MSDN documentation for ChartObjects.Top](http://msdn.microsoft.com/en-us/library/bb238410).
        */
        void SetTop(double top);

        /**
        Returns a Boolean value that determines whether the object is visible.

        [MSDN documentation for ChartObjects.Visible](http://msdn.microsoft.com/en-us/library/bb238417).
        */
        bool GetVisible();

        /**
        Sets a Boolean value that determines whether the object is visible.

        [MSDN documentation for ChartObjects.Visible](http://msdn.microsoft.com/en-us/library/bb238417).
        */
        void SetVisible(bool visible);

        /**
        Returns a Double value that represents the width, in points, of the object.

        [MSDN documentation for ChartObjects.Width](http://msdn.microsoft.com/en-us/library/bb238423).
        */
        double GetWidth();

        /**
        Sets a Double value that represents the width, in points, of the object.

        [MSDN documentation for ChartObjects.Width](http://msdn.microsoft.com/en-us/library/bb238423).
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
