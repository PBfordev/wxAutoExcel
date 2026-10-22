/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// License:     MIT license
/////////////////////////////////////////////////////////////////////////////


#ifndef _WXAUTOEXCEL_LINEFORMAT_H
#define _WXAUTOEXCEL_LINEFORMAT_H

#include "wx/wxAutoExcel_defs.h"

#if WXAUTOEXCEL_USE_SHAPES || WXAUTOEXCEL_USE_CHARTS

#include "wx/wxAutoExcel_object.h"
#include "wx/wxAutoExcel_enums.h"

namespace wxAutoExcel {
    /**
    @brief Represents Microsoft Excel LineFormat object.
    */
   class WXDLLIMPEXP_WXAUTOEXCEL wxExcelLineFormat  : public wxExcelObject
    {
    public:

        // ***** PROPERTIES *****

        /**
        Returns a ColorFormat Represents the specified fill background color.

        [MSDN documentation for LineFormat.BackColor](https://learn.microsoft.com/en-us/office/vba/api/excel.lineformat.backcolor).
        */
        wxExcelColorFormat GetBackColor();

        /**
        Returns the length of the arrowhead at the beginning of the specified line. Read/write MsoArrowheadLength.

        [MSDN documentation for LineFormat.BeginArrowheadLength](https://learn.microsoft.com/en-us/office/vba/api/excel.lineformat.beginarrowheadlength).
        */
        MsoArrowheadLength GetBeginArrowheadLength();

        /**
        Sets the length of the arrowhead at the beginning of the specified line. Read/write MsoArrowheadLength.

        [MSDN documentation for LineFormat.BeginArrowheadLength](https://learn.microsoft.com/en-us/office/vba/api/excel.lineformat.beginarrowheadlength).
        */
        void SetBeginArrowheadLength(MsoArrowheadLength beginArrowheadLength);

        /**
        Returns the style of the arrowhead at the beginning of the specified line. Read/write MsoArrowheadStyle.

        [MSDN documentation for LineFormat.BeginArrowheadStyle](https://learn.microsoft.com/en-us/office/vba/api/excel.lineformat.beginarrowheadstyle).
        */
        MsoArrowheadStyle GetBeginArrowheadStyle();

        /**
        Sets the style of the arrowhead at the beginning of the specified line. Read/write MsoArrowheadStyle.

        [MSDN documentation for LineFormat.BeginArrowheadStyle](https://learn.microsoft.com/en-us/office/vba/api/excel.lineformat.beginarrowheadstyle).
        */
        void SetBeginArrowheadStyle(MsoArrowheadStyle beginArrowheadStyle);

        /**
        Returns the width of the arrowhead at the beginning of the specified line. Read/write MsoArrowheadWidth.

        [MSDN documentation for LineFormat.BeginArrowheadWidth](https://learn.microsoft.com/en-us/office/vba/api/excel.lineformat.beginarrowheadwidth).
        */
        MsoArrowheadWidth GetBeginArrowheadWidth();

        /**
        Sets the width of the arrowhead at the beginning of the specified line. Read/write MsoArrowheadWidth.

        [MSDN documentation for LineFormat.BeginArrowheadWidth](https://learn.microsoft.com/en-us/office/vba/api/excel.lineformat.beginarrowheadwidth).
        */
        void SetBeginArrowheadWidth(MsoArrowheadWidth beginArrowheadWidth);

        /**
        Returns the dash style for the specified line. Can be one of the MsoLineDashStyle contants.

        [MSDN documentation for LineFormat.DashStyle](https://learn.microsoft.com/en-us/office/vba/api/excel.lineformat.dashstyle).
        */
        MsoLineDashStyle GetDashStyle();

        /**
        Sets the dash style for the specified line. Can be one of the MsoLineDashStyle contants.

        [MSDN documentation for LineFormat.DashStyle](https://learn.microsoft.com/en-us/office/vba/api/excel.lineformat.dashstyle).
        */
        void SetDashStyle(MsoLineDashStyle dashStyle);

        /**
        Returns the length of the arrowhead at the end of the specified line. Read/write MsoArrowheadLength.

        [MSDN documentation for LineFormat.EndArrowheadLength](https://learn.microsoft.com/en-us/office/vba/api/excel.lineformat.endarrowheadlength).
        */
        MsoArrowheadLength GetEndArrowheadLength();

        /**
        Sets the length of the arrowhead at the end of the specified line. Read/write MsoArrowheadLength.

        [MSDN documentation for LineFormat.EndArrowheadLength](https://learn.microsoft.com/en-us/office/vba/api/excel.lineformat.endarrowheadlength).
        */
        void SetEndArrowheadLength(MsoArrowheadLength endArrowheadLength);

        /**
        Returns the style of the arrowhead at the end of the specified line. Read/write MsoArrowheadStyle.

        [MSDN documentation for LineFormat.EndArrowheadStyle](https://learn.microsoft.com/en-us/office/vba/api/excel.lineformat.endarrowheadstyle).
        */
        MsoArrowheadStyle GetEndArrowheadStyle();

        /**
        Sets the style of the arrowhead at the end of the specified line. Read/write MsoArrowheadStyle.

        [MSDN documentation for LineFormat.EndArrowheadStyle](https://learn.microsoft.com/en-us/office/vba/api/excel.lineformat.endarrowheadstyle).
        */
        void SetEndArrowheadStyle(MsoArrowheadStyle endArrowheadStyle);

        /**
        Returns the width of the arrowhead at the end of the specified line. Read/write MsoArrowheadWidth.

        [MSDN documentation for LineFormat.EndArrowheadWidth](https://learn.microsoft.com/en-us/office/vba/api/excel.lineformat.endarrowheadwidth).
        */
        MsoArrowheadWidth GetEndArrowheadWidth();

        /**
        Sets the width of the arrowhead at the end of the specified line. Read/write MsoArrowheadWidth.

        [MSDN documentation for LineFormat.EndArrowheadWidth](https://learn.microsoft.com/en-us/office/vba/api/excel.lineformat.endarrowheadwidth).
        */
        void SetEndArrowheadWidth(MsoArrowheadWidth endArrowheadWidth);

        /**
        Returns a ColorFormat Represents the specified foreground fill or solid color.

        [MSDN documentation for LineFormat.ForeColor](https://learn.microsoft.com/en-us/office/vba/api/excel.lineformat.forecolor).
        */
        wxExcelColorFormat GetForeColor();

        /**
        Sets a ColorFormat Represents the specified foreground fill or solid color.

        [MSDN documentation for LineFormat.ForeColor](https://learn.microsoft.com/en-us/office/vba/api/excel.lineformat.forecolor).
        */
        void SetForeColor(const wxExcelColorFormat& foreColor);


         /**
        Whether lines are drawn inside the specified shape's boundaries. Since Excel 2010.

        [MSDN documentation for LineFormat.InsetPen](https://learn.microsoft.com/en-us/office/vba/api/excel.lineformat.insetpen).
        */
        MsoTriState GetInsetPen();

        /**
        Whether lines are drawn inside the specified shape's boundaries. Since Excel 2010.

        [MSDN documentation for LineFormat.InsetPen](https://learn.microsoft.com/en-us/office/vba/api/excel.lineformat.insetpen).
        */
        void SetInsetPen(MsoTriState insetPen);


        /**
        Returns an MsoPatternType value that represents the fill pattern.

        [MSDN documentation for LineFormat.Pattern](https://learn.microsoft.com/en-us/office/vba/api/excel.lineformat.pattern).
        */
        MsoPatternType GetPattern();

        /**
        Sets an MsoPatternType value that represents the fill pattern.

        [MSDN documentation for LineFormat.Pattern](https://learn.microsoft.com/en-us/office/vba/api/excel.lineformat.pattern).
        */
        void SetPattern(MsoPatternType pattern);

        /**
        Returns a MsoLineStyle value that represents the style of the line.

        [MSDN documentation for LineFormat.Style](https://learn.microsoft.com/en-us/office/vba/api/excel.lineformat.style).
        */
        MsoLineStyle  GetStyle();

        /**
        Sets a MsoLineStyle value that represents the style of the line.

        [MSDN documentation for LineFormat.Style](https://learn.microsoft.com/en-us/office/vba/api/excel.lineformat.style).
        */
        void SetStyle(MsoLineStyle  style);

        /**
        Returns the degree of transparency of the specified fill as a value from 0.0 (opaque) through 1.0 (clear). Read/write Double.

        [MSDN documentation for LineFormat.Transparency](https://learn.microsoft.com/en-us/office/vba/api/excel.lineformat.transparency).
        */
        double GetTransparency();

        /**
        Sets the degree of transparency of the specified fill as a value from 0.0 (opaque) through 1.0 (clear). Read/write Double.

        [MSDN documentation for LineFormat.Transparency](https://learn.microsoft.com/en-us/office/vba/api/excel.lineformat.transparency).
        */
        void SetTransparency(double transparency);

        /**
        Returns a MsoTriState value that determines whether the object is visible.

        [MSDN documentation for LineFormat.Visible](https://learn.microsoft.com/en-us/office/vba/api/excel.lineformat.visible).
        */
        MsoTriState GetVisible();

        /**
        Sets a MsoTriState value that determines whether the object is visible.

        [MSDN documentation for LineFormat.Visible](https://learn.microsoft.com/en-us/office/vba/api/excel.lineformat.visible).
        */
        void SetVisible(MsoTriState visible);

        /**
        Returns a Single value that represents the weight of the line.

        [MSDN documentation for LineFormat.Weight](https://learn.microsoft.com/en-us/office/vba/api/excel.lineformat.weight).
        */
        double GetWeight();

        /**
        Sets a Single value that represents the weight of the line.

        [MSDN documentation for LineFormat.Weight](https://learn.microsoft.com/en-us/office/vba/api/excel.lineformat.weight).
        */
        void SetWeight(double weight);

        /**
        Returns "LineFormat".
        */
        virtual wxString GetAutoExcelObjectName_() const { return wxS("LineFormat"); }
    };


} // namespace wxAutoExcel

#endif // #if WXAUTOEXCEL_USE_SHAPES || WXAUTOEXCEL_USE_CHARTS

#endif //_WXAUTOEXCEL_LINEFORMAT_H
