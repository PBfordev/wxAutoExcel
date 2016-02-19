/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Modified by:
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
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

        [MSDN documentation for LineFormat.BackColor](http://msdn.microsoft.com/en-us/library/bb148588).
        */
        wxExcelColorFormat GetBackColor();

        /**
        Returns the length of the arrowhead at the beginning of the specified line. Read/write MsoArrowheadLength.

        [MSDN documentation for LineFormat.BeginArrowheadLength](http://msdn.microsoft.com/en-us/library/bb220883).
        */
        MsoArrowheadLength GetBeginArrowheadLength();

        /**
        Sets the length of the arrowhead at the beginning of the specified line. Read/write MsoArrowheadLength.

        [MSDN documentation for LineFormat.BeginArrowheadLength](http://msdn.microsoft.com/en-us/library/bb220883).
        */
        void SetBeginArrowheadLength(MsoArrowheadLength beginArrowheadLength);

        /**
        Returns the style of the arrowhead at the beginning of the specified line. Read/write MsoArrowheadStyle.

        [MSDN documentation for LineFormat.BeginArrowheadStyle](http://msdn.microsoft.com/en-us/library/bb220884).
        */
        MsoArrowheadStyle GetBeginArrowheadStyle();

        /**
        Sets the style of the arrowhead at the beginning of the specified line. Read/write MsoArrowheadStyle.

        [MSDN documentation for LineFormat.BeginArrowheadStyle](http://msdn.microsoft.com/en-us/library/bb220884).
        */
        void SetBeginArrowheadStyle(MsoArrowheadStyle beginArrowheadStyle);

        /**
        Returns the width of the arrowhead at the beginning of the specified line. Read/write MsoArrowheadWidth.

        [MSDN documentation for LineFormat.BeginArrowheadWidth](http://msdn.microsoft.com/en-us/library/bb220885).
        */
        MsoArrowheadWidth GetBeginArrowheadWidth();

        /**
        Sets the width of the arrowhead at the beginning of the specified line. Read/write MsoArrowheadWidth.

        [MSDN documentation for LineFormat.BeginArrowheadWidth](http://msdn.microsoft.com/en-us/library/bb220885).
        */
        void SetBeginArrowheadWidth(MsoArrowheadWidth beginArrowheadWidth);

        /**
        Returns the dash style for the specified line. Can be one of the MsoLineDashStyle contants.

        [MSDN documentation for LineFormat.DashStyle](http://msdn.microsoft.com/en-us/library/bb177433).
        */
        MsoLineDashStyle GetDashStyle();

        /**
        Sets the dash style for the specified line. Can be one of the MsoLineDashStyle contants.

        [MSDN documentation for LineFormat.DashStyle](http://msdn.microsoft.com/en-us/library/bb177433).
        */
        void SetDashStyle(MsoLineDashStyle dashStyle);

        /**
        Returns the length of the arrowhead at the end of the specified line. Read/write MsoArrowheadLength.

        [MSDN documentation for LineFormat.EndArrowheadLength](http://msdn.microsoft.com/en-us/library/bb208443).
        */
        MsoArrowheadLength GetEndArrowheadLength();

        /**
        Sets the length of the arrowhead at the end of the specified line. Read/write MsoArrowheadLength.

        [MSDN documentation for LineFormat.EndArrowheadLength](http://msdn.microsoft.com/en-us/library/bb208443).
        */
        void SetEndArrowheadLength(MsoArrowheadLength endArrowheadLength);

        /**
        Returns the style of the arrowhead at the end of the specified line. Read/write MsoArrowheadStyle.

        [MSDN documentation for LineFormat.EndArrowheadStyle](http://msdn.microsoft.com/en-us/library/bb208446).
        */
        MsoArrowheadStyle GetEndArrowheadStyle();

        /**
        Sets the style of the arrowhead at the end of the specified line. Read/write MsoArrowheadStyle.

        [MSDN documentation for LineFormat.EndArrowheadStyle](http://msdn.microsoft.com/en-us/library/bb208446).
        */
        void SetEndArrowheadStyle(MsoArrowheadStyle endArrowheadStyle);

        /**
        Returns the width of the arrowhead at the end of the specified line. Read/write MsoArrowheadWidth.

        [MSDN documentation for LineFormat.EndArrowheadWidth](http://msdn.microsoft.com/en-us/library/bb208449).
        */
        MsoArrowheadWidth GetEndArrowheadWidth();

        /**
        Sets the width of the arrowhead at the end of the specified line. Read/write MsoArrowheadWidth.

        [MSDN documentation for LineFormat.EndArrowheadWidth](http://msdn.microsoft.com/en-us/library/bb208449).
        */
        void SetEndArrowheadWidth(MsoArrowheadWidth endArrowheadWidth);

        /**
        Returns a ColorFormat Represents the specified foreground fill or solid color.

        [MSDN documentation for LineFormat.ForeColor](http://msdn.microsoft.com/en-us/library/bb148591).
        */
        wxExcelColorFormat GetForeColor();

        /**
        Sets a ColorFormat Represents the specified foreground fill or solid color.

        [MSDN documentation for LineFormat.ForeColor](http://msdn.microsoft.com/en-us/library/bb148591).
        */
        void SetForeColor(const wxExcelColorFormat& foreColor);


         /**
        Whether lines are drawn inside the specified shape's boundaries. Since Excel 2010.

        [MSDN documentation for LineFormat.InsetPen](http://msdn.microsoft.com/en-us/library/office/ff834393%28v=office.14%29.aspx).
        */
        MsoTriState GetInsetPen();

        /**
        Whether lines are drawn inside the specified shape's boundaries. Since Excel 2010.

        [MSDN documentation for LineFormat.InsetPen](http://msdn.microsoft.com/en-us/library/office/ff834393%28v=office.14%29.aspx).
        */
        void SetInsetPen(MsoTriState insetPen);


        /**
        Returns an MsoPatternType value that represents the fill pattern.

        [MSDN documentation for LineFormat.Pattern](http://msdn.microsoft.com/en-us/library/bb148594).
        */
        MsoPatternType GetPattern();

        /**
        Sets an MsoPatternType value that represents the fill pattern.

        [MSDN documentation for LineFormat.Pattern](http://msdn.microsoft.com/en-us/library/bb148594).
        */
        void SetPattern(MsoPatternType pattern);

        /**
        Returns a MsoLineStyle value that represents the style of the line.

        [MSDN documentation for LineFormat.Style](http://msdn.microsoft.com/en-us/library/bb214835).
        */
        MsoLineStyle  GetStyle();

        /**
        Sets a MsoLineStyle value that represents the style of the line.

        [MSDN documentation for LineFormat.Style](http://msdn.microsoft.com/en-us/library/bb214835).
        */
        void SetStyle(MsoLineStyle  style);

        /**
        Returns the degree of transparency of the specified fill as a value from 0.0 (opaque) through 1.0 (clear). Read/write Double.

        [MSDN documentation for LineFormat.Transparency](http://msdn.microsoft.com/en-us/library/bb214838).
        */
        double GetTransparency();

        /**
        Sets the degree of transparency of the specified fill as a value from 0.0 (opaque) through 1.0 (clear). Read/write Double.

        [MSDN documentation for LineFormat.Transparency](http://msdn.microsoft.com/en-us/library/bb214838).
        */
        void SetTransparency(double transparency);

        /**
        Returns a MsoTriState value that determines whether the object is visible.

        [MSDN documentation for LineFormat.Visible](http://msdn.microsoft.com/en-us/library/bb214842).
        */
        MsoTriState GetVisible();

        /**
        Sets a MsoTriState value that determines whether the object is visible.

        [MSDN documentation for LineFormat.Visible](http://msdn.microsoft.com/en-us/library/bb214842).
        */
        void SetVisible(MsoTriState visible);

        /**
        Returns a Single value that represents the weight of the line.

        [MSDN documentation for LineFormat.Weight](http://msdn.microsoft.com/en-us/library/bb214845).
        */
        double GetWeight();

        /**
        Sets a Single value that represents the weight of the line.

        [MSDN documentation for LineFormat.Weight](http://msdn.microsoft.com/en-us/library/bb214845).
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
