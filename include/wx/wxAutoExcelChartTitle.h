/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Modified by:
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
/////////////////////////////////////////////////////////////////////////////


#ifndef _WXAUTOEXCEL_CHARTTITLE_H
#define _WXAUTOEXCEL_CHARTTITLE_H

#include "wx/wxAutoExcel_defs.h"

#if WXAUTOEXCEL_USE_CHARTS

#include "wx/wxAutoExcel_object.h"
#include "wx/wxAutoExcel_enums.h"

namespace wxAutoExcel {

    /**
    Represents Microsoft Excel ChartTitle object.
    */
   class WXDLLIMPEXP_WXAUTOEXCEL wxExcelChartTitle : public wxExcelObject
    {
    public:
        // ***** METHODS *****

        /**
        Deletes the object.

        [MSDN documentation for ChartTitle.Delete](http://msdn.microsoft.com/en-us/library/bb211698).
        */
        bool Delete();

        /**
        Selects the object.

        [MSDN documentation for ChartTitle.Select](http://msdn.microsoft.com/en-us/library/bb237769).
        */
        bool Select();

        // ***** PROPERTIES *****

        /**
        True if the text in the object changes font size when the object size changes. The default value is True. Read/write Variant.

        [MSDN documentation for ChartTitle.AutoScaleFont](http://msdn.microsoft.com/en-us/library/bb179498).
        */
        bool GetAutoScaleFont();

        /**
        True if the text in the object changes font size when the object size changes. The default value is True. Read/write Variant.

        [MSDN documentation for ChartTitle.AutoScaleFont](http://msdn.microsoft.com/en-us/library/bb179498).
        */
        void SetAutoScaleFont(bool autoScaleFont);

        /**
        Returns a String value that represents the chart title text.

        [MSDN documentation for ChartTitle.Caption](http://msdn.microsoft.com/en-us/library/bb179501).
        */
        wxString GetCaption();

        /**
        Sets a String value that represents the chart title text.

        [MSDN documentation for ChartTitle.Caption](http://msdn.microsoft.com/en-us/library/bb179501).
        */
        void SetCaption(const wxString& caption);

        /**
        Returns a Characters object that represents a range of characters within the object text. You can use the Characters object to format characters within a text string.

        [MSDN documentation for ChartTitle.Characters](http://msdn.microsoft.com/en-us/library/bb179504).
        */
        wxExcelCharacters GetCharacters(long start = 1, long* length = NULL);

        /**
        Returns the ChartFormat object.  Since Excel 2007.

        [MSDN documentation for ChartTitle.Format](http://msdn.microsoft.com/en-us/library/bb242519).
        */
        wxExcelChartFormat GetFormat();

         /**
        Returns a String value that represents the object's formula in A1-style notation and in the language of the macro. Since Excel 2010.

        [MSDN documentation for ChartTitle.Formula](http://msdn.microsoft.com/en-us/library/office/ff821962%28v=office.14%29.aspx).
        */
        wxString GetFormula();

        /**
        Sets a String value that represents the object's formula in A1-style notation and in the language of the macro. Since Excel 2010.

        [MSDN documentation for ChartTitle.Formula](http://msdn.microsoft.com/en-us/library/office/ff821962%28v=office.14%29.aspx).
        */
        void SetFormula(const wxString& formula);

        /**
        Returns the formula for the object, using A1-style references in the language of the user. Since Excel 2010.

        [MSDN documentation for ChartTitle.FormulaLocal](http://msdn.microsoft.com/en-us/library/office/ff193652%28v=office.14%29.aspx).
        */
        wxString GetFormulaLocal();

        /**
        Sets the formula for the object, using A1-style references in the language of the user. Since Excel 2010.

        [MSDN documentation for ChartTitle.FormulaLocal](http://msdn.microsoft.com/en-us/library/office/ff193652%28v=office.14%29.aspx).
        */
        void SetFormulaLocal(const wxString& formulaLocal);

        /**
        Returns the formula for the object, using R1C1-style notation in the language of the macro. Since Excel 2010.

        [MSDN documentation for ChartTitle.FormulaR1C1](http://msdn.microsoft.com/en-us/library/office/ff822658%28v=office.14%29.aspx).
        */
        wxString GetFormulaR1C1();

        /**
        Sets the formula for the object, using R1C1-style notation in the language of the macro. Since Excel 2010.

        [MSDN documentation for ChartTitle.FormulaR1C1](http://msdn.microsoft.com/en-us/library/office/ff822658%28v=office.14%29.aspx).
        */
        void SetFormulaR1C1(const wxString& formulaR1C1);

        /**
        Returns the formula for the object, using R1C1-style notation in the language of the user. Since Excel 2010.

        [MSDN documentation for ChartTitle.FormulaR1C1Local](http://msdn.microsoft.com/en-us/library/office/ff834688%28v=office.14%29.aspx).
        */
        wxString GetFormulaR1C1Local();

        /**
        Sets the formula for the object, using R1C1-style notation in the language of the user. Since Excel 2010.

        [MSDN documentation for ChartTitle.FormulaR1C1Local](http://msdn.microsoft.com/en-us/library/office/ff834688%28v=office.14%29.aspx).
        */
        void SetFormulaR1C1Local(const wxString& formulaR1C1Local);

         /**
        Returns the height, in points, of the object. Since Excel 2010.

        [MSDN documentation for ChartTitle.Height](http://msdn.microsoft.com/en-us/library/office/ff197548%28v=office.14%29.aspx).
        */
        double GetHeight();

        /**
        Returns a Variant value that represents the horizontal alignment for the specified object.

        [MSDN documentation for ChartTitle.HorizontalAlignment](http://msdn.microsoft.com/en-us/library/bb179505).
        */
        long GetHorizontalAlignment();

        /**
        Sets a Variant value that represents the horizontal alignment for the specified object.

        [MSDN documentation for ChartTitle.HorizontalAlignment](http://msdn.microsoft.com/en-us/library/bb179505).
        */
        void SetHorizontalAlignment(long horizontalAlignment);

        /**
        True if a chart title will occupy the chart layout space when a chart layout is being determined.  Since Excel 2007.

        [MSDN documentation for ChartTitle.IncludeInLayout](http://msdn.microsoft.com/en-us/library/bb179506).
        */
        bool GetIncludeInLayout();

        /**
        True if a chart title will occupy the chart layout space when a chart layout is being determined.  Since Excel 2007.

        [MSDN documentation for ChartTitle.IncludeInLayout](http://msdn.microsoft.com/en-us/library/bb179506).
        */
        void SetIncludeInLayout(bool includeInLayout);

        /**
        Returns a Double value that represents the distance, in points, from the left edge of the object to the left edge of column A (on a worksheet) or the left edge of the chart area (on a chart).

        [MSDN documentation for ChartTitle.Left](http://msdn.microsoft.com/en-us/library/bb179507).
        */
        double GetLeft();

        /**
        Sets a Double value that represents the distance, in points, from the left edge of the object to the left edge of column A (on a worksheet) or the left edge of the chart area (on a chart).

        [MSDN documentation for ChartTitle.Left](http://msdn.microsoft.com/en-us/library/bb179507).
        */
        void SetLeft(double left);

        /**
        Returns a String value that represents the name of the object.

        [MSDN documentation for ChartTitle.Name](http://msdn.microsoft.com/en-us/library/bb179511).
        */
        wxString GetName();

        /**
        Returns a Variant value that represents the text orientation.

        [MSDN documentation for ChartTitle.Orientation](http://msdn.microsoft.com/en-us/library/bb179513).
        */
        long GetOrientation();

        /**
        Sets a Variant value that represents the text orientation.

        [MSDN documentation for ChartTitle.Orientation](http://msdn.microsoft.com/en-us/library/bb179513).
        */
        void SetOrientation(long orientation);

        /**
        Returns the position of the chart title on the chart. Since Excel 2007.

        [MSDN documentation for ChartTitle.Position](http://msdn.microsoft.com/en-us/library/bb239958).
        */
        XlChartElementPosition GetPosition();

        /**
        Sets the position of the chart title on the chart. Since Excel 2007.

        [MSDN documentation for ChartTitle.Position](http://msdn.microsoft.com/en-us/library/bb239958).
        */
        void SetPosition(XlChartElementPosition position);

        /**
        Returns the reading order for the specified object. Can be one of the following constants: xlRTL (right-to-left), xlLTR (left-to-right), or xlContext.

        [MSDN documentation for ChartTitle.ReadingOrder](http://msdn.microsoft.com/en-us/library/bb179514).
        */
        long GetReadingOrder();

        /**
        Sets the reading order for the specified object. Can be one of the following constants: xlRTL (right-to-left), xlLTR (left-to-right), or xlContext.

        [MSDN documentation for ChartTitle.ReadingOrder](http://msdn.microsoft.com/en-us/library/bb179514).
        */
        void SetReadingOrder(long readingOrder);

        /**
        Returns a Boolean value that determines if the object has a shadow.

        [MSDN documentation for ChartTitle.Shadow](http://msdn.microsoft.com/en-us/library/bb238496).
        */
        bool GetShadow();

        /**
        Sets a Boolean value that determines if the object has a shadow.

        [MSDN documentation for ChartTitle.Shadow](http://msdn.microsoft.com/en-us/library/bb238496).
        */
        void SetShadow(bool shadow);

        /**
        Returns the text for the specified object.

        [MSDN documentation for ChartTitle.Text](http://msdn.microsoft.com/en-us/library/bb238502).
        */
        wxString GetText();

        /**
        Sets the text for the specified object.

        [MSDN documentation for ChartTitle.Text](http://msdn.microsoft.com/en-us/library/bb238502).
        */
        void SetText(const wxString& text);

        /**
        Returns a Double value that represents the distance, in points, from the top edge of the object to the top of row 1 (on a worksheet) or the top of the chart area (on a chart).

        [MSDN documentation for ChartTitle.Top](http://msdn.microsoft.com/en-us/library/bb238509).
        */
        double GetTop();

        /**
        Sets a Double value that represents the distance, in points, from the top edge of the object to the top of row 1 (on a worksheet) or the top of the chart area (on a chart).

        [MSDN documentation for ChartTitle.Top](http://msdn.microsoft.com/en-us/library/bb238509).
        */
        void SetTop(double top);

        /**
        Returns a Variant value that represents the vertical alignment of the specified object.

        [MSDN documentation for ChartTitle.VerticalAlignment](http://msdn.microsoft.com/en-us/library/bb238518).
        */
        long GetVerticalAlignment();

        /**
        Sets a Variant value that represents the vertical alignment of the specified object.

        [MSDN documentation for ChartTitle.VerticalAlignment](http://msdn.microsoft.com/en-us/library/bb238518).
        */
        void SetVerticalAlignment(long verticalAlignment);

         /**
        Returns the width, in points, of the object. Since Excel 2010.

        [MSDN documentation for ChartTitle.Width](http://msdn.microsoft.com/en-us/library/office/ff839248%28v=office.14%29.aspx).
        */
        double GetWidth();

        /**
        Returns "ChartTitle".
        */
        virtual wxString GetAutoExcelObjectName_() const { return wxS("ChartTitle"); }
    };


} // namespace wxAutoExcel

#endif // #if WXAUTOEXCEL_USE_CHARTS

#endif // _WXAUTOEXCEL_CHARTTITLE_H
