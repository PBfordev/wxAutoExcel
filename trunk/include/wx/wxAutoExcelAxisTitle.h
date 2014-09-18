/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Modified by:
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
/////////////////////////////////////////////////////////////////////////////


#ifndef _WXAUTOEXCEL_AXISTITLE_H
#define _WXAUTOEXCEL_AXISTITLE_H

#include "wx/wxAutoExcel_defs.h"

#if WXAUTOEXCEL_USE_CHARTS

#include "wx/wxAutoExcelObject.h"
#include "wx/wxAutoExcel_enums.h"

namespace wxAutoExcel {


    /**
    Represents Microsoft Excel AxisTitle object.
    */
   class WXDLLIMPEXP_WXAUTOEXCEL wxExcelAxisTitle : public wxExcelObject
    {
    public:
        // ***** METHODS *****

        /**
        Deletes the object.

        [MSDN documentation for AxisTitle.Delete](http://msdn.microsoft.com/en-us/library/bb211578).
        */
        bool Delete();

        /**
        Selects the object.

        [MSDN documentation for AxisTitle.Select](http://msdn.microsoft.com/en-us/library/bb213859).
        */
        bool Select();

        // ***** PROPERTIES *****

        /**
        True if the text in the object changes font size when the object size changes. The default value is True. Read/write Variant.

        [MSDN documentation for AxisTitle.AutoScaleFont](http://msdn.microsoft.com/en-us/library/bb179287).
        */
        bool GetAutoScaleFont();

        /**
        True if the text in the object changes font size when the object size changes. The default value is True. Read/write Variant.

        [MSDN documentation for AxisTitle.AutoScaleFont](http://msdn.microsoft.com/en-us/library/bb179287).
        */
        void SetAutoScaleFont(bool autoScaleFont);

        /**
        Returns a String value that represents the axis title text.

        [MSDN documentation for AxisTitle.Caption](http://msdn.microsoft.com/en-us/library/bb179289).
        */
        wxString GetCaption();

        /**
        Sets a String value that represents the axis title text.

        [MSDN documentation for AxisTitle.Caption](http://msdn.microsoft.com/en-us/library/bb179289).
        */
        void SetCaption(const wxString& caption);

        /**
        Returns a Characters object that represents a range of characters within the object text. You can use the Characters object to format characters within a text string.

        [MSDN documentation for AxisTitle.Characters](http://msdn.microsoft.com/en-us/library/bb179291).
        */
        wxExcelCharacters GetCharacters();

        /**
        Returns the ChartFormat object.  Since Excel 2007.

        [MSDN documentation for AxisTitle.Format](http://msdn.microsoft.com/en-us/library/bb242490).
        */
        wxExcelChartFormat GetFormat();

         /**
        Returns a String value that represents the object's formula in A1-style notation and in the language of the macro. Since Excel 2010.

        [MSDN documentation for AxisTitle.Formula](http://msdn.microsoft.com/en-us/library/office/ff821270%28v=office.14%29.aspx).
        */
        wxString GetFormula();

        /**
        Sets a String value that represents the object's formula in A1-style notation and in the language of the macro. Since Excel 2010.

        [MSDN documentation for AxisTitle.Formula](http://msdn.microsoft.com/en-us/library/office/ff821270%28v=office.14%29.aspx).
        */
        void SetFormula(const wxString& formula);

        /**
        Returns the formula for the object, using A1-style references in the language of the user. Since Excel 2010.

        [MSDN documentation for AxisTitle.FormulaLocal](http://msdn.microsoft.com/en-us/library/office/ff840929%28v=office.14%29.aspx).
        */
        wxString GetFormulaLocal();

        /**
        Sets the formula for the object, using A1-style references in the language of the user. Since Excel 2010.

        [MSDN documentation for AxisTitle.FormulaLocal](http://msdn.microsoft.com/en-us/library/office/ff840929%28v=office.14%29.aspx).
        */
        void SetFormulaLocal(const wxString& formulaLocal);

        /**
        Returns the formula for the object, using R1C1-style notation in the language of the macro. Since Excel 2010.

        [MSDN documentation for AxisTitle.FormulaR1C1](http://msdn.microsoft.com/en-us/library/office/ff822503%28v=office.14%29.aspx).
        */
        wxString GetFormulaR1C1();

        /**
        Sets the formula for the object, using R1C1-style notation in the language of the macro. Since Excel 2010.

        [MSDN documentation for AxisTitle.FormulaR1C1](http://msdn.microsoft.com/en-us/library/office/ff822503%28v=office.14%29.aspx).
        */
        void SetFormulaR1C1(const wxString& formulaR1C1);

        /**
        Returns the formula for the object, using R1C1-style notation in the language of the user. Since Excel 2010.

        [MSDN documentation for AxisTitle.FormulaR1C1Local](http://msdn.microsoft.com/en-us/library/office/ff839371%28v=office.14%29.aspx).
        */
        wxString GetFormulaR1C1Local();

        /**
        Sets the formula for the object, using R1C1-style notation in the language of the user. Since Excel 2010.

        [MSDN documentation for AxisTitle.FormulaR1C1Local](http://msdn.microsoft.com/en-us/library/office/ff839371%28v=office.14%29.aspx).
        */
        void SetFormulaR1C1Local(const wxString& formulaR1C1Local);

        /**
        Returns a Variant value that represents the horizontal alignment for the specified object.

        [MSDN documentation for AxisTitle.HorizontalAlignment](http://msdn.microsoft.com/en-us/library/bb179294).
        */
        long GetHorizontalAlignment();

        /**
        Sets a Variant value that represents the horizontal alignment for the specified object.

        [MSDN documentation for AxisTitle.HorizontalAlignment](http://msdn.microsoft.com/en-us/library/bb179294).
        */
        void SetHorizontalAlignment(long horizontalAlignment);

        /**
        True if an axis title will occupy the chart layout space when a chart layout is being determined. Since Excel 2007.

        [MSDN documentation for AxisTitle.IncludeInLayout](http://msdn.microsoft.com/en-us/library/bb179295).
        */
        bool GetIncludeInLayout();

        /**
        True if an axis title will occupy the chart layout space when a chart layout is being determined. Since Excel 2007.

        [MSDN documentation for AxisTitle.IncludeInLayout](http://msdn.microsoft.com/en-us/library/bb179295).
        */
        void SetIncludeInLayout(bool includeInLayout);

        /**
        Returns a Double value that represents the distance, in points, from the left edge of the object to the left edge of column A (on a worksheet) or the left edge of the chart area (on a chart).

        [MSDN documentation for AxisTitle.Left](http://msdn.microsoft.com/en-us/library/bb179297).
        */
        double GetLeft();

        /**
        Sets a Double value that represents the distance, in points, from the left edge of the object to the left edge of column A (on a worksheet) or the left edge of the chart area (on a chart).

        [MSDN documentation for AxisTitle.Left](http://msdn.microsoft.com/en-us/library/bb179297).
        */
        void SetLeft(double left);

        /**
        Returns a String value that represents the name of the object.

        [MSDN documentation for AxisTitle.Name](http://msdn.microsoft.com/en-us/library/bb179299).
        */
        wxString GetName();

        /**
        Returns a Variant value that represents the text orientation.

        [MSDN documentation for AxisTitle.Orientation](http://msdn.microsoft.com/en-us/library/bb179302).
        */
        long GetOrientation();

        /**
        Sets a Variant value that represents the text orientation.

        [MSDN documentation for AxisTitle.Orientation](http://msdn.microsoft.com/en-us/library/bb179302).
        */
        void SetOrientation(long orientation);

        /**
        Returns the position of the axis title on the chart.  Since Excel 2007.

        [MSDN documentation for AxisTitle.Position](http://msdn.microsoft.com/en-us/library/bb224822).
        */
        XlChartElementPosition GetPosition();

        /**
        Sets the position of the axis title on the chart.  Since Excel 2007.

        [MSDN documentation for AxisTitle.Position](http://msdn.microsoft.com/en-us/library/bb224822).
        */
        void SetPosition(XlChartElementPosition position);

        /**
        Returns the reading order for the specified object. Can be one of the following constants: xlRTL (right-to-left), xlLTR (left-to-right), or xlContext.

        [MSDN documentation for AxisTitle.ReadingOrder](http://msdn.microsoft.com/en-us/library/bb179305).
        */
        long GetReadingOrder();

        /**
        Sets the reading order for the specified object. Can be one of the following constants: xlRTL (right-to-left), xlLTR (left-to-right), or xlContext.

        [MSDN documentation for AxisTitle.ReadingOrder](http://msdn.microsoft.com/en-us/library/bb179305).
        */
        void SetReadingOrder(long readingOrder);

        /**
        Returns a Boolean value that determines if the object has a shadow.

        [MSDN documentation for AxisTitle.Shadow](http://msdn.microsoft.com/en-us/library/bb214469).
        */
        bool GetShadow();

        /**
        Sets a Boolean value that determines if the object has a shadow.

        [MSDN documentation for AxisTitle.Shadow](http://msdn.microsoft.com/en-us/library/bb214469).
        */
        void SetShadow(bool shadow);

        /**
        Returns the text for the specified object.

        [MSDN documentation for AxisTitle.Text](http://msdn.microsoft.com/en-us/library/bb214478).
        */
        wxString GetText();

        /**
        Sets the text for the specified object.

        [MSDN documentation for AxisTitle.Text](http://msdn.microsoft.com/en-us/library/bb214478).
        */
        void SetText(const wxString& text);

        /**
        Returns a Double value that represents the distance, in points, from the top edge of the object to the top of row 1 (on a worksheet) or the top of the chart area (on a chart).

        [MSDN documentation for AxisTitle.Top](http://msdn.microsoft.com/en-us/library/bb214482).
        */
        double GetTop();

        /**
        Sets a Double value that represents the distance, in points, from the top edge of the object to the top of row 1 (on a worksheet) or the top of the chart area (on a chart).

        [MSDN documentation for AxisTitle.Top](http://msdn.microsoft.com/en-us/library/bb214482).
        */
        void SetTop(double top);

        /**
        Returns a Variant value that represents the vertical alignment of the specified object.

        [MSDN documentation for AxisTitle.VerticalAlignment](http://msdn.microsoft.com/en-us/library/bb214491).
        */
        long GetVerticalAlignment();

        /**
        Sets a Variant value that represents the vertical alignment of the specified object.

        [MSDN documentation for AxisTitle.VerticalAlignment](http://msdn.microsoft.com/en-us/library/bb214491).
        */
        void SetVerticalAlignment(long verticalAlignment);

         /**
        Returns the width, in points, of the object. Since Excel 2010.

        [MSDN documentation for AxisTitle.Width](http://msdn.microsoft.com/en-us/library/office/ff834650%28v=office.14%29.aspx).
        */
        double GetWidth();

        /**
        Returns "AxisTitle".
        */
        virtual wxString GetAutoExcelObjectName_() const { return wxS("AxisTitle"); }
    };


} // namespace wxAutoExcel

#endif // #if WXAUTOEXCEL_USE_CHARTS

#endif // _WXAUTOEXCEL_AXISTITLE_H
