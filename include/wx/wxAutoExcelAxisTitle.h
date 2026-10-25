/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// License:     MIT license
/////////////////////////////////////////////////////////////////////////////


#ifndef _WXAUTOEXCEL_AXISTITLE_H
#define _WXAUTOEXCEL_AXISTITLE_H

#include "wx/wxAutoExcel_defs.h"

#if WXAUTOEXCEL_USE_CHARTS

#include "wx/wxAutoExcel_object.h"
#include "wx/wxAutoExcel_enums.h"

namespace wxAutoExcel {


    /**
    @brief Represents Microsoft Excel AxisTitle object.
    */
   class WXDLLIMPEXP_WXAUTOEXCEL wxExcelAxisTitle : public wxExcelObject
    {
    public:
        // ***** METHODS *****

        /**
        Deletes the object.

        [MSDN documentation for AxisTitle.Delete](https://learn.microsoft.com/en-us/office/vba/api/excel.axistitle.delete).
        */
        bool Delete();

        /**
        Selects the object.

        [MSDN documentation for AxisTitle.Select](https://learn.microsoft.com/en-us/office/vba/api/excel.axistitle.select).
        */
        bool Select();

        // ***** PROPERTIES *****

        /**
        True if the text in the object changes font size when the object size changes. The default value is True. Read/write Variant.

        [MSDN documentation for AxisTitle.AutoScaleFont](https://learn.microsoft.com/en-us/previous-versions/office/developer/office-2007/bb179287%28v=office.12%29).
        */
        bool GetAutoScaleFont();

        /**
        True if the text in the object changes font size when the object size changes. The default value is True. Read/write Variant.

        [MSDN documentation for AxisTitle.AutoScaleFont](https://learn.microsoft.com/en-us/previous-versions/office/developer/office-2007/bb179287%28v=office.12%29).
        */
        void SetAutoScaleFont(bool autoScaleFont);

        /**
        Returns a String value that represents the axis title text.

        [MSDN documentation for AxisTitle.Caption](https://learn.microsoft.com/en-us/office/vba/api/excel.axistitle.caption).
        */
        wxString GetCaption();

        /**
        Sets a String value that represents the axis title text.

        [MSDN documentation for AxisTitle.Caption](https://learn.microsoft.com/en-us/office/vba/api/excel.axistitle.caption).
        */
        void SetCaption(const wxString& caption);

        /**
        Returns a Characters object that represents a range of characters within the object text. You can use the Characters object to format characters within a text string.

        [MSDN documentation for AxisTitle.Characters](https://learn.microsoft.com/en-us/office/vba/api/excel.axistitle.characters).
        */
        wxExcelCharacters GetCharacters(long start = 1, long* length = NULL);

        /**
        Returns a ChartFillFormat object that contains fill formatting properties for the object.

        [MSDN documentation for AxisTitle.Fill](https://learn.microsoft.com/en-us/office/vba/api/excel.axistitle.fill).
        */
        wxExcelChartFillFormat GetFill();

        /**
        Returns the ChartFormat object.  Since Excel 2007.

        [MSDN documentation for AxisTitle.Format](https://learn.microsoft.com/en-us/office/vba/api/excel.axistitle.format).
        */
        wxExcelChartFormat GetFormat();

         /**
        Returns a String value that represents the object's formula in A1-style notation and in the language of the macro. Since Excel 2010.

        [MSDN documentation for AxisTitle.Formula](https://learn.microsoft.com/en-us/office/vba/api/excel.axistitle.formula).
        */
        wxString GetFormula();

        /**
        Sets a String value that represents the object's formula in A1-style notation and in the language of the macro. Since Excel 2010.

        [MSDN documentation for AxisTitle.Formula](https://learn.microsoft.com/en-us/office/vba/api/excel.axistitle.formula).
        */
        void SetFormula(const wxString& formula);

        /**
        Returns the formula for the object, using A1-style references in the language of the user. Since Excel 2010.

        [MSDN documentation for AxisTitle.FormulaLocal](https://learn.microsoft.com/en-us/office/vba/api/excel.axistitle.formulalocal).
        */
        wxString GetFormulaLocal();

        /**
        Sets the formula for the object, using A1-style references in the language of the user. Since Excel 2010.

        [MSDN documentation for AxisTitle.FormulaLocal](https://learn.microsoft.com/en-us/office/vba/api/excel.axistitle.formulalocal).
        */
        void SetFormulaLocal(const wxString& formulaLocal);

        /**
        Returns the formula for the object, using R1C1-style notation in the language of the macro. Since Excel 2010.

        [MSDN documentation for AxisTitle.FormulaR1C1](https://learn.microsoft.com/en-us/office/vba/api/excel.axistitle.formular1c1).
        */
        wxString GetFormulaR1C1();

        /**
        Sets the formula for the object, using R1C1-style notation in the language of the macro. Since Excel 2010.

        [MSDN documentation for AxisTitle.FormulaR1C1](https://learn.microsoft.com/en-us/office/vba/api/excel.axistitle.formular1c1).
        */
        void SetFormulaR1C1(const wxString& formulaR1C1);

        /**
        Returns the formula for the object, using R1C1-style notation in the language of the user. Since Excel 2010.

        [MSDN documentation for AxisTitle.FormulaR1C1Local](https://learn.microsoft.com/en-us/office/vba/api/excel.axistitle.formular1c1local).
        */
        wxString GetFormulaR1C1Local();

        /**
        Sets the formula for the object, using R1C1-style notation in the language of the user. Since Excel 2010.

        [MSDN documentation for AxisTitle.FormulaR1C1Local](https://learn.microsoft.com/en-us/office/vba/api/excel.axistitle.formular1c1local).
        */
        void SetFormulaR1C1Local(const wxString& formulaR1C1Local);

        /**
        Returns a Variant value that represents the horizontal alignment for the specified object.

        [MSDN documentation for AxisTitle.HorizontalAlignment](https://learn.microsoft.com/en-us/office/vba/api/excel.axistitle.horizontalalignment).
        */
        long GetHorizontalAlignment();

        /**
        Sets a Variant value that represents the horizontal alignment for the specified object.

        [MSDN documentation for AxisTitle.HorizontalAlignment](https://learn.microsoft.com/en-us/office/vba/api/excel.axistitle.horizontalalignment).
        */
        void SetHorizontalAlignment(long horizontalAlignment);

        /**
        True if an axis title will occupy the chart layout space when a chart layout is being determined. Since Excel 2007.

        [MSDN documentation for AxisTitle.IncludeInLayout](https://learn.microsoft.com/en-us/office/vba/api/excel.axistitle.includeinlayout).
        */
        bool GetIncludeInLayout();

        /**
        True if an axis title will occupy the chart layout space when a chart layout is being determined. Since Excel 2007.

        [MSDN documentation for AxisTitle.IncludeInLayout](https://learn.microsoft.com/en-us/office/vba/api/excel.axistitle.includeinlayout).
        */
        void SetIncludeInLayout(bool includeInLayout);

        /**
        Returns a Double value that represents the distance, in points, from the left edge of the object to the left edge of column A (on a worksheet) or the left edge of the chart area (on a chart).

        [MSDN documentation for AxisTitle.Left](https://learn.microsoft.com/en-us/office/vba/api/excel.axistitle.left).
        */
        double GetLeft();

        /**
        Sets a Double value that represents the distance, in points, from the left edge of the object to the left edge of column A (on a worksheet) or the left edge of the chart area (on a chart).

        [MSDN documentation for AxisTitle.Left](https://learn.microsoft.com/en-us/office/vba/api/excel.axistitle.left).
        */
        void SetLeft(double left);

        /**
        Returns a String value that represents the name of the object.

        [MSDN documentation for AxisTitle.Name](https://learn.microsoft.com/en-us/office/vba/api/excel.axistitle.name).
        */
        wxString GetName();

        /**
        Returns a Variant value that represents the text orientation.

        [MSDN documentation for AxisTitle.Orientation](https://learn.microsoft.com/en-us/office/vba/api/excel.axistitle.orientation).
        */
        long GetOrientation();

        /**
        Sets a Variant value that represents the text orientation.

        [MSDN documentation for AxisTitle.Orientation](https://learn.microsoft.com/en-us/office/vba/api/excel.axistitle.orientation).
        */
        void SetOrientation(long orientation);

        /**
        Returns the position of the axis title on the chart.  Since Excel 2007.

        [MSDN documentation for AxisTitle.Position](https://learn.microsoft.com/en-us/office/vba/api/excel.axistitle.position).
        */
        XlChartElementPosition GetPosition();

        /**
        Sets the position of the axis title on the chart.  Since Excel 2007.

        [MSDN documentation for AxisTitle.Position](https://learn.microsoft.com/en-us/office/vba/api/excel.axistitle.position).
        */
        void SetPosition(XlChartElementPosition position);

        /**
        Returns the reading order for the specified object. Can be one of the following constants: xlRTL (right-to-left), xlLTR (left-to-right), or xlContext.

        [MSDN documentation for AxisTitle.ReadingOrder](https://learn.microsoft.com/en-us/office/vba/api/excel.axistitle.readingorder).
        */
        long GetReadingOrder();

        /**
        Sets the reading order for the specified object. Can be one of the following constants: xlRTL (right-to-left), xlLTR (left-to-right), or xlContext.

        [MSDN documentation for AxisTitle.ReadingOrder](https://learn.microsoft.com/en-us/office/vba/api/excel.axistitle.readingorder).
        */
        void SetReadingOrder(long readingOrder);

        /**
        Returns a Boolean value that determines if the object has a shadow.

        [MSDN documentation for AxisTitle.Shadow](https://learn.microsoft.com/en-us/office/vba/api/excel.axistitle.shadow).
        */
        bool GetShadow();

        /**
        Sets a Boolean value that determines if the object has a shadow.

        [MSDN documentation for AxisTitle.Shadow](https://learn.microsoft.com/en-us/office/vba/api/excel.axistitle.shadow).
        */
        void SetShadow(bool shadow);

        /**
        Returns the text for the specified object.

        [MSDN documentation for AxisTitle.Text](https://learn.microsoft.com/en-us/office/vba/api/excel.axistitle.text).
        */
        wxString GetText();

        /**
        Sets the text for the specified object.

        [MSDN documentation for AxisTitle.Text](https://learn.microsoft.com/en-us/office/vba/api/excel.axistitle.text).
        */
        void SetText(const wxString& text);

        /**
        Returns a Double value that represents the distance, in points, from the top edge of the object to the top of row 1 (on a worksheet) or the top of the chart area (on a chart).

        [MSDN documentation for AxisTitle.Top](https://learn.microsoft.com/en-us/office/vba/api/excel.axistitle.top).
        */
        double GetTop();

        /**
        Sets a Double value that represents the distance, in points, from the top edge of the object to the top of row 1 (on a worksheet) or the top of the chart area (on a chart).

        [MSDN documentation for AxisTitle.Top](https://learn.microsoft.com/en-us/office/vba/api/excel.axistitle.top).
        */
        void SetTop(double top);

        /**
        Returns a Variant value that represents the vertical alignment of the specified object.

        [MSDN documentation for AxisTitle.VerticalAlignment](https://learn.microsoft.com/en-us/office/vba/api/excel.axistitle.verticalalignment).
        */
        long GetVerticalAlignment();

        /**
        Sets a Variant value that represents the vertical alignment of the specified object.

        [MSDN documentation for AxisTitle.VerticalAlignment](https://learn.microsoft.com/en-us/office/vba/api/excel.axistitle.verticalalignment).
        */
        void SetVerticalAlignment(long verticalAlignment);

         /**
        Returns the width, in points, of the object. Since Excel 2010.

        [MSDN documentation for AxisTitle.Width](https://learn.microsoft.com/en-us/office/vba/api/excel.axistitle.width).
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
