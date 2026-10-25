/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// License:     MIT license
/////////////////////////////////////////////////////////////////////////////


#ifndef _WXAUTOEXCEL_CHARTTITLE_H
#define _WXAUTOEXCEL_CHARTTITLE_H

#include "wx/wxAutoExcel_defs.h"

#if WXAUTOEXCEL_USE_CHARTS

#include "wx/wxAutoExcel_object.h"
#include "wx/wxAutoExcel_enums.h"

namespace wxAutoExcel {

    /**
    @brief Represents Microsoft Excel ChartTitle object.
    */
   class WXDLLIMPEXP_WXAUTOEXCEL wxExcelChartTitle : public wxExcelObject
    {
    public:
        // ***** METHODS *****

        /**
        Deletes the object.

        [MSDN documentation for ChartTitle.Delete](https://learn.microsoft.com/en-us/office/vba/api/excel.charttitle.delete).
        */
        bool Delete();

        /**
        Selects the object.

        [MSDN documentation for ChartTitle.Select](https://learn.microsoft.com/en-us/office/vba/api/excel.charttitle.select).
        */
        bool Select();

        // ***** PROPERTIES *****

        /**
        True if the text in the object changes font size when the object size changes. The default value is True. Read/write Variant.

        [MSDN documentation for ChartTitle.AutoScaleFont](https://learn.microsoft.com/en-us/previous-versions/office/developer/office-2007/bb179498%28v=office.12%29).
        */
        bool GetAutoScaleFont();

        /**
        True if the text in the object changes font size when the object size changes. The default value is True. Read/write Variant.

        [MSDN documentation for ChartTitle.AutoScaleFont](https://learn.microsoft.com/en-us/previous-versions/office/developer/office-2007/bb179498%28v=office.12%29).
        */
        void SetAutoScaleFont(bool autoScaleFont);

        /**
        Returns a String value that represents the chart title text.

        [MSDN documentation for ChartTitle.Caption](https://learn.microsoft.com/en-us/office/vba/api/excel.charttitle.caption).
        */
        wxString GetCaption();

        /**
        Sets a String value that represents the chart title text.

        [MSDN documentation for ChartTitle.Caption](https://learn.microsoft.com/en-us/office/vba/api/excel.charttitle.caption).
        */
        void SetCaption(const wxString& caption);

        /**
        Returns a Characters object that represents a range of characters within the object text. You can use the Characters object to format characters within a text string.

        [MSDN documentation for ChartTitle.Characters](https://learn.microsoft.com/en-us/office/vba/api/excel.charttitle.characters).
        */
        wxExcelCharacters GetCharacters(long start = 1, long* length = NULL);

        /**
        Returns a ChartFillFormat object that contains fill formatting properties for the object.

        [MSDN documentation for ChartTitle.Fill](https://learn.microsoft.com/en-us/office/vba/api/excel.charttitle.fill).
        */
        wxExcelChartFillFormat GetFill();

        /**
        Returns the ChartFormat object.  Since Excel 2007.

        [MSDN documentation for ChartTitle.Format](https://learn.microsoft.com/en-us/office/vba/api/excel.charttitle.format).
        */
        wxExcelChartFormat GetFormat();

         /**
        Returns a String value that represents the object's formula in A1-style notation and in the language of the macro. Since Excel 2010.

        [MSDN documentation for ChartTitle.Formula](https://learn.microsoft.com/en-us/office/vba/api/excel.charttitle.formula).
        */
        wxString GetFormula();

        /**
        Sets a String value that represents the object's formula in A1-style notation and in the language of the macro. Since Excel 2010.

        [MSDN documentation for ChartTitle.Formula](https://learn.microsoft.com/en-us/office/vba/api/excel.charttitle.formula).
        */
        void SetFormula(const wxString& formula);

        /**
        Returns the formula for the object, using A1-style references in the language of the user. Since Excel 2010.

        [MSDN documentation for ChartTitle.FormulaLocal](https://learn.microsoft.com/en-us/office/vba/api/excel.charttitle.formulalocal).
        */
        wxString GetFormulaLocal();

        /**
        Sets the formula for the object, using A1-style references in the language of the user. Since Excel 2010.

        [MSDN documentation for ChartTitle.FormulaLocal](https://learn.microsoft.com/en-us/office/vba/api/excel.charttitle.formulalocal).
        */
        void SetFormulaLocal(const wxString& formulaLocal);

        /**
        Returns the formula for the object, using R1C1-style notation in the language of the macro. Since Excel 2010.

        [MSDN documentation for ChartTitle.FormulaR1C1](https://learn.microsoft.com/en-us/office/vba/api/excel.charttitle.formular1c1).
        */
        wxString GetFormulaR1C1();

        /**
        Sets the formula for the object, using R1C1-style notation in the language of the macro. Since Excel 2010.

        [MSDN documentation for ChartTitle.FormulaR1C1](https://learn.microsoft.com/en-us/office/vba/api/excel.charttitle.formular1c1).
        */
        void SetFormulaR1C1(const wxString& formulaR1C1);

        /**
        Returns the formula for the object, using R1C1-style notation in the language of the user. Since Excel 2010.

        [MSDN documentation for ChartTitle.FormulaR1C1Local](https://learn.microsoft.com/en-us/office/vba/api/excel.charttitle.formular1c1local).
        */
        wxString GetFormulaR1C1Local();

        /**
        Sets the formula for the object, using R1C1-style notation in the language of the user. Since Excel 2010.

        [MSDN documentation for ChartTitle.FormulaR1C1Local](https://learn.microsoft.com/en-us/office/vba/api/excel.charttitle.formular1c1local).
        */
        void SetFormulaR1C1Local(const wxString& formulaR1C1Local);

         /**
        Returns the height, in points, of the object. Since Excel 2010.

        [MSDN documentation for ChartTitle.Height](https://learn.microsoft.com/en-us/office/vba/api/excel.charttitle.height).
        */
        double GetHeight();

        /**
        Returns a Variant value that represents the horizontal alignment for the specified object.

        [MSDN documentation for ChartTitle.HorizontalAlignment](https://learn.microsoft.com/en-us/office/vba/api/excel.charttitle.horizontalalignment).
        */
        long GetHorizontalAlignment();

        /**
        Sets a Variant value that represents the horizontal alignment for the specified object.

        [MSDN documentation for ChartTitle.HorizontalAlignment](https://learn.microsoft.com/en-us/office/vba/api/excel.charttitle.horizontalalignment).
        */
        void SetHorizontalAlignment(long horizontalAlignment);

        /**
        True if a chart title will occupy the chart layout space when a chart layout is being determined.  Since Excel 2007.

        [MSDN documentation for ChartTitle.IncludeInLayout](https://learn.microsoft.com/en-us/office/vba/api/excel.charttitle.includeinlayout).
        */
        bool GetIncludeInLayout();

        /**
        True if a chart title will occupy the chart layout space when a chart layout is being determined.  Since Excel 2007.

        [MSDN documentation for ChartTitle.IncludeInLayout](https://learn.microsoft.com/en-us/office/vba/api/excel.charttitle.includeinlayout).
        */
        void SetIncludeInLayout(bool includeInLayout);

        /**
        Returns a Double value that represents the distance, in points, from the left edge of the object to the left edge of column A (on a worksheet) or the left edge of the chart area (on a chart).

        [MSDN documentation for ChartTitle.Left](https://learn.microsoft.com/en-us/office/vba/api/excel.charttitle.left).
        */
        double GetLeft();

        /**
        Sets a Double value that represents the distance, in points, from the left edge of the object to the left edge of column A (on a worksheet) or the left edge of the chart area (on a chart).

        [MSDN documentation for ChartTitle.Left](https://learn.microsoft.com/en-us/office/vba/api/excel.charttitle.left).
        */
        void SetLeft(double left);

        /**
        Returns a String value that represents the name of the object.

        [MSDN documentation for ChartTitle.Name](https://learn.microsoft.com/en-us/office/vba/api/excel.charttitle.name).
        */
        wxString GetName();

        /**
        Returns a Variant value that represents the text orientation.

        [MSDN documentation for ChartTitle.Orientation](https://learn.microsoft.com/en-us/office/vba/api/excel.charttitle.orientation).
        */
        long GetOrientation();

        /**
        Sets a Variant value that represents the text orientation.

        [MSDN documentation for ChartTitle.Orientation](https://learn.microsoft.com/en-us/office/vba/api/excel.charttitle.orientation).
        */
        void SetOrientation(long orientation);

        /**
        Returns the position of the chart title on the chart. Since Excel 2007.

        [MSDN documentation for ChartTitle.Position](https://learn.microsoft.com/en-us/office/vba/api/excel.charttitle.position).
        */
        XlChartElementPosition GetPosition();

        /**
        Sets the position of the chart title on the chart. Since Excel 2007.

        [MSDN documentation for ChartTitle.Position](https://learn.microsoft.com/en-us/office/vba/api/excel.charttitle.position).
        */
        void SetPosition(XlChartElementPosition position);

        /**
        Returns the reading order for the specified object. Can be one of the following constants: xlRTL (right-to-left), xlLTR (left-to-right), or xlContext.

        [MSDN documentation for ChartTitle.ReadingOrder](https://learn.microsoft.com/en-us/office/vba/api/excel.charttitle.readingorder).
        */
        long GetReadingOrder();

        /**
        Sets the reading order for the specified object. Can be one of the following constants: xlRTL (right-to-left), xlLTR (left-to-right), or xlContext.

        [MSDN documentation for ChartTitle.ReadingOrder](https://learn.microsoft.com/en-us/office/vba/api/excel.charttitle.readingorder).
        */
        void SetReadingOrder(long readingOrder);

        /**
        Returns a Boolean value that determines if the object has a shadow.

        [MSDN documentation for ChartTitle.Shadow](https://learn.microsoft.com/en-us/office/vba/api/excel.charttitle.shadow).
        */
        bool GetShadow();

        /**
        Sets a Boolean value that determines if the object has a shadow.

        [MSDN documentation for ChartTitle.Shadow](https://learn.microsoft.com/en-us/office/vba/api/excel.charttitle.shadow).
        */
        void SetShadow(bool shadow);

        /**
        Returns the text for the specified object.

        [MSDN documentation for ChartTitle.Text](https://learn.microsoft.com/en-us/office/vba/api/excel.charttitle.text).
        */
        wxString GetText();

        /**
        Sets the text for the specified object.

        [MSDN documentation for ChartTitle.Text](https://learn.microsoft.com/en-us/office/vba/api/excel.charttitle.text).
        */
        void SetText(const wxString& text);

        /**
        Returns a Double value that represents the distance, in points, from the top edge of the object to the top of row 1 (on a worksheet) or the top of the chart area (on a chart).

        [MSDN documentation for ChartTitle.Top](https://learn.microsoft.com/en-us/office/vba/api/excel.charttitle.top).
        */
        double GetTop();

        /**
        Sets a Double value that represents the distance, in points, from the top edge of the object to the top of row 1 (on a worksheet) or the top of the chart area (on a chart).

        [MSDN documentation for ChartTitle.Top](https://learn.microsoft.com/en-us/office/vba/api/excel.charttitle.top).
        */
        void SetTop(double top);

        /**
        Returns a Variant value that represents the vertical alignment of the specified object.

        [MSDN documentation for ChartTitle.VerticalAlignment](https://learn.microsoft.com/en-us/office/vba/api/excel.charttitle.verticalalignment).
        */
        long GetVerticalAlignment();

        /**
        Sets a Variant value that represents the vertical alignment of the specified object.

        [MSDN documentation for ChartTitle.VerticalAlignment](https://learn.microsoft.com/en-us/office/vba/api/excel.charttitle.verticalalignment).
        */
        void SetVerticalAlignment(long verticalAlignment);

         /**
        Returns the width, in points, of the object. Since Excel 2010.

        [MSDN documentation for ChartTitle.Width](https://learn.microsoft.com/en-us/office/vba/api/excel.charttitle.width).
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
