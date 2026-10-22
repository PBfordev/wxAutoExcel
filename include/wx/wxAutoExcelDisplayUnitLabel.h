/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// License:     MIT license
/////////////////////////////////////////////////////////////////////////////


#ifndef _WXAUTOEXCEL_DISPLAYUNITLABEL_H
#define _WXAUTOEXCEL_DISPLAYUNITLABEL_H

#include "wx/wxAutoExcel_defs.h"

#if WXAUTOEXCEL_USE_CHARTS

#include "wx/wxAutoExcel_object.h"
#include "wx/wxAutoExcel_enums.h"

namespace wxAutoExcel {

    /**
    @brief Represents Microsoft Excel DisplayUnitLabel object.
    */
   class WXDLLIMPEXP_WXAUTOEXCEL wxExcelDisplayUnitLabel : public wxExcelObject
    {
    public:
        // ***** METHODS *****

        /**
        Deletes the object.

        [MSDN documentation for DisplayUnitLabel.Delete](https://learn.microsoft.com/en-us/office/vba/api/excel.displayunitlabel.delete).
        */
        bool Delete();

        /**
        Selects the object.

        [MSDN documentation for DisplayUnitLabel.Select](https://learn.microsoft.com/en-us/office/vba/api/excel.displayunitlabel.select).
        */
        bool Select();

        // ***** PROPERTIES *****

        /**
        True if the text in the object changes font size when the object size changes. The default value is True. Read/write Variant.

        [MSDN documentation for DisplayUnitLabel.AutoScaleFont](https://learn.microsoft.com/en-us/previous-versions/office/developer/office-2007/bb236939%28v=office.12%29).
        */
        bool GetAutoScaleFont();

        /**
        True if the text in the object changes font size when the object size changes. The default value is True. Read/write Variant.

        [MSDN documentation for DisplayUnitLabel.AutoScaleFont](https://learn.microsoft.com/en-us/previous-versions/office/developer/office-2007/bb236939%28v=office.12%29).
        */
        void SetAutoScaleFont(bool autoScaleFont);

        /**
        Returns a String value that represents the display unit label text.

        [MSDN documentation for DisplayUnitLabel.Caption](https://learn.microsoft.com/en-us/office/vba/api/excel.displayunitlabel.caption).
        */
        wxString GetCaption();

        /**
        Sets a String value that represents the display unit label text.

        [MSDN documentation for DisplayUnitLabel.Caption](https://learn.microsoft.com/en-us/office/vba/api/excel.displayunitlabel.caption).
        */
        void SetCaption(const wxString& caption);

        /**
        Returns a Characters object that represents a range of characters within the object text. You can use the Characters object to format characters within a text string.

        [MSDN documentation for DisplayUnitLabel.Characters](https://learn.microsoft.com/en-us/office/vba/api/excel.displayunitlabel.characters).
        */
        wxExcelCharacters  GetCharacters(long start = 1, long* length = NULL);

        /**
        Returns the ChartFormat object. Since Excel 2007.

        [MSDN documentation for DisplayUnitLabel.Format](https://learn.microsoft.com/en-us/office/vba/api/excel.displayunitlabel.format).
        */
        wxExcelChartFormat GetFormat();

        /**
        Returns a Variant value that represents the horizontal alignment for the specified object.

        [MSDN documentation for DisplayUnitLabel.HorizontalAlignment](https://learn.microsoft.com/en-us/office/vba/api/excel.displayunitlabel.horizontalalignment).
        */
        long GetHorizontalAlignment();

        /**
        Sets a Variant value that represents the horizontal alignment for the specified object.

        [MSDN documentation for DisplayUnitLabel.HorizontalAlignment](https://learn.microsoft.com/en-us/office/vba/api/excel.displayunitlabel.horizontalalignment).
        */
        void SetHorizontalAlignment(long horizontalAlignment);

        /**
        Returns a Double value that represents the distance, in points, from the left edge of the object to the left edge of column A (on a worksheet) or the left edge of the chart area (on a chart).

        [MSDN documentation for DisplayUnitLabel.Left](https://learn.microsoft.com/en-us/office/vba/api/excel.displayunitlabel.left).
        */
        double GetLeft();

        /**
        Sets a Double value that represents the distance, in points, from the left edge of the object to the left edge of column A (on a worksheet) or the left edge of the chart area (on a chart).

        [MSDN documentation for DisplayUnitLabel.Left](https://learn.microsoft.com/en-us/office/vba/api/excel.displayunitlabel.left).
        */
        void SetLeft(double left);

        /**
        Returns a String value that represents the name of the object.

        [MSDN documentation for DisplayUnitLabel.Name](https://learn.microsoft.com/en-us/office/vba/api/excel.displayunitlabel.name).
        */
        wxString GetName();

        /**
        Returns a Variant value that represents the text orientation.

        [MSDN documentation for DisplayUnitLabel.Orientation](https://learn.microsoft.com/en-us/office/vba/api/excel.displayunitlabel.orientation).
        */
        long GetOrientation();

        /**
        Sets a Variant value that represents the text orientation.

        [MSDN documentation for DisplayUnitLabel.Orientation](https://learn.microsoft.com/en-us/office/vba/api/excel.displayunitlabel.orientation).
        */
        void SetOrientation(long orientation);

        /**
        Returns the position of the unit label on an axis in the chart. Since Excel 2007.

        [MSDN documentation for DisplayUnitLabel.Position](https://learn.microsoft.com/en-us/office/vba/api/excel.displayunitlabel.position).
        */
        XlChartElementPosition GetPosition();

        /**
        Sets the position of the unit label on an axis in the chart. Since Excel 2007.

        [MSDN documentation for DisplayUnitLabel.Position](https://learn.microsoft.com/en-us/office/vba/api/excel.displayunitlabel.position).
        */
        void SetPosition(XlChartElementPosition position);

        /**
        Returns the reading order for the specified object. Can be one of the following constants: xlRTL (right-to-left), xlLTR (left-to-right), or xlContext.

        [MSDN documentation for DisplayUnitLabel.ReadingOrder](https://learn.microsoft.com/en-us/office/vba/api/excel.displayunitlabel.readingorder).
        */
        long GetReadingOrder();

        /**
        Sets the reading order for the specified object. Can be one of the following constants: xlRTL (right-to-left), xlLTR (left-to-right), or xlContext.

        [MSDN documentation for DisplayUnitLabel.ReadingOrder](https://learn.microsoft.com/en-us/office/vba/api/excel.displayunitlabel.readingorder).
        */
        void SetReadingOrder(long readingOrder);

        /**
        Returns a Boolean value that determines if the object has a shadow.

        [MSDN documentation for DisplayUnitLabel.Shadow](https://learn.microsoft.com/en-us/office/vba/api/excel.displayunitlabel.shadow).
        */
        bool GetShadow();

        /**
        Sets a Boolean value that determines if the object has a shadow.

        [MSDN documentation for DisplayUnitLabel.Shadow](https://learn.microsoft.com/en-us/office/vba/api/excel.displayunitlabel.shadow).
        */
        void SetShadow(bool shadow);

        /**
        Returns the text for the specified object.

        [MSDN documentation for DisplayUnitLabel.Text](https://learn.microsoft.com/en-us/office/vba/api/excel.displayunitlabel.text).
        */
        wxString GetText();

        /**
        Sets the text for the specified object.

        [MSDN documentation for DisplayUnitLabel.Text](https://learn.microsoft.com/en-us/office/vba/api/excel.displayunitlabel.text).
        */
        void SetText(const wxString& text);

        /**
        Returns a Double value that represents the distance, in points, from the top edge of the object to the top of row 1 (on a worksheet) or the top of the chart area (on a chart).

        [MSDN documentation for DisplayUnitLabel.Top](https://learn.microsoft.com/en-us/office/vba/api/excel.displayunitlabel.top).
        */
        double GetTop();

        /**
        Sets a Double value that represents the distance, in points, from the top edge of the object to the top of row 1 (on a worksheet) or the top of the chart area (on a chart).

        [MSDN documentation for DisplayUnitLabel.Top](https://learn.microsoft.com/en-us/office/vba/api/excel.displayunitlabel.top).
        */
        void SetTop(double top);

        /**
        Returns a Variant value that represents the vertical alignment of the specified object.

        [MSDN documentation for DisplayUnitLabel.VerticalAlignment](https://learn.microsoft.com/en-us/office/vba/api/excel.displayunitlabel.verticalalignment).
        */
        long GetVerticalAlignment();

        /**
        Sets a Variant value that represents the vertical alignment of the specified object.

        [MSDN documentation for DisplayUnitLabel.VerticalAlignment](https://learn.microsoft.com/en-us/office/vba/api/excel.displayunitlabel.verticalalignment).
        */
        void SetVerticalAlignment(long verticalAlignment);

        /**
        Returns "DisplayUnitLabel".
        */
        virtual wxString GetAutoExcelObjectName_() const { return wxS("DisplayUnitLabel"); }
    };


} // namespace wxAutoExcel

#endif // #if WXAUTOEXCEL_USE_CHARTS

#endif // _WXAUTOEXCEL_DISPLAYUNITLABEL_H
