/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Modified by:
// RCS-ID:      $Id: $
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
/////////////////////////////////////////////////////////////////////////////


#ifndef _WXAUTOEXCEL_DISPLAYUNITLABEL_H
#define _WXAUTOEXCEL_DISPLAYUNITLABEL_H

#include "wx/wxAutoExcel_defs.h"

#if WXAUTOEXCEL_USE_CHARTS

#include "wx/wxAutoExcelObject.h"
#include "wx/wxAutoExcel_enums.h"

namespace wxAutoExcel {

    /**
    Represents Microsoft Excel DisplayUnitLabel object.
    */
   class WXDLLIMPEXP_WXAUTOEXCEL wxExcelDisplayUnitLabel : public wxExcelObject
    {
    public:
        // ***** METHODS *****

        /**
        Deletes the object.

        [MSDN documentation for DisplayUnitLabel.Delete](http://msdn.microsoft.com/en-us/library/bb211745).
        */
        bool Delete();

        /**
        Selects the object.

        [MSDN documentation for DisplayUnitLabel.Select](http://msdn.microsoft.com/en-us/library/bb237868).
        */
        bool Select();

        // ***** PROPERTIES *****

        /**
        True if the text in the object changes font size when the object size changes. The default value is True. Read/write Variant.

        [MSDN documentation for DisplayUnitLabel.AutoScaleFont](http://msdn.microsoft.com/en-us/library/bb236939).
        */
        bool GetAutoScaleFont();

        /**
        True if the text in the object changes font size when the object size changes. The default value is True. Read/write Variant.

        [MSDN documentation for DisplayUnitLabel.AutoScaleFont](http://msdn.microsoft.com/en-us/library/bb236939).
        */
        void SetAutoScaleFont(bool autoScaleFont);

        /**
        Returns a String value that represents the display unit label text.

        [MSDN documentation for DisplayUnitLabel.Caption](http://msdn.microsoft.com/en-us/library/bb236942).
        */
        wxString GetCaption();

        /**
        Sets a String value that represents the display unit label text.

        [MSDN documentation for DisplayUnitLabel.Caption](http://msdn.microsoft.com/en-us/library/bb236942).
        */
        void SetCaption(const wxString& caption);

        /**
        Returns a Characters object that represents a range of characters within the object text. You can use the Characters object to format characters within a text string.

        [MSDN documentation for DisplayUnitLabel.Characters](http://msdn.microsoft.com/en-us/library/bb236944).
        */
        wxExcelCharacters  GetCharacters();

        /**
        Returns the ChartFormat object. Since Excel 2007.

        [MSDN documentation for DisplayUnitLabel.Format](http://msdn.microsoft.com/en-us/library/bb242525).
        */
        wxExcelChartFormat GetFormat();

        /**
        Returns a Variant value that represents the horizontal alignment for the specified object.

        [MSDN documentation for DisplayUnitLabel.HorizontalAlignment](http://msdn.microsoft.com/en-us/library/bb236947).
        */
        long GetHorizontalAlignment();

        /**
        Sets a Variant value that represents the horizontal alignment for the specified object.

        [MSDN documentation for DisplayUnitLabel.HorizontalAlignment](http://msdn.microsoft.com/en-us/library/bb236947).
        */
        void SetHorizontalAlignment(long horizontalAlignment);

        /**
        Returns a Double value that represents the distance, in points, from the left edge of the object to the left edge of column A (on a worksheet) or the left edge of the chart area (on a chart).

        [MSDN documentation for DisplayUnitLabel.Left](http://msdn.microsoft.com/en-us/library/bb236950).
        */
        double GetLeft();

        /**
        Sets a Double value that represents the distance, in points, from the left edge of the object to the left edge of column A (on a worksheet) or the left edge of the chart area (on a chart).

        [MSDN documentation for DisplayUnitLabel.Left](http://msdn.microsoft.com/en-us/library/bb236950).
        */
        void SetLeft(double left);

        /**
        Returns a String value that represents the name of the object.

        [MSDN documentation for DisplayUnitLabel.Name](http://msdn.microsoft.com/en-us/library/bb236951).
        */
        wxString GetName();

        /**
        Returns a Variant value that represents the text orientation.

        [MSDN documentation for DisplayUnitLabel.Orientation](http://msdn.microsoft.com/en-us/library/bb236953).
        */
        long GetOrientation();

        /**
        Sets a Variant value that represents the text orientation.

        [MSDN documentation for DisplayUnitLabel.Orientation](http://msdn.microsoft.com/en-us/library/bb236953).
        */
        void SetOrientation(long orientation);

        /**
        Returns the position of the unit label on an axis in the chart. Since Excel 2007.

        [MSDN documentation for DisplayUnitLabel.Position](http://msdn.microsoft.com/en-us/library/bb240037).
        */
        XlChartElementPosition GetPosition();

        /**
        Sets the position of the unit label on an axis in the chart. Since Excel 2007.

        [MSDN documentation for DisplayUnitLabel.Position](http://msdn.microsoft.com/en-us/library/bb240037).
        */
        void SetPosition(XlChartElementPosition position);

        /**
        Returns the reading order for the specified object. Can be one of the following constants: xlRTL (right-to-left), xlLTR (left-to-right), or xlContext.

        [MSDN documentation for DisplayUnitLabel.ReadingOrder](http://msdn.microsoft.com/en-us/library/bb236955).
        */
        long GetReadingOrder();

        /**
        Sets the reading order for the specified object. Can be one of the following constants: xlRTL (right-to-left), xlLTR (left-to-right), or xlContext.

        [MSDN documentation for DisplayUnitLabel.ReadingOrder](http://msdn.microsoft.com/en-us/library/bb236955).
        */
        void SetReadingOrder(long readingOrder);

        /**
        Returns a Boolean value that determines if the object has a shadow.

        [MSDN documentation for DisplayUnitLabel.Shadow](http://msdn.microsoft.com/en-us/library/bb214594).
        */
        bool GetShadow();

        /**
        Sets a Boolean value that determines if the object has a shadow.

        [MSDN documentation for DisplayUnitLabel.Shadow](http://msdn.microsoft.com/en-us/library/bb214594).
        */
        void SetShadow(bool shadow);

        /**
        Returns the text for the specified object.

        [MSDN documentation for DisplayUnitLabel.Text](http://msdn.microsoft.com/en-us/library/bb214596).
        */
        wxString GetText();

        /**
        Sets the text for the specified object.

        [MSDN documentation for DisplayUnitLabel.Text](http://msdn.microsoft.com/en-us/library/bb214596).
        */
        void SetText(const wxString& text);

        /**
        Returns a Double value that represents the distance, in points, from the top edge of the object to the top of row 1 (on a worksheet) or the top of the chart area (on a chart).

        [MSDN documentation for DisplayUnitLabel.Top](http://msdn.microsoft.com/en-us/library/bb214598).
        */
        double GetTop();

        /**
        Sets a Double value that represents the distance, in points, from the top edge of the object to the top of row 1 (on a worksheet) or the top of the chart area (on a chart).

        [MSDN documentation for DisplayUnitLabel.Top](http://msdn.microsoft.com/en-us/library/bb214598).
        */
        void SetTop(double top);

        /**
        Returns a Variant value that represents the vertical alignment of the specified object.

        [MSDN documentation for DisplayUnitLabel.VerticalAlignment](http://msdn.microsoft.com/en-us/library/bb214600).
        */
        long GetVerticalAlignment();

        /**
        Sets a Variant value that represents the vertical alignment of the specified object.

        [MSDN documentation for DisplayUnitLabel.VerticalAlignment](http://msdn.microsoft.com/en-us/library/bb214600).
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
