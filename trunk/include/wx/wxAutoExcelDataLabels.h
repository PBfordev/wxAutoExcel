/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Modified by:
// RCS-ID:      $Id: $
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
/////////////////////////////////////////////////////////////////////////////


#ifndef _WXAUTOEXCEL_DATALABELS_H
#define _WXAUTOEXCEL_DATALABELS_H

#include "wx/wxAutoExcel_defs.h"

#if WXAUTOEXCEL_USE_CHARTS

#include "wx/wxAutoExcelObject.h"
#include "wx/wxAutoExcel_enums.h"

namespace wxAutoExcel {

    /**
    Represents Microsoft Excel DataLabel object.
    */
   class WXDLLIMPEXP_WXAUTOEXCEL wxExcelDataLabel : public wxExcelObject
    {
    public:
        // ***** METHODS *****

        /**
        Deletes the object.

        [MSDN documentation for DataLabel.Delete](http://msdn.microsoft.com/en-us/library/bb211733).
        */
        bool Delete();

        /**
        Selects the object.

        [MSDN documentation for DataLabel.Select](http://msdn.microsoft.com/en-us/library/bb237826).
        */
        bool Select();

        // ***** PROPERTIES *****

        /**
        True if the text in the object changes font size when the object size changes. The default value is True. Read/write Variant.

        [MSDN documentation for DataLabel.AutoScaleFont](http://msdn.microsoft.com/en-us/library/bb179574).
        */
        bool GetAutoScaleFont();

        /**
        True if the text in the object changes font size when the object size changes. The default value is True. Read/write Variant.

        [MSDN documentation for DataLabel.AutoScaleFont](http://msdn.microsoft.com/en-us/library/bb179574).
        */
        void SetAutoScaleFont(bool autoScaleFont);

        /**
        True if the object automatically generates appropriate text based on context.

        [MSDN documentation for DataLabel.AutoText](http://msdn.microsoft.com/en-us/library/bb179575).
        */
        bool GetAutoText();

        /**
        True if the object automatically generates appropriate text based on context.

        [MSDN documentation for DataLabel.AutoText](http://msdn.microsoft.com/en-us/library/bb179575).
        */
        void SetAutoText(bool autoText);

        /**
        Returns a String value that represents the data label text.

        [MSDN documentation for DataLabel.Caption](http://msdn.microsoft.com/en-us/library/bb179578).
        */
        wxString GetCaption();

        /**
        Sets a String value that represents the data label text.

        [MSDN documentation for DataLabel.Caption](http://msdn.microsoft.com/en-us/library/bb179578).
        */
        void SetCaption(const wxString& caption);

        /**
        Returns a Characters object that represents a range of characters within the object text. You can use the Characters object to format characters within a text string.

        [MSDN documentation for DataLabel.Characters](http://msdn.microsoft.com/en-us/library/bb179582).
        */
        wxExcelCharacters GetCharacters();

        /**
        Returns the ChartFormat object.  Since Excel 2007.

        [MSDN documentation for DataLabel.Format](http://msdn.microsoft.com/en-us/library/bb242522).
        */
        wxExcelChartFormat GetFormat();

        /**
        Returns a Variant value that represents the horizontal alignment for the specified object.

        [MSDN documentation for DataLabel.HorizontalAlignment](http://msdn.microsoft.com/en-us/library/bb179586).
        */
        long GetHorizontalAlignment();

        /**
        Sets a Variant value that represents the horizontal alignment for the specified object.

        [MSDN documentation for DataLabel.HorizontalAlignment](http://msdn.microsoft.com/en-us/library/bb179586).
        */
        void SetHorizontalAlignment(long horizontalAlignment);

        /**
        Returns a Double value that represents the distance, in points, from the left edge of the object to the left edge of column A (on a worksheet) or the left edge of the chart area (on a chart).

        [MSDN documentation for DataLabel.Left](http://msdn.microsoft.com/en-us/library/bb179587).
        */
        double GetLeft();

        /**
        Sets a Double value that represents the distance, in points, from the left edge of the object to the left edge of column A (on a worksheet) or the left edge of the chart area (on a chart).

        [MSDN documentation for DataLabel.Left](http://msdn.microsoft.com/en-us/library/bb179587).
        */
        void SetLeft(double left);

        /**
        Returns a String value that represents the name of the object.

        [MSDN documentation for DataLabel.Name](http://msdn.microsoft.com/en-us/library/bb179590).
        */
        wxString GetName();

        /**
        Returns a String value that represents the format code for the object.

        [MSDN documentation for DataLabel.NumberFormat](http://msdn.microsoft.com/en-us/library/bb179592).
        */
        wxString GetNumberFormat();

        /**
        Sets a String value that represents the format code for the object.

        [MSDN documentation for DataLabel.NumberFormat](http://msdn.microsoft.com/en-us/library/bb179592).
        */
        void SetNumberFormat(const wxString& numberFormat);

        /**
        True if the number format is linked to the cells (so that the number format changes in the labels when it changes in the cells).

        [MSDN documentation for DataLabel.NumberFormatLinked](http://msdn.microsoft.com/en-us/library/bb179595).
        */
        bool GetNumberFormatLinked();

        /**
        True if the number format is linked to the cells (so that the number format changes in the labels when it changes in the cells).

        [MSDN documentation for DataLabel.NumberFormatLinked](http://msdn.microsoft.com/en-us/library/bb179595).
        */
        void SetNumberFormatLinked(bool numberFormatLinked);

        /**
        Returns a Variant value that represents the format code for the object as a string in the language of the user.

        [MSDN documentation for DataLabel.NumberFormatLocal](http://msdn.microsoft.com/en-us/library/bb179599).
        */
        wxString GetNumberFormatLocal();

        /**
        Sets a Variant value that represents the format code for the object as a string in the language of the user.

        [MSDN documentation for DataLabel.NumberFormatLocal](http://msdn.microsoft.com/en-us/library/bb179599).
        */
        void SetNumberFormatLocal(const wxString& numberFormatLocal);

        /**
        Returns a Variant value that represents the text orientation.

        [MSDN documentation for DataLabel.Orientation](http://msdn.microsoft.com/en-us/library/bb179602).
        */
        long GetOrientation();

        /**
        Sets a Variant value that represents the text orientation.

        [MSDN documentation for DataLabel.Orientation](http://msdn.microsoft.com/en-us/library/bb179602).
        */
        void SetOrientation(long orientation);

        /**
        Returns a XlDataLabelPosition value that represents the position of the data label.

        [MSDN documentation for DataLabel.Position](http://msdn.microsoft.com/en-us/library/bb179606).
        */
        XlDataLabelPosition GetPosition();

        /**
        Sets a XlDataLabelPosition value that represents the position of the data label.

        [MSDN documentation for DataLabel.Position](http://msdn.microsoft.com/en-us/library/bb179606).
        */
        void SetPosition(XlDataLabelPosition position);

        /**
        Returns the reading order for the specified object. Can be one of the following constants: xlRTL (right-to-left), xlLTR (left-to-right), or xlContext.

        [MSDN documentation for DataLabel.ReadingOrder](http://msdn.microsoft.com/en-us/library/bb179611).
        */
        long GetReadingOrder();

        /**
        Sets the reading order for the specified object. Can be one of the following constants: xlRTL (right-to-left), xlLTR (left-to-right), or xlContext.

        [MSDN documentation for DataLabel.ReadingOrder](http://msdn.microsoft.com/en-us/library/bb179611).
        */
        void SetReadingOrder(long readingOrder);

        /**
        Returns a Variant representing the separator used for the data labels on a chart.

        [MSDN documentation for DataLabel.Separator](http://msdn.microsoft.com/en-us/library/bb214539).
        */
        wxString GetSeparator();

        /**
        Sets a Variant representing the separator used for the data labels on a chart.

        [MSDN documentation for DataLabel.Separator](http://msdn.microsoft.com/en-us/library/bb214539).
        */
        void SetSeparator(const wxString& separator);

        /**
        Returns a Boolean value that determines if the object has a shadow.

        [MSDN documentation for DataLabel.Shadow](http://msdn.microsoft.com/en-us/library/bb214541).
        */
        bool GetShadow();

        /**
        Sets a Boolean value that determines if the object has a shadow.

        [MSDN documentation for DataLabel.Shadow](http://msdn.microsoft.com/en-us/library/bb214541).
        */
        void SetShadow(bool shadow);

        /**
        True to show the bubble size for the data labels on a chart. False to hide.

        [MSDN documentation for DataLabel.ShowBubbleSize](http://msdn.microsoft.com/en-us/library/bb214544).
        */
        bool GetShowBubbleSize();

        /**
        True to show the bubble size for the data labels on a chart. False to hide.

        [MSDN documentation for DataLabel.ShowBubbleSize](http://msdn.microsoft.com/en-us/library/bb214544).
        */
        void SetShowBubbleSize(bool showBubbleSize);

        /**
        True to display the category name for the data labels on a chart. False to hide.

        [MSDN documentation for DataLabel.ShowCategoryName](http://msdn.microsoft.com/en-us/library/bb214547).
        */
        bool GetShowCategoryName();

        /**
        True to display the category name for the data labels on a chart. False to hide.

        [MSDN documentation for DataLabel.ShowCategoryName](http://msdn.microsoft.com/en-us/library/bb214547).
        */
        void SetShowCategoryName(bool showCategoryName);

        /**
        True if the data label legend key is visible.

        [MSDN documentation for DataLabel.ShowLegendKey](http://msdn.microsoft.com/en-us/library/bb214551).
        */
        bool GetShowLegendKey();

        /**
        True if the data label legend key is visible.

        [MSDN documentation for DataLabel.ShowLegendKey](http://msdn.microsoft.com/en-us/library/bb214551).
        */
        void SetShowLegendKey(bool showLegendKey);

        /**
        True to display the percentage value for the data labels on a chart. False to hide.

        [MSDN documentation for DataLabel.ShowPercentage](http://msdn.microsoft.com/en-us/library/bb214558).
        */
        bool GetShowPercentage();

        /**
        True to display the percentage value for the data labels on a chart. False to hide.

        [MSDN documentation for DataLabel.ShowPercentage](http://msdn.microsoft.com/en-us/library/bb214558).
        */
        void SetShowPercentage(bool showPercentage);

        /**
        Returns a Boolean to indicate the series name display behavior for the data labels on a chart. True to show the series name. False to hide.

        [MSDN documentation for DataLabel.ShowSeriesName](http://msdn.microsoft.com/en-us/library/bb214561).
        */
        bool GetShowSeriesName();

        /**
        Sets a Boolean to indicate the series name display behavior for the data labels on a chart. True to show the series name. False to hide.

        [MSDN documentation for DataLabel.ShowSeriesName](http://msdn.microsoft.com/en-us/library/bb214561).
        */
        void SetShowSeriesName(bool showSeriesName);

        /**
        Returns a Boolean corresponding to a specified chart's data label values display behavior. True displays the values. False to hide.

        [MSDN documentation for DataLabel.ShowValue](http://msdn.microsoft.com/en-us/library/bb214565).
        */
        bool GetShowValue();

        /**
        Sets a Boolean corresponding to a specified chart's data label values display behavior. True displays the values. False to hide.

        [MSDN documentation for DataLabel.ShowValue](http://msdn.microsoft.com/en-us/library/bb214565).
        */
        void SetShowValue(bool showValue);

        /**
        Returns the text for the specified object.

        [MSDN documentation for DataLabel.Text](http://msdn.microsoft.com/en-us/library/bb214583).
        */
        wxString GetText();

        /**
        Sets the text for the specified object.

        [MSDN documentation for DataLabel.Text](http://msdn.microsoft.com/en-us/library/bb214583).
        */
        void SetText(const wxString& text);

        /**
        Returns a Double value that represents the distance, in points, from the top edge of the object to the top of row 1 (on a worksheet) or the top of the chart area (on a chart).

        [MSDN documentation for DataLabel.Top](http://msdn.microsoft.com/en-us/library/bb214585).
        */
        double GetTop();

        /**
        Sets a Double value that represents the distance, in points, from the top edge of the object to the top of row 1 (on a worksheet) or the top of the chart area (on a chart).

        [MSDN documentation for DataLabel.Top](http://msdn.microsoft.com/en-us/library/bb214585).
        */
        void SetTop(double top);

        /**
        Returns a Variant value that represents the vertical alignment of the specified object.

        [MSDN documentation for DataLabel.VerticalAlignment](http://msdn.microsoft.com/en-us/library/bb214587).
        */
        long GetVerticalAlignment();

        /**
        Sets a Variant value that represents the vertical alignment of the specified object.

        [MSDN documentation for DataLabel.VerticalAlignment](http://msdn.microsoft.com/en-us/library/bb214587).
        */
        void SetVerticalAlignment(long verticalAlignment);

        /**
        Returns "DataLabel".
        */
        virtual wxString GetAutoExcelObjectName_() const { return wxS("DataLabel"); }
    };

    /**
    Represents Microsoft Excel DataLabels collection.
    */
   class WXDLLIMPEXP_WXAUTOEXCEL wxExcelDataLabels : public wxExcelObject
    {
    public:
        // ***** METHODS *****

        /**
        Deletes the object.

        [MSDN documentation for DataLabels.Delete](http://msdn.microsoft.com/en-us/library/bb211735).
        */
        bool Delete();

        //@{
        /**
        Returns a single object from a collection.

        [MSDN documentation for DataLabels.Item](http://msdn.microsoft.com/en-us/library/bb211737).
        */
        wxExcelDataLabel Item(long index);
        wxExcelDataLabel operator[](long index);
        //@}

        /**
        Selects the object.

        [MSDN documentation for DataLabels.Select]().
        */
        bool Select();

        // ***** PROPERTIES *****
        

        /**
        True if the text in the object changes font size when the object size changes. The default value is True. Read/write Variant.

        [MSDN documentation for DataLabels.AutoScaleFont](http://msdn.microsoft.com/en-us/library/bb179615).
        */
        bool GetAutoScaleFont();

        /**
        True if the text in the object changes font size when the object size changes. The default value is True. Read/write Variant.

        [MSDN documentation for DataLabels.AutoScaleFont](http://msdn.microsoft.com/en-us/library/bb179615).
        */
        void SetAutoScaleFont(bool autoScaleFont);

        /**
        True if the object automatically generates appropriate text based on context.

        [MSDN documentation for DataLabels.AutoText](http://msdn.microsoft.com/en-us/library/bb179619).
        */
        bool GetAutoText();

        /**
        True if the object automatically generates appropriate text based on context.

        [MSDN documentation for DataLabels.AutoText](http://msdn.microsoft.com/en-us/library/bb179619).
        */
        void SetAutoText(bool autoText);

        /**
        Returns a Long value that represents the number of objects in the collection.

        [MSDN documentation for DataLabels.Count](http://msdn.microsoft.com/en-us/library/bb179622).
        */
        long GetCount();

        /**
        Returns the ChartFormat object.  Since Excel 2007.

        [MSDN documentation for DataLabels.Format](http://msdn.microsoft.com/en-us/library/bb242523).
        */
        wxExcelChartFormat GetFormat();

        /**
        Returns a Variant value that represents the horizontal alignment for the specified object.

        [MSDN documentation for DataLabels.HorizontalAlignment](http://msdn.microsoft.com/en-us/library/bb179626).
        */
        long GetHorizontalAlignment();

        /**
        Sets a Variant value that represents the horizontal alignment for the specified object.

        [MSDN documentation for DataLabels.HorizontalAlignment](http://msdn.microsoft.com/en-us/library/bb179626).
        */
        void SetHorizontalAlignment(long horizontalAlignment);

        /**
        Returns a String value that represents the name of the object.

        [MSDN documentation for DataLabels.Name](http://msdn.microsoft.com/en-us/library/bb179630).
        */
        wxString GetName();

        /**
        Returns a String value that represents the format code for the object.

        [MSDN documentation for DataLabels.NumberFormat](http://msdn.microsoft.com/en-us/library/bb179634).
        */
        wxString GetNumberFormat();

        /**
        Sets a String value that represents the format code for the object.

        [MSDN documentation for DataLabels.NumberFormat](http://msdn.microsoft.com/en-us/library/bb179634).
        */
        void SetNumberFormat(const wxString& numberFormat);

        /**
        True if the number format is linked to the cells (so that the number format changes in the labels when it changes in the cells).

        [MSDN documentation for DataLabels.NumberFormatLinked](http://msdn.microsoft.com/en-us/library/bb179638).
        */
        bool GetNumberFormatLinked();

        /**
        True if the number format is linked to the cells (so that the number format changes in the labels when it changes in the cells).

        [MSDN documentation for DataLabels.NumberFormatLinked](http://msdn.microsoft.com/en-us/library/bb179638).
        */
        void SetNumberFormatLinked(bool numberFormatLinked);

        /**
        Returns a Variant value that represents the format code for the object as a string in the language of the user.

        [MSDN documentation for DataLabels.NumberFormatLocal](http://msdn.microsoft.com/en-us/library/bb179640).
        */
        wxString GetNumberFormatLocal();

        /**
        Sets a Variant value that represents the format code for the object as a string in the language of the user.

        [MSDN documentation for DataLabels.NumberFormatLocal](http://msdn.microsoft.com/en-us/library/bb179640).
        */
        void SetNumberFormatLocal(const wxString& numberFormatLocal);

        /**
        Returns a Variant value that represents the text orientation.

        [MSDN documentation for DataLabels.Orientation](http://msdn.microsoft.com/en-us/library/bb179643).
        */
        long GetOrientation();

        /**
        Sets a Variant value that represents the text orientation.

        [MSDN documentation for DataLabels.Orientation](http://msdn.microsoft.com/en-us/library/bb179643).
        */
        void SetOrientation(long orientation);

        /**
        Returns a XlDataLabelPosition value that represents the position of the data label.

        [MSDN documentation for DataLabels.Position](http://msdn.microsoft.com/en-us/library/bb179648).
        */
        XlDataLabelPosition GetPosition();

        /**
        Sets a XlDataLabelPosition value that represents the position of the data label.

        [MSDN documentation for DataLabels.Position](http://msdn.microsoft.com/en-us/library/bb179648).
        */
        void SetPosition(XlDataLabelPosition position);

        /**
        Returns the reading order for the specified object. Can be one of the following constants: xlRTL (right-to-left), xlLTR (left-to-right), or xlContext.

        [MSDN documentation for DataLabels.ReadingOrder](http://msdn.microsoft.com/en-us/library/bb179650).
        */
        long GetReadingOrder();

        /**
        Sets the reading order for the specified object. Can be one of the following constants: xlRTL (right-to-left), xlLTR (left-to-right), or xlContext.

        [MSDN documentation for DataLabels.ReadingOrder](http://msdn.microsoft.com/en-us/library/bb179650).
        */
        void SetReadingOrder(long readingOrder);

        /**
        Returns a Variant representing the separator used for the data labels on a chart.

        [MSDN documentation for DataLabels.Separator](http://msdn.microsoft.com/en-us/library/bb214567).
        */
        wxString GetSeparator();

        /**
        Sets a Variant representing the separator used for the data labels on a chart.

        [MSDN documentation for DataLabels.Separator](http://msdn.microsoft.com/en-us/library/bb214567).
        */
        void SetSeparator(const wxString& separator);

        /**
        Returns a Boolean value that determines if the object has a shadow.

        [MSDN documentation for DataLabels.Shadow](http://msdn.microsoft.com/en-us/library/bb214569).
        */
        bool GetShadow();

        /**
        Sets a Boolean value that determines if the object has a shadow.

        [MSDN documentation for DataLabels.Shadow](http://msdn.microsoft.com/en-us/library/bb214569).
        */
        void SetShadow(bool shadow);

        /**
        True to show the bubble size for the data labels on a chart. False to hide.

        [MSDN documentation for DataLabels.ShowBubbleSize](http://msdn.microsoft.com/en-us/library/bb214570).
        */
        bool GetShowBubbleSize();

        /**
        True to show the bubble size for the data labels on a chart. False to hide.

        [MSDN documentation for DataLabels.ShowBubbleSize](http://msdn.microsoft.com/en-us/library/bb214570).
        */
        void SetShowBubbleSize(bool showBubbleSize);

        /**
        True to display the category name for the data labels on a chart. False to hide.

        [MSDN documentation for DataLabels.ShowCategoryName](http://msdn.microsoft.com/en-us/library/bb214572).
        */
        bool GetShowCategoryName();

        /**
        True to display the category name for the data labels on a chart. False to hide.

        [MSDN documentation for DataLabels.ShowCategoryName](http://msdn.microsoft.com/en-us/library/bb214572).
        */
        void SetShowCategoryName(bool showCategoryName);

        /**
        True if the data label legend key is visible.

        [MSDN documentation for DataLabels.ShowLegendKey](http://msdn.microsoft.com/en-us/library/bb214573).
        */
        bool GetShowLegendKey();

        /**
        True if the data label legend key is visible.

        [MSDN documentation for DataLabels.ShowLegendKey](http://msdn.microsoft.com/en-us/library/bb214573).
        */
        void SetShowLegendKey(bool showLegendKey);

        /**
        True to display the percentage value for the data labels on a chart. False to hide.

        [MSDN documentation for DataLabels.ShowPercentage](http://msdn.microsoft.com/en-us/library/bb214576).
        */
        bool GetShowPercentage();

        /**
        True to display the percentage value for the data labels on a chart. False to hide.

        [MSDN documentation for DataLabels.ShowPercentage](http://msdn.microsoft.com/en-us/library/bb214576).
        */
        void SetShowPercentage(bool showPercentage);

        /**
        Returns a Boolean to indicate the series name display behavior for the data labels on a chart. True to show the series name. False to hide.

        [MSDN documentation for DataLabels.ShowSeriesName](http://msdn.microsoft.com/en-us/library/bb214577).
        */
        bool GetShowSeriesName();

        /**
        Sets a Boolean to indicate the series name display behavior for the data labels on a chart. True to show the series name. False to hide.

        [MSDN documentation for DataLabels.ShowSeriesName](http://msdn.microsoft.com/en-us/library/bb214577).
        */
        void SetShowSeriesName(bool showSeriesName);

        /**
        Returns a Boolean corresponding to a specified chart's data label values display behavior. True displays the values. False to hide.

        [MSDN documentation for DataLabels.ShowValue](http://msdn.microsoft.com/en-us/library/bb214579).
        */
        bool GetShowValue();

        /**
        Sets a Boolean corresponding to a specified chart's data label values display behavior. True displays the values. False to hide.

        [MSDN documentation for DataLabels.ShowValue](http://msdn.microsoft.com/en-us/library/bb214579).
        */
        void SetShowValue(bool showValue);

        /**
        Returns a Variant value that represents the vertical alignment of the specified object.

        [MSDN documentation for DataLabels.VerticalAlignment](http://msdn.microsoft.com/en-us/library/bb214581).
        */
        long GetVerticalAlignment();

        /**
        Sets a Variant value that represents the vertical alignment of the specified object.

        [MSDN documentation for DataLabels.VerticalAlignment](http://msdn.microsoft.com/en-us/library/bb214581).
        */
        void SetVerticalAlignment(long verticalAlignment);

        /**
        Returns "DataLabels".
        */
        virtual wxString GetAutoExcelObjectName_() const { return wxS("DataLabels"); }
    };


} // namespace wxAutoExcel

#endif // #if WXAUTOEXCEL_USE_CHARTS

#endif // _WXAUTOEXCEL_DATALABELS_H
