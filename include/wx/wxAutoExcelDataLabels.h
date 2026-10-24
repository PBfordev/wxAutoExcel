/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// License:     MIT license
/////////////////////////////////////////////////////////////////////////////


#ifndef _WXAUTOEXCEL_DATALABELS_H
#define _WXAUTOEXCEL_DATALABELS_H

#include "wx/wxAutoExcel_defs.h"

#if WXAUTOEXCEL_USE_CHARTS

#include "wx/wxAutoExcel_object.h"
#include "wx/wxAutoExcel_enums.h"

namespace wxAutoExcel {

    /**
    @brief Represents Microsoft Excel DataLabel object.
    */
   class WXDLLIMPEXP_WXAUTOEXCEL wxExcelDataLabel : public wxExcelObject
    {
    public:
        // ***** METHODS *****

        /**
        Deletes the object.

        [MSDN documentation for DataLabel.Delete](https://learn.microsoft.com/en-us/office/vba/api/excel.datalabel.delete).
        */
        bool Delete();

        /**
        Selects the object.

        [MSDN documentation for DataLabel.Select](https://learn.microsoft.com/en-us/office/vba/api/excel.datalabel.select).
        */
        bool Select();

        // ***** PROPERTIES *****

        /**
        True if the text in the object changes font size when the object size changes. The default value is True. Read/write Variant.

        [MSDN documentation for DataLabel.AutoScaleFont](https://learn.microsoft.com/en-us/previous-versions/office/developer/office-2007/bb179574%28v=office.12%29).
        */
        bool GetAutoScaleFont();

        /**
        True if the text in the object changes font size when the object size changes. The default value is True. Read/write Variant.

        [MSDN documentation for DataLabel.AutoScaleFont](https://learn.microsoft.com/en-us/previous-versions/office/developer/office-2007/bb179574%28v=office.12%29).
        */
        void SetAutoScaleFont(bool autoScaleFont);

        /**
        True if the object automatically generates appropriate text based on context.

        [MSDN documentation for DataLabel.AutoText](https://learn.microsoft.com/en-us/office/vba/api/excel.datalabel.autotext).
        */
        bool GetAutoText();

        /**
        True if the object automatically generates appropriate text based on context.

        [MSDN documentation for DataLabel.AutoText](https://learn.microsoft.com/en-us/office/vba/api/excel.datalabel.autotext).
        */
        void SetAutoText(bool autoText);

        /**
        Returns a String value that represents the data label text.

        [MSDN documentation for DataLabel.Caption](https://learn.microsoft.com/en-us/office/vba/api/excel.datalabel.caption).
        */
        wxString GetCaption();

        /**
        Sets a String value that represents the data label text.

        [MSDN documentation for DataLabel.Caption](https://learn.microsoft.com/en-us/office/vba/api/excel.datalabel.caption).
        */
        void SetCaption(const wxString& caption);

        /**
        Returns a Characters object that represents a range of characters within the object text. You can use the Characters object to format characters within a text string.

        Excel fails when DataLabel.Characters gets any argument, so only the default values of start and length work.

        [MSDN documentation for DataLabel.Characters](https://learn.microsoft.com/en-us/office/vba/api/excel.datalabel.characters).
        */
        wxExcelCharacters GetCharacters(long start = 1, long* length = NULL);

        /**
        Returns the ChartFormat object.  Since Excel 2007.

        [MSDN documentation for DataLabel.Format](https://learn.microsoft.com/en-us/office/vba/api/excel.datalabel.format).
        */
        wxExcelChartFormat GetFormat();

        /**
        Returns a Variant value that represents the horizontal alignment for the specified object.

        [MSDN documentation for DataLabel.HorizontalAlignment](https://learn.microsoft.com/en-us/office/vba/api/excel.datalabel.horizontalalignment).
        */
        long GetHorizontalAlignment();

        /**
        Sets a Variant value that represents the horizontal alignment for the specified object.

        [MSDN documentation for DataLabel.HorizontalAlignment](https://learn.microsoft.com/en-us/office/vba/api/excel.datalabel.horizontalalignment).
        */
        void SetHorizontalAlignment(long horizontalAlignment);

        /**
        Returns a Double value that represents the distance, in points, from the left edge of the object to the left edge of column A (on a worksheet) or the left edge of the chart area (on a chart).

        [MSDN documentation for DataLabel.Left](https://learn.microsoft.com/en-us/office/vba/api/excel.datalabel.left).
        */
        double GetLeft();

        /**
        Sets a Double value that represents the distance, in points, from the left edge of the object to the left edge of column A (on a worksheet) or the left edge of the chart area (on a chart).

        [MSDN documentation for DataLabel.Left](https://learn.microsoft.com/en-us/office/vba/api/excel.datalabel.left).
        */
        void SetLeft(double left);

        /**
        Returns a String value that represents the name of the object.

        [MSDN documentation for DataLabel.Name](https://learn.microsoft.com/en-us/office/vba/api/excel.datalabel.name).
        */
        wxString GetName();

        /**
        Returns a String value that represents the format code for the object.

        The name is spelled Numberformat rather than NumberFormat as in Excel, because the Windows headers define GetNumberFormat as a macro, which would otherwise rename the method.

        [MSDN documentation for DataLabel.NumberFormat](https://learn.microsoft.com/en-us/office/vba/api/excel.datalabel.numberformat).
        */
        wxString GetNumberformat();

        /**
        Sets a String value that represents the format code for the object.

        The name is spelled Numberformat rather than NumberFormat as in Excel, because the Windows headers define GetNumberFormat as a macro, which would otherwise rename the method.

        [MSDN documentation for DataLabel.NumberFormat](https://learn.microsoft.com/en-us/office/vba/api/excel.datalabel.numberformat).
        */
        void SetNumberformat(const wxString& numberFormat);

        /**
        True if the number format is linked to the cells (so that the number format changes in the labels when it changes in the cells).

        [MSDN documentation for DataLabel.NumberFormatLinked](https://learn.microsoft.com/en-us/office/vba/api/excel.datalabel.numberformatlinked).
        */
        bool GetNumberFormatLinked();

        /**
        True if the number format is linked to the cells (so that the number format changes in the labels when it changes in the cells).

        [MSDN documentation for DataLabel.NumberFormatLinked](https://learn.microsoft.com/en-us/office/vba/api/excel.datalabel.numberformatlinked).
        */
        void SetNumberFormatLinked(bool numberFormatLinked);

        /**
        Returns a Variant value that represents the format code for the object as a string in the language of the user.

        The name is spelled Numberformat rather than NumberFormat as in Excel, because the Windows headers define GetNumberFormat as a macro, which would otherwise rename the method.

        [MSDN documentation for DataLabel.NumberFormatLocal](https://learn.microsoft.com/en-us/office/vba/api/excel.datalabel.numberformatlocal).
        */
        wxString GetNumberformatLocal();

        /**
        Sets a Variant value that represents the format code for the object as a string in the language of the user.

        The name is spelled Numberformat rather than NumberFormat as in Excel, because the Windows headers define GetNumberFormat as a macro, which would otherwise rename the method.

        [MSDN documentation for DataLabel.NumberFormatLocal](https://learn.microsoft.com/en-us/office/vba/api/excel.datalabel.numberformatlocal).
        */
        void SetNumberformatLocal(const wxString& numberFormatLocal);

        /**
        Returns a Variant value that represents the text orientation.

        [MSDN documentation for DataLabel.Orientation](https://learn.microsoft.com/en-us/office/vba/api/excel.datalabel.orientation).
        */
        long GetOrientation();

        /**
        Sets a Variant value that represents the text orientation.

        [MSDN documentation for DataLabel.Orientation](https://learn.microsoft.com/en-us/office/vba/api/excel.datalabel.orientation).
        */
        void SetOrientation(long orientation);

        /**
        Returns a XlDataLabelPosition value that represents the position of the data label.

        [MSDN documentation for DataLabel.Position](https://learn.microsoft.com/en-us/office/vba/api/excel.datalabel.position).
        */
        XlDataLabelPosition GetPosition();

        /**
        Sets a XlDataLabelPosition value that represents the position of the data label.

        [MSDN documentation for DataLabel.Position](https://learn.microsoft.com/en-us/office/vba/api/excel.datalabel.position).
        */
        void SetPosition(XlDataLabelPosition position);

        /**
        Returns the reading order for the specified object. Can be one of the following constants: xlRTL (right-to-left), xlLTR (left-to-right), or xlContext.

        [MSDN documentation for DataLabel.ReadingOrder](https://learn.microsoft.com/en-us/office/vba/api/excel.datalabel.readingorder).
        */
        long GetReadingOrder();

        /**
        Sets the reading order for the specified object. Can be one of the following constants: xlRTL (right-to-left), xlLTR (left-to-right), or xlContext.

        [MSDN documentation for DataLabel.ReadingOrder](https://learn.microsoft.com/en-us/office/vba/api/excel.datalabel.readingorder).
        */
        void SetReadingOrder(long readingOrder);

        /**
        Returns a Variant representing the separator used for the data labels on a chart.

        Returns an empty string when Excel selects the separator, i.e. when Excel returns xlDataLabelSeparatorDefault.

        [MSDN documentation for DataLabel.Separator](https://learn.microsoft.com/en-us/office/vba/api/excel.datalabel.separator).
        */
        wxString GetSeparator();

        /**
        Sets a Variant representing the separator used for the data labels on a chart.

        [MSDN documentation for DataLabel.Separator](https://learn.microsoft.com/en-us/office/vba/api/excel.datalabel.separator).
        */
        void SetSeparator(const wxString& separator);

        /**
        Returns a Boolean value that determines if the object has a shadow.

        [MSDN documentation for DataLabel.Shadow](https://learn.microsoft.com/en-us/office/vba/api/excel.datalabel.shadow).
        */
        bool GetShadow();

        /**
        Sets a Boolean value that determines if the object has a shadow.

        [MSDN documentation for DataLabel.Shadow](https://learn.microsoft.com/en-us/office/vba/api/excel.datalabel.shadow).
        */
        void SetShadow(bool shadow);

        /**
        True to show the bubble size for the data labels on a chart. False to hide.

        [MSDN documentation for DataLabel.ShowBubbleSize](https://learn.microsoft.com/en-us/office/vba/api/excel.datalabel.showbubblesize).
        */
        bool GetShowBubbleSize();

        /**
        True to show the bubble size for the data labels on a chart. False to hide.

        [MSDN documentation for DataLabel.ShowBubbleSize](https://learn.microsoft.com/en-us/office/vba/api/excel.datalabel.showbubblesize).
        */
        void SetShowBubbleSize(bool showBubbleSize);

        /**
        True to display the category name for the data labels on a chart. False to hide.

        [MSDN documentation for DataLabel.ShowCategoryName](https://learn.microsoft.com/en-us/office/vba/api/excel.datalabel.showcategoryname).
        */
        bool GetShowCategoryName();

        /**
        True to display the category name for the data labels on a chart. False to hide.

        [MSDN documentation for DataLabel.ShowCategoryName](https://learn.microsoft.com/en-us/office/vba/api/excel.datalabel.showcategoryname).
        */
        void SetShowCategoryName(bool showCategoryName);

        /**
        True if the data label legend key is visible.

        [MSDN documentation for DataLabel.ShowLegendKey](https://learn.microsoft.com/en-us/office/vba/api/excel.datalabel.showlegendkey).
        */
        bool GetShowLegendKey();

        /**
        True if the data label legend key is visible.

        [MSDN documentation for DataLabel.ShowLegendKey](https://learn.microsoft.com/en-us/office/vba/api/excel.datalabel.showlegendkey).
        */
        void SetShowLegendKey(bool showLegendKey);

        /**
        True to display the percentage value for the data labels on a chart. False to hide.

        [MSDN documentation for DataLabel.ShowPercentage](https://learn.microsoft.com/en-us/office/vba/api/excel.datalabel.showpercentage).
        */
        bool GetShowPercentage();

        /**
        True to display the percentage value for the data labels on a chart. False to hide.

        [MSDN documentation for DataLabel.ShowPercentage](https://learn.microsoft.com/en-us/office/vba/api/excel.datalabel.showpercentage).
        */
        void SetShowPercentage(bool showPercentage);

        /**
        Toggles the Range field if it exists on the data label range.

        [Excel VBA documentation for DataLabel.ShowRange](https://learn.microsoft.com/en-us/office/vba/api/excel.datalabel.showrange)
        */
        bool GetShowRange();

        /**
        Toggles the Range field if it exists on the data label range.

        [Excel VBA documentation for DataLabel.ShowRange](https://learn.microsoft.com/en-us/office/vba/api/excel.datalabel.showrange)
        */
        void SetShowRange(bool showRange);

        /**
        Returns a Boolean to indicate the series name display behavior for the data labels on a chart. True to show the series name. False to hide.

        [MSDN documentation for DataLabel.ShowSeriesName](https://learn.microsoft.com/en-us/office/vba/api/excel.datalabel.showseriesname).
        */
        bool GetShowSeriesName();

        /**
        Sets a Boolean to indicate the series name display behavior for the data labels on a chart. True to show the series name. False to hide.

        [MSDN documentation for DataLabel.ShowSeriesName](https://learn.microsoft.com/en-us/office/vba/api/excel.datalabel.showseriesname).
        */
        void SetShowSeriesName(bool showSeriesName);

        /**
        Returns a Boolean corresponding to a specified chart's data label values display behavior. True displays the values. False to hide.

        [MSDN documentation for DataLabel.ShowValue](https://learn.microsoft.com/en-us/office/vba/api/excel.datalabel.showvalue).
        */
        bool GetShowValue();

        /**
        Sets a Boolean corresponding to a specified chart's data label values display behavior. True displays the values. False to hide.

        [MSDN documentation for DataLabel.ShowValue](https://learn.microsoft.com/en-us/office/vba/api/excel.datalabel.showvalue).
        */
        void SetShowValue(bool showValue);

        /**
        Returns the text for the specified object.

        [MSDN documentation for DataLabel.Text](https://learn.microsoft.com/en-us/office/vba/api/excel.datalabel.text).
        */
        wxString GetText();

        /**
        Sets the text for the specified object.

        [MSDN documentation for DataLabel.Text](https://learn.microsoft.com/en-us/office/vba/api/excel.datalabel.text).
        */
        void SetText(const wxString& text);

        /**
        Returns a Double value that represents the distance, in points, from the top edge of the object to the top of row 1 (on a worksheet) or the top of the chart area (on a chart).

        [MSDN documentation for DataLabel.Top](https://learn.microsoft.com/en-us/office/vba/api/excel.datalabel.top).
        */
        double GetTop();

        /**
        Sets a Double value that represents the distance, in points, from the top edge of the object to the top of row 1 (on a worksheet) or the top of the chart area (on a chart).

        [MSDN documentation for DataLabel.Top](https://learn.microsoft.com/en-us/office/vba/api/excel.datalabel.top).
        */
        void SetTop(double top);

        /**
        Returns a Variant value that represents the vertical alignment of the specified object.

        [MSDN documentation for DataLabel.VerticalAlignment](https://learn.microsoft.com/en-us/office/vba/api/excel.datalabel.verticalalignment).
        */
        long GetVerticalAlignment();

        /**
        Sets a Variant value that represents the vertical alignment of the specified object.

        [MSDN documentation for DataLabel.VerticalAlignment](https://learn.microsoft.com/en-us/office/vba/api/excel.datalabel.verticalalignment).
        */
        void SetVerticalAlignment(long verticalAlignment);

        /**
        Returns "DataLabel".
        */
        virtual wxString GetAutoExcelObjectName_() const { return wxS("DataLabel"); }
    };

    /**
    @brief Represents Microsoft Excel DataLabels collection.
    */
   class WXDLLIMPEXP_WXAUTOEXCEL wxExcelDataLabels : public wxExcelObject
    {
    public:
        // ***** METHODS *****

        /**
        Deletes the object.

        [MSDN documentation for DataLabels.Delete](https://learn.microsoft.com/en-us/office/vba/api/excel.datalabels.delete).
        */
        bool Delete();

        //@{
        /**
        Returns a single object from a collection.

        [MSDN documentation for DataLabels.Item](https://learn.microsoft.com/en-us/office/vba/api/excel.datalabels.item).
        */
        wxExcelDataLabel Item(long index);
        wxExcelDataLabel operator[](long index);
        //@}

         /**
        Enables to take the contents and formatting of a single data label and apply it to every other data label in the series.

        [Excel VBA documentation for DataLabels.Propagate](https://learn.microsoft.com/en-us/office/vba/api/excel.datalabels.propagate)
        */
        void Propagate(long index);

        /**
        Selects the object.

        [MSDN documentation for DataLabels.Select]().
        */
        bool Select();

        // ***** PROPERTIES *****


        /**
        True if the text in the object changes font size when the object size changes. The default value is True. Read/write Variant.

        [MSDN documentation for DataLabels.AutoScaleFont](https://learn.microsoft.com/en-us/previous-versions/office/developer/office-2007/bb179615%28v=office.12%29).
        */
        bool GetAutoScaleFont();

        /**
        True if the text in the object changes font size when the object size changes. The default value is True. Read/write Variant.

        [MSDN documentation for DataLabels.AutoScaleFont](https://learn.microsoft.com/en-us/previous-versions/office/developer/office-2007/bb179615%28v=office.12%29).
        */
        void SetAutoScaleFont(bool autoScaleFont);

        /**
        True if the object automatically generates appropriate text based on context.

        [MSDN documentation for DataLabels.AutoText](https://learn.microsoft.com/en-us/office/vba/api/excel.datalabels.autotext).
        */
        bool GetAutoText();

        /**
        True if the object automatically generates appropriate text based on context.

        [MSDN documentation for DataLabels.AutoText](https://learn.microsoft.com/en-us/office/vba/api/excel.datalabels.autotext).
        */
        void SetAutoText(bool autoText);

        /**
        Returns a Long value that represents the number of objects in the collection.

        [MSDN documentation for DataLabels.Count](https://learn.microsoft.com/en-us/office/vba/api/excel.datalabels.count).
        */
        long GetCount();

        /**
        Returns the ChartFormat object.  Since Excel 2007.

        [MSDN documentation for DataLabels.Format](https://learn.microsoft.com/en-us/office/vba/api/excel.datalabels.format).
        */
        wxExcelChartFormat GetFormat();

        /**
        Returns a Variant value that represents the horizontal alignment for the specified object.

        [MSDN documentation for DataLabels.HorizontalAlignment](https://learn.microsoft.com/en-us/office/vba/api/excel.datalabels.horizontalalignment).
        */
        long GetHorizontalAlignment();

        /**
        Sets a Variant value that represents the horizontal alignment for the specified object.

        [MSDN documentation for DataLabels.HorizontalAlignment](https://learn.microsoft.com/en-us/office/vba/api/excel.datalabels.horizontalalignment).
        */
        void SetHorizontalAlignment(long horizontalAlignment);

        /**
        Returns a String value that represents the name of the object.

        [MSDN documentation for DataLabels.Name](https://learn.microsoft.com/en-us/office/vba/api/excel.datalabels.name).
        */
        wxString GetName();

        /**
        Returns a String value that represents the format code for the object.

        The name is spelled Numberformat rather than NumberFormat as in Excel, because the Windows headers define GetNumberFormat as a macro, which would otherwise rename the method.

        [MSDN documentation for DataLabels.NumberFormat](https://learn.microsoft.com/en-us/office/vba/api/excel.datalabels.numberformat).
        */
        wxString GetNumberformat();

        /**
        Sets a String value that represents the format code for the object.

        The name is spelled Numberformat rather than NumberFormat as in Excel, because the Windows headers define GetNumberFormat as a macro, which would otherwise rename the method.

        [MSDN documentation for DataLabels.NumberFormat](https://learn.microsoft.com/en-us/office/vba/api/excel.datalabels.numberformat).
        */
        void SetNumberformat(const wxString& numberFormat);

        /**
        True if the number format is linked to the cells (so that the number format changes in the labels when it changes in the cells).

        [MSDN documentation for DataLabels.NumberFormatLinked](https://learn.microsoft.com/en-us/office/vba/api/excel.datalabels.numberformatlinked).
        */
        bool GetNumberFormatLinked();

        /**
        True if the number format is linked to the cells (so that the number format changes in the labels when it changes in the cells).

        [MSDN documentation for DataLabels.NumberFormatLinked](https://learn.microsoft.com/en-us/office/vba/api/excel.datalabels.numberformatlinked).
        */
        void SetNumberFormatLinked(bool numberFormatLinked);

        /**
        Returns a Variant value that represents the format code for the object as a string in the language of the user.

        The name is spelled Numberformat rather than NumberFormat as in Excel, because the Windows headers define GetNumberFormat as a macro, which would otherwise rename the method.

        [MSDN documentation for DataLabels.NumberFormatLocal](https://learn.microsoft.com/en-us/office/vba/api/excel.datalabels.numberformatlocal).
        */
        wxString GetNumberformatLocal();

        /**
        Sets a Variant value that represents the format code for the object as a string in the language of the user.

        The name is spelled Numberformat rather than NumberFormat as in Excel, because the Windows headers define GetNumberFormat as a macro, which would otherwise rename the method.

        [MSDN documentation for DataLabels.NumberFormatLocal](https://learn.microsoft.com/en-us/office/vba/api/excel.datalabels.numberformatlocal).
        */
        void SetNumberformatLocal(const wxString& numberFormatLocal);

        /**
        Returns a Variant value that represents the text orientation.

        [MSDN documentation for DataLabels.Orientation](https://learn.microsoft.com/en-us/office/vba/api/excel.datalabels.orientation).
        */
        long GetOrientation();

        /**
        Sets a Variant value that represents the text orientation.

        [MSDN documentation for DataLabels.Orientation](https://learn.microsoft.com/en-us/office/vba/api/excel.datalabels.orientation).
        */
        void SetOrientation(long orientation);

        /**
        Returns a XlDataLabelPosition value that represents the position of the data label.

        [MSDN documentation for DataLabels.Position](https://learn.microsoft.com/en-us/office/vba/api/excel.datalabels.position).
        */
        XlDataLabelPosition GetPosition();

        /**
        Sets a XlDataLabelPosition value that represents the position of the data label.

        [MSDN documentation for DataLabels.Position](https://learn.microsoft.com/en-us/office/vba/api/excel.datalabels.position).
        */
        void SetPosition(XlDataLabelPosition position);

        /**
        Returns the reading order for the specified object. Can be one of the following constants: xlRTL (right-to-left), xlLTR (left-to-right), or xlContext.

        [MSDN documentation for DataLabels.ReadingOrder](https://learn.microsoft.com/en-us/office/vba/api/excel.datalabels.readingorder).
        */
        long GetReadingOrder();

        /**
        Sets the reading order for the specified object. Can be one of the following constants: xlRTL (right-to-left), xlLTR (left-to-right), or xlContext.

        [MSDN documentation for DataLabels.ReadingOrder](https://learn.microsoft.com/en-us/office/vba/api/excel.datalabels.readingorder).
        */
        void SetReadingOrder(long readingOrder);

        /**
        Returns a Variant representing the separator used for the data labels on a chart.

        Returns an empty string when Excel selects the separator, i.e. when Excel returns xlDataLabelSeparatorDefault.

        [MSDN documentation for DataLabels.Separator](https://learn.microsoft.com/en-us/office/vba/api/excel.datalabels.separator).
        */
        wxString GetSeparator();

        /**
        Sets a Variant representing the separator used for the data labels on a chart.

        [MSDN documentation for DataLabels.Separator](https://learn.microsoft.com/en-us/office/vba/api/excel.datalabels.separator).
        */
        void SetSeparator(const wxString& separator);

        /**
        Returns a Boolean value that determines if the object has a shadow.

        [MSDN documentation for DataLabels.Shadow](https://learn.microsoft.com/en-us/office/vba/api/excel.datalabels.shadow).
        */
        bool GetShadow();

        /**
        Sets a Boolean value that determines if the object has a shadow.

        [MSDN documentation for DataLabels.Shadow](https://learn.microsoft.com/en-us/office/vba/api/excel.datalabels.shadow).
        */
        void SetShadow(bool shadow);

        /**
        True to show the bubble size for the data labels on a chart. False to hide.

        [MSDN documentation for DataLabels.ShowBubbleSize](https://learn.microsoft.com/en-us/office/vba/api/excel.datalabels.showbubblesize).
        */
        bool GetShowBubbleSize();

        /**
        True to show the bubble size for the data labels on a chart. False to hide.

        [MSDN documentation for DataLabels.ShowBubbleSize](https://learn.microsoft.com/en-us/office/vba/api/excel.datalabels.showbubblesize).
        */
        void SetShowBubbleSize(bool showBubbleSize);

        /**
        True to display the category name for the data labels on a chart. False to hide.

        [MSDN documentation for DataLabels.ShowCategoryName](https://learn.microsoft.com/en-us/office/vba/api/excel.datalabels.showcategoryname).
        */
        bool GetShowCategoryName();

        /**
        True to display the category name for the data labels on a chart. False to hide.

        [MSDN documentation for DataLabels.ShowCategoryName](https://learn.microsoft.com/en-us/office/vba/api/excel.datalabels.showcategoryname).
        */
        void SetShowCategoryName(bool showCategoryName);

        /**
        True if the data label legend key is visible.

        [MSDN documentation for DataLabels.ShowLegendKey](https://learn.microsoft.com/en-us/office/vba/api/excel.datalabels.showlegendkey).
        */
        bool GetShowLegendKey();

        /**
        True if the data label legend key is visible.

        [MSDN documentation for DataLabels.ShowLegendKey](https://learn.microsoft.com/en-us/office/vba/api/excel.datalabels.showlegendkey).
        */
        void SetShowLegendKey(bool showLegendKey);

        /**
        True to display the percentage value for the data labels on a chart. False to hide.

        [MSDN documentation for DataLabels.ShowPercentage](https://learn.microsoft.com/en-us/office/vba/api/excel.datalabels.showpercentage).
        */
        bool GetShowPercentage();

        /**
        True to display the percentage value for the data labels on a chart. False to hide.

        [MSDN documentation for DataLabels.ShowPercentage](https://learn.microsoft.com/en-us/office/vba/api/excel.datalabels.showpercentage).
        */
        void SetShowPercentage(bool showPercentage);

        /**
        Toggles the Range field if it exists on the data label range.

        [Excel VBA documentation for DataLabels.ShowRange](https://learn.microsoft.com/en-us/office/vba/api/excel.datalabels.showrange)
        */
        bool GetShowRange();

        /**
        Toggles the Range field if it exists on the data label range.

        [Excel VBA documentation for DataLabels.ShowRange](https://learn.microsoft.com/en-us/office/vba/api/excel.datalabels.showrange)
        */
        void SetShowRange(bool showRange);

        /**
        Returns a Boolean to indicate the series name display behavior for the data labels on a chart. True to show the series name. False to hide.

        [MSDN documentation for DataLabels.ShowSeriesName](https://learn.microsoft.com/en-us/office/vba/api/excel.datalabels.showseriesname).
        */
        bool GetShowSeriesName();

        /**
        Sets a Boolean to indicate the series name display behavior for the data labels on a chart. True to show the series name. False to hide.

        [MSDN documentation for DataLabels.ShowSeriesName](https://learn.microsoft.com/en-us/office/vba/api/excel.datalabels.showseriesname).
        */
        void SetShowSeriesName(bool showSeriesName);

        /**
        Returns a Boolean corresponding to a specified chart's data label values display behavior. True displays the values. False to hide.

        [MSDN documentation for DataLabels.ShowValue](https://learn.microsoft.com/en-us/office/vba/api/excel.datalabels.showvalue).
        */
        bool GetShowValue();

        /**
        Sets a Boolean corresponding to a specified chart's data label values display behavior. True displays the values. False to hide.

        [MSDN documentation for DataLabels.ShowValue](https://learn.microsoft.com/en-us/office/vba/api/excel.datalabels.showvalue).
        */
        void SetShowValue(bool showValue);

        /**
        Returns a Variant value that represents the vertical alignment of the specified object.

        [MSDN documentation for DataLabels.VerticalAlignment](https://learn.microsoft.com/en-us/office/vba/api/excel.datalabels.verticalalignment).
        */
        long GetVerticalAlignment();

        /**
        Sets a Variant value that represents the vertical alignment of the specified object.

        [MSDN documentation for DataLabels.VerticalAlignment](https://learn.microsoft.com/en-us/office/vba/api/excel.datalabels.verticalalignment).
        */
        void SetVerticalAlignment(long verticalAlignment);

        /**
        The iterator over the items, see wxExcelCollectionIterator.
        */
        typedef wxExcelCollectionIterator<wxExcelDataLabels, wxExcelDataLabel> iterator;

        //@{
        /**
        Returns the iterator referring to the first item, or past the last item,
        which allows iterating over the items with a range-based for loop,
        see wxExcelCollectionIterator.
        */
        iterator begin() { return iterator(this, 1); }
        iterator end() { return iterator(this, GetCount() + 1); }
        //@}

        /**
        Returns "DataLabels".
        */
        virtual wxString GetAutoExcelObjectName_() const { return wxS("DataLabels"); }
    };


} // namespace wxAutoExcel

#endif // #if WXAUTOEXCEL_USE_CHARTS

#endif // _WXAUTOEXCEL_DATALABELS_H
