/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// License:     MIT license
/////////////////////////////////////////////////////////////////////////////


#ifndef _WXAUTOEXCEL_TICKLABELS_H
#define _WXAUTOEXCEL_TICKLABELS_H

#include "wx/wxAutoExcel_defs.h"

#if WXAUTOEXCEL_USE_CHARTS

#include "wx/wxAutoExcel_object.h"

namespace wxAutoExcel {


    /**
    @brief Represents Microsoft Excel TickLabels object.
    */
   class WXDLLIMPEXP_WXAUTOEXCEL wxExcelTickLabels : public wxExcelObject
    {
    public:
        // ***** METHODS *****

        /**
        Deletes the object.

        [MSDN documentation for TickLabels.Delete](https://learn.microsoft.com/en-us/office/vba/api/excel.ticklabels.delete).
        */
        bool Delete();

        /**
        Selects the object.

        [MSDN documentation for TickLabels.Select](https://learn.microsoft.com/en-us/office/vba/api/excel.ticklabels.select).
        */
        bool Select();

        // ***** PROPERTIES *****

        /**
        Returns a Long value that represents the alignment for the specified phonetic text or tick label.

        [MSDN documentation for TickLabels.Alignment](https://learn.microsoft.com/en-us/office/vba/api/excel.ticklabels.alignment).
        */
        long GetAlignment();

        /**
        Sets a Long value that represents the alignment for the specified phonetic text or tick label.

        [MSDN documentation for TickLabels.Alignment](https://learn.microsoft.com/en-us/office/vba/api/excel.ticklabels.alignment).
        */
        void SetAlignment(long alignment);

        /**
        True if the text in the object changes font size when the object size changes. The default value is True. Read/write Variant.

        [MSDN documentation for TickLabels.AutoScaleFont](https://learn.microsoft.com/en-us/previous-versions/office/developer/office-2007/bb213978%28v=office.12%29).
        */
        bool GetAutoScaleFont();

        /**
        True if the text in the object changes font size when the object size changes. The default value is True. Read/write Variant.

        [MSDN documentation for TickLabels.AutoScaleFont](https://learn.microsoft.com/en-us/previous-versions/office/developer/office-2007/bb213978%28v=office.12%29).
        */
        void SetAutoScaleFont(bool autoScaleFont);

        /**
        Returns a Long value that represents the number of levels of category tick labels.

        [MSDN documentation for TickLabels.Depth](https://learn.microsoft.com/en-us/office/vba/api/excel.ticklabels.depth).
        */
        long GetDepth();

        /**
        Returns a Font object that represents the font of the specified object.

        [MSDN documentation for TickLabels.Font](https://learn.microsoft.com/en-us/office/vba/api/excel.ticklabels.font).
        */
        wxExcelFont GetFont();

        /**
        Returns the ChartFormat object.  Since Excel 2007.

        [MSDN documentation for TickLabels.Format](https://learn.microsoft.com/en-us/office/vba/api/excel.ticklabels.format).
        */
        wxExcelChartFormat GetFormat();

        /**
        Sets whether an axis is multilevel or not.  Since Excel 2007.

        [MSDN documentation for TickLabels.MultiLevel](https://learn.microsoft.com/en-us/office/vba/api/excel.ticklabels.multilevel).
        */
        bool GetMultiLevel();

        /**
        Sets whether an axis is multilevel or not.  Since Excel 2007.

        [MSDN documentation for TickLabels.MultiLevel](https://learn.microsoft.com/en-us/office/vba/api/excel.ticklabels.multilevel).
        */
        void SetMultiLevel(bool multiLevel);

        /**
        Returns a String value that represents the name of the object.

        [MSDN documentation for TickLabels.Name](https://learn.microsoft.com/en-us/office/vba/api/excel.ticklabels.name).
        */
        wxString GetName();

        /**
        Returns a String value that represents the format code for the object.

        The name is spelled Numberformat rather than NumberFormat as in Excel, because the Windows headers define GetNumberFormat as a macro, which would otherwise rename the method.

        [MSDN documentation for TickLabels.NumberFormat](https://learn.microsoft.com/en-us/office/vba/api/excel.ticklabels.numberformat).
        */
        wxString GetNumberformat();

        /**
        Sets a String value that represents the format code for the object.

        The name is spelled Numberformat rather than NumberFormat as in Excel, because the Windows headers define GetNumberFormat as a macro, which would otherwise rename the method.

        [MSDN documentation for TickLabels.NumberFormat](https://learn.microsoft.com/en-us/office/vba/api/excel.ticklabels.numberformat).
        */
        void SetNumberformat(const wxString& numberFormat);

        /**
        True if the number format is linked to the cells (so that the number format changes in the labels when it changes in the cells).

        [MSDN documentation for TickLabels.NumberFormatLinked](https://learn.microsoft.com/en-us/office/vba/api/excel.ticklabels.numberformatlinked).
        */
        bool GetNumberFormatLinked();

        /**
        True if the number format is linked to the cells (so that the number format changes in the labels when it changes in the cells).

        [MSDN documentation for TickLabels.NumberFormatLinked](https://learn.microsoft.com/en-us/office/vba/api/excel.ticklabels.numberformatlinked).
        */
        void SetNumberFormatLinked(bool numberFormatLinked);

        /**
        Returns a Variant value that represents the format code for the object as a string in the language of the user.

        The name is spelled Numberformat rather than NumberFormat as in Excel, because the Windows headers define GetNumberFormat as a macro, which would otherwise rename the method.

        [MSDN documentation for TickLabels.NumberFormatLocal](https://learn.microsoft.com/en-us/office/vba/api/excel.ticklabels.numberformatlocal).
        */
        wxString GetNumberformatLocal();

        /**
        Sets a Variant value that represents the format code for the object as a string in the language of the user.

        The name is spelled Numberformat rather than NumberFormat as in Excel, because the Windows headers define GetNumberFormat as a macro, which would otherwise rename the method.

        [MSDN documentation for TickLabels.NumberFormatLocal](https://learn.microsoft.com/en-us/office/vba/api/excel.ticklabels.numberformatlocal).
        */
        void SetNumberformatLocal(const wxString& numberFormatLocal);

        /**
        Returns a Long value that represents the distance between the levels of labels, and the distance between the first level and the axis line.

        [MSDN documentation for TickLabels.Offset](https://learn.microsoft.com/en-us/office/vba/api/excel.ticklabels.offset).
        */
        long GetOffset();

        /**
        Sets a Long value that represents the distance between the levels of labels, and the distance between the first level and the axis line.

        [MSDN documentation for TickLabels.Offset](https://learn.microsoft.com/en-us/office/vba/api/excel.ticklabels.offset).
        */
        void SetOffset(long offset);

        /**
        Returns a Long value that represents the text orientation.

        [MSDN documentation for TickLabels.Orientation](https://learn.microsoft.com/en-us/office/vba/api/excel.ticklabels.orientation).
        */
        long GetOrientation();

        /**
        Sets a Long value that represents the text orientation.

        [MSDN documentation for TickLabels.Orientation](https://learn.microsoft.com/en-us/office/vba/api/excel.ticklabels.orientation).
        */
        void SetOrientation(long orientation);
        /**
        Returns the reading order for the specified object. Can be one of the following constants: xlRTL (right-to-left), xlLTR (left-to-right), or xlContext.

        [MSDN documentation for TickLabels.ReadingOrder](https://learn.microsoft.com/en-us/office/vba/api/excel.ticklabels.readingorder).
        */
        long GetReadingOrder();

        /**
        Sets the reading order for the specified object. Can be one of the following constants: xlRTL (right-to-left), xlLTR (left-to-right), or xlContext.

        [MSDN documentation for TickLabels.ReadingOrder](https://learn.microsoft.com/en-us/office/vba/api/excel.ticklabels.readingorder).
        */
        void SetReadingOrder(long readingOrder);

        /**
        Returns "TickLabels".
        */
        virtual wxString GetAutoExcelObjectName_() const { return wxS("TickLabels"); }
    };


} // namespace wxAutoExcel


#endif // #if WXAUTOEXCEL_USE_CHARTS

#endif // _WXAUTOEXCEL_TICKLABELS_H
