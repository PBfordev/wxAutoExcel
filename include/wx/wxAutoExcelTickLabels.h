/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Modified by:
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
/////////////////////////////////////////////////////////////////////////////


#ifndef _WXAUTOEXCEL_TICKLABELS_H
#define _WXAUTOEXCEL_TICKLABELS_H

#include "wx/wxAutoExcel_defs.h"

#if WXAUTOEXCEL_USE_CHARTS

#include "wx/wxAutoExcel_object.h"

namespace wxAutoExcel {


    /**
    Represents Microsoft Excel TickLabels object.
    */
   class WXDLLIMPEXP_WXAUTOEXCEL wxExcelTickLabels : public wxExcelObject
    {
    public:
        // ***** METHODS *****

        /**
        Deletes the object.

        [MSDN documentation for TickLabels.Delete](http://msdn.microsoft.com/en-us/library/bb179058).
        */
        bool Delete();

        /**
        Selects the object.

        [MSDN documentation for TickLabels.Select](http://msdn.microsoft.com/en-us/library/bb214081).
        */
        bool Select();

        // ***** PROPERTIES *****

        /**
        Returns a Long value that represents the alignment for the specified phonetic text or tick label.

        [MSDN documentation for TickLabels.Alignment](http://msdn.microsoft.com/en-us/library/bb238261).
        */
        long GetAlignment();

        /**
        Sets a Long value that represents the alignment for the specified phonetic text or tick label.

        [MSDN documentation for TickLabels.Alignment](http://msdn.microsoft.com/en-us/library/bb238261).
        */
        void SetAlignment(long alignment);

        /**
        True if the text in the object changes font size when the object size changes. The default value is True. Read/write Variant.

        [MSDN documentation for TickLabels.AutoScaleFont](http://msdn.microsoft.com/en-us/library/bb213978).
        */
        bool GetAutoScaleFont();

        /**
        True if the text in the object changes font size when the object size changes. The default value is True. Read/write Variant.

        [MSDN documentation for TickLabels.AutoScaleFont](http://msdn.microsoft.com/en-us/library/bb213978).
        */
        void SetAutoScaleFont(bool autoScaleFont);

        /**
        Returns a Long value that represents the number of levels of category tick labels.

        [MSDN documentation for TickLabels.Depth](http://msdn.microsoft.com/en-us/library/bb213989).
        */
        long GetDepth();

        /**
        Returns a Font object that represents the font of the specified object.

        [MSDN documentation for TickLabels.Font](http://msdn.microsoft.com/en-us/library/bb213994).
        */
        wxExcelFont GetFont();

        /**
        Returns the ChartFormat object.  Since Excel 2007.

        [MSDN documentation for TickLabels.Format](http://msdn.microsoft.com/en-us/library/bb242540).
        */
        wxExcelChartFormat GetFormat();

        /**
        Sets whether an axis is multilevel or not.  Since Excel 2007.

        [MSDN documentation for TickLabels.MultiLevel](http://msdn.microsoft.com/en-us/library/bb215931).
        */
        bool GetMultiLevel();

        /**
        Sets whether an axis is multilevel or not.  Since Excel 2007.

        [MSDN documentation for TickLabels.MultiLevel](http://msdn.microsoft.com/en-us/library/bb215931).
        */
        void SetMultiLevel(bool multiLevel);

        /**
        Returns a String value that represents the name of the object.

        [MSDN documentation for TickLabels.Name](http://msdn.microsoft.com/en-us/library/bb214002).
        */
        wxString GetName();

        /**
        Returns a String value that represents the format code for the object.

        [MSDN documentation for TickLabels.NumberFormat](http://msdn.microsoft.com/en-us/library/bb214006).
        */
        wxString GetNumberFormat();

        /**
        Sets a String value that represents the format code for the object.

        [MSDN documentation for TickLabels.NumberFormat](http://msdn.microsoft.com/en-us/library/bb214006).
        */
        void SetNumberFormat(const wxString& numberFormat);

        /**
        True if the number format is linked to the cells (so that the number format changes in the labels when it changes in the cells).

        [MSDN documentation for TickLabels.NumberFormatLinked](http://msdn.microsoft.com/en-us/library/bb214011).
        */
        bool GetNumberFormatLinked();

        /**
        True if the number format is linked to the cells (so that the number format changes in the labels when it changes in the cells).

        [MSDN documentation for TickLabels.NumberFormatLinked](http://msdn.microsoft.com/en-us/library/bb214011).
        */
        void SetNumberFormatLinked(bool numberFormatLinked);

        /**
        Returns a Variant value that represents the format code for the object as a string in the language of the user.

        [MSDN documentation for TickLabels.NumberFormatLocal](http://msdn.microsoft.com/en-us/library/bb214017).
        */
        wxString GetNumberFormatLocal();

        /**
        Sets a Variant value that represents the format code for the object as a string in the language of the user.

        [MSDN documentation for TickLabels.NumberFormatLocal](http://msdn.microsoft.com/en-us/library/bb214017).
        */
        void SetNumberFormatLocal(const wxString& numberFormatLocal);

        /**
        Returns a Long value that represents the distance between the levels of labels, and the distance between the first level and the axis line.

        [MSDN documentation for TickLabels.Offset](http://msdn.microsoft.com/en-us/library/bb214027).
        */
        long GetOffset();

        /**
        Sets a Long value that represents the distance between the levels of labels, and the distance between the first level and the axis line.

        [MSDN documentation for TickLabels.Offset](http://msdn.microsoft.com/en-us/library/bb214027).
        */
        void SetOffset(long offset);

        /**
        Returns a Long value that represents the text orientation.

        [MSDN documentation for TickLabels.Orientation](http://msdn.microsoft.com/en-us/library/bb214031).
        */
        long GetOrientation();

        /**
        Sets a Long value that represents the text orientation.

        [MSDN documentation for TickLabels.Orientation](http://msdn.microsoft.com/en-us/library/bb214031).
        */
        void SetOrientation(long orientation);
        /**
        Returns the reading order for the specified object. Can be one of the following constants: xlRTL (right-to-left), xlLTR (left-to-right), or xlContext.

        [MSDN documentation for TickLabels.ReadingOrder](http://msdn.microsoft.com/en-us/library/bb214040).
        */
        long GetReadingOrder();

        /**
        Sets the reading order for the specified object. Can be one of the following constants: xlRTL (right-to-left), xlLTR (left-to-right), or xlContext.

        [MSDN documentation for TickLabels.ReadingOrder](http://msdn.microsoft.com/en-us/library/bb214040).
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
