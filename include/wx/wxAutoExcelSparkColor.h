/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// License:     MIT license
/////////////////////////////////////////////////////////////////////////////


#ifndef _WXAUTOEXCEL_SPARKCOLOR_H
#define _WXAUTOEXCEL_SPARKCOLOR_H

#include "wx/wxAutoExcel_defs.h"

#if WXAUTOEXCEL_USE_CHARTS

#include "wx/wxAutoExcel_object.h"
#include "wx/wxAutoExcel_enums.h"


namespace wxAutoExcel {


    /**
    @brief Represents Microsoft Excel SparkColor object. Since Excel 2010.
    */
   class WXDLLIMPEXP_WXAUTOEXCEL wxExcelSparkColor: public wxExcelObject
    {
    public:
        // ***** PROPERTIES *****

        /**
        Returns a FormatColor object that you can use to set the color of the markers for points in a sparkline.

        [MSDN documentation for SparkColor.Color](https://learn.microsoft.com/en-us/office/vba/api/excel.sparkcolor.color).
        */
        wxExcelFormatColor GetColor();

        /**
        Returns the parent SparklineGroup object for the specified object.

        [MSDN documentation for SparkColor.Parent](https://learn.microsoft.com/en-us/office/vba/api/excel.sparkcolor.parent).
        */
        wxExcelSparklineGroup GetParent();

        /**
        Returns of sets whether the point is visible.

        [MSDN documentation for SparkColor.Visible](https://learn.microsoft.com/en-us/office/vba/api/excel.sparkcolor.visible).
        */
        bool GetVisible();

        /**
        Returns of sets whether the point is visible.

        [MSDN documentation for SparkColor.Visible](https://learn.microsoft.com/en-us/office/vba/api/excel.sparkcolor.visible).
        */
        void SetVisible(bool visible);


        /**
        Returns "SparkColor".
        */
        virtual wxString GetAutoExcelObjectName_() const { return wxS("SparkColor"); }
    };


} // namespace wxAutoExcel

#endif // #if WXAUTOEXCEL_USE_CHARTS

#endif // _WXAUTOEXCEL_SPARKCOLOR_H
