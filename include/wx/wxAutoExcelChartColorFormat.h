/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// License:     MIT license
/////////////////////////////////////////////////////////////////////////////


#ifndef _WXAUTOEXCEL_CHARTCOLORFORMAT_H
#define _WXAUTOEXCEL_CHARTCOLORFORMAT_H

#include "wx/wxAutoExcel_defs.h"

#if WXAUTOEXCEL_USE_CHARTS

#include "wx/wxAutoExcel_object.h"
#include "wx/wxAutoExcel_enums.h"

namespace wxAutoExcel {

    /**
    @brief Represents Microsoft Excel ChartColorFormat object.
    */
   class WXDLLIMPEXP_WXAUTOEXCEL wxExcelChartColorFormat : public wxExcelObject
    {
    public:
        // ***** PROPERTIES *****

        /**
        Returns a Long value that represents the red-green-blue value of the specified color.

        [MSDN documentation for ChartColorFormat.RGB](http://msdn.microsoft.com/en-us/library/bb148840).
        */
        wxColour GetRGB();

        /**
        Returns a Long value that represents the color of a Color object, as an index in the current color scheme.

        [MSDN documentation for ChartColorFormat.SchemeColor](http://msdn.microsoft.com/en-us/library/bb148845).
        */
        long GetSchemeColor();

        /**
        Returns a Long value that that represents the color format type.

        [MSDN documentation for ChartColorFormat.Type](http://msdn.microsoft.com/en-us/library/bb148852).
        */
        MsoColorType GetType();

        /**
        Returns "ChartColorFormat".
        */
        virtual wxString GetAutoExcelObjectName_() const { return wxS("ChartColorFormat"); }
    };


} // namespace wxAutoExcel

#endif // #if WXAUTOEXCEL_USE_CHARTS

#endif // _WXAUTOEXCEL_CHARTCOLORFORMAT_H
