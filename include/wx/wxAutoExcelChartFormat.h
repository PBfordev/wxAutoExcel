/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Modified by:
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
/////////////////////////////////////////////////////////////////////////////


#ifndef _WXAUTOEXCEL_CHARTFORMAT_H
#define _WXAUTOEXCEL_CHARTFORMAT_H

#include "wx/wxAutoExcel_defs.h"

#if WXAUTOEXCEL_USE_CHARTS

#include "wx/wxAutoExcel_object.h"
#include "wx/wxAutoExcel_enums.h"

namespace wxAutoExcel {

    /**
    Represents Microsoft Excel ChartFormat object.
    */
   class WXDLLIMPEXP_WXAUTOEXCEL wxExcelChartFormat: public wxExcelObject
    {
    public:
        // ***** PROPERTIES *****

        /**
        Returns a FillFormat object for the parent chart element that contains fill formatting properties for the chart element.  Since Excel 2007.

        [MSDN documentation for ChartFormat.Fill](http://msdn.microsoft.com/en-us/library/bb242497).
        */
        wxExcelFillFormat GetFill();

        /**
        Returns a GlowFormat object for a specified chart that contains glow formatting properties for the chart element.  Since Excel 2007.

        [MSDN documentation for ChartFormat.Glow](http://msdn.microsoft.com/en-us/library/bb242499).
        */
        wxExcelGlowFormat GetGlow();

        /**
        Returns a LineFormat object that contains line formatting properties for the specified chart element. Since Excel 2007.

        [MSDN documentation for ChartFormat.Line](http://msdn.microsoft.com/en-us/library/bb242502).
        */
        wxExcelLineFormat GetLine();

        /**
        Returns a PictureFormat object for a specified chart that contains pictures. Since Excel 2007.

        [MSDN documentation for ChartFormat.PictureFormat](http://msdn.microsoft.com/en-us/library/bb242504).
        */
        wxExcelPictureFormat GetPictureFormat();

        /**
        Returns a ReflectionFormat object for a specified chart that contains reflection formatting properties for the chart. 

        [MSDN documentation for ChartFormat.Reflection]().
        */
        wxExcelReflectionFormat GetReflection();

        /**
        Returns a ShadowFormat object that contains shadow formatting properties for the chart element.  Since Excel 2007.

        [MSDN documentation for ChartFormat.Shadow](http://msdn.microsoft.com/en-us/library/bb242509).
        */
        wxExcelShadowFormat GetShadow();

        /**
        Returns a SoftEdgeFormat object for a specified chart that contains soft edge formatting properties for the chart.  Since Excel 2007.

        [MSDN documentation for ChartFormat.SoftEdge](http://msdn.microsoft.com/en-us/library/bb242512).
        */
        wxExcelSoftEdgeFormat GetSoftEdge();

        /**
        Returns a TextFrame object that contains the alignment and anchoring properties for the specified chart. 

        [MSDN documentation for ChartFormat.TextFrame]().
        */
        wxExcelTextFrame GetTextFrame();

        /**
        Returns a ThreeDFormat object that contains 3-D–effect formatting properties for the specified chart.  Since Excel 2007.

        [MSDN documentation for ChartFormat.ThreeD](http://msdn.microsoft.com/en-us/library/bb242517).
        */
        wxExcelThreeDFormat GetThreeD();

        /**
        Returns "ChartFormat".
        */
        virtual wxString GetAutoExcelObjectName_() const { return wxS("ChartFormat"); }
    };


} // namespace wxAutoExcel

#endif // #if WXAUTOEXCEL_USE_CHARTS

#endif // _WXAUTOEXCEL_CHARTFORMAT_H
