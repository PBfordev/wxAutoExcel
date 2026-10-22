/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// License:     MIT license
/////////////////////////////////////////////////////////////////////////////


#ifndef _WXAUTOEXCEL_CHARTFORMAT_H
#define _WXAUTOEXCEL_CHARTFORMAT_H

#include "wx/wxAutoExcel_defs.h"

#if WXAUTOEXCEL_USE_CHARTS

#include "wx/wxAutoExcel_object.h"
#include "wx/wxAutoExcel_enums.h"

namespace wxAutoExcel {

    /**
    @brief Represents Microsoft Excel ChartFormat object.
    */
   class WXDLLIMPEXP_WXAUTOEXCEL wxExcelChartFormat: public wxExcelObject
    {
    public:
        // ***** PROPERTIES *****

        /**
        Returns an Adjustments Object.

        [Excel VBA documentation for ChartFormat.Adjustments](https://learn.microsoft.com/en-us/office/vba/api/excel.chartformat.adjustments)
        */
        wxExcelAdjustments GetAdjustments();

        /**
        Returns the type of the specified shape

        [Excel VBA documentation for ChartFormat.AutoShapeType](https://learn.microsoft.com/en-us/office/vba/api/excel.chartformat.autoshapetype)
        */
        MsoAutoShapeType GetAutoShapeType();


        /**
        Returns a FillFormat object for the parent chart element that contains fill formatting properties for the chart element.  Since Excel 2007.

        [MSDN documentation for ChartFormat.Fill](https://learn.microsoft.com/en-us/office/vba/api/excel.chartformat.fill).
        */
        wxExcelFillFormat GetFill();

        /**
        Returns a GlowFormat object for a specified chart that contains glow formatting properties for the chart element.  Since Excel 2007.

        [MSDN documentation for ChartFormat.Glow](https://learn.microsoft.com/en-us/office/vba/api/excel.chartformat.glow).
        */
        wxExcelGlowFormat GetGlow();

        /**
        Returns a LineFormat object that contains line formatting properties for the specified chart element. Since Excel 2007.

        [MSDN documentation for ChartFormat.Line](https://learn.microsoft.com/en-us/office/vba/api/excel.chartformat.line).
        */
        wxExcelLineFormat GetLine();

        /**
        Returns a PictureFormat object for a specified chart that contains pictures. Since Excel 2007.

        [MSDN documentation for ChartFormat.PictureFormat](https://learn.microsoft.com/en-us/office/vba/api/excel.chartformat.pictureformat).
        */
        wxExcelPictureFormat GetPictureFormat();

        /**
        Returns a ShadowFormat object that contains shadow formatting properties for the chart element.  Since Excel 2007.

        [MSDN documentation for ChartFormat.Shadow](https://learn.microsoft.com/en-us/office/vba/api/excel.chartformat.shadow).
        */
        wxExcelShadowFormat GetShadow();

        /**
        Returns a SoftEdgeFormat object for a specified chart that contains soft edge formatting properties for the chart.  Since Excel 2007.

        [MSDN documentation for ChartFormat.SoftEdge](https://learn.microsoft.com/en-us/office/vba/api/excel.chartformat.softedge).
        */
        wxExcelSoftEdgeFormat GetSoftEdge();

        /**
        Returns a TextFrame2 object that contains the text formatting for the specified chart.

        [MSDN documentation for ChartFormat.TextFrame2](https://learn.microsoft.com/en-us/office/vba/api/excel.chartformat.textframe2).
        */
        wxExcelTextFrame2 GetTextFrame2();

        /**
        Returns a ThreeDFormat object that contains 3-D–effect formatting properties for the specified chart.  Since Excel 2007.

        [MSDN documentation for ChartFormat.ThreeD](https://learn.microsoft.com/en-us/office/vba/api/excel.chartformat.threed).
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
