/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Modified by:
// RCS-ID:      $Id: $
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
/////////////////////////////////////////////////////////////////////////////


#ifndef _WXAUTOEXCEL_SPARKCOLOR_H
#define _WXAUTOEXCEL_SPARKCOLOR_H

#include "wx/wxAutoExcel_defs.h"

#if WXAUTOEXCEL_USE_CHARTS

#include "wx/wxAutoExcelObject.h"
#include "wx/wxAutoExcel_enums.h"


namespace wxAutoExcel {


    /**
    Represents Microsoft Excel SparkColor object. Since Excel 2010.
    */
   class WXDLLIMPEXP_WXAUTOEXCEL wxExcelSparkColor: public wxExcelObject
    {
    public:
        // ***** PROPERTIES *****

        /**
        Returns a FormatColor object that you can use to set the color of the markers for points in a sparkline. 

        [MSDN documentation for SparkColor.Color](http://msdn.microsoft.com/en-us/library/office/ff822339(v=office.14).aspx).
        */
        wxExcelFormatColor GetColor();

        /**
        Returns the parent SparklineGroup object for the specified object. 

        [MSDN documentation for SparkColor.Parent](http://msdn.microsoft.com/en-us/library/office/ff839399(v=office.14).aspx).
        */
        wxExcelSparklineGroup GetParent();

        /**
        Returns of sets whether the point is visible. 

        [MSDN documentation for SparkColor.Visible](http://msdn.microsoft.com/en-us/library/office/ff834408(v=office.14).aspx).
        */
        bool GetVisible();

        /**
        Returns of sets whether the point is visible.

        [MSDN documentation for SparkColor.Visible](http://msdn.microsoft.com/en-us/library/office/ff834408(v=office.14).aspx).
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
