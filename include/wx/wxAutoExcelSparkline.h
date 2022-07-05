/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Modified by:
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
/////////////////////////////////////////////////////////////////////////////


#ifndef _WXAUTOEXCEL_SPARKLINE_H
#define _WXAUTOEXCEL_SPARKLINE_H

#include "wx/wxAutoExcel_defs.h"

#if WXAUTOEXCEL_USE_CHARTS

#include "wx/wxAutoExcel_object.h"
#include "wx/wxAutoExcel_enums.h"


namespace wxAutoExcel {


    /**
    @brief Represents Microsoft Excel Sparkline object. Since Excel 2010.
    */
   class WXDLLIMPEXP_WXAUTOEXCEL wxExcelSparkline: public wxExcelObject
    {
    public:
        // ***** METHODS *****

        /**
        Modifies the location of a single sparkline.

        [MSDN documentation for Sparkline.ModifyLocation](http://msdn.microsoft.com/en-us/library/office/ff196726(v=office.14).aspx).
        */
        void ModifyLocation(wxExcelRange location);

        /**
        Modifies the source data for a single sparkline.

        [MSDN documentation for Sparkline.ModifySourceData](http://msdn.microsoft.com/en-us/library/office/ff839062(v=office.14).aspx).
        */
        void ModifySourceData(const wxString& formula);

        // ***** PROPERTIES *****

        /**
        Returns the location of a single sparkline. Read/write

        [MSDN documentation for Sparkline.Location]().
        */
        wxExcelRange GetLocation();

        /**
        Sets the location of a single sparkline. Read/write

        [MSDN documentation for Sparkline.Location]().
        */
        void SetLocation(const wxExcelRange& location);

        /**
        Returns the parent SparklineGroup object for the specified object. Since Excel 2007.

        [MSDN documentation for Sparkline.Parent](http://msdn.microsoft.com/en-us/library/office/ff837439(v=office.14).aspx).
        */
        wxExcelSparklineGroup GetParent();

        /**
        Returns the range the contains the source data for a single sparkline. Read/write

        [MSDN documentation for Sparkline.SourceData]().
        */
        wxString GetSourceData();

        /**
        Sets the range the contains the source data for a single sparkline. Read/write

        [MSDN documentation for Sparkline.SourceData]().
        */
        void SetSourceData(const wxString& sourceData);


        /**
        Returns "Sparkline".
        */
        virtual wxString GetAutoExcelObjectName_() const { return wxS("Sparkline"); }
    };


} // namespace wxAutoExcel

#endif // #if WXAUTOEXCEL_USE_CHARTS

#endif // _WXAUTOEXCEL_SPARKLINE_H
