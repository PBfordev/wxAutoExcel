/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// License:     MIT license
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

        [MSDN documentation for Sparkline.ModifyLocation](https://learn.microsoft.com/en-us/office/vba/api/excel.sparkline.modifylocation).
        */
        void ModifyLocation(wxExcelRange location);

        /**
        Modifies the source data for a single sparkline.

        [MSDN documentation for Sparkline.ModifySourceData](https://learn.microsoft.com/en-us/office/vba/api/excel.sparkline.modifysourcedata).
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

        [MSDN documentation for Sparkline.Parent](https://learn.microsoft.com/en-us/office/vba/api/excel.sparkline.parent).
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
