/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// License:     MIT license
/////////////////////////////////////////////////////////////////////////////


#ifndef _WXAUTOEXCEL_OLEFORMAT_H
#define _WXAUTOEXCEL_OLEFORMAT_H

#include "wx/wxAutoExcel_defs.h"

#if WXAUTOEXCEL_USE_SHAPES

#include "wx/wxAutoExcel_object.h"
#include "wx/wxAutoExcel_enums.h"

namespace wxAutoExcel {

    /**
    @brief Represents Microsoft Excel OLEFormat object.
    */
   class WXDLLIMPEXP_WXAUTOEXCEL wxExcelOLEFormat : public wxExcelObject
    {
    public:

        // ***** METHODS *****

        /**
        Makes the current chart the active chart.

        [MSDN documentation for OLEFormat.Activate](https://learn.microsoft.com/en-us/office/vba/api/excel.oleformat.activate).
        */
        void Activate();

        /**
        Sends a verb to the server of the specified OLE object.

        [MSDN documentation for OLEFormat.Verb](https://learn.microsoft.com/en-us/office/vba/api/excel.oleformat.verb).
        */
        void Verb(XlOLEVerb* verb = NULL);

        // ***** PROPERTIES *****

        /**
        Returns the programmatic identifiers for the object.

        [MSDN documentation for OLEFormat.progID](https://learn.microsoft.com/en-us/office/vba/api/excel.oleformat.progid).
        */
        wxString GetprogID();


        /**
        Returns "OLEFormat".
        */
        virtual wxString GetAutoExcelObjectName_() const { return wxS("OLEFormat"); }
    };


} // namespace wxAutoExcel

#endif // #if WXAUTOEXCEL_USE_SHAPES || WXAUTOEXCEL_USE_CHARTS

#endif //_WXAUTOEXCEL_OLEFORMAT_H
