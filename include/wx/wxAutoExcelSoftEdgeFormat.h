/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// License:     MIT license
/////////////////////////////////////////////////////////////////////////////


#ifndef _WXAUTOEXCEL_SOFTEDGEFORMAT_H
#define _WXAUTOEXCEL_SOFTEDGEFORMAT_H

#include "wx/wxAutoExcel_defs.h"

#if WXAUTOEXCEL_USE_SHAPES || WXAUTOEXCEL_USE_CHARTS

#include "wx/wxAutoExcel_object.h"
#include "wx/wxAutoExcel_enums.h"

namespace wxAutoExcel {

    /**
    @brief Represents Microsoft Excel SoftEdgeFormat object.
    */
   class WXDLLIMPEXP_WXAUTOEXCEL wxExcelSoftEdgeFormat : public wxExcelObject
    {
    public:

        // ***** PROPERTIES *****

        /**
        Gets or sets the type of the SoftEdgeFormat object.

        [MSDN documentation for SoftEdgeFormat.Type](ttp://msdn.microsoft.com/en-us/library/aa434504).
        */
        MsoSoftEdgeType GetType();

        /**
        Gets or sets the type of the SoftEdgeFormat object.

        [MSDN documentation for SoftEdgeFormat.Type](ttp://msdn.microsoft.com/en-us/library/aa434504).
        */
        void SetType(MsoSoftEdgeType type);

        /**
        Returns "SoftEdgeFormat".
        */
        virtual wxString GetAutoExcelObjectName_() const { return wxS("SoftEdgeFormat"); }
    };


} // namespace wxAutoExcel


#endif // #if WXAUTOEXCEL_USE_SHAPES || WXAUTOEXCEL_USE_CHARTS

#endif //_WXAUTOEXCEL_SOFTEDGEFORMAT_H
