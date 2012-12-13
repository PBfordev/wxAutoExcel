/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Modified by:
// RCS-ID:      $Id: $
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
/////////////////////////////////////////////////////////////////////////////


#ifndef _WXAUTOEXCEL_REFLECTIONFORMAT_H
#define _WXAUTOEXCEL_REFLECTIONFORMAT_H

#include "wx/wxAutoExcel_defs.h"

#if WXAUTOEXCEL_USE_SHAPES || WXAUTOEXCEL_USE_CHARTS

#include "wx/wxAutoExcelObject.h"
#include "wx/wxAutoExcel_enums.h"

namespace wxAutoExcel {

    /**
    @brief Represents Microsoft Excel ReflectionFormat object.
    */
   class WXDLLIMPEXP_WXAUTOEXCEL wxExcelReflectionFormat : public wxExcelObject
    {
    public:        
        // ***** PROPERTIES *****

        /**
        Gets or sets the type of the ReflectionFormat object.

        [MSDN documentation for ReflectionFormat.Type](http://msdn.microsoft.comhttp://msdn.microsoft.com/en-us/library/aa434501).
        */
        MsoReflectionType GetType();

        /**
        Gets or sets the type of the ReflectionFormat object.

        [MSDN documentation for ReflectionFormat.Type](http://msdn.microsoft.comhttp://msdn.microsoft.com/en-us/library/aa434501).
        */
        void SetType(MsoReflectionType type);

        /**
        Returns "ReflectionFormat".
        */
        virtual wxString GetAutoExcelObjectName_() const { return wxS("ReflectionFormat"); }
    };


} // namespace wxAutoExcel

#endif // #if WXAUTOEXCEL_USE_SHAPES || WXAUTOEXCEL_USE_CHARTS

#endif //_WXAUTOEXCEL_REFLECTIONFORMAT_H
