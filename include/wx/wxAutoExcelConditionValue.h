/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Modified by:
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
/////////////////////////////////////////////////////////////////////////////


#ifndef _WXAUTOEXCEL_CONDITIONVALUE_H
#define _WXAUTOEXCEL_CONDITIONVALUE_H

#include "wx/wxAutoExcel_defs.h"

#if WXAUTOEXCEL_USE_CONDFORMAT

#include "wx/wxAutoExcelObject.h"
#include "wx/wxAutoExcel_enums.h"

namespace wxAutoExcel {
    /**
    @brief Represents a Microsoft Excel ConditionValue object.
    */
   class WXDLLIMPEXP_WXAUTOEXCEL wxExcelConditionValue : public wxExcelObject
    {
    public:
        // ***** METHODS *****

        /**
        Modifies how the longest bar or shortest bar is evaluated for a data bar conditional formatting rule.

        [MSDN documentation for ConditionValue.Modify](http://msdn.microsoft.com/en-us/library/bb178613.aspx).
        */
        void Modify(XlConditionValueTypes newType, const wxVariant& newValue = wxNullVariant);

        // ***** PROPERTIES *****

        /**
        Returns one of the constants of the XlConditionValueTypes enumeration, which specifies how the threshold values for a data bar, color scale, or icon set conditional format are determined. Since Excel 2007.

        [MSDN documentation for ConditionValue.Type](http://msdn.microsoft.com/en-us/library/bb224384.aspx).
        */
        XlConditionValueTypes GetType();

        /**
        Returns the shortest bar or longest bar threshold value for a data bar conditional format. Read/write Variant. Since Excel 2007.

        [MSDN documentation for ConditionValue.Value](http://msdn.microsoft.com/en-us/library/bb224388.aspx).
        */
        wxVariant GetValue();

        /**
        Sets the shortest bar or longest bar threshold value for a data bar conditional format. Read/write Variant. Since Excel 2007.

        [MSDN documentation for ConditionValue.Value](http://msdn.microsoft.com/en-us/library/bb224388.aspx).
        */
        void SetValue(const wxVariant& value);
        
        
        /**
        Returns "ConditionValue".
        */
        virtual wxString GetAutoExcelObjectName_() const { return wxS("ConditionValue"); }    
    };


} // namespace wxAutoExcel

#endif // WXAUTOEXCEL_USE_CONDFORMAT

#endif //_WXAUTOEXCEL_CONDITIONVALUE_H
