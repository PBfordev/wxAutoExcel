/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Modified by:
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
/////////////////////////////////////////////////////////////////////////////


#ifndef _WXAUTOEXCEL_COLORSCALECRITERIA_H
#define _WXAUTOEXCEL_COLORSCALECRITERIA_H

#include "wx/wxAutoExcel_defs.h"

#if WXAUTOEXCEL_USE_CONDFORMAT

#include "wx/wxAutoExcel_object.h"
#include "wx/wxAutoExcel_enums.h"

namespace wxAutoExcel {

    /**
    @brief Represents a Microsoft Excel ColorScaleCriterion object.
    */
   class WXDLLIMPEXP_WXAUTOEXCEL wxExcelColorScaleCriterion : public wxExcelObject
    {
    public:
        // ***** PROPERTIES *****

        /**
        Returns a FormatColor object which specifies the color assigned to the threshold of a color scale conditional format. Since Excel 2007.

        [MSDN documentation for ColorScaleCriterion.FormatColor](http://msdn.microsoft.com/en-us/library/bb224545.aspx).
        */
        wxExcelFormatColor GetFormatColor();

        /**
        Returns a Long value indicating which threshold the criteria represents. Since Excel 2007.

        [MSDN documentation for ColorScaleCriterion.Index](http://msdn.microsoft.com/en-us/library/bb224547.aspx).
        */
        long GetIndex();

        /**
        Returns one of the constants of the XlConditionValueTypes enumeration, which specifies how the threshold values for a data bar or color scale conditional format are determined.

        [MSDN documentation for ColorScaleCriterion.Type]().
        */
        XlConditionValueTypes GetType();

        /**
        Returns the minimum, midpoint, or maximum threshold value for a color scale conditional format. Read/write Variant. Since Excel 2007.

        [MSDN documentation for ColorScaleCriterion.Value](http://msdn.microsoft.com/en-us/library/bb224549.aspx).
        */
        wxVariant GetValue();

        /**
        Sets the minimum, midpoint, or maximum threshold value for a color scale conditional format. Read/write Variant. Since Excel 2007.

        [MSDN documentation for ColorScaleCriterion.Value](http://msdn.microsoft.com/en-us/library/bb224549.aspx).
        */
        void SetValue(const wxVariant& value);


        /**
        Returns "ColorScaleCriterion".
        */
        virtual wxString GetAutoExcelObjectName_() const { return wxS("ColorScaleCriterion"); }    
    };

    /**
    @brief Represents a Microsoft Excel ColorScaleCriteria collection.
    */
   class WXDLLIMPEXP_WXAUTOEXCEL wxExcelColorScaleCriteria : public wxExcelObject
    {
    public:

        // ***** PROPERTIES *****

        /**
        Returns a Long value which specifies the number of criteria for an color scale conditional format rule. Since Excel 2007.

        [MSDN documentation for ColorScaleCriteria.Count](http://msdn.microsoft.com/en-us/library/bb224540.aspx).
        */
        long GetCount();

        //@{
        /**
        Returns a single ColorScaleCriterion object from the ColorScaleCriteria collection. Since Excel 2007.

        [MSDN documentation for ColorScaleCriteria.Item](http://msdn.microsoft.com/en-us/library/bb224543.aspx).
        */
        wxExcelColorScaleCriterion GetItem(long index);
        wxExcelColorScaleCriterion operator[](long index);
        //@}
                        
        /**
        Returns "ColorScaleCriteria".
        */
        virtual wxString GetAutoExcelObjectName_() const { return wxS("ColorScaleCriteria"); }    
    };


} // namespace wxAutoExcel

#endif // WXAUTOEXCEL_USE_CONDFORMAT

#endif //_WXAUTOEXCEL_COLORSCALECRITERIA_H
