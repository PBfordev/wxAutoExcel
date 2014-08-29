/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Modified by:
// RCS-ID:      $Id: $
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
/////////////////////////////////////////////////////////////////////////////


#ifndef _WXAUTOEXCEL_ICONCRITERIA_H
#define _WXAUTOEXCEL_ICONCRITERIA_H

#include "wx/wxAutoExcel_defs.h"

#if WXAUTOEXCEL_USE_CONDFORMAT

#include "wx/wxAutoExcelObject.h"
#include "wx/wxAutoExcel_enums.h"

namespace wxAutoExcel {

    /**
    @brief Represents a Microsoft Excel IconCriterion object.
    */
   class WXDLLIMPEXP_WXAUTOEXCEL wxExcelIconCriterion : public wxExcelObject
    {
    public:
        // ***** PROPERTIES *****

        
        /**
        The icon for a criterion in an icon set conditional formatting rule.. Since Excel 2010.

        [MSDN documentation for IconCriterion.Icon](http://msdn.microsoft.com/en-us/library/office/ff838436%28v=office.14%29.aspx).
        */
        wxExcelIcon GetIcon();

        /**
        The icon for a criterion in an icon set conditional formatting rule.. Since Excel 2010.

        [MSDN documentation for IconCriterion.Icon](http://msdn.microsoft.com/en-us/library/office/ff838436%28v=office.14%29.aspx).
        */
        void SetIcon(const wxExcelIcon& icon);

        /**
        Returns a Long value indicating which threshold the criteria represents. Since Excel 2007.

        [MSDN documentation for IconCriterion.Index](http://msdn.microsoft.com/en-us/library/bb224557.aspx).
        */
        long GetIndex();

        /**
        Returns one of the constants of the XlFormatConditionOperator enumeration which specifes if the threshold is "greater than" or "greater than or equal" to the threshold value. Since Excel 2007.

        [MSDN documentation for IconCriterion.Operator](http://msdn.microsoft.com/en-us/library/bb224559.aspx).
        */
        XlFormatConditionOperator GetOperator();

        /**
        Sets one of the constants of the XlFormatConditionOperator enumeration which specifes if the threshold is "greater than" or "greater than or equal" to the threshold value. Since Excel 2007.

        [MSDN documentation for IconCriterion.Operator](http://msdn.microsoft.com/en-us/library/bb224559.aspx).
        */
        void SetOperator(XlFormatConditionOperator conditionOperator);

        /**
        Returns one of the constants of the XlConditionValueTypes enumeration, which specifies how the threshold value for an icon set is determined. Since Excel 2007.

        [MSDN documentation for IconCriterion.Type](http://msdn.microsoft.com/en-us/library/bb224560.aspx).
        */
        XlConditionValueTypes GetType();

        /**
        Returns the threshold value for an icon in a conditional format. Read/write Variant. Since Excel 2007.

        [MSDN documentation for IconCriterion.Value](http://msdn.microsoft.com/en-us/library/bb224563.aspx).
        */
        wxVariant GetValue();

        /**
        Sets the threshold value for an icon in a conditional format. Read/write Variant. Since Excel 2007.

        [MSDN documentation for IconCriterion.Value](http://msdn.microsoft.com/en-us/library/bb224563.aspx).
        */
        void SetValue(const wxVariant& value);


        /**
        Returns "IconCriterion".
        */
        virtual wxString GetAutoExcelObjectName_() const { return wxS("IconCriterion"); }    
    };

    /**
    @brief Represents a Microsoft Excel IconCriteria collection.
    */
   class WXDLLIMPEXP_WXAUTOEXCEL wxExcelIconCriteria : public wxExcelObject
    {
    public:
        // ***** PROPERTIES *****

        /**
        Returns a Long value which specifies the number of criteria for an icon set conditional format rule. Since Excel 2007.

        [MSDN documentation for IconCriteria.Count](http://msdn.microsoft.com/en-us/library/bb224553.aspx).
        */
        long GetCount();

         //@{
        /**
        Returns a single IconCriterion object from the IconCriteria collection. Since Excel 2007.

        [MSDN documentation for IconCriteria.Item](http://msdn.microsoft.com/en-us/library/bb224556.aspx).
        */
        wxExcelIconCriterion GetItem(long index);
        wxExcelIconCriterion operator[](long index);
        //@}        
                        
        /**
        Returns "IconCriteria".
        */
        virtual wxString GetAutoExcelObjectName_() const { return wxS("IconCriteria"); }    
    };


} // namespace wxAutoExcel

#endif // WXAUTOEXCEL_USE_CONDFORMAT

#endif //_WXAUTOEXCEL_ICONCRITERIA_H
