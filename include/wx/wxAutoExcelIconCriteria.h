/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// License:     MIT license
/////////////////////////////////////////////////////////////////////////////


#ifndef _WXAUTOEXCEL_ICONCRITERIA_H
#define _WXAUTOEXCEL_ICONCRITERIA_H

#include "wx/wxAutoExcel_defs.h"

#if WXAUTOEXCEL_USE_CONDFORMAT

#include "wx/wxAutoExcel_object.h"
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

        [MSDN documentation for IconCriterion.Icon](https://learn.microsoft.com/en-us/office/vba/api/excel.iconcriterion.icon).
        */
        XlIcon GetIcon();

        /**
        The icon for a criterion in an icon set conditional formatting rule.. Since Excel 2010.

        [MSDN documentation for IconCriterion.Icon](https://learn.microsoft.com/en-us/office/vba/api/excel.iconcriterion.icon).
        */
        void SetIcon(XlIcon icon);

        /**
        Returns a Long value indicating which threshold the criteria represents. Since Excel 2007.

        [MSDN documentation for IconCriterion.Index](https://learn.microsoft.com/en-us/office/vba/api/excel.iconcriterion.index).
        */
        long GetIndex();

        /**
        Returns one of the constants of the XlFormatConditionOperator enumeration which specifes if the threshold is "greater than" or "greater than or equal" to the threshold value. Since Excel 2007.

        [MSDN documentation for IconCriterion.Operator](https://learn.microsoft.com/en-us/office/vba/api/excel.iconcriterion.operator).
        */
        XlFormatConditionOperator GetOperator();

        /**
        Sets one of the constants of the XlFormatConditionOperator enumeration which specifes if the threshold is "greater than" or "greater than or equal" to the threshold value. Since Excel 2007.

        [MSDN documentation for IconCriterion.Operator](https://learn.microsoft.com/en-us/office/vba/api/excel.iconcriterion.operator).
        */
        void SetOperator(XlFormatConditionOperator conditionOperator);

        /**
        Returns one of the constants of the XlConditionValueTypes enumeration, which specifies how the threshold value for an icon set is determined. Since Excel 2007.

        [MSDN documentation for IconCriterion.Type](https://learn.microsoft.com/en-us/office/vba/api/excel.iconcriterion.type).
        */
        XlConditionValueTypes GetType();

        /**
        Returns the threshold value for an icon in a conditional format. Read/write Variant. Since Excel 2007.

        [MSDN documentation for IconCriterion.Value](https://learn.microsoft.com/en-us/office/vba/api/excel.iconcriterion.value).
        */
        wxVariant GetValue();

        /**
        Sets the threshold value for an icon in a conditional format. Read/write Variant. Since Excel 2007.

        [MSDN documentation for IconCriterion.Value](https://learn.microsoft.com/en-us/office/vba/api/excel.iconcriterion.value).
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

        [MSDN documentation for IconCriteria.Count](https://learn.microsoft.com/en-us/office/vba/api/excel.iconcriteria.count).
        */
        long GetCount();

         //@{
        /**
        Returns a single IconCriterion object from the IconCriteria collection. Since Excel 2007.

        [MSDN documentation for IconCriteria.Item](https://learn.microsoft.com/en-us/office/vba/api/excel.iconcriteria.item).
        */
        wxExcelIconCriterion GetItem(long index);
        wxExcelIconCriterion operator[](long index);
        //@}

        /**
        The iterator over the items, see wxExcelCollectionIterator.
        */
        typedef wxExcelCollectionIterator<wxExcelIconCriteria, wxExcelIconCriterion> iterator;

        //@{
        /**
        Returns the iterator referring to the first item, or past the last item,
        which allows iterating over the items with a range-based for loop,
        see wxExcelCollectionIterator.
        */
        iterator begin() { return iterator(this, 1); }
        iterator end() { return iterator(this, GetCount() + 1); }
        //@}

        /**
        Returns "IconCriteria".
        */
        virtual wxString GetAutoExcelObjectName_() const { return wxS("IconCriteria"); }
    };


} // namespace wxAutoExcel

#endif // WXAUTOEXCEL_USE_CONDFORMAT

#endif //_WXAUTOEXCEL_ICONCRITERIA_H
