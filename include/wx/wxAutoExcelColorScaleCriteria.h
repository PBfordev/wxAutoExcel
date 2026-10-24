/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// License:     MIT license
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

        [MSDN documentation for ColorScaleCriterion.FormatColor](https://learn.microsoft.com/en-us/office/vba/api/excel.colorscalecriterion.formatcolor).
        */
        wxExcelFormatColor GetFormatColor();

        /**
        Returns a Long value indicating which threshold the criteria represents. Since Excel 2007.

        [MSDN documentation for ColorScaleCriterion.Index](https://learn.microsoft.com/en-us/office/vba/api/excel.colorscalecriterion.index).
        */
        long GetIndex();

        /**
        Returns one of the constants of the XlConditionValueTypes enumeration, which specifies how the threshold values for a data bar or color scale conditional format are determined.

        [MSDN documentation for ColorScaleCriterion.Type](https://learn.microsoft.com/en-us/office/vba/api/excel.colorscalecriterion.type).
        */
        XlConditionValueTypes GetType();

        /**
        Sets one of the constants of the XlConditionValueTypes enumeration, which specifies how the threshold values for a data bar or color scale conditional format are determined.

        [MSDN documentation for ColorScaleCriterion.Type](https://learn.microsoft.com/en-us/office/vba/api/excel.colorscalecriterion.type).
        */
        void SetType(XlConditionValueTypes type);

        /**
        Returns the minimum, midpoint, or maximum threshold value for a color scale conditional format. Read/write Variant. Since Excel 2007.

        [MSDN documentation for ColorScaleCriterion.Value](https://learn.microsoft.com/en-us/office/vba/api/excel.colorscalecriterion.value).
        */
        wxVariant GetValue();

        /**
        Sets the minimum, midpoint, or maximum threshold value for a color scale conditional format. Read/write Variant. Since Excel 2007.

        [MSDN documentation for ColorScaleCriterion.Value](https://learn.microsoft.com/en-us/office/vba/api/excel.colorscalecriterion.value).
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

        [MSDN documentation for ColorScaleCriteria.Count](https://learn.microsoft.com/en-us/office/vba/api/excel.colorscalecriteria.count).
        */
        long GetCount();

        //@{
        /**
        Returns a single ColorScaleCriterion object from the ColorScaleCriteria collection. Since Excel 2007.

        [MSDN documentation for ColorScaleCriteria.Item](https://learn.microsoft.com/en-us/office/vba/api/excel.colorscalecriteria.item).
        */
        wxExcelColorScaleCriterion GetItem(long index);
        wxExcelColorScaleCriterion operator[](long index);
        //@}

        /**
        The iterator over the items, see wxExcelCollectionIterator.
        */
        typedef wxExcelCollectionIterator<wxExcelColorScaleCriteria, wxExcelColorScaleCriterion> iterator;

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
        Returns "ColorScaleCriteria".
        */
        virtual wxString GetAutoExcelObjectName_() const { return wxS("ColorScaleCriteria"); }
    };


} // namespace wxAutoExcel

#endif // WXAUTOEXCEL_USE_CONDFORMAT

#endif //_WXAUTOEXCEL_COLORSCALECRITERIA_H
