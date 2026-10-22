/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// License:     MIT license
/////////////////////////////////////////////////////////////////////////////


#ifndef _WXAUTOEXCEL_ICONSETS_H
#define _WXAUTOEXCEL_ICONSETS_H

#include "wx/wxAutoExcel_defs.h"

#if WXAUTOEXCEL_USE_CONDFORMAT

#include "wx/wxAutoExcel_object.h"
#include "wx/wxAutoExcel_enums.h"

namespace wxAutoExcel {

    /**
    @brief Represents a Microsoft Excel IconSet object.
    */
   class WXDLLIMPEXP_WXAUTOEXCEL wxExcelIconSet : public wxExcelObject
    {
    public:
        // ***** PROPERTIES *****

        /**
        Returns a Long value which specifies the number of icons in the icon set. Since Excel 2007.

        [MSDN documentation for IconSet.Count](https://learn.microsoft.com/en-us/office/vba/api/excel.iconset.count).
        */
        long GetCount();

        /**
        Returns one of the constants of the XlIconSet enumeration which specifies the name of the icon set used in an icon set conditional formatting rule. Since Excel 2007.

        [MSDN documentation for IconSet.ID](https://learn.microsoft.com/en-us/office/vba/api/excel.iconset.id).
        */
        XlIconSet GetID();

        /**
        Returns an Icon object which represents a single icon from an icon set. Since Excel 2007.

        [MSDN documentation for IconSet.Item](https://learn.microsoft.com/en-us/office/vba/api/excel.iconset.item).
        */
        wxExcelIcon GetItem(long index);

        /**
        Returns "IconSet".
        */
        virtual wxString GetAutoExcelObjectName_() const { return wxS("IconSet"); }
    };

    /**
    @brief Represents a Microsoft Excel IconSets collection.
    */
   class WXDLLIMPEXP_WXAUTOEXCEL wxExcelIconSets : public wxExcelObject
    {
    public:
        // ***** PROPERTIES *****

        /**
        Returns a Long value which specifies the number of icon sets available in the workbook. Since Excel 2007.

        [MSDN documentation for IconSets.Count](https://learn.microsoft.com/en-us/office/vba/api/excel.iconsets.count).
        */
        long GetCount();

        //@{
        /**
        Returns a single IconSet object from the IconSets collection. Since Excel 2007.

        [MSDN documentation for IconSets.Item](https://learn.microsoft.com/en-us/office/vba/api/excel.iconsets.item).
        */
        wxExcelIconSet GetItem(long index);
        wxExcelIconSet operator[](long index);
        //@}


        /**
        Returns "IconSets".
        */
        virtual wxString GetAutoExcelObjectName_() const { return wxS("IconSets"); }
    };


} // namespace wxAutoExcel

#endif // WXAUTOEXCEL_USE_CONDFORMAT

#endif //_WXAUTOEXCEL_ICONSETS_H
