/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Modified by:
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
/////////////////////////////////////////////////////////////////////////////


#ifndef _WXAUTOEXCEL_ICONSETS_H
#define _WXAUTOEXCEL_ICONSETS_H

#include "wx/wxAutoExcel_defs.h"

#if WXAUTOEXCEL_USE_CONDFORMAT

#include "wx/wxAutoExcelObject.h"
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

        [MSDN documentation for IconSet.Count](http://msdn.microsoft.com/en-us/library/bb240130.aspx).
        */
        long GetCount();

        /**
        Returns one of the constants of the XlIconSetE enumeration which specifies the name of the icon set used in an icon set conditional formatting rule. Since Excel 2007.

        [MSDN documentation for IconSet.ID](http://msdn.microsoft.com/en-us/library/bb240132.aspx).
        */
        XlIconSetE GetID();

        /**
        Returns an Icon object which represents a single icon from an icon set. Since Excel 2007.

        [MSDN documentation for IconSet.Item](http://msdn.microsoft.com/en-us/library/bb240135.aspx).
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
        // ***** METHODS *****

        /**
        Deletes the specified conditional formatting rule object.

        [MSDN documentation for IconSets.Delete]().
        */
        void Delete();

        /**
        Sets the cell range to which this formatting rule applies.

        [MSDN documentation for IconSets.ModifyAppliesToRange]().
        */
        void ModifyAppliesToRange(wxExcelRange range);

        /**
        Sets the priority value for this conditional formatting rule to "1" so that it will be evaluated before all other rules on the worksheet.

        [MSDN documentation for IconSets.SetFirstPriority]().
        */
        void SetFirstPriority();

        /**
        Sets the evaluation order for this conditional formatting rule so it is evaluated after all other rules on the worksheet.

        [MSDN documentation for IconSets.SetLastPriority]().
        */
        void SetLastPriority();

        // ***** PROPERTIES *****

        /**
        Returns a Long value which specifies the number of icon sets available in the workbook. Since Excel 2007.

        [MSDN documentation for IconSets.Count](http://msdn.microsoft.com/en-us/library/bb224759.aspx).
        */
        long GetCount();

        //@{
        /**
        Returns a single IconSet object from the IconSets collection. Since Excel 2007.

        [MSDN documentation for IconSets.Item](http://msdn.microsoft.com/en-us/library/bb224761.aspx).
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
