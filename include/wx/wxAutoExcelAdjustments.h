/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Modified by:
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
/////////////////////////////////////////////////////////////////////////////

#ifndef _WXAUTOEXCEL_ADJUSTMENTS_H
#define _WXAUTOEXCEL_ADJUSTMENTS_H

#include "wx/wxAutoExcel_defs.h"

#if WXAUTOEXCEL_USE_SHAPES

#include "wx/wxAutoExcel_object.h"

namespace wxAutoExcel {

    /**
    @brief Represents Microsoft Excel Adjustments object.
    */
   class WXDLLIMPEXP_WXAUTOEXCEL wxExcelAdjustments : public wxExcelObject
    {
    public:

        // ***** PROPERTIES *****

        /**
        Returns the number of objects in the collection.

        [MSDN documentation for Adjustments.Count](http://msdn.microsoft.com/en-us/library/bb212468).
        */
        long GetCount();

        //@{
        /**
        Returns an item from the collection.

        [MSDN documentation for Adjustments.Item](http://msdn.microsoft.com/en-us/library/bb212473).
        */
        double GetItem(long index);
        double operator[](long index);
        //@}

        /**
        Returns "Adjustments".
        */
        virtual wxString GetAutoExcelObjectName_() const { return wxS("Adjustments"); }
    };


} // namespace wxAutoExcel

#endif // #if WXAUTOEXCEL_USE_SHAPES

#endif //_WXAUTOEXCEL_ADJUSTMENTS_H
