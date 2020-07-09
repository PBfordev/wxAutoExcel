/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Modified by:
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
/////////////////////////////////////////////////////////////////////////////


#ifndef _WXAUTOEXCEL_LINKFORMAT_H
#define _WXAUTOEXCEL_LINKFORMAT_H

#include "wx/wxAutoExcel_defs.h"

#if WXAUTOEXCEL_USE_SHAPES

#include "wx/wxAutoExcel_object.h"

namespace wxAutoExcel {

    /**
    @brief Represents Microsoft Excel LinkFormat object.
    */
   class WXDLLIMPEXP_WXAUTOEXCEL wxExcelLinkFormat  : public wxExcelObject
    {
    public:
        // ***** METHODS *****

        /**
        Updates the link.

        [MSDN documentation for LinkFormat.Update](http://msdn.microsoft.com/en-us/library/bb237997).
        */
        void Update();

        // ***** PROPERTIES *****

        /**
        True if the LinkFormat object is updated automatically when the source changes.

        [MSDN documentation for LinkFormat.AutoUpdate](http://msdn.microsoft.com/en-us/library/bb148596).
        */
        bool GetAutoUpdate();

        /**
        Returns a Boolean value that indicates if the object is locked.

        [MSDN documentation for LinkFormat.Locked](http://msdn.microsoft.com/en-us/library/bb148600).
        */
        bool GetLocked();

        /**
        Sets a Boolean value that indicates if the object is locked.

        [MSDN documentation for LinkFormat.Locked](http://msdn.microsoft.com/en-us/library/bb148600).
        */
        void SetLocked(bool locked);

        /**
        Returns "LinkFormat".
        */
        virtual wxString GetAutoExcelObjectName_() const { return wxS("LinkFormat"); }
    };


} // namespace wxAutoExcel

#endif // #if WXAUTOEXCEL_USE_SHAPES

#endif //_WXAUTOEXCEL_LINKFORMAT_H
