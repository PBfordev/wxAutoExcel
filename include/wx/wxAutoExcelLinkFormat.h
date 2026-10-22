/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// License:     MIT license
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

        [MSDN documentation for LinkFormat.Update](https://learn.microsoft.com/en-us/office/vba/api/excel.linkformat.update).
        */
        void Update();

        // ***** PROPERTIES *****

        /**
        True if the LinkFormat object is updated automatically when the source changes.

        [MSDN documentation for LinkFormat.AutoUpdate](https://learn.microsoft.com/en-us/office/vba/api/excel.linkformat.autoupdate).
        */
        bool GetAutoUpdate();

        /**
        Returns a Boolean value that indicates if the object is locked.

        [MSDN documentation for LinkFormat.Locked](https://learn.microsoft.com/en-us/office/vba/api/excel.linkformat.locked).
        */
        bool GetLocked();

        /**
        Sets a Boolean value that indicates if the object is locked.

        [MSDN documentation for LinkFormat.Locked](https://learn.microsoft.com/en-us/office/vba/api/excel.linkformat.locked).
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
