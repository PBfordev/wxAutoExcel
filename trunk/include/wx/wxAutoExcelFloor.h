/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Modified by:
// RCS-ID:      $Id: $
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
/////////////////////////////////////////////////////////////////////////////


#ifndef _WXAUTOEXCEL_FLOOR_H
#define _WXAUTOEXCEL_FLOOR_H

#include "wx/wxAutoExcel_defs.h"

#if WXAUTOEXCEL_USE_CHARTS

#include "wx/wxAutoExcelObject.h"

namespace wxAutoExcel {

    /**
    Represents Microsoft Excel Floor object.
    */
   class WXDLLIMPEXP_WXAUTOEXCEL wxExcelFloor : public wxExcelObject
    {
    public:

        // ***** METHODS *****

        /**
        Clears the formatting of the object.

        [MSDN documentation for Floor.ClearFormats](http://msdn.microsoft.com/en-us/library/bb211768).
        */
        bool ClearFormats();

        /**
        Pastes a picture from the Clipboard on the floor of the specified chart.

        [MSDN documentation for Floor.Paste](http://msdn.microsoft.com/en-us/library/bb211771).
        */
        bool Paste();

        /**
        Selects the object.

        [MSDN documentation for Floor.Select](http://msdn.microsoft.com/en-us/library/bb237939).
        */
        bool Select();

        // ***** PROPERTIES *****

        /**
        Returns the ChartFormat object.  Since Excel 2007.

        [MSDN documentation for Floor.Format](http://msdn.microsoft.com/en-us/library/bb242529).
        */
        wxExcelChartFormat GetFormat();

        /**
        Returns a String value that represents the name of the object.

        [MSDN documentation for Floor.Name](http://msdn.microsoft.com/en-us/library/bb213176).
        */
        wxString GetName();

        /**
        Returns a Variant value that represents the way pictures are displayed on the walls and faces of a 3-D chart.

        [MSDN documentation for Floor.PictureType](http://msdn.microsoft.com/en-us/library/bb213179).
        */
        long GetPictureType();

        /**
        Sets a Variant value that represents the way pictures are displayed on the walls and faces of a 3-D chart.

        [MSDN documentation for Floor.PictureType](http://msdn.microsoft.com/en-us/library/bb213179).
        */
        void SetPictureType(long pictureType);

        /**
        Returns a Long specifying the thickness of the wall.  Since Excel 2007.

        [MSDN documentation for Floor.Thickness](http://msdn.microsoft.com/en-us/library/bb214614).
        */
        long GetThickness();

        /**
        Sets a Long specifying the thickness of the wall.  Since Excel 2007.

        [MSDN documentation for Floor.Thickness](http://msdn.microsoft.com/en-us/library/bb214614).
        */
        void SetThickness(long thickness);

        /**
        Returns "Floor".
        */
        virtual wxString GetAutoExcelObjectName_() const { return wxS("Floor"); }
    };


} // namespace wxAutoExcel

#endif // #if WXAUTOEXCEL_USE_CHARTS

#endif // _WXAUTOEXCEL_FLOOR_H
