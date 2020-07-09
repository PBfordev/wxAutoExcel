/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Modified by:
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
/////////////////////////////////////////////////////////////////////////////


#ifndef _WXAUTOEXCEL_WALLS_H
#define _WXAUTOEXCEL_WALLS_H

#include "wx/wxAutoExcel_defs.h"

#if WXAUTOEXCEL_USE_CHARTS

#include "wx/wxAutoExcel_object.h"

namespace wxAutoExcel {


    /**
    Represents Microsoft Excel Walls object.
    */
   class WXDLLIMPEXP_WXAUTOEXCEL wxExcelWalls : public wxExcelObject
    {
    public:
        // ***** METHODS *****

        /**
        Clears the formatting of the object.

        [MSDN documentation for Walls.ClearFormats](http://msdn.microsoft.com/en-us/library/bb179105).
        */
        bool ClearFormats();

        /**
        Pastes a picture from the Clipboard on the walls of the specified chart.

        [MSDN documentation for Walls.Paste](http://msdn.microsoft.com/en-us/library/bb179109).
        */
        bool Paste();

        /**
        Selects the object.

        [MSDN documentation for Walls.Select](http://msdn.microsoft.com/en-us/library/bb214100).
        */
        bool Select();

        // ***** PROPERTIES *****


        /**
        Returns the ChartFormat object.  Since Excel 2007.

        [MSDN documentation for Walls.Format](http://msdn.microsoft.com/en-us/library/bb242543).
        */
        wxExcelChartFormat GetFormat();

        /**
        Returns a String value that represents the name of the object.

        [MSDN documentation for Walls.Name](http://msdn.microsoft.com/en-us/library/bb214388).
        */
        wxString GetName();

        /**
        Returns a Variant value that represents the way pictures are displayed on the walls and faces of a 3-D chart.

        [MSDN documentation for Walls.PictureType](http://msdn.microsoft.com/en-us/library/bb214394).
        */
        long GetPictureType();

        /**
        Sets a Variant value that represents the way pictures are displayed on the walls and faces of a 3-D chart.

        [MSDN documentation for Walls.PictureType](http://msdn.microsoft.com/en-us/library/bb214394).
        */
        void SetPictureType(long pictureType);

        /**
        Returns the unit for each picture on the chart if the PictureType property is set to xlStackScale (if not, this property is ignored).

        [MSDN documentation for Walls.PictureUnit](http://msdn.microsoft.com/en-us/library/bb214400).
        */
        long GetPictureUnit();

        /**
        Sets the unit for each picture on the chart if the PictureType property is set to xlStackScale (if not, this property is ignored).

        [MSDN documentation for Walls.PictureUnit](http://msdn.microsoft.com/en-us/library/bb214400).
        */
        void SetPictureUnit(long pictureUnit);

        /**
        Returns a Long specifying the thickness of the wall.  Since Excel 2007.

        [MSDN documentation for Walls.Thickness](http://msdn.microsoft.com/en-us/library/bb215174).
        */
        long GetThickness();

        /**
        Sets a Long specifying the thickness of the wall.  Since Excel 2007.

        [MSDN documentation for Walls.Thickness](http://msdn.microsoft.com/en-us/library/bb215174).
        */
        void SetThickness(long thickness);

        /**
        Returns "Walls".
        */
        virtual wxString GetAutoExcelObjectName_() const { return wxS("Walls"); }
    };


} // namespace wxAutoExcel

#endif // #if WXAUTOEXCEL_USE_CHARTS

#endif // _WXAUTOEXCEL_WALLS_H
