/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// License:     MIT license
/////////////////////////////////////////////////////////////////////////////


#ifndef _WXAUTOEXCEL_WALLS_H
#define _WXAUTOEXCEL_WALLS_H

#include "wx/wxAutoExcel_defs.h"

#if WXAUTOEXCEL_USE_CHARTS

#include "wx/wxAutoExcel_object.h"

namespace wxAutoExcel {


    /**
    @brief Represents Microsoft Excel Walls object.
    */
   class WXDLLIMPEXP_WXAUTOEXCEL wxExcelWalls : public wxExcelObject
    {
    public:
        // ***** METHODS *****

        /**
        Clears the formatting of the object.

        [MSDN documentation for Walls.ClearFormats](https://learn.microsoft.com/en-us/office/vba/api/excel.walls.clearformats).
        */
        bool ClearFormats();

        /**
        Pastes a picture from the Clipboard on the walls of the specified chart.

        [MSDN documentation for Walls.Paste](https://learn.microsoft.com/en-us/office/vba/api/excel.walls.paste).
        */
        bool Paste();

        /**
        Selects the object.

        [MSDN documentation for Walls.Select](https://learn.microsoft.com/en-us/office/vba/api/excel.walls.select).
        */
        bool Select();

        // ***** PROPERTIES *****


        /**
        Returns a ChartFillFormat object that contains fill formatting properties for the object.

        [MSDN documentation for Walls.Fill](https://learn.microsoft.com/en-us/office/vba/api/excel.walls.fill).
        */
        wxExcelChartFillFormat GetFill();

        /**
        Returns the ChartFormat object.  Since Excel 2007.

        [MSDN documentation for Walls.Format](https://learn.microsoft.com/en-us/office/vba/api/excel.walls.format).
        */
        wxExcelChartFormat GetFormat();

        /**
        Returns a String value that represents the name of the object.

        [MSDN documentation for Walls.Name](https://learn.microsoft.com/en-us/office/vba/api/excel.walls.name).
        */
        wxString GetName();

        /**
        Returns a Variant value that represents the way pictures are displayed on the walls and faces of a 3-D chart.

        [MSDN documentation for Walls.PictureType](https://learn.microsoft.com/en-us/office/vba/api/excel.walls.picturetype).
        */
        long GetPictureType();

        /**
        Sets a Variant value that represents the way pictures are displayed on the walls and faces of a 3-D chart.

        [MSDN documentation for Walls.PictureType](https://learn.microsoft.com/en-us/office/vba/api/excel.walls.picturetype).
        */
        void SetPictureType(long pictureType);

        /**
        Returns the unit for each picture on the chart if the PictureType property is set to xlStackScale (if not, this property is ignored).

        [MSDN documentation for Walls.PictureUnit](https://learn.microsoft.com/en-us/office/vba/api/excel.walls.pictureunit).
        */
        long GetPictureUnit();

        /**
        Sets the unit for each picture on the chart if the PictureType property is set to xlStackScale (if not, this property is ignored).

        [MSDN documentation for Walls.PictureUnit](https://learn.microsoft.com/en-us/office/vba/api/excel.walls.pictureunit).
        */
        void SetPictureUnit(long pictureUnit);

        /**
        Returns a Long specifying the thickness of the wall.  Since Excel 2007.

        [MSDN documentation for Walls.Thickness](https://learn.microsoft.com/en-us/office/vba/api/excel.walls.thickness).
        */
        long GetThickness();

        /**
        Sets a Long specifying the thickness of the wall.  Since Excel 2007.

        [MSDN documentation for Walls.Thickness](https://learn.microsoft.com/en-us/office/vba/api/excel.walls.thickness).
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
