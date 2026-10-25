/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// License:     MIT license
/////////////////////////////////////////////////////////////////////////////


#ifndef _WXAUTOEXCEL_FLOOR_H
#define _WXAUTOEXCEL_FLOOR_H

#include "wx/wxAutoExcel_defs.h"

#if WXAUTOEXCEL_USE_CHARTS

#include "wx/wxAutoExcel_object.h"

namespace wxAutoExcel {

    /**
    @brief Represents Microsoft Excel Floor object.
    */
   class WXDLLIMPEXP_WXAUTOEXCEL wxExcelFloor : public wxExcelObject
    {
    public:

        // ***** METHODS *****

        /**
        Clears the formatting of the object.

        [MSDN documentation for Floor.ClearFormats](https://learn.microsoft.com/en-us/office/vba/api/excel.floor.clearformats).
        */
        bool ClearFormats();

        /**
        Pastes a picture from the Clipboard on the floor of the specified chart.

        [MSDN documentation for Floor.Paste](https://learn.microsoft.com/en-us/office/vba/api/excel.floor.paste).
        */
        bool Paste();

        /**
        Selects the object.

        [MSDN documentation for Floor.Select](https://learn.microsoft.com/en-us/office/vba/api/excel.floor.select).
        */
        bool Select();

        // ***** PROPERTIES *****

        /**
        Returns a ChartFillFormat object that contains fill formatting properties for the object.

        [MSDN documentation for Floor.Fill](https://learn.microsoft.com/en-us/office/vba/api/excel.floor.fill).
        */
        wxExcelChartFillFormat GetFill();

        /**
        Returns the ChartFormat object.  Since Excel 2007.

        [MSDN documentation for Floor.Format](https://learn.microsoft.com/en-us/office/vba/api/excel.floor.format).
        */
        wxExcelChartFormat GetFormat();

        /**
        Returns a String value that represents the name of the object.

        [MSDN documentation for Floor.Name](https://learn.microsoft.com/en-us/office/vba/api/excel.floor.name).
        */
        wxString GetName();

        /**
        Returns a Variant value that represents the way pictures are displayed on the walls and faces of a 3-D chart.

        [MSDN documentation for Floor.PictureType](https://learn.microsoft.com/en-us/office/vba/api/excel.floor.picturetype).
        */
        long GetPictureType();

        /**
        Sets a Variant value that represents the way pictures are displayed on the walls and faces of a 3-D chart.

        [MSDN documentation for Floor.PictureType](https://learn.microsoft.com/en-us/office/vba/api/excel.floor.picturetype).
        */
        void SetPictureType(long pictureType);

        /**
        Returns a Long specifying the thickness of the wall.  Since Excel 2007.

        [MSDN documentation for Floor.Thickness](https://learn.microsoft.com/en-us/office/vba/api/excel.floor.thickness).
        */
        long GetThickness();

        /**
        Sets a Long specifying the thickness of the wall.  Since Excel 2007.

        [MSDN documentation for Floor.Thickness](https://learn.microsoft.com/en-us/office/vba/api/excel.floor.thickness).
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
