/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// License:     MIT license
/////////////////////////////////////////////////////////////////////////////


#ifndef _WXAUTOEXCEL_GRAPHIC_H
#define _WXAUTOEXCEL_GRAPHIC_H

#include "wx/wxAutoExcel_defs.h"

#include "wx/wxAutoExcel_object.h"
#include "wx/wxAutoExcel_enums.h"

namespace wxAutoExcel {


    /**
    @brief Represents Microsoft Excel Graphic object.
    */
   class WXDLLIMPEXP_WXAUTOEXCEL wxExcelGraphic: public wxExcelObject
    {
    public:
        // ***** PROPERTIES *****
        /**
        Returns the brightness of the specified picture or OLE object. The value for this property must be a number from 0.0 (dimmest) to 1.0 (brightest). Read/write Single.

        [MSDN documentation for Graphic.Brightness](https://learn.microsoft.com/en-us/office/vba/api/excel.graphic.brightness).
        */
        double GetBrightness();

        /**
        Sets the brightness of the specified picture or OLE object. The value for this property must be a number from 0.0 (dimmest) to 1.0 (brightest). Read/write Single.

        [MSDN documentation for Graphic.Brightness](https://learn.microsoft.com/en-us/office/vba/api/excel.graphic.brightness).
        */
        void SetBrightness(double brightness);

        /**
        Returns the type of color transformation applied to the specified picture or OLE object. Read/write MsoPictureColorType.

        [MSDN documentation for Graphic.ColorType](https://learn.microsoft.com/en-us/office/vba/api/excel.graphic.colortype).
        */
        MsoPictureColorType GetColorType();

        /**
        Sets the type of color transformation applied to the specified picture or OLE object. Read/write MsoPictureColorType.

        [MSDN documentation for Graphic.ColorType](https://learn.microsoft.com/en-us/office/vba/api/excel.graphic.colortype).
        */
        void SetColorType(MsoPictureColorType colorType);

        /**
        Returns the contrast for the specified picture or OLE object. The value for this property must be a number from 0.0 (the least contrast) to 1.0 (the greatest contrast). Read/write Single.

        [MSDN documentation for Graphic.Contrast](https://learn.microsoft.com/en-us/office/vba/api/excel.graphic.contrast).
        */
        double GetContrast();

        /**
        Sets the contrast for the specified picture or OLE object. The value for this property must be a number from 0.0 (the least contrast) to 1.0 (the greatest contrast). Read/write Single.

        [MSDN documentation for Graphic.Contrast](https://learn.microsoft.com/en-us/office/vba/api/excel.graphic.contrast).
        */
        void SetContrast(double contrast);

        /**
        Returns the number of points that are cropped off the bottom of the specified picture or OLE object. Read/write Single.

        [MSDN documentation for Graphic.CropBottom](https://learn.microsoft.com/en-us/office/vba/api/excel.graphic.cropbottom).
        */
        double GetCropBottom();

        /**
        Sets the number of points that are cropped off the bottom of the specified picture or OLE object. Read/write Single.

        [MSDN documentation for Graphic.CropBottom](https://learn.microsoft.com/en-us/office/vba/api/excel.graphic.cropbottom).
        */
        void SetCropBottom(double cropBottom);

        /**
        Returns the number of points that are cropped off the left side of the specified picture or OLE object. Read/write Single.

        [MSDN documentation for Graphic.CropLeft](https://learn.microsoft.com/en-us/office/vba/api/excel.graphic.cropleft).
        */
        double GetCropLeft();

        /**
        Sets the number of points that are cropped off the left side of the specified picture or OLE object. Read/write Single.

        [MSDN documentation for Graphic.CropLeft](https://learn.microsoft.com/en-us/office/vba/api/excel.graphic.cropleft).
        */
        void SetCropLeft(double cropLeft);

        /**
        Returns the number of points that are cropped off the right side of the specified picture or OLE object. Read/write Single.

        [MSDN documentation for Graphic.CropRight](https://learn.microsoft.com/en-us/office/vba/api/excel.graphic.cropright).
        */
        double GetCropRight();

        /**
        Sets the number of points that are cropped off the right side of the specified picture or OLE object. Read/write Single.

        [MSDN documentation for Graphic.CropRight](https://learn.microsoft.com/en-us/office/vba/api/excel.graphic.cropright).
        */
        void SetCropRight(double cropRight);

        /**
        Returns the number of points that are cropped off the top of the specified picture or OLE object. Read/write Single.

        [MSDN documentation for Graphic.CropTop](https://learn.microsoft.com/en-us/office/vba/api/excel.graphic.croptop).
        */
        double GetCropTop();

        /**
        Sets the number of points that are cropped off the top of the specified picture or OLE object. Read/write Single.

        [MSDN documentation for Graphic.CropTop](https://learn.microsoft.com/en-us/office/vba/api/excel.graphic.croptop).
        */
        void SetCropTop(double cropTop);

        /**
        Returns the URL (on the intranet or the Web) or path (local or network) to the location where the specified source object was saved.

        [MSDN documentation for Graphic.Filename](https://learn.microsoft.com/en-us/office/vba/api/excel.graphic.filename).
        */
        wxString GetFilename();

        /**
        Sets the URL (on the intranet or the Web) or path (local or network) to the location where the specified source object was saved.

        [MSDN documentation for Graphic.Filename](https://learn.microsoft.com/en-us/office/vba/api/excel.graphic.filename).
        */
        void SetFilename(const wxString& filename);

        /**
        Returns a Single value that represents the height, in points, of the object.

        [MSDN documentation for Graphic.Height](https://learn.microsoft.com/en-us/office/vba/api/excel.graphic.height).
        */
        double GetHeight();

        /**
        Sets a Single value that represents the height, in points, of the object.

        [MSDN documentation for Graphic.Height](https://learn.microsoft.com/en-us/office/vba/api/excel.graphic.height).
        */
        void SetHeight(double height);

        /**
        True if the specified shape retains its original proportions when you resize it. False if you can change the height and width of the shape independently of one another when you resize it. Read/write MsoTriState.

        [MSDN documentation for Graphic.LockAspectRatio](https://learn.microsoft.com/en-us/office/vba/api/excel.graphic.lockaspectratio).
        */
        MsoTriState GetLockAspectRatio();

        /**
        True if the specified shape retains its original proportions when you resize it. False if you can change the height and width of the shape independently of one another when you resize it. Read/write MsoTriState.

        [MSDN documentation for Graphic.LockAspectRatio](https://learn.microsoft.com/en-us/office/vba/api/excel.graphic.lockaspectratio).
        */
        void SetLockAspectRatio(MsoTriState lockAspectRatio);

        /**
        Returns a Single value that represents the width, in points, of the object.

        [MSDN documentation for Graphic.Width](https://learn.microsoft.com/en-us/office/vba/api/excel.graphic.width).
        */
        double GetWidth();

        /**
        Sets a Single value that represents the width, in points, of the object.

        [MSDN documentation for Graphic.Width](https://learn.microsoft.com/en-us/office/vba/api/excel.graphic.width).
        */
        void SetWidth(double width);

        /**
        Returns "Graphic".
        */
        virtual wxString GetAutoExcelObjectName_() const { return wxS("Graphic"); }
    };


} // namespace wxAutoExcel

#endif //_WXAUTOEXCEL_GRAPHIC_H
