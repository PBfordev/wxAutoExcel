/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// License:     MIT license
/////////////////////////////////////////////////////////////////////////////


#ifndef _WXAUTOEXCEL_PICTUREFORMAT_H
#define _WXAUTOEXCEL_PICTUREFORMAT_H

#include "wx/wxAutoExcel_defs.h"

#if WXAUTOEXCEL_USE_SHAPES || WXAUTOEXCEL_USE_CHARTS

#include "wx/wxAutoExcel_object.h"
#include "wx/wxAutoExcel_enums.h"

class WXDLLIMPEXP_FWD_CORE wxColour;

namespace wxAutoExcel {


    /**
    @brief Represents Microsoft Excel PictureFormat object.
    */
   class WXDLLIMPEXP_WXAUTOEXCEL wxExcelPictureFormat : public wxExcelObject
    {
    public:
        // ***** METHODS *****

        /**
        Changes the brightness of the picture by the specified amount. Use the Brightness property to set the absolute brightness of the picture.

        [MSDN documentation for PictureFormat.IncrementBrightness](https://learn.microsoft.com/en-us/office/vba/api/excel.pictureformat.incrementbrightness).
        */
        void IncrementBrightness(double increment);

        /**
        Changes the contrast of the picture by the specified amount. Use the Contrast property to set the absolute contrast for the picture.

        [MSDN documentation for PictureFormat.IncrementContrast](https://learn.microsoft.com/en-us/office/vba/api/excel.pictureformat.incrementcontrast).
        */
        void IncrementContrast(double increment);

        // ***** PROPERTIES *****

        /**
        Returns the brightness of the specified picture or OLE object. The value for this property must be a number from 0.0 (dimmest) to 1.0 (brightest). Read/write Single.

        [MSDN documentation for PictureFormat.Brightness](https://learn.microsoft.com/en-us/office/vba/api/excel.pictureformat.brightness).
        */
        double GetBrightness();

        /**
        Sets the brightness of the specified picture or OLE object. The value for this property must be a number from 0.0 (dimmest) to 1.0 (brightest). Read/write Single.

        [MSDN documentation for PictureFormat.Brightness](https://learn.microsoft.com/en-us/office/vba/api/excel.pictureformat.brightness).
        */
        void SetBrightness(double brightness);

        /**
        Returns the type of color transformation applied to the specified picture or OLE object. Read/write MsoPictureColorType.

        [MSDN documentation for PictureFormat.ColorType](https://learn.microsoft.com/en-us/office/vba/api/excel.pictureformat.colortype).
        */
        MsoPictureColorType GetColorType();

        /**
        Sets the type of color transformation applied to the specified picture or OLE object. Read/write MsoPictureColorType.

        [MSDN documentation for PictureFormat.ColorType](https://learn.microsoft.com/en-us/office/vba/api/excel.pictureformat.colortype).
        */
        void SetColorType(MsoPictureColorType colorType);

        /**
        Returns the contrast for the specified picture or OLE object. The value for this property must be a number from 0.0 (the least contrast) to 1.0 (the greatest contrast). Read/write Single.

        [MSDN documentation for PictureFormat.Contrast](https://learn.microsoft.com/en-us/office/vba/api/excel.pictureformat.contrast).
        */
        double GetContrast();

        /**
        Sets the contrast for the specified picture or OLE object. The value for this property must be a number from 0.0 (the least contrast) to 1.0 (the greatest contrast). Read/write Single.

        [MSDN documentation for PictureFormat.Contrast](https://learn.microsoft.com/en-us/office/vba/api/excel.pictureformat.contrast).
        */
        void SetContrast(double contrast);

        /**
        Returns the number of points that are cropped off the bottom of the specified picture or OLE object. Read/write Single.

        [MSDN documentation for PictureFormat.CropBottom](https://learn.microsoft.com/en-us/office/vba/api/excel.pictureformat.cropbottom).
        */
        double GetCropBottom();

        /**
        Sets the number of points that are cropped off the bottom of the specified picture or OLE object. Read/write Single.

        [MSDN documentation for PictureFormat.CropBottom](https://learn.microsoft.com/en-us/office/vba/api/excel.pictureformat.cropbottom).
        */
        void SetCropBottom(double cropBottom);

        /**
        Returns the number of points that are cropped off the left side of the specified picture or OLE object. Read/write Single.

        [MSDN documentation for PictureFormat.CropLeft](https://learn.microsoft.com/en-us/office/vba/api/excel.pictureformat.cropleft).
        */
        double GetCropLeft();

        /**
        Sets the number of points that are cropped off the left side of the specified picture or OLE object. Read/write Single.

        [MSDN documentation for PictureFormat.CropLeft](https://learn.microsoft.com/en-us/office/vba/api/excel.pictureformat.cropleft).
        */
        void SetCropLeft(double cropLeft);

        /**
        Returns the number of points that are cropped off the right side of the specified picture or OLE object. Read/write Single.

        [MSDN documentation for PictureFormat.CropRight](https://learn.microsoft.com/en-us/office/vba/api/excel.pictureformat.cropright).
        */
        double GetCropRight();

        /**
        Sets the number of points that are cropped off the right side of the specified picture or OLE object. Read/write Single.

        [MSDN documentation for PictureFormat.CropRight](https://learn.microsoft.com/en-us/office/vba/api/excel.pictureformat.cropright).
        */
        void SetCropRight(double cropRight);

        /**
        Returns the number of points that are cropped off the top of the specified picture or OLE object. Read/write Single.

        [MSDN documentation for PictureFormat.CropTop](https://learn.microsoft.com/en-us/office/vba/api/excel.pictureformat.croptop).
        */
        double GetCropTop();

        /**
        Sets the number of points that are cropped off the top of the specified picture or OLE object. Read/write Single.

        [MSDN documentation for PictureFormat.CropTop](https://learn.microsoft.com/en-us/office/vba/api/excel.pictureformat.croptop).
        */
        void SetCropTop(double cropTop);

        /**
        Returns the transparent color for the specified picture as a red-green-blue (RGB) value. For this property to take effect, the TransparentBackground property must be set to True. Applies to bitmaps only.

        [MSDN documentation for PictureFormat.TransparencyColor](https://learn.microsoft.com/en-us/office/vba/api/excel.pictureformat.transparencycolor).
        */
        wxColour GetTransparencyColor();

        /**
        Sets the transparent color for the specified picture as a red-green-blue (RGB) value. For this property to take effect, the TransparentBackground property must be set to True. Applies to bitmaps only.

        [MSDN documentation for PictureFormat.TransparencyColor](https://learn.microsoft.com/en-us/office/vba/api/excel.pictureformat.transparencycolor).
        */
        void SetTransparencyColor(const wxColour& transparencyColor);

        /**
        Use the TransparencyColor property to set the transparent color. Applies to bitmaps only. Read/write MsoTriState.

        [MSDN documentation for PictureFormat.TransparentBackground](https://learn.microsoft.com/en-us/office/vba/api/excel.pictureformat.transparentbackground).
        */
        MsoTriState GetTransparentBackground();

        /**
        Use the TransparencyColor property to set the transparent color. Applies to bitmaps only. Read/write MsoTriState.

        [MSDN documentation for PictureFormat.TransparentBackground](https://learn.microsoft.com/en-us/office/vba/api/excel.pictureformat.transparentbackground).
        */
        void SetTransparentBackground(MsoTriState transparentBackground);


        /**
        Returns "PictureFormat".
        */
        virtual wxString GetAutoExcelObjectName_() const { return wxS("PictureFormat"); }
    };


} // namespace wxAutoExcel

#endif

#endif //_WXAUTOEXCEL_PICTUREFORMAT_H
