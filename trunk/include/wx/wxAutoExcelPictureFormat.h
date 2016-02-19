/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Modified by:
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
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

        [MSDN documentation for PictureFormat.IncrementBrightness](http://msdn.microsoft.com/en-us/library/bb209932).
        */
        void IncrementBrightness(double increment);

        /**
        Changes the contrast of the picture by the specified amount. Use the Contrast property to set the absolute contrast for the picture.

        [MSDN documentation for PictureFormat.IncrementContrast](http://msdn.microsoft.com/en-us/library/bb209935).
        */
        void IncrementContrast(double increment);

        // ***** PROPERTIES *****

        /**
        Returns the brightness of the specified picture or OLE object. The value for this property must be a number from 0.0 (dimmest) to 1.0 (brightest). Read/write Single.

        [MSDN documentation for PictureFormat.Brightness](http://msdn.microsoft.com/en-us/library/bb213258).
        */
        double GetBrightness();

        /**
        Sets the brightness of the specified picture or OLE object. The value for this property must be a number from 0.0 (dimmest) to 1.0 (brightest). Read/write Single.

        [MSDN documentation for PictureFormat.Brightness](http://msdn.microsoft.com/en-us/library/bb213258).
        */
        void SetBrightness(double brightness);

        /**
        Returns the type of color transformation applied to the specified picture or OLE object. Read/write MsoPictureColorType.

        [MSDN documentation for PictureFormat.ColorType](http://msdn.microsoft.com/en-us/library/bb213260).
        */
        MsoPictureColorType GetColorType();

        /**
        Sets the type of color transformation applied to the specified picture or OLE object. Read/write MsoPictureColorType.

        [MSDN documentation for PictureFormat.ColorType](http://msdn.microsoft.com/en-us/library/bb213260).
        */
        void SetColorType(MsoPictureColorType colorType);

        /**
        Returns the contrast for the specified picture or OLE object. The value for this property must be a number from 0.0 (the least contrast) to 1.0 (the greatest contrast). Read/write Single.

        [MSDN documentation for PictureFormat.Contrast](http://msdn.microsoft.com/en-us/library/bb213262).
        */
        double GetContrast();

        /**
        Sets the contrast for the specified picture or OLE object. The value for this property must be a number from 0.0 (the least contrast) to 1.0 (the greatest contrast). Read/write Single.

        [MSDN documentation for PictureFormat.Contrast](http://msdn.microsoft.com/en-us/library/bb213262).
        */
        void SetContrast(double contrast);

        /**
        Returns the number of points that are cropped off the bottom of the specified picture or OLE object. Read/write Single.

        [MSDN documentation for PictureFormat.CropBottom](http://msdn.microsoft.com/en-us/library/bb213263).
        */
        double GetCropBottom();

        /**
        Sets the number of points that are cropped off the bottom of the specified picture or OLE object. Read/write Single.

        [MSDN documentation for PictureFormat.CropBottom](http://msdn.microsoft.com/en-us/library/bb213263).
        */
        void SetCropBottom(double cropBottom);

        /**
        Returns the number of points that are cropped off the left side of the specified picture or OLE object. Read/write Single.

        [MSDN documentation for PictureFormat.CropLeft](http://msdn.microsoft.com/en-us/library/bb213265).
        */
        double GetCropLeft();

        /**
        Sets the number of points that are cropped off the left side of the specified picture or OLE object. Read/write Single.

        [MSDN documentation for PictureFormat.CropLeft](http://msdn.microsoft.com/en-us/library/bb213265).
        */
        void SetCropLeft(double cropLeft);

        /**
        Returns the number of points that are cropped off the right side of the specified picture or OLE object. Read/write Single.

        [MSDN documentation for PictureFormat.CropRight](http://msdn.microsoft.com/en-us/library/bb213267).
        */
        double GetCropRight();

        /**
        Sets the number of points that are cropped off the right side of the specified picture or OLE object. Read/write Single.

        [MSDN documentation for PictureFormat.CropRight](http://msdn.microsoft.com/en-us/library/bb213267).
        */
        void SetCropRight(double cropRight);

        /**
        Returns the number of points that are cropped off the top of the specified picture or OLE object. Read/write Single.

        [MSDN documentation for PictureFormat.CropTop](http://msdn.microsoft.com/en-us/library/bb213269).
        */
        double GetCropTop();

        /**
        Sets the number of points that are cropped off the top of the specified picture or OLE object. Read/write Single.

        [MSDN documentation for PictureFormat.CropTop](http://msdn.microsoft.com/en-us/library/bb213269).
        */
        void SetCropTop(double cropTop);

        /**
        Returns the transparent color for the specified picture as a red-green-blue (RGB) value. For this property to take effect, the TransparentBackground property must be set to True. Applies to bitmaps only.

        [MSDN documentation for PictureFormat.TransparencyColor](http://msdn.microsoft.com/en-us/library/bb221909).
        */
        wxColour GetTransparencyColor();

        /**
        Sets the transparent color for the specified picture as a red-green-blue (RGB) value. For this property to take effect, the TransparentBackground property must be set to True. Applies to bitmaps only.

        [MSDN documentation for PictureFormat.TransparencyColor](http://msdn.microsoft.com/en-us/library/bb221909).
        */
        void SetTransparencyColor(const wxColour& transparencyColor);

        /**
        Use the TransparencyColor property to set the transparent color. Applies to bitmaps only. Read/write MsoTriState.

        [MSDN documentation for PictureFormat.TransparentBackground](http://msdn.microsoft.com/en-us/library/bb221914).
        */
        MsoTriState GetTransparentBackground();

        /**
        Use the TransparencyColor property to set the transparent color. Applies to bitmaps only. Read/write MsoTriState.

        [MSDN documentation for PictureFormat.TransparentBackground](http://msdn.microsoft.com/en-us/library/bb221914).
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
