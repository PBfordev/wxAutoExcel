/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Modified by:
// RCS-ID:      $Id: $
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
/////////////////////////////////////////////////////////////////////////////


#ifndef _WXAUTOEXCEL_GRAPHIC_H
#define _WXAUTOEXCEL_GRAPHIC_H

#include "wx/wxAutoExcel_defs.h"

#include "wx/wxAutoExcelObject.h"
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

        [MSDN documentation for Graphic.Brightness](http://msdn.microsoft.com/en-us/library/bb148460).
        */
        double GetBrightness();

        /**
        Sets the brightness of the specified picture or OLE object. The value for this property must be a number from 0.0 (dimmest) to 1.0 (brightest). Read/write Single.

        [MSDN documentation for Graphic.Brightness](http://msdn.microsoft.com/en-us/library/bb148460).
        */
        void SetBrightness(double brightness);

        /**
        Returns the type of color transformation applied to the specified picture or OLE object. Read/write MsoPictureColorType.

        [MSDN documentation for Graphic.ColorType](http://msdn.microsoft.com/en-us/library/bb148464).
        */
        MsoPictureColorType GetColorType();

        /**
        Sets the type of color transformation applied to the specified picture or OLE object. Read/write MsoPictureColorType.

        [MSDN documentation for Graphic.ColorType](http://msdn.microsoft.com/en-us/library/bb148464).
        */
        void SetColorType(MsoPictureColorType colorType);

        /**
        Returns the contrast for the specified picture or OLE object. The value for this property must be a number from 0.0 (the least contrast) to 1.0 (the greatest contrast). Read/write Single.

        [MSDN documentation for Graphic.Contrast](http://msdn.microsoft.com/en-us/library/bb148465).
        */
        double GetContrast();

        /**
        Sets the contrast for the specified picture or OLE object. The value for this property must be a number from 0.0 (the least contrast) to 1.0 (the greatest contrast). Read/write Single.

        [MSDN documentation for Graphic.Contrast](http://msdn.microsoft.com/en-us/library/bb148465).
        */
        void SetContrast(double contrast);

        /**
        Returns the number of points that are cropped off the bottom of the specified picture or OLE object. Read/write Single.

        [MSDN documentation for Graphic.CropBottom](http://msdn.microsoft.com/en-us/library/bb148470).
        */
        double GetCropBottom();

        /**
        Sets the number of points that are cropped off the bottom of the specified picture or OLE object. Read/write Single.

        [MSDN documentation for Graphic.CropBottom](http://msdn.microsoft.com/en-us/library/bb148470).
        */
        void SetCropBottom(double cropBottom);

        /**
        Returns the number of points that are cropped off the left side of the specified picture or OLE object. Read/write Single.

        [MSDN documentation for Graphic.CropLeft](http://msdn.microsoft.com/en-us/library/bb148473).
        */
        double GetCropLeft();

        /**
        Sets the number of points that are cropped off the left side of the specified picture or OLE object. Read/write Single.

        [MSDN documentation for Graphic.CropLeft](http://msdn.microsoft.com/en-us/library/bb148473).
        */
        void SetCropLeft(double cropLeft);

        /**
        Returns the number of points that are cropped off the right side of the specified picture or OLE object. Read/write Single.

        [MSDN documentation for Graphic.CropRight](http://msdn.microsoft.com/en-us/library/bb148476).
        */
        double GetCropRight();

        /**
        Sets the number of points that are cropped off the right side of the specified picture or OLE object. Read/write Single.

        [MSDN documentation for Graphic.CropRight](http://msdn.microsoft.com/en-us/library/bb148476).
        */
        void SetCropRight(double cropRight);

        /**
        Returns the number of points that are cropped off the top of the specified picture or OLE object. Read/write Single.

        [MSDN documentation for Graphic.CropTop](http://msdn.microsoft.com/en-us/library/bb148480).
        */
        double GetCropTop();

        /**
        Sets the number of points that are cropped off the top of the specified picture or OLE object. Read/write Single.

        [MSDN documentation for Graphic.CropTop](http://msdn.microsoft.com/en-us/library/bb148480).
        */
        void SetCropTop(double cropTop);

        /**
        Returns the URL (on the intranet or the Web) or path (local or network) to the location where the specified source object was saved.

        [MSDN documentation for Graphic.Filename](http://msdn.microsoft.com/en-us/library/bb148485).
        */
        wxString GetFilename();

        /**
        Sets the URL (on the intranet or the Web) or path (local or network) to the location where the specified source object was saved.

        [MSDN documentation for Graphic.Filename](http://msdn.microsoft.com/en-us/library/bb148485).
        */
        void SetFilename(const wxString& filename);

        /**
        Returns a Single value that represents the height, in points, of the object.

        [MSDN documentation for Graphic.Height](http://msdn.microsoft.com/en-us/library/bb148490).
        */
        double GetHeight();

        /**
        Sets a Single value that represents the height, in points, of the object.

        [MSDN documentation for Graphic.Height](http://msdn.microsoft.com/en-us/library/bb148490).
        */
        void SetHeight(double height);

        /**
        True if the specified shape retains its original proportions when you resize it. False if you can change the height and width of the shape independently of one another when you resize it. Read/write MsoTriState.

        [MSDN documentation for Graphic.LockAspectRatio](http://msdn.microsoft.com/en-us/library/bb148491).
        */
        MsoTriState GetLockAspectRatio();

        /**
        True if the specified shape retains its original proportions when you resize it. False if you can change the height and width of the shape independently of one another when you resize it. Read/write MsoTriState.

        [MSDN documentation for Graphic.LockAspectRatio](http://msdn.microsoft.com/en-us/library/bb148491).
        */
        void SetLockAspectRatio(MsoTriState lockAspectRatio);

        /**
        Returns a Single value that represents the width, in points, of the object.

        [MSDN documentation for Graphic.Width](http://msdn.microsoft.com/en-us/library/bb214625).
        */
        double GetWidth();

        /**
        Sets a Single value that represents the width, in points, of the object.

        [MSDN documentation for Graphic.Width](http://msdn.microsoft.com/en-us/library/bb214625).
        */
        void SetWidth(double width);

        /**
        Returns "Graphic".
        */
        virtual wxString GetAutoExcelObjectName_() const { return wxS("Graphic"); }
    };


} // namespace wxAutoExcel

#endif //_WXAUTOEXCEL_GRAPHIC_H
