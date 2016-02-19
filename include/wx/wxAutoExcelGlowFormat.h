/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Modified by:
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
/////////////////////////////////////////////////////////////////////////////


#ifndef _WXAUTOEXCEL_GLOWFORMAT_H
#define _WXAUTOEXCEL_GLOWFORMAT_H

#include "wx/wxAutoExcel_defs.h"

#if WXAUTOEXCEL_USE_SHAPES || WXAUTOEXCEL_USE_CHARTS

#include "wx/wxAutoExcel_object.h"

namespace wxAutoExcel {

    /**
    @brief Represents Microsoft Excel GlowFormat object.
    */
   class WXDLLIMPEXP_WXAUTOEXCEL wxExcelGlowFormat : public wxExcelObject
    {
    public:        

        // ***** PROPERTIES *****

        /**
        Gets a ColorFormat Represents the color of text formatted as glow.

        [MSDN documentation for GlowFormat.Color](http://msdn.microsoft.com/en-us/library/aa434496).
        */
        wxExcelColorFormat GetColor();

        /**
        Gets or sets the radius value of the glow effect for the GlowFormat object.

        [MSDN documentation for GlowFormat.Radius](http://msdn.microsoft.com/en-us/library/aa434498).
        */
        double GetRadius();

        /**
        Gets or sets the radius value of the glow effect for the GlowFormat object.

        [MSDN documentation for GlowFormat.Radius](http://msdn.microsoft.com/en-us/library/aa434498).
        */
        void SetRadius(double radius);


        /**
        Returns "GlowFormat".
        */
        virtual wxString GetAutoExcelObjectName_() const { return wxS("GlowFormat"); }
    };


} // namespace wxAutoExcel

#endif // #if WXAUTOEXCEL_USE_SHAPES || WXAUTOEXCEL_USE_CHARTS

#endif //_WXAUTOEXCEL_GLOWFORMAT_H
