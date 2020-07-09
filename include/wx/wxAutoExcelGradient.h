/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Modified by:
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
/////////////////////////////////////////////////////////////////////////////


#ifndef _WXAUTOEXCEL_GRADIENT_H
#define _WXAUTOEXCEL_GRADIENT_H

#include "wx/wxAutoExcel_defs.h"

#include "wx/wxAutoExcel_object.h"
#include "wx/wxAutoExcel_enums.h"

class WXDLLIMPEXP_FWD_CORE wxColour;

namespace wxAutoExcel {

    /**
    @brief Represents Microsoft Excel ColorStop object.
    */
   class WXDLLIMPEXP_WXAUTOEXCEL wxExcelColorStop: public wxExcelObject
    {
    public:
        // ***** METHODS *****

        /**
        Deletes the object.

        [MSDN documentation for ColorStop.Delete](http://msdn.microsoft.com/en-us/library/bb242962).
        */
        bool Delete();

        // ***** PROPERTIES *****

        /**
        Returns the color.

        [MSDN documentation for ColorStop.Color]().
        */
        wxColour GetColor();

        /**
        Sets the color.

        [MSDN documentation for ColorStop.Color]().
        */
        void SetColor(const wxColour& color);

        /**
        Returns its position.

        [MSDN documentation for ColorStop.Position]().
        */
        double GetPosition();

        /**
        Sets the position.

        [MSDN documentation for ColorStop.Position]().
        */
        void SetPosition(double position);

        /**
        Returns the theme color.

        [MSDN documentation for ColorStop.ThemeColor]().
        */
        XlThemeColor GetThemeColor();

        /**
        Sets the theme color.

        [MSDN documentation for ColorStop.ThemeColor]().
        */
        void SetThemeColor(XlThemeColor themeColor);

        /**
        A number from -1 (darkest) to 1 (lightest) or 0 for neutral.

        [MSDN documentation for ColorStop.TintAndShade]().
        */
        double GetTintAndShade();

        /**
        A number from -1 (darkest) to 1 (lightest) or 0 for neutral.

        [MSDN documentation for ColorStop.TintAndShade]().
        */
        void SetTintAndShade(double tintAndShade);

        /**
        Returns "ColorStop".
        */
        virtual wxString GetAutoExcelObjectName_() const { return wxS("ColorStops"); }
    };

    /**
    @brief Represents Microsoft Excel ColorStops collection.
    */
   class WXDLLIMPEXP_WXAUTOEXCEL wxExcelColorStops: public wxExcelObject
    {
    public:
        // ***** METHODS *****

        /**
        Adds a ColorStop object.

        [MSDN documentation for ColorStops.Add](http://msdn.microsoft.com/en-us/library/bb257094).
        */
        wxExcelColorStop Add(double position);

        /**
        Clears the object.

        [MSDN documentation for ColorStops.Clear](http://msdn.microsoft.com/en-us/library/bb242935).
        */
        void Clear();

        // ***** PROPERTIES *****

        /**
        Returns a ColorStop object.

        [MSDN documentation for ColorStops.Item](http://msdn.microsoft.com/en-us/library/bb242942).
        */
        wxExcelColorStop GetItem(long index);


        /**
        Returns the number of ColorStops. Since Excel 2007.

        [MSDN documentation for ColorStops.Count](http://msdn.microsoft.com/en-us/library/bb242953).
        */
        long GetCount();
        /**
        Returns "ColorStops".
        */
        virtual wxString GetAutoExcelObjectName_() const { return wxS("ColorStops"); }
    };

    /**
    @brief Represents Microsoft Excel LinearGradient object.
    */
   class WXDLLIMPEXP_WXAUTOEXCEL wxExcelLinearGradient: public wxExcelObject
    {
    public:
        // ***** PROPERTIES *****

        /**
        Returns the ColorStops collection. Read-only  Since Excel 2007.

        [MSDN documentation for LinearGradient.ColorStops](http://msdn.microsoft.com/en-us/library/bb243113).
        */
        wxExcelColorStops GetColorStops();

        /**
        The angle of the linear gradient fill within a selection.

        [MSDN documentation for LinearGradient.Degree](http://msdn.microsoft.com/en-us/library/bb243119).
        */
        long GetDegree();

        /**
        The angle of the linear gradient fill within a selection.

        [MSDN documentation for LinearGradient.Degree](http://msdn.microsoft.com/en-us/library/bb243119).
        */
        void SetDegree(long degree);

        /**
        Returns "LinearGradient".
        */
        virtual wxString GetAutoExcelObjectName_() const { return wxS("LinearGradient"); }
    };

    /**
    @brief Represents Microsoft Excel RectangularGradient object.
    */
   class WXDLLIMPEXP_WXAUTOEXCEL wxExcelRectangularGradient: public wxExcelObject
    {
    public:
        // ***** PROPERTIES *****

        /**
        Returns the ColorStops collection. Since Excel 2007.

        [MSDN documentation for RectangularGradient.ColorStops](http://msdn.microsoft.com/en-us/library/bb243140).
        */
        wxExcelColorStops GetColorStops();

        /**
        Represents the point or vector that the gradient fill converges to. Since Excel 2007.

        [MSDN documentation for RectangularGradient.RectangleBottom](http://msdn.microsoft.com/en-us/library/bb243151).
        */
        double GetRectangleBottom();

        /**
        Represents the point or vector that the gradient fill converges to.

        [MSDN documentation for RectangularGradient.RectangleBottom](http://msdn.microsoft.com/en-us/library/bb243151).
        */
        void SetRectangleBottom(double rectangleBottom);

        /**
        Represents the point or vector that the gradient fill converges to. Since Excel 2007.

        [MSDN documentation for RectangularGradient.RectangleLeft](http://msdn.microsoft.com/en-us/library/bb243156).
        */
        double GetRectangleLeft();

        /**
        Represents the point or vector that the gradient fill converges to.

        [MSDN documentation for RectangularGradient.RectangleLeft](http://msdn.microsoft.com/en-us/library/bb243156).
        */
        void SetRectangleLeft(double rectangleLeft);

        /**
        Represents the point or vector that the gradient fill converges to. Since Excel 2007.

        [MSDN documentation for RectangularGradient.RectangleRight](http://msdn.microsoft.com/en-us/library/bb243159).
        */
        double GetRectangleRight();

        /**
        Represents the point or vector that the gradient fill converges to.

        [MSDN documentation for RectangularGradient.RectangleRight](http://msdn.microsoft.com/en-us/library/bb243159).
        */
        void SetRectangleRight(double rectangleRight);

        /**
        Represents the point or vector that the gradient fill converges to. Since Excel 2007.

        [MSDN documentation for RectangularGradient.RectangleTop](http://msdn.microsoft.com/en-us/library/bb243165).
        */
        double GetRectangleTop();

        /**
        Represents the point or vector that the gradient fill converges to.

        [MSDN documentation for RectangularGradient.RectangleTop](http://msdn.microsoft.com/en-us/library/bb243165).
        */
        void SetRectangleTop(double rectangleTop);

        /**
        Returns "RectangularGradient".
        */
        virtual wxString GetAutoExcelObjectName_() const { return wxS("RectangularGradient"); }
    };

} // namespace wxAutoExcel

#endif //_WXAUTOEXCEL_GRADIENT_H
