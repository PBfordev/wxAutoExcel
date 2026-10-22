/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// License:     MIT license
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

        [MSDN documentation for ColorStop.Delete](https://learn.microsoft.com/en-us/office/vba/api/excel.colorstop.delete).
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

        [MSDN documentation for ColorStops.Add](https://learn.microsoft.com/en-us/office/vba/api/excel.colorstops.add).
        */
        wxExcelColorStop Add(double position);

        /**
        Clears the object.

        [MSDN documentation for ColorStops.Clear](https://learn.microsoft.com/en-us/office/vba/api/excel.colorstops.clear).
        */
        void Clear();

        // ***** PROPERTIES *****

        /**
        Returns a ColorStop object.

        [MSDN documentation for ColorStops.Item](https://learn.microsoft.com/en-us/office/vba/api/excel.colorstops.item).
        */
        wxExcelColorStop GetItem(long index);


        /**
        Returns the number of ColorStops. Since Excel 2007.

        [MSDN documentation for ColorStops.Count](https://learn.microsoft.com/en-us/office/vba/api/excel.colorstops.count).
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

        [MSDN documentation for LinearGradient.ColorStops](https://learn.microsoft.com/en-us/office/vba/api/excel.lineargradient.colorstops).
        */
        wxExcelColorStops GetColorStops();

        /**
        The angle of the linear gradient fill within a selection.

        [MSDN documentation for LinearGradient.Degree](https://learn.microsoft.com/en-us/office/vba/api/excel.lineargradient.degree).
        */
        double GetDegree();

        /**
        The angle of the linear gradient fill within a selection.

        [MSDN documentation for LinearGradient.Degree](https://learn.microsoft.com/en-us/office/vba/api/excel.lineargradient.degree).
        */
        void SetDegree(double degree);

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

        [MSDN documentation for RectangularGradient.ColorStops](https://learn.microsoft.com/en-us/office/vba/api/excel.rectangulargradient.colorstops).
        */
        wxExcelColorStops GetColorStops();

        /**
        Represents the point or vector that the gradient fill converges to. Since Excel 2007.

        [MSDN documentation for RectangularGradient.RectangleBottom](https://learn.microsoft.com/en-us/office/vba/api/excel.rectangulargradient.rectanglebottom).
        */
        double GetRectangleBottom();

        /**
        Represents the point or vector that the gradient fill converges to.

        [MSDN documentation for RectangularGradient.RectangleBottom](https://learn.microsoft.com/en-us/office/vba/api/excel.rectangulargradient.rectanglebottom).
        */
        void SetRectangleBottom(double rectangleBottom);

        /**
        Represents the point or vector that the gradient fill converges to. Since Excel 2007.

        [MSDN documentation for RectangularGradient.RectangleLeft](https://learn.microsoft.com/en-us/office/vba/api/excel.rectangulargradient.rectangleleft).
        */
        double GetRectangleLeft();

        /**
        Represents the point or vector that the gradient fill converges to.

        [MSDN documentation for RectangularGradient.RectangleLeft](https://learn.microsoft.com/en-us/office/vba/api/excel.rectangulargradient.rectangleleft).
        */
        void SetRectangleLeft(double rectangleLeft);

        /**
        Represents the point or vector that the gradient fill converges to. Since Excel 2007.

        [MSDN documentation for RectangularGradient.RectangleRight](https://learn.microsoft.com/en-us/office/vba/api/excel.rectangulargradient.rectangleright).
        */
        double GetRectangleRight();

        /**
        Represents the point or vector that the gradient fill converges to.

        [MSDN documentation for RectangularGradient.RectangleRight](https://learn.microsoft.com/en-us/office/vba/api/excel.rectangulargradient.rectangleright).
        */
        void SetRectangleRight(double rectangleRight);

        /**
        Represents the point or vector that the gradient fill converges to. Since Excel 2007.

        [MSDN documentation for RectangularGradient.RectangleTop](https://learn.microsoft.com/en-us/office/vba/api/excel.rectangulargradient.rectangletop).
        */
        double GetRectangleTop();

        /**
        Represents the point or vector that the gradient fill converges to.

        [MSDN documentation for RectangularGradient.RectangleTop](https://learn.microsoft.com/en-us/office/vba/api/excel.rectangulargradient.rectangletop).
        */
        void SetRectangleTop(double rectangleTop);

        /**
        Returns "RectangularGradient".
        */
        virtual wxString GetAutoExcelObjectName_() const { return wxS("RectangularGradient"); }
    };

} // namespace wxAutoExcel

#endif //_WXAUTOEXCEL_GRADIENT_H
