/////////////////////////////////////////////////////////////////////////////
// Author:      Claude Opus 5.5 (Anthropic), under the direction of PB
// Copyright:   (c) 2026 PB <pbfordev@gmail.com>
// License:     MIT license
/////////////////////////////////////////////////////////////////////////////


#ifndef _WXAUTOEXCEL_GRADIENTSTOPS_H
#define _WXAUTOEXCEL_GRADIENTSTOPS_H

#include "wx/wxAutoExcel_defs.h"

#if WXAUTOEXCEL_USE_SHAPES || WXAUTOEXCEL_USE_CHARTS

#include "wx/wxAutoExcel_object.h"

class WXDLLIMPEXP_FWD_CORE wxColour;

namespace wxAutoExcel {

    /**
    @brief Represents Microsoft Office GradientStop object, one of the stops of a gradient fill.
    */
    class WXDLLIMPEXP_WXAUTOEXCEL wxExcelGradientStop : public wxExcelObject
    {
    public:
        // ***** PROPERTIES *****

        /**
        Returns a ColorFormat object that represents the color of the gradient stop.

        [MSDN documentation for GradientStop.Color](https://learn.microsoft.com/en-us/office/vba/api/office.gradientstop.color).
        */
        wxExcelColorFormat GetColor();

        /**
        Returns the position of the stop within the gradient, from 0 to 1.

        [MSDN documentation for GradientStop.Position](https://learn.microsoft.com/en-us/office/vba/api/office.gradientstop.position).
        */
        double GetPosition();

        /**
        Sets the position of the stop within the gradient, from 0 to 1.

        [MSDN documentation for GradientStop.Position](https://learn.microsoft.com/en-us/office/vba/api/office.gradientstop.position).
        */
        void SetPosition(double position);

        /**
        Returns the opacity of the stop color, from 0 (opaque) to 1 (clear).

        [MSDN documentation for GradientStop.Transparency](https://learn.microsoft.com/en-us/office/vba/api/office.gradientstop.transparency).
        */
        double GetTransparency();

        /**
        Sets the opacity of the stop color, from 0 (opaque) to 1 (clear).

        [MSDN documentation for GradientStop.Transparency](https://learn.microsoft.com/en-us/office/vba/api/office.gradientstop.transparency).
        */
        void SetTransparency(double transparency);

        /**
        Returns "GradientStop".
        */
        virtual wxString GetAutoExcelObjectName_() const { return wxS("GradientStop"); }
    };


    /**
    @brief Represents Microsoft Office GradientStops collection, the stops of a gradient fill.
    */
    class WXDLLIMPEXP_WXAUTOEXCEL wxExcelGradientStops : public wxExcelObject
    {
    public:
        // ***** METHODS *****

        /**
        Removes the gradient stop at @a index, or the last one when @a index is -1.

        [MSDN documentation for GradientStops.Delete](https://learn.microsoft.com/en-us/office/vba/api/office.gradientstops.delete).
        */
        void Delete(long index = -1);

        /**
        Adds a stop of @a color at @a position, from 0 to 1, with the given @a transparency,
        at @a index of the stops, or after the last one when @a index is -1.

        [MSDN documentation for GradientStops.Insert](https://learn.microsoft.com/en-us/office/vba/api/office.gradientstops.insert).
        */
        void Insert(const wxColour& color, double position, double transparency = 0, long index = -1);

        /**
        Like Insert(), also setting the @a brightness of the stop, from -1 (darkest) to 1 (lightest).

        [MSDN documentation for GradientStops.Insert2](https://learn.microsoft.com/en-us/office/vba/api/office.gradientstops.insert2).
        */
        void Insert2(const wxColour& color, double position, double transparency = 0, long index = -1,
                     double brightness = 0);

        // ***** PROPERTIES *****

        /**
        Returns the number of gradient stops.

        [MSDN documentation for GradientStops.Count](https://learn.microsoft.com/en-us/office/vba/api/office.gradientstops.count).
        */
        long GetCount();

        //@{
        /**
        Returns the gradient stop at @a index.

        [MSDN documentation for GradientStops.Item](https://learn.microsoft.com/en-us/office/vba/api/office.gradientstops.item).
        */
        wxExcelGradientStop GetItem(long index);
        wxExcelGradientStop operator[](long index);
        //@}

        /**
        The iterator over the items, see wxExcelCollectionIterator.
        */
        typedef wxExcelCollectionIterator<wxExcelGradientStops, wxExcelGradientStop> iterator;

        //@{
        /**
        Returns the iterator referring to the first item, or past the last item,
        which allows iterating over the items with a range-based for loop,
        see wxExcelCollectionIterator.
        */
        iterator begin() { return iterator(this, 1); }
        iterator end() { return iterator(this, GetCount() + 1); }
        //@}

        /**
        Returns "GradientStops".
        */
        virtual wxString GetAutoExcelObjectName_() const { return wxS("GradientStops"); }
    };

} // namespace wxAutoExcel

#endif // #if WXAUTOEXCEL_USE_SHAPES || WXAUTOEXCEL_USE_CHARTS

#endif //_WXAUTOEXCEL_GRADIENTSTOPS_H
