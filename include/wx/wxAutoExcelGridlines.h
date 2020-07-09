/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Modified by:
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
/////////////////////////////////////////////////////////////////////////////


#ifndef _WXAUTOEXCEL_GRIDLINES_H
#define _WXAUTOEXCEL_GRIDLINES_H

#include "wx/wxAutoExcel_defs.h"

#if WXAUTOEXCEL_USE_CHARTS

#include "wx/wxAutoExcel_object.h"

namespace wxAutoExcel {


    /**
    Represents Microsoft Excel Gridlines object.
    */
   class WXDLLIMPEXP_WXAUTOEXCEL wxExcelGridlines : public wxExcelObject
    {
    public:
        // ***** METHODS *****

        /**
        Deletes the object.

        [MSDN documentation for Gridlines.Delete](http://msdn.microsoft.com/en-us/library/bb211789).
        */
        bool Delete();

        /**
        Selects the object.

        [MSDN documentation for Gridlines.Select](http://msdn.microsoft.com/en-us/library/bb237952).
        */
        bool Select();

        // ***** PROPERTIES *****


        /**
        Returns a Border object that represents the border of the object.

        [MSDN documentation for Gridlines.Border](http://msdn.microsoft.com/en-us/library//bb148496).
        */
        wxExcelBorder GetBorder();

        /**
        Returns the ChartFormat object.

        [MSDN documentation for Gridlines.Format](http://msdn.microsoft.com/en-us/library/bb242530).
        */
        wxExcelChartFormat GetFormat();

        /**
        Returns a String value that represents the name of the object.

        [MSDN documentation for Gridlines.Name](http://msdn.microsoft.com/en-us/library//bb148497).
        */
        wxString GetName();

        /**
        Returns "Gridlines".
        */
        virtual wxString GetAutoExcelObjectName_() const { return wxS("Gridlines"); }
    };


} // namespace wxAutoExcel


#endif // #if WXAUTOEXCEL_USE_CHARTS

#endif // _WXAUTOEXCEL_GRIDLINES_H
