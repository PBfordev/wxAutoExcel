/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Modified by:
// RCS-ID:      $Id: $
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
/////////////////////////////////////////////////////////////////////////////


#ifndef _WXAUTOEXCEL_TABSTOPS2_H
#define _WXAUTOEXCEL_TABSTOPS2_H

#include "wx/wxAutoExcel_defs.h"

#if WXAUTOEXCEL_USE_SHAPES || WXAUTOEXCEL_USE_CHARTS

#include "wx/wxAutoExcelObject.h"

namespace wxAutoExcel {

    /**
    @brief Represents Microsoft Excel TabStop2 object.
    */
   class WXDLLIMPEXP_WXAUTOEXCEL wxExcelTabStop2 : public wxExcelObject
    {
    public:        
        // ***** METHODS *****

        /**
        Removes the specified custom tab stop

        [MSDN documentation for TabStop2.Clear](http://msdn.microsoft.com/en-us/library/aa434138.aspx).
        */
        void Clear();

        // ***** PROPERTIES *****

        /**
        Gets or sets the position of a tab stop relative to the left margin.

        [MSDN documentation for TabStop2.Position](http://msdn.microsoft.com/en-us/library/aa434414.aspx).
        */
        double GetPosition();

        /**
        Gets or sets the position of a tab stop relative to the left margin.

        [MSDN documentation for TabStop2.Position](http://msdn.microsoft.com/en-us/library/aa434414.aspx).
        */
        void SetPosition(double position);

        /**
        Gets or sets the type of the TabStop2 object.

        [MSDN documentation for TabStop2.Type](http://msdn.microsoft.com/en-us/library/aa434415.aspx).
        */
        MsoTabStopType GetType();

        /**
        Gets or sets the type of the TabStop2 object.

        [MSDN documentation for TabStop2.Type](http://msdn.microsoft.com/en-us/library/aa434415.aspx).
        */
        void SetType(MsoTabStopType type);

        /**
        Returns "TabStop2".
        */
        virtual wxString GetAutoExcelObjectName_() const { return wxS("TabStop2"); }
    };



    /**
    @brief Represents Microsoft Excel TabStops2 collection.
    */
   class WXDLLIMPEXP_WXAUTOEXCEL wxExcelTabStops2 : public wxExcelObject
    {
    public:        
        // ***** METHODS *****

        /**
        Adds a new tab stop to the specified TabStops2 object.

        [MSDN documentation for TabStops2.Add](http://msdn.microsoft.com/en-us/library/aa434139.aspx).
        */
        wxExcelTabStop2 Add(MsoTabStopType type, double position);

        //@{
        /**
        Gets an individual object from the TabStops2 collection.

        [MSDN documentation for TabStops2.Item](http://msdn.microsoft.com/en-us/library/aa434140.aspx).
        */
        wxExcelTabStop2 Item(long index);
        wxExcelTabStop2 operator[](long index);
        //@}

        // ***** PROPERTIES *****
        

        /**
        Gets a Long indicating the number of items in the TabStops2 collection.

        [MSDN documentation for TabStops2.Count](http://msdn.microsoft.com/en-us/library/aa434417.aspx).
        */
        long GetCount();

        /**
        Gets or sets the default spacing between tab stops.

        [MSDN documentation for TabStops2.DefaultSpacing](http://msdn.microsoft.com/en-us/library/aa434419.aspx).
        */
        double GetDefaultSpacing();

        /**
        Gets or sets the default spacing between tab stops.

        [MSDN documentation for TabStops2.DefaultSpacing](http://msdn.microsoft.com/en-us/library/aa434419.aspx).
        */
        void SetDefaultSpacing(double defaultSpacing);

        /**
        Returns "TabStops2".
        */
        virtual wxString GetAutoExcelObjectName_() const { return wxS("TabStops2"); }
    };

} // namespace wxAutoExcel

#endif // #if WXAUTOEXCEL_USE_SHAPES || WXAUTOEXCEL_USE_CHARTS

#endif //_WXAUTOEXCEL_TABSTOPS2_H

