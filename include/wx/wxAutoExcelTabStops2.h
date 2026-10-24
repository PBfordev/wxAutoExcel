/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// License:     MIT license
/////////////////////////////////////////////////////////////////////////////


#ifndef _WXAUTOEXCEL_TABSTOPS2_H
#define _WXAUTOEXCEL_TABSTOPS2_H

#include "wx/wxAutoExcel_defs.h"

#if WXAUTOEXCEL_USE_SHAPES || WXAUTOEXCEL_USE_CHARTS

#include "wx/wxAutoExcel_object.h"

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

        [MSDN documentation for TabStop2.Clear](https://learn.microsoft.com/en-us/office/vba/api/office.tabstop2.clear).
        */
        void Clear();

        // ***** PROPERTIES *****

        /**
        Gets or sets the position of a tab stop relative to the left margin.

        [MSDN documentation for TabStop2.Position](https://learn.microsoft.com/en-us/office/vba/api/office.tabstop2.position).
        */
        double GetPosition();

        /**
        Gets or sets the position of a tab stop relative to the left margin.

        [MSDN documentation for TabStop2.Position](https://learn.microsoft.com/en-us/office/vba/api/office.tabstop2.position).
        */
        void SetPosition(double position);

        /**
        Gets or sets the type of the TabStop2 object.

        [MSDN documentation for TabStop2.Type](https://learn.microsoft.com/en-us/office/vba/api/office.tabstop2.type).
        */
        MsoTabStopType GetType();

        /**
        Gets or sets the type of the TabStop2 object.

        [MSDN documentation for TabStop2.Type](https://learn.microsoft.com/en-us/office/vba/api/office.tabstop2.type).
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

        [MSDN documentation for TabStops2.Add](https://learn.microsoft.com/en-us/office/vba/api/office.tabstops2.add).
        */
        wxExcelTabStop2 Add(MsoTabStopType type, double position);

        //@{
        /**
        Gets an individual object from the TabStops2 collection.

        [MSDN documentation for TabStops2.Item](https://learn.microsoft.com/en-us/office/vba/api/office.tabstops2.item).
        */
        wxExcelTabStop2 Item(long index);
        wxExcelTabStop2 operator[](long index);
        //@}

        // ***** PROPERTIES *****


        /**
        Gets a Long indicating the number of items in the TabStops2 collection.

        [MSDN documentation for TabStops2.Count](https://learn.microsoft.com/en-us/office/vba/api/office.tabstops2.count).
        */
        long GetCount();

        /**
        Gets or sets the default spacing between tab stops.

        [MSDN documentation for TabStops2.DefaultSpacing](https://learn.microsoft.com/en-us/office/vba/api/office.tabstops2.defaultspacing).
        */
        double GetDefaultSpacing();

        /**
        Gets or sets the default spacing between tab stops.

        [MSDN documentation for TabStops2.DefaultSpacing](https://learn.microsoft.com/en-us/office/vba/api/office.tabstops2.defaultspacing).
        */
        void SetDefaultSpacing(double defaultSpacing);

        /**
        The iterator over the items, see wxExcelCollectionIterator.
        */
        typedef wxExcelCollectionIterator<wxExcelTabStops2, wxExcelTabStop2> iterator;

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
        Returns "TabStops2".
        */
        virtual wxString GetAutoExcelObjectName_() const { return wxS("TabStops2"); }
    };

} // namespace wxAutoExcel

#endif // #if WXAUTOEXCEL_USE_SHAPES || WXAUTOEXCEL_USE_CHARTS

#endif //_WXAUTOEXCEL_TABSTOPS2_H

