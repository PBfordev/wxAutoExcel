/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Modified by:
// RCS-ID:      $Id: $
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
/////////////////////////////////////////////////////////////////////////////


#ifndef _WXAUTOEXCEL_PAGEBREAKS_H
#define _WXAUTOEXCEL_PAGEBREAKS_H

#include "wx/wxAutoExcel_defs.h"

#include "wx/wxAutoExcelObject.h"
#include "wx/wxAutoExcel_enums.h"

namespace wxAutoExcel {

    /**
    @brief Represents Microsoft Excel HPageBreak or VPageBreak object.
    */
   class WXDLLIMPEXP_WXAUTOEXCEL wxExcelPageBreak : public wxExcelObject
    {
    public:        
        // ***** METHODS *****

        /**
        Deletes the object.

        [MSDN documentation for HPageBreak.Delete](http://msdn.microsoft.com/en-us/library/bb211801).
        [MSDN documentation for VPageBreak.Delete](http://msdn.microsoft.com/en-us/library/bb179095).
        */
        void Delete();

        /**
        Drags a page break out of the print area.

        [MSDN documentation for HPageBreak.DragOff](http://msdn.microsoft.com/en-us/library/bb211802).
        [MSDN documentation for VPageBreak.DragOff](http://msdn.microsoft.com/en-us/library/bb179099).
        */
        void DragOff(XlDirection direction, long regionIndex);

        // ***** PROPERTIES *****


        /**
        Returns the type of the specified page break: full-screen or only within a print area. Can be either of the following XlPageBreakExtent constants: xlPageBreakFull or xlPageBreakPartial.

        [MSDN documentation for HPageBreak.Extent](http://msdn.microsoft.com/en-us/library/bb148510).
        [MSDN documentation for VPageBreak.Extent](http://msdn.microsoft.com/en-us/library/bb214181).
        */
        XlPageBreakExtent GetExtent();

        /**
        Returns the cell (a Range object) that defines the page-break location. Horizontal page breaks are aligned with the top edge of the location cell; vertical page breaks are aligned with the left edge of the location cell. 

        [MSDN documentation for HPageBreak.Location](http://msdn.microsoft.com/en-us/library/bb148512).
        [MSDN documentation for VPageBreak.Location](http://msdn.microsoft.com/en-us/library/bb214186).
        */
        wxExcelRange GetLocation();

        /**
        Sets the cell (a Range object) that defines the page-break location. Horizontal page breaks are aligned with the top edge of the location cell; vertical page breaks are aligned with the left edge of the location cell.

        [MSDN documentation for HPageBreak.Location](http://msdn.microsoft.com/en-us/library/bb148512).
        [MSDN documentation for VPageBreak.Location](http://msdn.microsoft.com/en-us/library/bb214186).
        */
        void SetLocation(wxExcelRange location);

        /**
        Returns a XlPageBreak value that represents the page break type.

        [MSDN documentation for HPageBreak.Type](http://msdn.microsoft.com/en-us/library/bb214627).
        [MSDN documentation for VPageBreak.Type](http://msdn.microsoft.com/en-us/library/bb215170).
        */
        XlPageBreak GetType();

        /**
        Sets a XlPageBreak value that represents the page break type.

        [MSDN documentation for HPageBreak.Type](http://msdn.microsoft.com/en-us/library/bb214627).
        [MSDN documentation for VPageBreak.Type](http://msdn.microsoft.com/en-us/library/bb215170).
        */
        void SetType(XlPageBreak type);

        /**
        Returns "PageBreak".
        */
        virtual wxString GetAutoExcelObjectName_() const { return wxS("PageBreak"); }
    };


    /**
    @brief Represents Microsoft Excel HPageBreaks or VPageBreaks collection.
    */
   class WXDLLIMPEXP_WXAUTOEXCEL wxExcelPageBreaks : public wxExcelObject
    {
    public:        
        // ***** METHODS *****

        /**
        Adds a horizontal page break.

        [MSDN documentation for HPageBreaks.Add](http://msdn.microsoft.com/en-us/library/bb211808).
        [MSDN documentation for VPageBreaks.Add](http://msdn.microsoft.com/en-us/library/bb179103).
        */
        wxExcelPageBreak Add(wxExcelRange before);

        // ***** PROPERTIES *****


        /**
        Returns a Long value that represents the number of objects in the collection.

        [MSDN documentation for HPageBreaks.Count](http://msdn.microsoft.com/en-us/library/bb148514).
        [MSDN documentation for VPageBreaks.Count](http://msdn.microsoft.com/en-us/library/bb214191).
        */
        long GetCount();

        //@{
        /**
        Returns a single object from a collection.

        [MSDN documentation for HPageBreaks.Item](http://msdn.microsoft.com/en-us/library/bb148516).
        [MSDN documentation for VPageBreaks.Item](http://msdn.microsoft.com/en-us/library/bb214196).
        */
        wxExcelPageBreak GetItem(long index);
        wxExcelPageBreak operator[](long index);
        //@}

        /**
        Returns "PageBreaks".
        */
        virtual wxString GetAutoExcelObjectName_() const { return wxS("PageBreaks"); }
    };


} // namespace wxAutoExcel

#endif //_WXAUTOEXCEL_PAGEBREAKS_H
