/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Modified by:
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
/////////////////////////////////////////////////////////////////////////////


#ifndef _WXAUTOEXCEL_SORT_H
#define _WXAUTOEXCEL_SORT_H

#include "wx/wxAutoExcel_defs.h"

#include "wx/wxAutoExcelObject.h"
#include "wx/wxAutoExcel_enums.h"

namespace wxAutoExcel {

    /**
    @brief Represents Microsoft Excel Sort object.
    */
   class WXDLLIMPEXP_WXAUTOEXCEL wxExcelSort : public wxExcelObject
    {
    public:        
        // ***** METHODS *****

        /**
        Sorts the range based on the currently applied sort states.

        [MSDN documentation for Sort.Apply](http://msdn.microsoft.com/en-us/library/bb210537).
        */
        void Apply();

        /**
        Sets the starting and ending character positions for Sort object.

        [MSDN documentation for Sort.SetRange](http://msdn.microsoft.com/en-us/library/bb210561).
        */
        void SetRange(wxExcelRange rng);

        // ***** PROPERTIES *****

        /**
        Specifies whether the first row contains header information. Read/write XlYesNoGuess. Since Excel 2007.

        [MSDN documentation for Sort.Header](http://msdn.microsoft.com/en-us/library/bb148034).
        */
        XlYesNoGuess GetHeader();

        /**
        Specifies whether the first row contains header information. Read/write XlYesNoGuess. Since Excel 2007.

        [MSDN documentation for Sort.Header](http://msdn.microsoft.com/en-us/library/bb148034).
        */
        void SetHeader(XlYesNoGuess header);

        /**
        Set to True to perform a case-sensitive sort or set to False to perform non-case sensitive sort. Since Excel 2007.

        [MSDN documentation for Sort.MatchCase](http://msdn.microsoft.com/en-us/library/bb148036).
        */
        bool GetMatchCase();

        /**
        Set to True to perform a case-sensitive sort or set to False to perform non-case sensitive sort. Since Excel 2007.

        [MSDN documentation for Sort.MatchCase](http://msdn.microsoft.com/en-us/library/bb148036).
        */
        void SetMatchCase(bool matchCase);

        /**
        Specifies the orientation for the sort. Read/write XlSortOrientation. Since Excel 2007.

        [MSDN documentation for Sort.Orientation](http://msdn.microsoft.com/en-us/library/bb148040).
        */
        XlSortOrientation GetOrientation();

        /**
        Specifies the orientation for the sort. Read/write XlSortOrientation. Since Excel 2007.

        [MSDN documentation for Sort.Orientation](http://msdn.microsoft.com/en-us/library/bb148040).
        */
        void SetOrientation(XlSortOrientation orientation);

        /**
        Return the range of values on which the sort is performed. Since Excel 2007.

        [MSDN documentation for Sort.Rng](http://msdn.microsoft.com/en-us/library/bb148043).
        */
        wxExcelRange GetRng();

        /**
        Stores the sort state for workbooks, lists, and autofilters. Since Excel 2007.

        [MSDN documentation for Sort.SortFields](http://msdn.microsoft.com/en-us/library/bb148046).
        */
        wxExcelSortFields GetSortFields();

        /**
        Specifies the sort method for Chinese languages. Read/write XlSortMethod. Since Excel 2007.

        [MSDN documentation for Sort.SortMethod](http://msdn.microsoft.com/en-us/library/bb148054).
        */
        XlSortMethod GetSortMethod();

        /**
        Returns "Sort".
        */
        virtual wxString GetAutoExcelObjectName_() const { return wxS("Sort"); }
    };


} // namespace wxAutoExcel

#endif //_WXAUTOEXCEL_SORT_H
