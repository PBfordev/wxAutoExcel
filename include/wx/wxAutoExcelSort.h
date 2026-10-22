/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// License:     MIT license
/////////////////////////////////////////////////////////////////////////////


#ifndef _WXAUTOEXCEL_SORT_H
#define _WXAUTOEXCEL_SORT_H

#include "wx/wxAutoExcel_defs.h"

#include "wx/wxAutoExcel_object.h"
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

        [MSDN documentation for Sort.Apply](https://learn.microsoft.com/en-us/office/vba/api/excel.sort.apply).
        */
        void Apply();

        /**
        Sets the starting and ending character positions for Sort object.

        [MSDN documentation for Sort.SetRange](https://learn.microsoft.com/en-us/office/vba/api/excel.sort.setrange).
        */
        void SetRange(wxExcelRange rng);

        // ***** PROPERTIES *****

        /**
        Specifies whether the first row contains header information. Read/write XlYesNoGuess. Since Excel 2007.

        [MSDN documentation for Sort.Header](https://learn.microsoft.com/en-us/office/vba/api/excel.sort.header).
        */
        XlYesNoGuess GetHeader();

        /**
        Specifies whether the first row contains header information. Read/write XlYesNoGuess. Since Excel 2007.

        [MSDN documentation for Sort.Header](https://learn.microsoft.com/en-us/office/vba/api/excel.sort.header).
        */
        void SetHeader(XlYesNoGuess header);

        /**
        Set to True to perform a case-sensitive sort or set to False to perform non-case sensitive sort. Since Excel 2007.

        [MSDN documentation for Sort.MatchCase](https://learn.microsoft.com/en-us/office/vba/api/excel.sort.matchcase).
        */
        bool GetMatchCase();

        /**
        Set to True to perform a case-sensitive sort or set to False to perform non-case sensitive sort. Since Excel 2007.

        [MSDN documentation for Sort.MatchCase](https://learn.microsoft.com/en-us/office/vba/api/excel.sort.matchcase).
        */
        void SetMatchCase(bool matchCase);

        /**
        Specifies the orientation for the sort. Read/write XlSortOrientation. Since Excel 2007.

        [MSDN documentation for Sort.Orientation](https://learn.microsoft.com/en-us/office/vba/api/excel.sort.orientation).
        */
        XlSortOrientation GetOrientation();

        /**
        Specifies the orientation for the sort. Read/write XlSortOrientation. Since Excel 2007.

        [MSDN documentation for Sort.Orientation](https://learn.microsoft.com/en-us/office/vba/api/excel.sort.orientation).
        */
        void SetOrientation(XlSortOrientation orientation);

        /**
        Return the range of values on which the sort is performed. Since Excel 2007.

        [MSDN documentation for Sort.Rng](https://learn.microsoft.com/en-us/office/vba/api/excel.sort.rng).
        */
        wxExcelRange GetRng();

        /**
        Stores the sort state for workbooks, lists, and autofilters. Since Excel 2007.

        [MSDN documentation for Sort.SortFields](https://learn.microsoft.com/en-us/office/vba/api/excel.sort.sortfields).
        */
        wxExcelSortFields GetSortFields();

        /**
        Specifies the sort method for Chinese languages. Read/write XlSortMethod. Since Excel 2007.

        [MSDN documentation for Sort.SortMethod](https://learn.microsoft.com/en-us/office/vba/api/excel.sort.sortmethod).
        */
        XlSortMethod GetSortMethod();

        /**
        Returns "Sort".
        */
        virtual wxString GetAutoExcelObjectName_() const { return wxS("Sort"); }
    };


} // namespace wxAutoExcel

#endif //_WXAUTOEXCEL_SORT_H
