/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// License:     MIT license
/////////////////////////////////////////////////////////////////////////////


#ifndef _WXAUTOEXCEL_TEXTCOLUMN2_H
#define _WXAUTOEXCEL_TEXTCOLUMN2_H

#include "wx/wxAutoExcel_defs.h"

#if WXAUTOEXCEL_USE_SHAPES || WXAUTOEXCEL_USE_CHARTS

#include "wx/wxAutoExcel_object.h"
#include "wx/wxAutoExcel_enums.h"

namespace wxAutoExcel {

    /**
    @brief Represents Microsoft Excel TextColumn2 object.
    */
   class WXDLLIMPEXP_WXAUTOEXCEL wxExcelTextColumn2: public wxExcelObject
    {
    public:
        // ***** PROPERTIES *****


        /**
        Gets or sets the index of the TextColumn2 object. Read/write

        [MSDN documentation for TextColumn2.Number](https://learn.microsoft.com/en-us/office/vba/api/office.textcolumn2.number).
        */
        long GetNumber();

        /**
        Gets or sets the index of the TextColumn2 object. Read/write

        [MSDN documentation for TextColumn2.Number](https://learn.microsoft.com/en-us/office/vba/api/office.textcolumn2.number).
        */
        void SetNumber(long number);

        /**
        Gets or sets the spacing between text columns in a TextColumn2 object.

        [MSDN documentation for TextColumn2.Spacing]().
        */
        double GetSpacing();

        /**
        Gets or sets the spacing between text columns in a TextColumn2 object.

        [MSDN documentation for TextColumn2.Spacing]().
        */
        void SetSpacing(double spacing);

        /**
        Gets or sets the direction of text in the TextColumn2 object.

        [MSDN documentation for TextColumn2.TextDirection](https://learn.microsoft.com/en-us/office/vba/api/office.textcolumn2.textdirection).
        */
        MsoTextDirection GetTextDirection();

        /**
        Gets or sets the direction of text in the TextColumn2 object.

        [MSDN documentation for TextColumn2.TextDirection](https://learn.microsoft.com/en-us/office/vba/api/office.textcolumn2.textdirection).
        */
        void SetTextDirection(MsoTextDirection textDirection);


        /**
        Returns "TextColumn2".
        */
        virtual wxString GetAutoExcelObjectName_() const { return wxS("TextColumn2"); }
    };


} // namespace wxAutoExcel

#endif // #if WXAUTOEXCEL_USE_SHAPES || WXAUTOEXCEL_USE_CHARTS

#endif //_WXAUTOEXCEL_TEXTCOLUMN2_H
