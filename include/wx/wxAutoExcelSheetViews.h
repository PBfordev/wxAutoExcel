/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// License:     MIT license
/////////////////////////////////////////////////////////////////////////////


#ifndef _WXAUTOEXCEL_SHEETSVIEW_H
#define _WXAUTOEXCEL_SHEETSVIEW_H

#include "wx/wxAutoExcel_defs.h"

#include "wx/wxAutoExcel_object.h"


namespace wxAutoExcel {

    /**
    @brief Represents Microsoft Excel SheetView object.
    */
   class WXDLLIMPEXP_WXAUTOEXCEL wxExcelSheetView : public wxExcelObject
    {
    public:
        enum View { ChartView, DialogSheetView, WorksheetView };

        // ***** PROPERTIES *****

        /**
        Returns true if the SheetView is of a view type.
        */
        bool IsView(View view) const;

#if WXAUTOEXCEL_USE_CHARTS
        /**
            If possible, returns the SheetView as a ChartView (the SheetView object is still preserved)
        */
        wxExcelChartView ToChartView();
#endif // #if WXAUTOEXCEL_USE_CHARTS

        /**
            If possible, returns the SheetView as a WorksheetView (the SheetView object is still preserved)
        */
        wxExcelWorksheetView ToWorksheetView();

        /**
        Returns "SheetView".
        */
        virtual wxString GetAutoExcelObjectName_() const { return wxS("SheetView"); }

    };



    /**
    @brief Represents Microsoft Excel SheetViews collection.
    */
   class WXDLLIMPEXP_WXAUTOEXCEL wxExcelSheetViews : public wxExcelObject
    {
    public:

        // ***** PROPERTIES *****

        /**
        The number of objects in the collection. Since Excel 2007.

        [MSDN documentation for SheetViews.Count](https://learn.microsoft.com/en-us/office/vba/api/excel.sheetviews.count).
        */
        long GetCount();

        /**
        Returns a SheetView object that represents views in a workbook. Since Excel 2007.

        [MSDN documentation for SheetViews.Item](https://learn.microsoft.com/en-us/office/vba/api/excel.sheetviews.item).
        */
        wxExcelSheetView GetItem(long index);

        /**
        Returns "SheetViews".
        */
        virtual wxString GetAutoExcelObjectName_() const { return wxS("SheetViews"); }

    };


} // namespace wxAutoExcel

#endif // _WXAUTOEXCEL_SHEETSVIEW_H
