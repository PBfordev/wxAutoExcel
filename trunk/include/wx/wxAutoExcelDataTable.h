/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Modified by:
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
/////////////////////////////////////////////////////////////////////////////


#ifndef _WXAUTOEXCEL_DATATABLE_H
#define _WXAUTOEXCEL_DATATABLE_H

#include "wx/wxAutoExcel_defs.h"

#if WXAUTOEXCEL_USE_CHARTS

#include "wx/wxAutoExcel_object.h"

namespace wxAutoExcel {

    /**
    Represents Microsoft Excel DataTable object.
    */
   class WXDLLIMPEXP_WXAUTOEXCEL wxExcelDataTable : public wxExcelObject    
   {
   public:
        // ***** METHODS *****

        /**
        Deletes the object.

        [MSDN documentation for DataTable.Delete](http://msdn.microsoft.com/en-us/library/office/ff837424.aspx).
        */
        void Delete();

        /**
        Selects the object.

        [MSDN documentation for DataTable.Select](http://msdn.microsoft.com/en-us/library/office/ff836757.aspx).
        */
        void Select();

        // ***** PROPERTIES *****

        /**
        Returns a Border object that represents the border of the object.

        [MSDN documentation for DataTable.Border]().
        */
        wxExcelBorder GetBorder();

        /**
        Returns a Font object that represents the font of the specified object.

        [MSDN documentation for DataTable.Font](http://msdn.microsoft.com/en-us/library/office/ff839457(v=office.14).aspx).
        */
        wxExcelFont GetFont();

        /**
        Returns the ChartFormat object. Since Excel 2007.

        [MSDN documentation for DataTable.Format](http://msdn.microsoft.com/en-us/library/office/ff195438(v=office.14).aspx).
        */
        wxExcelChartFormat GetFormat();

        /**
        True if the chart data table has horizontal cell borders.

        [MSDN documentation for DataTable.HasBorderHorizontal](http://msdn.microsoft.com/en-us/library/office/ff836793(v=office.14).aspx).
        */
        bool GetHasBorderHorizontal();

        /**
        True if the chart data table has horizontal cell borders.

        [MSDN documentation for DataTable.HasBorderHorizontal](http://msdn.microsoft.com/en-us/library/office/ff836793(v=office.14).aspx).
        */
        void SetHasBorderHorizontal(bool hasBorderHorizontal);

        /**
        True if the chart data table has outline borders.

        [MSDN documentation for DataTable.HasBorderOutline](http://msdn.microsoft.com/en-us/library/office/ff840438(v=office.14).aspx).
        */
        bool GetHasBorderOutline();

        /**
        True if the chart data table has outline borders.

        [MSDN documentation for DataTable.HasBorderOutline](http://msdn.microsoft.com/en-us/library/office/ff840438(v=office.14).aspx).
        */
        void SetHasBorderOutline(bool hasBorderOutline);

        /**
        True if the chart data table has vertical cell borders.

        [MSDN documentation for DataTable.HasBorderVertical](http://msdn.microsoft.com/en-us/library/office/ff194369(v=office.14).aspx).
        */
        bool GetHasBorderVertical();

        /**
        True if the chart data table has vertical cell borders.

        [MSDN documentation for DataTable.HasBorderVertical](http://msdn.microsoft.com/en-us/library/office/ff194369(v=office.14).aspx).
        */
        void SetHasBorderVertical(bool hasBorderVertical);

        /**
        True if the data label legend key is visible.

        [MSDN documentation for DataTable.ShowLegendKey](http://msdn.microsoft.com/en-us/library/office/ff840154(v=office.14).aspx).
        */
        bool GetShowLegendKey();

        /**
        True if the data label legend key is visible.

        [MSDN documentation for DataTable.ShowLegendKey](http://msdn.microsoft.com/en-us/library/office/ff840154(v=office.14).aspx).
        */
        void SetShowLegendKey(bool showLegendKey);

        /**
        Returns "DataTable".
        */
        virtual wxString GetAutoExcelObjectName_() const { return wxS("DataTable"); }
    };


} // namespace wxAutoExcel

#endif // #if WXAUTOEXCEL_USE_CHARTS

#endif // _WXAUTOEXCEL_DATATABLE_H
