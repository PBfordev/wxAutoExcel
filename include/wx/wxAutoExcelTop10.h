/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
/////////////////////////////////////////////////////////////////////////////


#ifndef _WXAUTOEXCEL_TOPTEN_H
#define _WXAUTOEXCEL_TOPTEN_H

#include "wx/wxAutoExcel_defs.h"

#if WXAUTOEXCEL_USE_CONDFORMAT

#include "wx/wxAutoExcel_object.h"
#include "wx/wxAutoExcel_enums.h"

namespace wxAutoExcel {
    /**
    @brief Represents a Microsoft Excel Top10 object.
    */
   class WXDLLIMPEXP_WXAUTOEXCEL wxExcelTop10 : public wxExcelObject
    {
    public:
        // ***** METHODS *****

        /**
        Deletes the specified conditional formatting rule object.

        [MSDN documentation for Top10.Delete](http://msdn.microsoft.com/en-us/library/bb210589.aspx).
        */
        void Delete();

        /**
        Sets the cell range to which this formatting rule applies.

        [MSDN documentation for Top10.ModifyAppliesToRange](http://msdn.microsoft.com/en-us/library/bb210593.aspx).
        */
        void ModifyAppliesToRange(wxExcelRange range);

        /**
        Sets the priority value for this conditional formatting rule to "1" so that it will be evaluated before all other rules on the worksheet.

        [MSDN documentation for Top10.SetFirstPriority](http://msdn.microsoft.com/en-us/library/bb210595.aspx).
        */
        void SetFirstPriority();

        /**
        Sets the evaluation order for this conditional formatting rule so it is evaluated after all other rules on the worksheet.

        [MSDN documentation for Top10.SetLastPriority](http://msdn.microsoft.com/en-us/library/bb210600.aspx).
        */
        void SetLastPriority();

        // ***** PROPERTIES *****

        /**
        Returns a Range object specifying the cell range to which the formatting rule is applied. Since Excel 2007.

        [MSDN documentation for Top10.AppliesTo](http://msdn.microsoft.com/en-us/library/bb148101.aspx).
        */
        wxExcelRange GetAppliesTo();

        /**
        Returns a Borders collection that specifies the formatting of cell borders if the conditional formatting rule evaluates to True. Since Excel 2007.

        [MSDN documentation for Top10.Borders](http://msdn.microsoft.com/en-us/library/bb148105.aspx).
        */
        wxExcelBorders GetBorders();

        /**
        Returns one of the constants of XlCalcFor enumeration which specifies the scope of data to be evaluated for the conditional format in a PivotTable report. Since Excel 2007.

        [MSDN documentation for Top10.CalcFor](http://msdn.microsoft.com/en-us/library/bb148109.aspx).
        */
        XlCalcFor GetCalcFor();

        /**
        Sets one of the constants of XlCalcFor enumeration which specifies the scope of data to be evaluated for the conditional format in a PivotTable report. Since Excel 2007.

        [MSDN documentation for Top10.CalcFor](http://msdn.microsoft.com/en-us/library/bb148109.aspx).
        */
        void SetCalcFor(XlCalcFor calcFor);

        /**
        Returns a Font object that specifies the font formatting if the conditional formatting rule evaluates to True. Since Excel 2007.

        [MSDN documentation for Top10.Font](http://msdn.microsoft.com/en-us/library/bb148112.aspx).
        */
        wxExcelFont GetFont();

        /**
        Returns an Interior object that specifies a cell's interior attributes for a conditional formatting rule that evaluates to True. Since Excel 2007.

        [MSDN documentation for Top10.Interior](http://msdn.microsoft.com/en-us/library/bb148120.aspx).
        */
        wxExcelInterior GetInterior();

        /**
        Returns the number format applied to a cell if the conditional formatting rule evaluates to True. Read/write Variant. Since Excel 2007.

        [MSDN documentation for Top10.NumberFormat](http://msdn.microsoft.com/en-us/library/bb148122.aspx).
        */
        wxString GetNumberFormat();

        /**
        Sets the number format applied to a cell if the conditional formatting rule evaluates to True. Read/write Variant. Since Excel 2007.

        [MSDN documentation for Top10.NumberFormat](http://msdn.microsoft.com/en-us/library/bb148122.aspx).
        */
        void SetNumberFormat(const wxString& numberFormat);

        /**
        Returns a Boolean value specifying if the rank is determined by a percentage value. Since Excel 2007.

        [MSDN documentation for Top10.Percent](http://msdn.microsoft.com/en-us/library/bb148126.aspx).
        */
        bool GetPercent();

        /**
        Sets a Boolean value specifying if the rank is determined by a percentage value. Since Excel 2007.

        [MSDN documentation for Top10.Percent](http://msdn.microsoft.com/en-us/library/bb148126.aspx).
        */
        void SetPercent(bool percent);

        /**
        Returns the priority value of the conditional formatting rule. The priority determines the order of evaluation when multiple conditional formatting rules exist in a worksheet. Since Excel 2007.

        [MSDN documentation for Top10.Priority](http://msdn.microsoft.com/en-us/library/bb148128.aspx).
        */
        long GetPriority();

        /**
        Sets the priority value of the conditional formatting rule. The priority determines the order of evaluation when multiple conditional formatting rules exist in a worksheet. Since Excel 2007.

        [MSDN documentation for Top10.Priority](http://msdn.microsoft.com/en-us/library/bb148128.aspx).
        */
        void SetPriority(long priority);

        /**
        Returns a Boolean value indicating if the conditional format is being applied to a PivotTable chart. Since Excel 2007.

        [MSDN documentation for Top10.PTCondition](http://msdn.microsoft.com/en-us/library/bb148132.aspx).
        */
        bool GetPTCondition();

        /**
        Returns a Long value specifying either the number or percentage of the rank value for the conditional format rule. Since Excel 2007.

        [MSDN documentation for Top10.Rank](http://msdn.microsoft.com/en-us/library/bb148137.aspx).
        */
        long GetRank();

        /**
        Sets a Long value specifying either the number or percentage of the rank value for the conditional format rule. Since Excel 2007.

        [MSDN documentation for Top10.Rank](http://msdn.microsoft.com/en-us/library/bb148137.aspx).
        */
        void SetRank(long rank);

        /**
        Returns one of the constants of the XlPivotConditionScope enumeration, which determines the scope of the conditional format when it is applied to a PivotTable chart. Since Excel 2007.

        [MSDN documentation for Top10.ScopeType](http://msdn.microsoft.com/en-us/library/bb148139.aspx).
        */
        XlPivotConditionScope GetScopeType();

        /**
        Sets one of the constants of the XlPivotConditionScope enumeration, which determines the scope of the conditional format when it is applied to a PivotTable chart. Since Excel 2007.

        [MSDN documentation for Top10.ScopeType](http://msdn.microsoft.com/en-us/library/bb148139.aspx).
        */
        void SetScopeType(XlPivotConditionScope scopeType);

        /**
        Returns a Boolean value that determines if additional formatting rules on the cell should be evaluated if the current rule evaluates to True. Since Excel 2007.

        [MSDN documentation for Top10.StopIfTrue](http://msdn.microsoft.com/en-us/library/bb148141.aspx).
        */
        bool GetStopIfTrue();

        /**
        Sets a Boolean value that determines if additional formatting rules on the cell should be evaluated if the current rule evaluates to True. Since Excel 2007.

        [MSDN documentation for Top10.StopIfTrue](http://msdn.microsoft.com/en-us/library/bb148141.aspx).
        */
        void SetStopIfTrue(bool stopIfTrue);

        /**
        Returns one of the constants of XlTopBottom enumeration determining if the ranking is evaluated from the top or bottom. Since Excel 2007.

        [MSDN documentation for Top10.TopBottom](http://msdn.microsoft.com/en-us/library/bb148143.aspx).
        */
        XlTopBottom GetTopBottom();

        /**
        Sets one of the constants of XlTopBottom enumeration determining if the ranking is evaluated from the top or bottom. Since Excel 2007.

        [MSDN documentation for Top10.TopBottom](http://msdn.microsoft.com/en-us/library/bb148143.aspx).
        */
        void SetTopBottom(XlTopBottom topBottom);

        /**
        Returns one of the constants of the XlFormatConditionType enumeration, which specifies the type of conditional format. Since Excel 2007.

        [MSDN documentation for Top10.Type](http://msdn.microsoft.com/en-us/library/bb148146.aspx).
        */
        XlFormatConditionType GetType();


        /**
        Returns "Top10".
        */
        virtual wxString GetAutoExcelObjectName_() const { return wxS("Top10"); }
    };


} // namespace wxAutoExcel

#endif // WXAUTOEXCEL_USE_CONDFORMAT

#endif //_WXAUTOEXCEL_TOPTEN_H
