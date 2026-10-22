/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// License:     MIT license
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

        [MSDN documentation for Top10.Delete](https://learn.microsoft.com/en-us/office/vba/api/excel.top10.delete).
        */
        void Delete();

        /**
        Sets the cell range to which this formatting rule applies.

        [MSDN documentation for Top10.ModifyAppliesToRange](https://learn.microsoft.com/en-us/office/vba/api/excel.top10.modifyappliestorange).
        */
        void ModifyAppliesToRange(wxExcelRange range);

        /**
        Sets the priority value for this conditional formatting rule to "1" so that it will be evaluated before all other rules on the worksheet.

        [MSDN documentation for Top10.SetFirstPriority](https://learn.microsoft.com/en-us/office/vba/api/excel.top10.setfirstpriority).
        */
        void SetFirstPriority();

        /**
        Sets the evaluation order for this conditional formatting rule so it is evaluated after all other rules on the worksheet.

        [MSDN documentation for Top10.SetLastPriority](https://learn.microsoft.com/en-us/office/vba/api/excel.top10.setlastpriority).
        */
        void SetLastPriority();

        // ***** PROPERTIES *****

        /**
        Returns a Range object specifying the cell range to which the formatting rule is applied. Since Excel 2007.

        [MSDN documentation for Top10.AppliesTo](https://learn.microsoft.com/en-us/office/vba/api/excel.top10.appliesto).
        */
        wxExcelRange GetAppliesTo();

        /**
        Returns a Borders collection that specifies the formatting of cell borders if the conditional formatting rule evaluates to True. Since Excel 2007.

        [MSDN documentation for Top10.Borders](https://learn.microsoft.com/en-us/office/vba/api/excel.top10.borders).
        */
        wxExcelBorders GetBorders();

        /**
        Returns one of the constants of XlCalcFor enumeration which specifies the scope of data to be evaluated for the conditional format in a PivotTable report. Since Excel 2007.

        [MSDN documentation for Top10.CalcFor](https://learn.microsoft.com/en-us/office/vba/api/excel.top10.calcfor).
        */
        XlCalcFor GetCalcFor();

        /**
        Sets one of the constants of XlCalcFor enumeration which specifies the scope of data to be evaluated for the conditional format in a PivotTable report. Since Excel 2007.

        [MSDN documentation for Top10.CalcFor](https://learn.microsoft.com/en-us/office/vba/api/excel.top10.calcfor).
        */
        void SetCalcFor(XlCalcFor calcFor);

        /**
        Returns a Font object that specifies the font formatting if the conditional formatting rule evaluates to True. Since Excel 2007.

        [MSDN documentation for Top10.Font](https://learn.microsoft.com/en-us/office/vba/api/excel.top10.font).
        */
        wxExcelFont GetFont();

        /**
        Returns an Interior object that specifies a cell's interior attributes for a conditional formatting rule that evaluates to True. Since Excel 2007.

        [MSDN documentation for Top10.Interior](https://learn.microsoft.com/en-us/office/vba/api/excel.top10.interior).
        */
        wxExcelInterior GetInterior();

        /**
        Returns the number format applied to a cell if the conditional formatting rule evaluates to True. Read/write Variant. Since Excel 2007.

        The name is spelled Numberformat rather than NumberFormat as in Excel, because the Windows headers define GetNumberFormat as a macro, which would otherwise rename the method.

        [MSDN documentation for Top10.NumberFormat](https://learn.microsoft.com/en-us/office/vba/api/excel.top10.numberformat).
        */
        wxString GetNumberformat();

        /**
        Sets the number format applied to a cell if the conditional formatting rule evaluates to True. Read/write Variant. Since Excel 2007.

        The name is spelled Numberformat rather than NumberFormat as in Excel, because the Windows headers define GetNumberFormat as a macro, which would otherwise rename the method.

        [MSDN documentation for Top10.NumberFormat](https://learn.microsoft.com/en-us/office/vba/api/excel.top10.numberformat).
        */
        void SetNumberformat(const wxString& numberFormat);

        /**
        Returns a Boolean value specifying if the rank is determined by a percentage value. Since Excel 2007.

        [MSDN documentation for Top10.Percent](https://learn.microsoft.com/en-us/office/vba/api/excel.top10.percent).
        */
        bool GetPercent();

        /**
        Sets a Boolean value specifying if the rank is determined by a percentage value. Since Excel 2007.

        [MSDN documentation for Top10.Percent](https://learn.microsoft.com/en-us/office/vba/api/excel.top10.percent).
        */
        void SetPercent(bool percent);

        /**
        Returns the priority value of the conditional formatting rule. The priority determines the order of evaluation when multiple conditional formatting rules exist in a worksheet. Since Excel 2007.

        [MSDN documentation for Top10.Priority](https://learn.microsoft.com/en-us/office/vba/api/excel.top10.priority).
        */
        long GetPriority();

        /**
        Sets the priority value of the conditional formatting rule. The priority determines the order of evaluation when multiple conditional formatting rules exist in a worksheet. Since Excel 2007.

        [MSDN documentation for Top10.Priority](https://learn.microsoft.com/en-us/office/vba/api/excel.top10.priority).
        */
        void SetPriority(long priority);

        /**
        Returns a Boolean value indicating if the conditional format is being applied to a PivotTable chart. Since Excel 2007.

        [MSDN documentation for Top10.PTCondition](https://learn.microsoft.com/en-us/office/vba/api/excel.top10.ptcondition).
        */
        bool GetPTCondition();

        /**
        Returns a Long value specifying either the number or percentage of the rank value for the conditional format rule. Since Excel 2007.

        [MSDN documentation for Top10.Rank](https://learn.microsoft.com/en-us/office/vba/api/excel.top10.rank).
        */
        long GetRank();

        /**
        Sets a Long value specifying either the number or percentage of the rank value for the conditional format rule. Since Excel 2007.

        [MSDN documentation for Top10.Rank](https://learn.microsoft.com/en-us/office/vba/api/excel.top10.rank).
        */
        void SetRank(long rank);

        /**
        Returns one of the constants of the XlPivotConditionScope enumeration, which determines the scope of the conditional format when it is applied to a PivotTable chart. Since Excel 2007.

        [MSDN documentation for Top10.ScopeType](https://learn.microsoft.com/en-us/office/vba/api/excel.top10.scopetype).
        */
        XlPivotConditionScope GetScopeType();

        /**
        Sets one of the constants of the XlPivotConditionScope enumeration, which determines the scope of the conditional format when it is applied to a PivotTable chart. Since Excel 2007.

        [MSDN documentation for Top10.ScopeType](https://learn.microsoft.com/en-us/office/vba/api/excel.top10.scopetype).
        */
        void SetScopeType(XlPivotConditionScope scopeType);

        /**
        Returns a Boolean value that determines if additional formatting rules on the cell should be evaluated if the current rule evaluates to True. Since Excel 2007.

        [MSDN documentation for Top10.StopIfTrue](https://learn.microsoft.com/en-us/office/vba/api/excel.top10.stopiftrue).
        */
        bool GetStopIfTrue();

        /**
        Sets a Boolean value that determines if additional formatting rules on the cell should be evaluated if the current rule evaluates to True. Since Excel 2007.

        [MSDN documentation for Top10.StopIfTrue](https://learn.microsoft.com/en-us/office/vba/api/excel.top10.stopiftrue).
        */
        void SetStopIfTrue(bool stopIfTrue);

        /**
        Returns one of the constants of XlTopBottom enumeration determining if the ranking is evaluated from the top or bottom. Since Excel 2007.

        [MSDN documentation for Top10.TopBottom](https://learn.microsoft.com/en-us/office/vba/api/excel.top10.topbottom).
        */
        XlTopBottom GetTopBottom();

        /**
        Sets one of the constants of XlTopBottom enumeration determining if the ranking is evaluated from the top or bottom. Since Excel 2007.

        [MSDN documentation for Top10.TopBottom](https://learn.microsoft.com/en-us/office/vba/api/excel.top10.topbottom).
        */
        void SetTopBottom(XlTopBottom topBottom);

        /**
        Returns one of the constants of the XlFormatConditionType enumeration, which specifies the type of conditional format. Since Excel 2007.

        [MSDN documentation for Top10.Type](https://learn.microsoft.com/en-us/office/vba/api/excel.top10.type).
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
