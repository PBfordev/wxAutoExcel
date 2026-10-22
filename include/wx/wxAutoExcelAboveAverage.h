/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// License:     MIT license
/////////////////////////////////////////////////////////////////////////////


#ifndef _WXAUTOEXCEL_ABOVEAVERAGE_H
#define _WXAUTOEXCEL_ABOVEAVERAGE_H

#include "wx/wxAutoExcel_defs.h"

#if WXAUTOEXCEL_USE_CONDFORMAT

#include "wx/wxAutoExcel_object.h"
#include "wx/wxAutoExcel_enums.h"

namespace wxAutoExcel {
    /**
    @brief Represents a Microsoft Excel AboveAverage object.
    */
   class WXDLLIMPEXP_WXAUTOEXCEL wxExcelAboveAverage : public wxExcelObject
    {
    public:
        // ***** METHODS *****

        /**
        Deletes the specified conditional formatting rule object.

        [MSDN documentation for AboveAverage.Delete](https://learn.microsoft.com/en-us/office/vba/api/excel.aboveaverage.delete).
        */
        void Delete();

        /**
        Sets the cell range to which this formatting rule applies.

        [MSDN documentation for AboveAverage.ModifyAppliesToRange](https://learn.microsoft.com/en-us/office/vba/api/excel.aboveaverage.modifyappliestorange).
        */
        void ModifyAppliesToRange(wxExcelRange range);

        /**
        Sets the priority value for this conditional formatting rule to "1" so that it will be evaluated before all other rules on the worksheet.

        [MSDN documentation for AboveAverage.SetFirstPriority](https://learn.microsoft.com/en-us/office/vba/api/excel.aboveaverage.setfirstpriority).
        */
        void SetFirstPriority();

        /**
        Sets the evaluation order for this conditional formatting rule so it is evaluated after all other rules on the worksheet.

        [MSDN documentation for AboveAverage.SetLastPriority](https://learn.microsoft.com/en-us/office/vba/api/excel.aboveaverage.setlastpriority).
        */
        void SetLastPriority();

        // ***** PROPERTIES *****

        /**
        Returns one of the constants of the XlAboveBelow enumeration specifying if the conditional formatting rule looks for cell values above or below the range average. Since Excel 2007.

        [MSDN documentation for AboveAverage.AboveBelow](https://learn.microsoft.com/en-us/office/vba/api/excel.aboveaverage.abovebelow).
        */
        XlAboveBelow GetAboveBelow();

        /**
        Sets one of the constants of the XlAboveBelow enumeration specifying if the conditional formatting rule looks for cell values above or below the range average. Since Excel 2007.

        [MSDN documentation for AboveAverage.AboveBelow](https://learn.microsoft.com/en-us/office/vba/api/excel.aboveaverage.abovebelow).
        */
        void SetAboveBelow(XlAboveBelow aboveBelow);

        /**
        Returns a Range object specifying the cell range to which the formatting rule is applied. Since Excel 2007.

        [MSDN documentation for AboveAverage.AppliesTo](https://learn.microsoft.com/en-us/office/vba/api/excel.aboveaverage.appliesto).
        */
        wxExcelRange GetAppliesTo();

        /**
        Returns a Borders collection that specifies the formatting of cell borders if the conditional formatting rule evaluates to True. Since Excel 2007.

        [MSDN documentation for AboveAverage.Borders](https://learn.microsoft.com/en-us/office/vba/api/excel.aboveaverage.borders).
        */
        wxExcelBorders GetBorders();

        /**
        Returns one of the constants of XlCalcFor enumeration which specifies the scope of data to be evaluated for the conditional format in a PivotTable report. Since Excel 2007.

        [MSDN documentation for AboveAverage.CalcFor](https://learn.microsoft.com/en-us/office/vba/api/excel.aboveaverage.calcfor).
        */
        XlCalcFor GetCalcFor();

        /**
        Sets one of the constants of XlCalcFor enumeration which specifies the scope of data to be evaluated for the conditional format in a PivotTable report. Since Excel 2007.

        [MSDN documentation for AboveAverage.CalcFor](https://learn.microsoft.com/en-us/office/vba/api/excel.aboveaverage.calcfor).
        */
        void SetCalcFor(XlCalcFor calcFor);

        /**
        Returns a Font object that specifies the font formatting if the conditional formatting rule evaluates to True. Since Excel 2007.

        [MSDN documentation for AboveAverage.Font](https://learn.microsoft.com/en-us/office/vba/api/excel.aboveaverage.font).
        */
        wxExcelFont GetFont();

        /**
        Returns an Interior object that specifies a cell's interior attributes for a conditional formatting rule that evaluates to True. Since Excel 2007.

        [MSDN documentation for AboveAverage.Interior](https://learn.microsoft.com/en-us/office/vba/api/excel.aboveaverage.interior).
        */
        wxExcelInterior GetInterior();

        /**
        Returns the number format applied to a cell if the conditional formatting rule evaluates to True. Read/write Variant. Since Excel 2007.

        The name is spelled Numberformat rather than NumberFormat as in Excel, because the Windows headers define GetNumberFormat as a macro, which would otherwise rename the method.

        [MSDN documentation for AboveAverage.NumberFormat](https://learn.microsoft.com/en-us/office/vba/api/excel.aboveaverage.numberformat).
        */
        wxString GetNumberformat();

        /**
        Sets the number format applied to a cell if the conditional formatting rule evaluates to True. Read/write Variant. Since Excel 2007.

        The name is spelled Numberformat rather than NumberFormat as in Excel, because the Windows headers define GetNumberFormat as a macro, which would otherwise rename the method.

        [MSDN documentation for AboveAverage.NumberFormat](https://learn.microsoft.com/en-us/office/vba/api/excel.aboveaverage.numberformat).
        */
        void SetNumberformat(const wxString& numberFormat);

        /**
        Returns the numeric standard deviation for an AboveAverage object.

        [MSDN documentation for AboveAverage.NumStdDev](https://learn.microsoft.com/en-us/office/vba/api/excel.aboveaverage.numstddev).
        */
        long GetNumStDev();

        /**
        Sets the numeric standard deviation for an AboveAverage object.

        [MSDN documentation for AboveAverage.NumStdDev](https://learn.microsoft.com/en-us/office/vba/api/excel.aboveaverage.numstddev).
        */
        void SetNumStDev(long numStDev);

        /**
        Returns the priority value of the conditional formatting rule. The priority determines the order of evaluation when multiple conditional formatting rules exist in a worksheet. Since Excel 2007.

        [MSDN documentation for AboveAverage.Priority](https://learn.microsoft.com/en-us/office/vba/api/excel.aboveaverage.priority).
        */
        long GetPriority();

        /**
        Sets the priority value of the conditional formatting rule. The priority determines the order of evaluation when multiple conditional formatting rules exist in a worksheet. Since Excel 2007.

        [MSDN documentation for AboveAverage.Priority](https://learn.microsoft.com/en-us/office/vba/api/excel.aboveaverage.priority).
        */
        void SetPriority(long priority);

        /**
        Returns a Boolean value indicating if the conditional format is being applied to a PivotTable chart. Since Excel 2007.

        [MSDN documentation for AboveAverage.PTCondition](https://learn.microsoft.com/en-us/office/vba/api/excel.aboveaverage.ptcondition).
        */
        bool GetPTCondition();

        /**
        Returns one of the constants of the XlPivotConditionScope enumeration, which determines the scope of the conditional format when it is applied to a PivotTable chart. Since Excel 2007.

        [MSDN documentation for AboveAverage.ScopeType](https://learn.microsoft.com/en-us/office/vba/api/excel.aboveaverage.scopetype).
        */
        XlPivotConditionScope GetScopeType();

        /**
        Sets one of the constants of the XlPivotConditionScope enumeration, which determines the scope of the conditional format when it is applied to a PivotTable chart. Since Excel 2007.

        [MSDN documentation for AboveAverage.ScopeType](https://learn.microsoft.com/en-us/office/vba/api/excel.aboveaverage.scopetype).
        */
        void SetScopeType(XlPivotConditionScope scopeType);

        /**
        Returns a Boolean value that determines if additional formatting rules on the cell should be evaluated if the current rule evaluates to True. Since Excel 2007.

        [MSDN documentation for AboveAverage.StopIfTrue](https://learn.microsoft.com/en-us/office/vba/api/excel.aboveaverage.stopiftrue).
        */
        bool GetStopIfTrue();

        /**
        Sets a Boolean value that determines if additional formatting rules on the cell should be evaluated if the current rule evaluates to True. Since Excel 2007.

        [MSDN documentation for AboveAverage.StopIfTrue](https://learn.microsoft.com/en-us/office/vba/api/excel.aboveaverage.stopiftrue).
        */
        void SetStopIfTrue(bool stopIfTrue);

        /**
        Returns one of the constants of the XlFormatConditionType enumeration, which specifies the type of conditional format. Since Excel 2007.

        [MSDN documentation for AboveAverage.Type](https://learn.microsoft.com/en-us/office/vba/api/excel.aboveaverage.type).
        */
        XlFormatConditionType GetType();

        /**
        Returns "AboveAverage".
        */
        virtual wxString GetAutoExcelObjectName_() const { return wxS("AboveAverage"); }
    };


} // namespace wxAutoExcel


#endif // WXAUTOEXCEL_USE_CONDFORMAT

#endif //_WXAUTOEXCEL_ABOVEAVERAGE_H

