/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Modified by:
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
/////////////////////////////////////////////////////////////////////////////


#ifndef _WXAUTOEXCEL_ABOVEAVERAGE_H
#define _WXAUTOEXCEL_ABOVEAVERAGE_H

#include "wx/wxAutoExcel_defs.h"

#if WXAUTOEXCEL_USE_CONDFORMAT

#include "wx/wxAutoExcelObject.h"
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

        [MSDN documentation for AboveAverage.Delete](http://msdn.microsoft.com/en-us/library/bb178556.aspx).
        */
        void Delete();

        /**
        Sets the cell range to which this formatting rule applies.

        [MSDN documentation for AboveAverage.ModifyAppliesToRange](http://msdn.microsoft.com/en-us/library/bb224180.aspx).
        */
        void ModifyAppliesToRange(wxExcelRange range);

        /**
        Sets the priority value for this conditional formatting rule to "1" so that it will be evaluated before all other rules on the worksheet.

        [MSDN documentation for AboveAverage.SetFirstPriority](http://msdn.microsoft.com/en-us/library/bb178558.aspx).
        */
        void SetFirstPriority();

        /**
        Sets the evaluation order for this conditional formatting rule so it is evaluated after all other rules on the worksheet.

        [MSDN documentation for AboveAverage.SetLastPriority](http://msdn.microsoft.com/en-us/library/bb178563.aspx).
        */
        void SetLastPriority();

        // ***** PROPERTIES *****

        /**
        Returns one of the constants of the XlAboveBelow enumeration specifying if the conditional formatting rule looks for cell values above or below the range average. Since Excel 2007.

        [MSDN documentation for AboveAverage.AboveBelow](http://msdn.microsoft.com/en-us/library/bb210618.aspx).
        */
        XlAboveBelow GetAboveBelow();

        /**
        Sets one of the constants of the XlAboveBelow enumeration specifying if the conditional formatting rule looks for cell values above or below the range average. Since Excel 2007.

        [MSDN documentation for AboveAverage.AboveBelow](http://msdn.microsoft.com/en-us/library/bb210618.aspx).
        */
        void SetAboveBelow(XlAboveBelow aboveBelow);

        /**
        Returns a Range object specifying the cell range to which the formatting rule is applied. Since Excel 2007.

        [MSDN documentation for AboveAverage.AppliesTo](http://msdn.microsoft.com/en-us/library/bb210620.aspx).
        */
        wxExcelRange GetAppliesTo();

        /**
        Returns a Borders collection that specifies the formatting of cell borders if the conditional formatting rule evaluates to True. Since Excel 2007.

        [MSDN documentation for AboveAverage.Borders](http://msdn.microsoft.com/en-us/library/bb210621.aspx).
        */
        wxExcelBorders GetBorders();

        /**
        Returns one of the constants of XlCalcFor enumeration which specifies the scope of data to be evaluated for the conditional format in a PivotTable report. Since Excel 2007.

        [MSDN documentation for AboveAverage.CalcFor](http://msdn.microsoft.com/en-us/library/bb256339.aspx).
        */
        XlCalcFor GetCalcFor();

        /**
        Sets one of the constants of XlCalcFor enumeration which specifies the scope of data to be evaluated for the conditional format in a PivotTable report. Since Excel 2007.

        [MSDN documentation for AboveAverage.CalcFor](http://msdn.microsoft.com/en-us/library/bb256339.aspx).
        */
        void SetCalcFor(XlCalcFor calcFor);

        /**
        Returns a Font object that specifies the font formatting if the conditional formatting rule evaluates to True. Since Excel 2007.

        [MSDN documentation for AboveAverage.Font](http://msdn.microsoft.com/en-us/library/bb210623.aspx).
        */
        wxExcelFont GetFont();

        /**
        Returns an Interior object that specifies a cell's interior attributes for a conditional formatting rule that evaluates to True. Since Excel 2007.

        [MSDN documentation for AboveAverage.Interior](http://msdn.microsoft.com/en-us/library/bb224177.aspx).
        */
        wxExcelInterior GetInterior();

        /**
        Returns the number format applied to a cell if the conditional formatting rule evaluates to True. Read/write Variant. Since Excel 2007.

        [MSDN documentation for AboveAverage.NumberFormat](http://msdn.microsoft.com/en-us/library/bb210740.aspx).
        */
        wxString GetNumberFormat();

        /**
        Sets the number format applied to a cell if the conditional formatting rule evaluates to True. Read/write Variant. Since Excel 2007.

        [MSDN documentation for AboveAverage.NumberFormat](http://msdn.microsoft.com/en-us/library/bb210740.aspx).
        */
        void SetNumberFormat(const wxString& numberFormat);

        /**
        Returns the numeric standard deviation for an AboveAverage object.

        [MSDN documentation for AboveAverage.NumStDev]().
        */
        long GetNumStDev();

        /**
        Sets the numeric standard deviation for an AboveAverage object.

        [MSDN documentation for AboveAverage.NumStDev]().
        */
        void SetNumStDev(long numStDev);

        /**
        Returns the priority value of the conditional formatting rule. The priority determines the order of evaluation when multiple conditional formatting rules exist in a worksheet. Since Excel 2007.

        [MSDN documentation for AboveAverage.Priority](http://msdn.microsoft.com/en-us/library/bb210744.aspx).
        */
        long GetPriority();

        /**
        Sets the priority value of the conditional formatting rule. The priority determines the order of evaluation when multiple conditional formatting rules exist in a worksheet. Since Excel 2007.

        [MSDN documentation for AboveAverage.Priority](http://msdn.microsoft.com/en-us/library/bb210744.aspx).
        */
        void SetPriority(long priority);

        /**
        Returns a Boolean value indicating if the conditional format is being applied to a PivotTable chart. Since Excel 2007.

        [MSDN documentation for AboveAverage.PTCondition](http://msdn.microsoft.com/en-us/library/bb210747.aspx).
        */
        bool GetPTCondition();

        /**
        Returns one of the constants of the XlPivotConditionScope enumeration, which determines the scope of the conditional format when it is applied to a PivotTable chart. Since Excel 2007.

        [MSDN documentation for AboveAverage.ScopeType](http://msdn.microsoft.com/en-us/library/bb210749.aspx).
        */
        XlPivotConditionScope GetScopeType();

        /**
        Sets one of the constants of the XlPivotConditionScope enumeration, which determines the scope of the conditional format when it is applied to a PivotTable chart. Since Excel 2007.

        [MSDN documentation for AboveAverage.ScopeType](http://msdn.microsoft.com/en-us/library/bb210749.aspx).
        */
        void SetScopeType(XlPivotConditionScope scopeType);

        /**
        Returns a Boolean value that determines if additional formatting rules on the cell should be evaluated if the current rule evaluates to True. Since Excel 2007.

        [MSDN documentation for AboveAverage.StopIfTrue](http://msdn.microsoft.com/en-us/library/bb210753.aspx).
        */
        bool GetStopIfTrue();

        /**
        Sets a Boolean value that determines if additional formatting rules on the cell should be evaluated if the current rule evaluates to True. Since Excel 2007.

        [MSDN documentation for AboveAverage.StopIfTrue](http://msdn.microsoft.com/en-us/library/bb210753.aspx).
        */
        void SetStopIfTrue(bool stopIfTrue);

        /**
        Returns one of the constants of the XlFormatConditionType enumeration, which specifies the type of conditional format. Since Excel 2007.

        [MSDN documentation for AboveAverage.Type](http://msdn.microsoft.com/en-us/library/bb210755.aspx).
        */
        XlFormatConditionType GetType();

        /**
        Returns one of the constants of the XlFormatConditionType enumeration, which specifies the type of conditional format. Since Excel 2007.

        [MSDN documentation for AboveAverage.Type](http://msdn.microsoft.com/en-us/library/bb210755.aspx).
        */
        void SetType(XlFormatConditionType type);
                        
        /**
        Returns "AboveAverage".
        */
        virtual wxString GetAutoExcelObjectName_() const { return wxS("AboveAverage"); }    
    };


} // namespace wxAutoExcel


#endif // WXAUTOEXCEL_USE_CONDFORMAT

#endif //_WXAUTOEXCEL_ABOVEAVERAGE_H

