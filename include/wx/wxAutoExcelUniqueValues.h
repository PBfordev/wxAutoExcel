/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// License:     MIT license
/////////////////////////////////////////////////////////////////////////////


#ifndef _WXAUTOEXCEL_UNIQUEVALUES_H
#define _WXAUTOEXCEL_UNIQUEVALUES_H

#include "wx/wxAutoExcel_defs.h"

#if WXAUTOEXCEL_USE_CONDFORMAT

#include "wx/wxAutoExcel_object.h"
#include "wx/wxAutoExcel_enums.h"

namespace wxAutoExcel {
    /**
    @brief Represents a Microsoft Excel UniqueValues object.
    */
   class WXDLLIMPEXP_WXAUTOEXCEL wxExcelUniqueValues : public wxExcelObject
    {
    public:
        // ***** METHODS *****

        /**
        Deletes the specified conditional formatting rule object.

        [MSDN documentation for UniqueValues.Delete](https://learn.microsoft.com/en-us/office/vba/api/excel.uniquevalues.delete).
        */
        void Delete();

        /**
        Sets the cell range to which this formatting rule applies.

        [MSDN documentation for UniqueValues.ModifyAppliesToRange](https://learn.microsoft.com/en-us/office/vba/api/excel.uniquevalues.modifyappliestorange).
        */
        void ModifyAppliesToRange(wxExcelRange range);

        /**
        Sets the priority value for this conditional formatting rule to "1" so that it will be evaluated before all other rules on the worksheet.

        [MSDN documentation for UniqueValues.SetFirstPriority](https://learn.microsoft.com/en-us/office/vba/api/excel.uniquevalues.setfirstpriority).
        */
        void SetFirstPriority();

        /**
        Sets the evaluation order for this conditional formatting rule so it is evaluated after all other rules on the worksheet.

        [MSDN documentation for UniqueValues.SetLastPriority](https://learn.microsoft.com/en-us/office/vba/api/excel.uniquevalues.setlastpriority).
        */
        void SetLastPriority();

        // ***** PROPERTIES *****

        /**
        Returns a Range object specifying the cell range to which the formatting rule is applied. Since Excel 2007.

        [MSDN documentation for UniqueValues.AppliesTo](https://learn.microsoft.com/en-us/office/vba/api/excel.uniquevalues.appliesto).
        */
        wxExcelRange GetAppliesTo();

        /**
        Returns a Borders collection that specifies the formatting of cell borders if the conditional formatting rule evaluates to True. Since Excel 2007.

        [MSDN documentation for UniqueValues.Borders](https://learn.microsoft.com/en-us/office/vba/api/excel.uniquevalues.borders).
        */
        wxExcelBorders GetBorders();

        /**
        Returns of sets one of the constants of XlDupeUnique enumeration specifying if the conditional format rule is looking for unique or duplicate values. Since Excel 2007.

        [MSDN documentation for UniqueValues.DupeUnique](https://learn.microsoft.com/en-us/office/vba/api/excel.uniquevalues.dupeunique).
        */
        XlDupeUnique GetDupeUnique();

        /**
        Returns of sets one of the constants of XlDupeUnique enumeration specifying if the conditional format rule is looking for unique or duplicate values. Since Excel 2007.

        [MSDN documentation for UniqueValues.DupeUnique](https://learn.microsoft.com/en-us/office/vba/api/excel.uniquevalues.dupeunique).
        */
        void SetDupeUnique(XlDupeUnique dupeUnique);

        /**
        Returns a Font object that specifies the font formatting if the conditional formatting rule evaluates to True. Since Excel 2007.

        [MSDN documentation for UniqueValues.Font](https://learn.microsoft.com/en-us/office/vba/api/excel.uniquevalues.font).
        */
        wxExcelFont GetFont();

        /**
        Returns an Interior object that specifies a cell's interior attributes for a conditional formatting rule that evaluates to True. Since Excel 2007.

        [MSDN documentation for UniqueValues.Interior](https://learn.microsoft.com/en-us/office/vba/api/excel.uniquevalues.interior).
        */
        wxExcelInterior GetInterior();

        /**
        Returns the number format applied to a cell if the conditional formatting rule evaluates to True. Read/write Variant. Since Excel 2007.

        The name is spelled Numberformat rather than NumberFormat as in Excel, because the Windows headers define GetNumberFormat as a macro, which would otherwise rename the method.

        [MSDN documentation for UniqueValues.NumberFormat](https://learn.microsoft.com/en-us/office/vba/api/excel.uniquevalues.numberformat).
        */
        wxString GetNumberformat();

        /**
        Sets the number format applied to a cell if the conditional formatting rule evaluates to True. Read/write Variant. Since Excel 2007.

        The name is spelled Numberformat rather than NumberFormat as in Excel, because the Windows headers define GetNumberFormat as a macro, which would otherwise rename the method.

        [MSDN documentation for UniqueValues.NumberFormat](https://learn.microsoft.com/en-us/office/vba/api/excel.uniquevalues.numberformat).
        */
        void SetNumberformat(const wxString& numberFormat);

        /**
        Returns the priority value of the conditional formatting rule. The priority determines the order of evaluation when multiple conditional formatting rules exist in a worksheet. Since Excel 2007.

        [MSDN documentation for UniqueValues.Priority](https://learn.microsoft.com/en-us/office/vba/api/excel.uniquevalues.priority).
        */
        long GetPriority();

        /**
        Sets the priority value of the conditional formatting rule. The priority determines the order of evaluation when multiple conditional formatting rules exist in a worksheet. Since Excel 2007.

        [MSDN documentation for UniqueValues.Priority](https://learn.microsoft.com/en-us/office/vba/api/excel.uniquevalues.priority).
        */
        void SetPriority(long priority);

        /**
        Returns a Boolean value indicating if the conditional format is being applied to a PivotTable chart. Since Excel 2007.

        [MSDN documentation for UniqueValues.PTCondition](https://learn.microsoft.com/en-us/office/vba/api/excel.uniquevalues.ptcondition).
        */
        bool GetPTCondition();

        /**
        Returns one of the constants of the XlPivotConditionScope enumeration, which determines the scope of the conditional format when it is applied to a PivotTable chart. Since Excel 2007.

        [MSDN documentation for UniqueValues.ScopeType](https://learn.microsoft.com/en-us/office/vba/api/excel.uniquevalues.scopetype).
        */
        XlPivotConditionScope GetScopeType();

        /**
        Sets one of the constants of the XlPivotConditionScope enumeration, which determines the scope of the conditional format when it is applied to a PivotTable chart. Since Excel 2007.

        [MSDN documentation for UniqueValues.ScopeType](https://learn.microsoft.com/en-us/office/vba/api/excel.uniquevalues.scopetype).
        */
        void SetScopeType(XlPivotConditionScope scopeType);

        /**
        Returns a Boolean value that determines if additional formatting rules on the cell should be evaluated if the current rule evaluates to True. Since Excel 2007.

        [MSDN documentation for UniqueValues.StopIfTrue](https://learn.microsoft.com/en-us/office/vba/api/excel.uniquevalues.stopiftrue).
        */
        bool GetStopIfTrue();

        /**
        Sets a Boolean value that determines if additional formatting rules on the cell should be evaluated if the current rule evaluates to True. Since Excel 2007.

        [MSDN documentation for UniqueValues.StopIfTrue](https://learn.microsoft.com/en-us/office/vba/api/excel.uniquevalues.stopiftrue).
        */
        void SetStopIfTrue(bool stopIfTrue);

        /**
        Returns one of the constants of the XlFormatConditionType enumeration, which specifies the type of conditional format. Since Excel 2007.

        [MSDN documentation for UniqueValues.Type](https://learn.microsoft.com/en-us/office/vba/api/excel.uniquevalues.type).
        */
        XlFormatConditionType GetType();

        /**
        Returns "UniqueValues".
        */
        virtual wxString GetAutoExcelObjectName_() const { return wxS("UniqueValues"); }
    };


} // namespace wxAutoExcel

#endif // WXAUTOEXCEL_USE_CONDFORMAT

#endif //_WXAUTOEXCEL_UNIQUEVALUES_H
