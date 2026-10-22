/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// License:     MIT license
/////////////////////////////////////////////////////////////////////////////


#ifndef _WXAUTOEXCEL_DATABAR_H
#define _WXAUTOEXCEL_DATABAR_H

#include "wx/wxAutoExcel_defs.h"

#if WXAUTOEXCEL_USE_CONDFORMAT

#include "wx/wxAutoExcel_object.h"
#include "wx/wxAutoExcel_enums.h"

namespace wxAutoExcel {
    /**
    @brief Represents a Microsoft Excel Databar object.
    */
   class WXDLLIMPEXP_WXAUTOEXCEL wxExcelDatabar : public wxExcelObject
    {
    public:
        // ***** METHODS *****

        /**
        Deletes the specified conditional formatting rule object.

        [MSDN documentation for Databar.Delete](https://learn.microsoft.com/en-us/office/vba/api/excel.databar.delete).
        */
        void Delete();

        /**
        Sets the cell range to which this formatting rule applies.

        [MSDN documentation for Databar.ModifyAppliesToRange](https://learn.microsoft.com/en-us/office/vba/api/excel.databar.modifyappliestorange).
        */
        void ModifyAppliesToRange(wxExcelRange range);

        /**
        Sets the priority value for this conditional formatting rule to "1" so that it will be evaluated before all other rules on the worksheet.

        [MSDN documentation for Databar.SetFirstPriority](https://learn.microsoft.com/en-us/office/vba/api/excel.databar.setfirstpriority).
        */
        void SetFirstPriority();

        /**
        Sets the evaluation order for this conditional formatting rule so it is evaluated after all other rules on the worksheet.

        [MSDN documentation for Databar.SetLastPriority](https://learn.microsoft.com/en-us/office/vba/api/excel.databar.setlastpriority).
        */
        void SetLastPriority();

        // ***** PROPERTIES *****

        /**
        Returns a Range object specifying the cell range to which the formatting rule is applied. Since Excel 2007.

        [MSDN documentation for Databar.AppliesTo](https://learn.microsoft.com/en-us/office/vba/api/excel.databar.appliesto).
        */
        wxExcelRange GetAppliesTo();

         /**
        Returns the color of the axis for cells with conditional formatting as data bars. Since Excel 2010.

        [MSDN documentation for Databar.AxisColor](https://learn.microsoft.com/en-us/office/vba/api/excel.databar.axiscolor).
        */
        wxExcelFormatColor GetAxisColor();

        /**
        The position of the axis of the data bars specified by a conditional formatting rule. Since Excel 2010.

        [MSDN documentation for Databar.AxisPosition](https://learn.microsoft.com/en-us/office/vba/api/excel.databar.axisposition).
        */
        XlDataBarAxisPosition GetAxisPosition();

        /**
        The position of the axis of the data bars specified by a conditional formatting rule. Since Excel 2010.

        [MSDN documentation for Databar.AxisPosition](https://learn.microsoft.com/en-us/office/vba/api/excel.databar.axisposition).
        */
        void SetAxisPosition(XlDataBarAxisPosition axisPosition);

        /**
        Returns an object that specifies the border of a data bar. Since Excel 2007.

        [MSDN documentation for Databar.BarBorder](https://learn.microsoft.com/en-us/office/vba/api/excel.databar.barborder).
        */
        wxExcelDataBarBorder GetBarBorder();


        /**
        Returns a FormatColor object that you can use to modify the color of the bars in a data bar conditional format. Since Excel 2007.

        [MSDN documentation for Databar.BarColor](https://learn.microsoft.com/en-us/office/vba/api/excel.databar.barcolor).
        */
        wxExcelFormatColor GetBarColor();


        /**
        How a data bar is filled with color. Since Excel 2010.

        [MSDN documentation for Databar.BarFillType](https://learn.microsoft.com/en-us/office/vba/api/excel.databar.barfilltype).
        */
        XlDataBarFillType GetBarFillType();

        /**
        How a data bar is filled with color. Since Excel 2010.

        [MSDN documentation for Databar.BarFillType](https://learn.microsoft.com/en-us/office/vba/api/excel.databar.barfilltype).
        */
        void SetBarFillType(XlDataBarFillType fillType);

        /**
        Returns a String representing a formula which determines the values the icon set will be applied to. Since Excel 2007.

        [MSDN documentation for Databar.Formula](https://learn.microsoft.com/en-us/office/vba/api/excel.databar.formula).
        */
        wxString GetFormula();

        /**
        Sets a String representing a formula which determines the values the icon set will be applied to. Since Excel 2007.

        [MSDN documentation for Databar.Formula](https://learn.microsoft.com/en-us/office/vba/api/excel.databar.formula).
        */
        void SetFormula(const wxString& formula);

        /**
        Returns a ConditionValue object which specifies how the longest bar is evaluated for a data bar conditional format. Since Excel 2007.

        [MSDN documentation for Databar.MaxPoint](https://learn.microsoft.com/en-us/office/vba/api/excel.databar.maxpoint).
        */
        wxExcelConditionValue GetMaxPoint();

        /**
        Returns a ConditionValue object which specifies how the shortest bar is evaluated for a data bar conditional format. Since Excel 2007.

        [MSDN documentation for Databar.MinPoint](https://learn.microsoft.com/en-us/office/vba/api/excel.databar.minpoint).
        */
        wxExcelConditionValue GetMinPoint();

        /**
        Returns Returns the NegativeBarFormat object associated with a data bar conditional formatting rule. Since Excel 2010.

        [MSDN documentation for Databar.NegativeBarFormat](https://learn.microsoft.com/en-us/office/vba/api/excel.databar.negativebarformat).
        */
        wxExcelNegativeBarFormat GetNegativeBarFormat();

        /**
        Returns a Long value that specifies the length of the longest data bar as a percentage of cell width. Since Excel 2007.

        [MSDN documentation for Databar.PercentMax](https://learn.microsoft.com/en-us/office/vba/api/excel.databar.percentmax).
        */
        long GetPercentMax();

        /**
        Sets a Long value that specifies the length of the longest data bar as a percentage of cell width. Since Excel 2007.

        [MSDN documentation for Databar.PercentMax](https://learn.microsoft.com/en-us/office/vba/api/excel.databar.percentmax).
        */
        void SetPercentMax(long percentMax);

        /**
        Returns a Long value that specifies the length of the shortest data bar as a percentage of cell width. Since Excel 2007.

        [MSDN documentation for Databar.PercentMin](https://learn.microsoft.com/en-us/office/vba/api/excel.databar.percentmin).
        */
        long GetPercentMin();

        /**
        Sets a Long value that specifies the length of the shortest data bar as a percentage of cell width. Since Excel 2007.

        [MSDN documentation for Databar.PercentMin](https://learn.microsoft.com/en-us/office/vba/api/excel.databar.percentmin).
        */
        void SetPercentMin(long percentMin);

        /**
        Returns the priority value of the conditional formatting rule. The priority determines the order of evaluation when multiple conditional formatting rules exist in a worksheet. Since Excel 2007.

        [MSDN documentation for Databar.Priority](https://learn.microsoft.com/en-us/office/vba/api/excel.databar.priority).
        */
        long GetPriority();

        /**
        Sets the priority value of the conditional formatting rule. The priority determines the order of evaluation when multiple conditional formatting rules exist in a worksheet. Since Excel 2007.

        [MSDN documentation for Databar.Priority](https://learn.microsoft.com/en-us/office/vba/api/excel.databar.priority).
        */
        void SetPriority(long priority);

        /**
        Returns a Boolean value indicating if the conditional format is being applied to a PivotTable chart. Since Excel 2007.

        [MSDN documentation for Databar.PTCondition](https://learn.microsoft.com/en-us/office/vba/api/excel.databar.ptcondition).
        */
        bool GetPTCondition();

        /**
        Returns one of the constants of the XlPivotConditionScope enumeration, which determines the scope of the conditional format when it is applied to a PivotTable chart. Since Excel 2007.

        [MSDN documentation for Databar.ScopeType](https://learn.microsoft.com/en-us/office/vba/api/excel.databar.scopetype).
        */
        XlPivotConditionScope GetScopeType();

        /**
        Sets one of the constants of the XlPivotConditionScope enumeration, which determines the scope of the conditional format when it is applied to a PivotTable chart. Since Excel 2007.

        [MSDN documentation for Databar.ScopeType](https://learn.microsoft.com/en-us/office/vba/api/excel.databar.scopetype).
        */
        void SetScopeType(XlPivotConditionScope scopeType);


        /**
        Returns a Boolean value that specifies if the value in the cell is displayed if the data bar conditional format is applied to the range. Since Excel 2007.

        [MSDN documentation for Databar.ShowValue](https://learn.microsoft.com/en-us/office/vba/api/excel.databar.showvalue).
        */
        bool GetShowValue();

        /**
        Sets a Boolean value that specifies if the value in the cell is displayed if the data bar conditional format is applied to the range. Since Excel 2007.

        [MSDN documentation for Databar.ShowValue](https://learn.microsoft.com/en-us/office/vba/api/excel.databar.showvalue).
        */
        void SetShowValue(bool showValue);

        /**
        Returns a Boolean value that determines if additional formatting rules on the cell should be evaluated if the current rule evaluates to True. Since Excel 2007.

        [MSDN documentation for Databar.StopIfTrue](https://learn.microsoft.com/en-us/office/vba/api/excel.databar.stopiftrue).
        */
        bool GetStopIfTrue();

        /**
        Returns one of the constants of the XlFormatConditionType enumeration, which specifies the type of conditional format. Since Excel 2007.

        [MSDN documentation for Databar.Type](https://learn.microsoft.com/en-us/office/vba/api/excel.databar.type).
        */
        XlFormatConditionType GetType();


        /**
        Returns "Databar".
        */
        virtual wxString GetAutoExcelObjectName_() const { return wxS("Databar"); }
    };


} // namespace wxAutoExcel

#endif // WXAUTOEXCEL_USE_CONDFORMAT

#endif //_WXAUTOEXCEL_DATABAR_H