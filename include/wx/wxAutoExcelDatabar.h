/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Modified by:
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
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

        [MSDN documentation for Databar.?Delete](http://msdn.microsoft.com/en-us/library/bb178631.aspx).
        */
        void Delete();

        /**
        Sets the cell range to which this formatting rule applies.

        [MSDN documentation for Databar.ModifyAppliesToRange](http://msdn.microsoft.com/en-us/library/bb178636.aspx).
        */
        void ModifyAppliesToRange(wxExcelRange range);

        /**
        Sets the priority value for this conditional formatting rule to "1" so that it will be evaluated before all other rules on the worksheet.

        [MSDN documentation for Databar.SetFirstPriority](http://msdn.microsoft.com/en-us/library/bb178638.aspx).
        */
        void SetFirstPriority();

        /**
        Sets the evaluation order for this conditional formatting rule so it is evaluated after all other rules on the worksheet.

        [MSDN documentation for Databar.SetLastPriority](http://msdn.microsoft.com/en-us/library/bb178644.aspx).
        */
        void SetLastPriority();

        // ***** PROPERTIES *****

        /**
        Returns a Range object specifying the cell range to which the formatting rule is applied. Since Excel 2007.

        [MSDN documentation for Databar.AppliesTo](http://msdn.microsoft.com/en-us/library/bb224396.aspx).
        */
        wxExcelRange GetAppliesTo();

         /**
        Returns the color of the axis for cells with conditional formatting as data bars. Since Excel 2010.

        [MSDN documentation for Databar.AxisColor](http://msdn.microsoft.com/en-us/library/office/ff193665%28v=office.14%29.aspx).
        */
        wxExcelFormatColor GetAxisColor();

        /**
        The position of the axis of the data bars specified by a conditional formatting rule. Since Excel 2010.

        [MSDN documentation for Databar.AxisPosition](http://msdn.microsoft.com/en-us/library/office/ff193799%28v=office.14%29.aspx).
        */
        XlDataBarAxisPosition GetAxisPosition();

        /**
        The position of the axis of the data bars specified by a conditional formatting rule. Since Excel 2010.

        [MSDN documentation for Databar.AxisPosition](http://msdn.microsoft.com/en-us/library/office/ff193799%28v=office.14%29.aspx).
        */
        void SetAxisPosition(XlDataBarAxisPosition axisPosition);

        /**
        Returns an object that specifies the border of a data bar. Since Excel 2007.

        [MSDN documentation for Databar.BarBorder](https://docs.microsoft.com/office/vba/api/excel.databar.barborder).
        */
        wxExcelDataBarBorder GetBarBorder();


        /**
        Returns a FormatColor object that you can use to modify the color of the bars in a data bar conditional format. Since Excel 2007.

        [MSDN documentation for Databar.BarColor](http://msdn.microsoft.com/en-us/library/bb224398.aspx).
        */
        wxExcelFormatColor GetBarColor();


        /**
        How a data bar is filled with color. Since Excel 2010.

        [MSDN documentation for Databar.BarFillType](http://msdn.microsoft.com/en-us/library/office/ff839004%28v=office.14%29.aspx).
        */
        XlDataBarFillType GetBarFillType();

        /**
        How a data bar is filled with color. Since Excel 2010.

        [MSDN documentation for Databar.BarFillType](http://msdn.microsoft.com/en-us/library/office/ff839004%28v=office.14%29.aspx).
        */
        void SetBarFillType(XlDataBarFillType fillType);

        /**
        Returns a String representing a formula which determines the values the icon set will be applied to. Since Excel 2007.

        [MSDN documentation for Databar.Formula](http://msdn.microsoft.com/en-us/library/bb224403.aspx).
        */
        wxString GetFormula();

        /**
        Sets a String representing a formula which determines the values the icon set will be applied to. Since Excel 2007.

        [MSDN documentation for Databar.Formula](http://msdn.microsoft.com/en-us/library/bb224403.aspx).
        */
        void SetFormula(const wxString& formula);

        /**
        Returns a ConditionValue object which specifies how the longest bar is evaluated for a data bar conditional format. Since Excel 2007.

        [MSDN documentation for Databar.MaxPoint](http://msdn.microsoft.com/en-us/library/bb224408.aspx).
        */
        wxExcelConditionValue GetMaxPoint();

        /**
        Returns a ConditionValue object which specifies how the shortest bar is evaluated for a data bar conditional format. Since Excel 2007.

        [MSDN documentation for Databar.MinPoint](http://msdn.microsoft.com/en-us/library/bb224410.aspx).
        */
        wxExcelConditionValue GetMinPoint();

        /**
        Returns Returns the NegativeBarFormat object associated with a data bar conditional formatting rule. Since Excel 2010.

        [MSDN documentation for Databar.NegativeBarFormat](http://msdn.microsoft.com/en-us/library/ff839392.aspx).
        */
        wxExcelNegativeBarFormat GetNegativeBarFormat();

        /**
        Returns a Long value that specifies the length of the longest data bar as a percentage of cell width. Since Excel 2007.

        [MSDN documentation for Databar.PercentMax](http://msdn.microsoft.com/en-us/library/bb224413.aspx).
        */
        long GetPercentMax();

        /**
        Sets a Long value that specifies the length of the longest data bar as a percentage of cell width. Since Excel 2007.

        [MSDN documentation for Databar.PercentMax](http://msdn.microsoft.com/en-us/library/bb224413.aspx).
        */
        void SetPercentMax(long percentMax);

        /**
        Returns a Long value that specifies the length of the shortest data bar as a percentage of cell width. Since Excel 2007.

        [MSDN documentation for Databar.PercentMin](http://msdn.microsoft.com/en-us/library/bb224417.aspx).
        */
        long GetPercentMin();

        /**
        Sets a Long value that specifies the length of the shortest data bar as a percentage of cell width. Since Excel 2007.

        [MSDN documentation for Databar.PercentMin](http://msdn.microsoft.com/en-us/library/bb224417.aspx).
        */
        void SetPercentMin(long percentMin);

        /**
        Returns the priority value of the conditional formatting rule. The priority determines the order of evaluation when multiple conditional formatting rules exist in a worksheet. Since Excel 2007.

        [MSDN documentation for Databar.Priority](http://msdn.microsoft.com/en-us/library/bb224418.aspx).
        */
        long GetPriority();

        /**
        Sets the priority value of the conditional formatting rule. The priority determines the order of evaluation when multiple conditional formatting rules exist in a worksheet. Since Excel 2007.

        [MSDN documentation for Databar.Priority](http://msdn.microsoft.com/en-us/library/bb224418.aspx).
        */
        void SetPriority(long priority);

        /**
        Returns a Boolean value indicating if the conditional format is being applied to a PivotTable chart. Since Excel 2007.

        [MSDN documentation for Databar.PTCondition](http://msdn.microsoft.com/en-us/library/bb224424.aspx).
        */
        bool GetPTCondition();

        /**
        Returns one of the constants of the XlPivotConditionScope enumeration, which determines the scope of the conditional format when it is applied to a PivotTable chart. Since Excel 2007.

        [MSDN documentation for Databar.ScopeType](http://msdn.microsoft.com/en-us/library/bb224429.aspx).
        */
        XlPivotConditionScope GetScopeType();

        /**
        Sets one of the constants of the XlPivotConditionScope enumeration, which determines the scope of the conditional format when it is applied to a PivotTable chart. Since Excel 2007.

        [MSDN documentation for Databar.ScopeType](http://msdn.microsoft.com/en-us/library/bb224429.aspx).
        */
        void SetScopeType(XlPivotConditionScope scopeType);


        /**
        Returns a Boolean value that specifies if the value in the cell is displayed if the data bar conditional format is applied to the range. Since Excel 2007.

        [MSDN documentation for Databar.ShowValue](http://msdn.microsoft.com/en-us/library/bb224430.aspx).
        */
        bool GetShowValue();

        /**
        Sets a Boolean value that specifies if the value in the cell is displayed if the data bar conditional format is applied to the range. Since Excel 2007.

        [MSDN documentation for Databar.ShowValue](http://msdn.microsoft.com/en-us/library/bb224430.aspx).
        */
        void SetShowValue(bool showValue);

        /**
        Returns a Boolean value that determines if additional formatting rules on the cell should be evaluated if the current rule evaluates to True. Since Excel 2007.

        [MSDN documentation for Databar.StopIfTrue](http://msdn.microsoft.com/en-us/library/bb224434.aspx).
        */
        bool GetStopIfTrue();

        /**
        Sets a Boolean value that determines if additional formatting rules on the cell should be evaluated if the current rule evaluates to True. Since Excel 2007.

        [MSDN documentation for Databar.StopIfTrue](http://msdn.microsoft.com/en-us/library/bb224434.aspx).
        */
        void SetStopIfTrue(bool stopIfTrue);

        /**
        Returns one of the constants of the XlFormatConditionType enumeration, which specifies the type of conditional format. Since Excel 2007.

        [MSDN documentation for Databar.Type](http://msdn.microsoft.com/en-us/library/bb224438.aspx).
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