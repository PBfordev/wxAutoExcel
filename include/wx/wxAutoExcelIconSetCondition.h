/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Modified by:
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
/////////////////////////////////////////////////////////////////////////////


#ifndef _WXAUTOEXCEL_ICONSETCONDITION_H
#define _WXAUTOEXCEL_ICONSETCONDITION_H

#include "wx/wxAutoExcel_defs.h"

#if WXAUTOEXCEL_USE_CONDFORMAT

#include "wx/wxAutoExcel_object.h"
#include "wx/wxAutoExcel_enums.h"

namespace wxAutoExcel {
    /**
    @brief Represents a Microsoft Excel IconSetCondition object.
    */
   class WXDLLIMPEXP_WXAUTOEXCEL wxExcelIconSetCondition : public wxExcelObject
    {
    public:
        // ***** METHODS *****

        /**
        Deletes the specified conditional formatting rule object.

        [MSDN documentation for IconSetCondition.Delete](http://msdn.microsoft.com/en-us/library/bb224533.aspx).
        */
        void Delete();

        /**
        Sets the cell range to which this formatting rule applies.

        [MSDN documentation for IconSetCondition.ModifyAppliesToRange](http://msdn.microsoft.com/en-us/library/bb224536.aspx).
        */
        void ModifyAppliesToRange(wxExcelRange range);

        /**
        Sets the priority value for this conditional formatting rule to "1" so that it will be evaluated before all other rules on the worksheet.

        [MSDN documentation for IconSetCondition.SetFirstPriority](http://msdn.microsoft.com/en-us/library/bb224537.aspx).
        */
        void SetFirstPriority();

        /**
        Sets the evaluation order for this conditional formatting rule so it is evaluated after all other rules on the worksheet.

        [MSDN documentation for IconSetCondition.SetLastPriority](http://msdn.microsoft.com/en-us/library/bb224539.aspx).
        */
        void SetLastPriority();

        // ***** PROPERTIES *****

        /**
        Returns a Range object specifying the cell range to which the formatting rule is applied. Since Excel 2007.

        [MSDN documentation for IconSetCondition.AppliesTo](http://msdn.microsoft.com/en-us/library/bb224572.aspx).
        */
        wxExcelRange GetAppliesTo();

        /**
        Returns a String representing a formula which determines the values the icon set will be applied to. Since Excel 2007.

        [MSDN documentation for IconSetCondition.Formula](http://msdn.microsoft.com/en-us/library/bb242076.aspx).
        */
        wxString GetFormula();

        /**
        Sets a String representing a formula which determines the values the icon set will be applied to. Since Excel 2007.

        [MSDN documentation for IconSetCondition.Formula](http://msdn.microsoft.com/en-us/library/bb242076.aspx).
        */
        void SetFormula(const wxString& formula);

        /**
        Returns an IconCriteria collection which represents the set of criteria for an icon set conditional formatting rule. Since Excel 2007.

        [MSDN documentation for IconSetCondition.IconCriteria](http://msdn.microsoft.com/en-us/library/bb224722.aspx).
        */
        wxExcelIconCriteria GetIconCriteria();

        /**
        Returns an IconSets collection which specifies the icon set used in the conditional format. Since Excel 2007.

        [MSDN documentation for IconSetCondition.IconSet](http://msdn.microsoft.com/en-us/library/bb224724.aspx).
        */
        wxExcelIconSets GetIconSet();

        /**
        Sets an IconSets collection which specifies the icon set used in the conditional format. Since Excel 2007.

        [MSDN documentation for IconSetCondition.IconSet](http://msdn.microsoft.com/en-us/library/bb224724.aspx).
        */
        void SetIconSet(const wxExcelIconSets& iconSet);

        /**
        Returns a Boolean value indicating if the thresholds for an icon set conditional format are determined using percentiles. Since Excel 2007.

        [MSDN documentation for IconSetCondition.PercentileValues](http://msdn.microsoft.com/en-us/library/bb224731.aspx).
        */
        bool GetPercentileValues();

        /**
        Sets a Boolean value indicating if the thresholds for an icon set conditional format are determined using percentiles. Since Excel 2007.

        [MSDN documentation for IconSetCondition.PercentileValues](http://msdn.microsoft.com/en-us/library/bb224731.aspx).
        */
        void SetPercentileValues(bool percentileValues);

        /**
        Returns the priority value of the conditional formatting rule. The priority determines the order of evaluation when multiple conditional formatting rules exist in a worksheet. Since Excel 2007.

        [MSDN documentation for IconSetCondition.Priority](http://msdn.microsoft.com/en-us/library/bb224734.aspx).
        */
        long GetPriority();

        /**
        Sets the priority value of the conditional formatting rule. The priority determines the order of evaluation when multiple conditional formatting rules exist in a worksheet. Since Excel 2007.

        [MSDN documentation for IconSetCondition.Priority](http://msdn.microsoft.com/en-us/library/bb224734.aspx).
        */
        void SetPriority(long priority);

        /**
        Returns a Boolean value indicating if the conditional format is being applied to a PivotTable chart. Since Excel 2007.

        [MSDN documentation for IconSetCondition.PTCondition](http://msdn.microsoft.com/en-us/library/bb224738.aspx).
        */
        bool GetPTCondition();

        /**
        Returns a Boolean value indicating if the order of icons is reversed for an icon set. Since Excel 2007.

        [MSDN documentation for IconSetCondition.ReverseOrder](http://msdn.microsoft.com/en-us/library/bb224741.aspx).
        */
        bool GetReverseOrder();

        /**
        Sets a Boolean value indicating if the order of icons is reversed for an icon set. Since Excel 2007.

        [MSDN documentation for IconSetCondition.ReverseOrder](http://msdn.microsoft.com/en-us/library/bb224741.aspx).
        */
        void SetReverseOrder(bool reverseOrder);

        /**
        Returns one of the constants of the XlPivotConditionScope enumeration, which determines the scope of the conditional format when it is applied to a PivotTable chart. Since Excel 2007.

        [MSDN documentation for IconSetCondition.ScopeType](http://msdn.microsoft.com/en-us/library/bb224744.aspx).
        */
        XlPivotConditionScope GetScopeType();

        /**
        Sets one of the constants of the XlPivotConditionScope enumeration, which determines the scope of the conditional format when it is applied to a PivotTable chart. Since Excel 2007.

        [MSDN documentation for IconSetCondition.ScopeType](http://msdn.microsoft.com/en-us/library/bb224744.aspx).
        */
        void SetScopeType(XlPivotConditionScope scopeType);

        /**
        Returns a Boolean value indicating if only the icon is displayed for an icon set conditional format. Since Excel 2007.

        [MSDN documentation for IconSetCondition.ShowIconOnly](http://msdn.microsoft.com/en-us/library/bb224747.aspx).
        */
        bool GetShowIconOnly();

        /**
        Sets a Boolean value indicating if only the icon is displayed for an icon set conditional format. Since Excel 2007.

        [MSDN documentation for IconSetCondition.ShowIconOnly](http://msdn.microsoft.com/en-us/library/bb224747.aspx).
        */
        void SetShowIconOnly(bool showIconOnly);

        /**
        Returns a Boolean value that determines if additional formatting rules on the cell should be evaluated if the current rule evaluates to True. Since Excel 2007.

        [MSDN documentation for IconSetCondition.StopIfTrue](http://msdn.microsoft.com/en-us/library/bb224750.aspx).
        */
        bool GetStopIfTrue();

        /**
        Sets a Boolean value that determines if additional formatting rules on the cell should be evaluated if the current rule evaluates to True. Since Excel 2007.

        [MSDN documentation for IconSetCondition.StopIfTrue](http://msdn.microsoft.com/en-us/library/bb224750.aspx).
        */
        void SetStopIfTrue(bool stopIfTrue);

        /**
        Returns one of the constants of the XlFormatConditionType enumeration, which specifies the type of conditional format. Since Excel 2007.

        [MSDN documentation for IconSetCondition.Type](http://msdn.microsoft.com/en-us/library/bb224753.aspx).
        */
        XlFormatConditionType  GetType();                        
        /**
        Returns "IconSetCondition".
        */
        virtual wxString GetAutoExcelObjectName_() const { return wxS("IconSetCondition"); }    
    };


} // namespace wxAutoExcel

#endif // WXAUTOEXCEL_USE_CONDFORMAT

#endif //_WXAUTOEXCEL_ICONSETCONDITION_H
