/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// License:     MIT license
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

        [MSDN documentation for IconSetCondition.Delete](https://learn.microsoft.com/en-us/office/vba/api/excel.iconsetcondition.delete).
        */
        void Delete();

        /**
        Sets the cell range to which this formatting rule applies.

        [MSDN documentation for IconSetCondition.ModifyAppliesToRange](https://learn.microsoft.com/en-us/office/vba/api/excel.iconsetcondition.modifyappliestorange).
        */
        void ModifyAppliesToRange(wxExcelRange range);

        /**
        Sets the priority value for this conditional formatting rule to "1" so that it will be evaluated before all other rules on the worksheet.

        [MSDN documentation for IconSetCondition.SetFirstPriority](https://learn.microsoft.com/en-us/office/vba/api/excel.iconsetcondition.setfirstpriority).
        */
        void SetFirstPriority();

        /**
        Sets the evaluation order for this conditional formatting rule so it is evaluated after all other rules on the worksheet.

        [MSDN documentation for IconSetCondition.SetLastPriority](https://learn.microsoft.com/en-us/office/vba/api/excel.iconsetcondition.setlastpriority).
        */
        void SetLastPriority();

        // ***** PROPERTIES *****

        /**
        Returns a Range object specifying the cell range to which the formatting rule is applied. Since Excel 2007.

        [MSDN documentation for IconSetCondition.AppliesTo](https://learn.microsoft.com/en-us/office/vba/api/excel.iconsetcondition.appliesto).
        */
        wxExcelRange GetAppliesTo();

        /**
        Returns a String representing a formula which determines the values the icon set will be applied to. Since Excel 2007.

        [MSDN documentation for IconSetCondition.Formula](https://learn.microsoft.com/en-us/office/vba/api/excel.iconsetcondition.formula).
        */
        wxString GetFormula();

        /**
        Sets a String representing a formula which determines the values the icon set will be applied to. Since Excel 2007.

        [MSDN documentation for IconSetCondition.Formula](https://learn.microsoft.com/en-us/office/vba/api/excel.iconsetcondition.formula).
        */
        void SetFormula(const wxString& formula);

        /**
        Returns an IconCriteria collection which represents the set of criteria for an icon set conditional formatting rule. Since Excel 2007.

        [MSDN documentation for IconSetCondition.IconCriteria](https://learn.microsoft.com/en-us/office/vba/api/excel.iconsetcondition.iconcriteria).
        */
        wxExcelIconCriteria GetIconCriteria();

        /**
        Returns an IconSet object which specifies the icon set used in the conditional format. Since Excel 2007.

        [MSDN documentation for IconSetCondition.IconSet](https://learn.microsoft.com/en-us/office/vba/api/excel.iconsetcondition.iconset).
        */
        wxExcelIconSet GetIconSet();

        /**
        Sets an IconSet object which specifies the icon set used in the conditional format. Since Excel 2007.

        [MSDN documentation for IconSetCondition.IconSet](https://learn.microsoft.com/en-us/office/vba/api/excel.iconsetcondition.iconset).
        */
        void SetIconSet(const wxExcelIconSet& iconSet);

        /**
        Returns a Boolean value indicating if the thresholds for an icon set conditional format are determined using percentiles. Since Excel 2007.

        [MSDN documentation for IconSetCondition.PercentileValues](https://learn.microsoft.com/en-us/office/vba/api/excel.iconsetcondition.percentilevalues).
        */
        bool GetPercentileValues();

        /**
        Sets a Boolean value indicating if the thresholds for an icon set conditional format are determined using percentiles. Since Excel 2007.

        [MSDN documentation for IconSetCondition.PercentileValues](https://learn.microsoft.com/en-us/office/vba/api/excel.iconsetcondition.percentilevalues).
        */
        void SetPercentileValues(bool percentileValues);

        /**
        Returns the priority value of the conditional formatting rule. The priority determines the order of evaluation when multiple conditional formatting rules exist in a worksheet. Since Excel 2007.

        [MSDN documentation for IconSetCondition.Priority](https://learn.microsoft.com/en-us/office/vba/api/excel.iconsetcondition.priority).
        */
        long GetPriority();

        /**
        Sets the priority value of the conditional formatting rule. The priority determines the order of evaluation when multiple conditional formatting rules exist in a worksheet. Since Excel 2007.

        [MSDN documentation for IconSetCondition.Priority](https://learn.microsoft.com/en-us/office/vba/api/excel.iconsetcondition.priority).
        */
        void SetPriority(long priority);

        /**
        Returns a Boolean value indicating if the conditional format is being applied to a PivotTable chart. Since Excel 2007.

        [MSDN documentation for IconSetCondition.PTCondition](https://learn.microsoft.com/en-us/office/vba/api/excel.iconsetcondition.ptcondition).
        */
        bool GetPTCondition();

        /**
        Returns a Boolean value indicating if the order of icons is reversed for an icon set. Since Excel 2007.

        [MSDN documentation for IconSetCondition.ReverseOrder](https://learn.microsoft.com/en-us/office/vba/api/excel.iconsetcondition.reverseorder).
        */
        bool GetReverseOrder();

        /**
        Sets a Boolean value indicating if the order of icons is reversed for an icon set. Since Excel 2007.

        [MSDN documentation for IconSetCondition.ReverseOrder](https://learn.microsoft.com/en-us/office/vba/api/excel.iconsetcondition.reverseorder).
        */
        void SetReverseOrder(bool reverseOrder);

        /**
        Returns one of the constants of the XlPivotConditionScope enumeration, which determines the scope of the conditional format when it is applied to a PivotTable chart. Since Excel 2007.

        [MSDN documentation for IconSetCondition.ScopeType](https://learn.microsoft.com/en-us/office/vba/api/excel.iconsetcondition.scopetype).
        */
        XlPivotConditionScope GetScopeType();

        /**
        Sets one of the constants of the XlPivotConditionScope enumeration, which determines the scope of the conditional format when it is applied to a PivotTable chart. Since Excel 2007.

        [MSDN documentation for IconSetCondition.ScopeType](https://learn.microsoft.com/en-us/office/vba/api/excel.iconsetcondition.scopetype).
        */
        void SetScopeType(XlPivotConditionScope scopeType);

        /**
        Returns a Boolean value indicating if only the icon is displayed for an icon set conditional format. Since Excel 2007.

        [MSDN documentation for IconSetCondition.ShowIconOnly](https://learn.microsoft.com/en-us/office/vba/api/excel.iconsetcondition.showicononly).
        */
        bool GetShowIconOnly();

        /**
        Sets a Boolean value indicating if only the icon is displayed for an icon set conditional format. Since Excel 2007.

        [MSDN documentation for IconSetCondition.ShowIconOnly](https://learn.microsoft.com/en-us/office/vba/api/excel.iconsetcondition.showicononly).
        */
        void SetShowIconOnly(bool showIconOnly);

        /**
        Returns a Boolean value that determines if additional formatting rules on the cell should be evaluated if the current rule evaluates to True. Since Excel 2007.

        [MSDN documentation for IconSetCondition.StopIfTrue](https://learn.microsoft.com/en-us/office/vba/api/excel.iconsetcondition.stopiftrue).
        */
        bool GetStopIfTrue();

        /**
        Returns one of the constants of the XlFormatConditionType enumeration, which specifies the type of conditional format. Since Excel 2007.

        [MSDN documentation for IconSetCondition.Type](https://learn.microsoft.com/en-us/office/vba/api/excel.iconsetcondition.type).
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
