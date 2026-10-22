/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// License:     MIT license
/////////////////////////////////////////////////////////////////////////////


#ifndef _WXAUTOEXCEL_COLORSCALE_H
#define _WXAUTOEXCEL_COLORSCALE_H

#include "wx/wxAutoExcel_defs.h"

#if WXAUTOEXCEL_USE_CONDFORMAT

#include "wx/wxAutoExcel_object.h"
#include "wx/wxAutoExcel_enums.h"

namespace wxAutoExcel {
    /**
    @brief Represents a Microsoft Excel ColorScale object.
    */
   class WXDLLIMPEXP_WXAUTOEXCEL wxExcelColorScale : public wxExcelObject
    {
    public:
        // ***** METHODS *****

        /**
        Deletes the specified conditional formatting rule object.

        [MSDN documentation for ColorScale.Delete](https://learn.microsoft.com/en-us/office/vba/api/excel.colorscale.delete).
        */
        void Delete();

        /**
        Sets the cell range to which this formatting rule applies.

        [MSDN documentation for ColorScale.ModifyAppliesToRange](https://learn.microsoft.com/en-us/office/vba/api/excel.colorscale.modifyappliestorange).
        */
        void ModifyAppliesToRange(wxExcelRange range);

        /**
        Sets the priority value for this conditional formatting rule to "1" so that it will be evaluated before all other rules on the worksheet.

        [MSDN documentation for ColorScale.SetFirstPriority](https://learn.microsoft.com/en-us/office/vba/api/excel.colorscale.setfirstpriority).
        */
        void SetFirstPriority();

        /**
        Sets the evaluation order for this conditional formatting rule so it is evaluated after all other rules on the worksheet.

        [MSDN documentation for ColorScale.SetLastPriority](https://learn.microsoft.com/en-us/office/vba/api/excel.colorscale.setlastpriority).
        */
        void SetLastPriority();

        // ***** PROPERTIES *****

        /**
        When used without an object qualifier, this property returns an Application object that represents the Microsoft Office Excel application. When used with an object qualifier, this property returns an Application object that represents the creator of the specified object. Since Excel 2007.

        [MSDN documentation for ColorScale.Application](https://learn.microsoft.com/en-us/office/vba/api/excel.colorscale.application).
        */

        /**
        Returns a Range object specifying the cell range to which the formatting rule is applied. Since Excel 2007.

        [MSDN documentation for ColorScale.AppliesTo](https://learn.microsoft.com/en-us/office/vba/api/excel.colorscale.appliesto).
        */
        wxExcelRange GetAppliesTo();

        /**
        Returns a ColorScaleCriteria object which is a collection of individual ColorScaleCriterion objects. The ColorScaleCriterion object specifies the type, value, and the color of threshold criteria used in the color scale conditional format. Since Excel 2007.

        [MSDN documentation for ColorScale.ColorScaleCriteria](https://learn.microsoft.com/en-us/office/vba/api/excel.colorscale.colorscalecriteria).
        */
        wxExcelColorScaleCriteria GetColorScaleCriteria();

        /**
        Returns a String representing a formula which determines the values the icon set will be applied to. Since Excel 2007.

        [MSDN documentation for ColorScale.Formula](https://learn.microsoft.com/en-us/office/vba/api/excel.colorscale.formula).
        */
        wxString GetFormula();

        /**
        Sets a String representing a formula which determines the values the icon set will be applied to. Since Excel 2007.

        [MSDN documentation for ColorScale.Formula](https://learn.microsoft.com/en-us/office/vba/api/excel.colorscale.formula).
        */
        void SetFormula(const wxString& formula);

        /**
        Returns the priority value of the conditional formatting rule. The priority determines the order of evaluation when multiple conditional formatting rules exist in a worksheet. Since Excel 2007.

        [MSDN documentation for ColorScale.Priority](https://learn.microsoft.com/en-us/office/vba/api/excel.colorscale.priority).
        */
        long GetPriority();

        /**
        Sets the priority value of the conditional formatting rule. The priority determines the order of evaluation when multiple conditional formatting rules exist in a worksheet. Since Excel 2007.

        [MSDN documentation for ColorScale.Priority](https://learn.microsoft.com/en-us/office/vba/api/excel.colorscale.priority).
        */
        void SetPriority(long priority);

        /**
        Returns a Boolean value indicating if the conditional format is being applied to a PivotTable chart. Since Excel 2007.

        [MSDN documentation for ColorScale.PTCondition](https://learn.microsoft.com/en-us/office/vba/api/excel.colorscale.ptcondition).
        */
        bool GetPTCondition();

        /**
        Returns one of the constants of the XlPivotConditionScope enumeration, which determines the scope of the conditional format when it is applied to a PivotTable chart. Since Excel 2007.

        [MSDN documentation for ColorScale.ScopeType](https://learn.microsoft.com/en-us/office/vba/api/excel.colorscale.scopetype).
        */
        XlPivotConditionScope GetScopeType();

        /**
        Returns a Boolean value that determines if additional formatting rules on the cell should be evaluated if the current rule evaluates to True. Since Excel 2007.

        [MSDN documentation for ColorScale.StopIfTrue](https://learn.microsoft.com/en-us/office/vba/api/excel.colorscale.stopiftrue).
        */
        bool GetStopIfTrue();

        /**
        Returns one of the constants of the XlFormatConditionType enumeration, which specifies the type of conditional format. Since Excel 2007.

        [MSDN documentation for ColorScale.Type](https://learn.microsoft.com/en-us/office/vba/api/excel.colorscale.type).
        */
        XlFormatConditionType GetType();


        /**
        Returns "ColorScale".
        */
        virtual wxString GetAutoExcelObjectName_() const { return wxS("ColorScale"); }
    };


} // namespace wxAutoExcel

#endif // WXAUTOEXCEL_USE_CONDFORMAT

#endif //_WXAUTOEXCEL_COLORSCALE_H
