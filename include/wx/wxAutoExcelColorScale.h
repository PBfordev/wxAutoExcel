/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Modified by:
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
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

        [MSDN documentation for ColorScale.Delete](http://msdn.microsoft.com/en-us/library/bb178571.aspx).
        */
        void Delete();

        /**
        Sets the cell range to which this formatting rule applies.

        [MSDN documentation for ColorScale.ModifyAppliesToRange](http://msdn.microsoft.com/en-us/library/bb178573.aspx).
        */
        void ModifyAppliesToRange(wxExcelRange range);

        /**
        Sets the priority value for this conditional formatting rule to "1" so that it will be evaluated before all other rules on the worksheet.

        [MSDN documentation for ColorScale.SetFirstPriority](http://msdn.microsoft.com/en-us/library/bb178582.aspx).
        */
        void SetFirstPriority();

        /**
        Sets the evaluation order for this conditional formatting rule so it is evaluated after all other rules on the worksheet.

        [MSDN documentation for ColorScale.SetLastPriority](http://msdn.microsoft.com/en-us/library/bb178586.aspx).
        */
        void SetLastPriority();

        // ***** PROPERTIES *****

        /**
        When used without an object qualifier, this property returns an Application object that represents the Microsoft Office Excel application. When used with an object qualifier, this property returns an Application object that represents the creator of the specified object. Since Excel 2007.

        [MSDN documentation for ColorScale.Application](http://msdn.microsoft.com/en-us/library/bb239968.aspx).
        */

        /**
        Returns a Range object specifying the cell range to which the formatting rule is applied. Since Excel 2007.

        [MSDN documentation for ColorScale.AppliesTo](http://msdn.microsoft.com/en-us/library/bb224268.aspx).
        */
        wxExcelRange GetAppliesTo();

        /**
        Returns a ColorScaleCriteria object which is a collection of individual ColorScaleCriterion objects. The ColorScaleCriterion object specifies the type, value, and the color of threshold criteria used in the color scale conditional format. Since Excel 2007.

        [MSDN documentation for ColorScale.ColorScaleCriteria](http://msdn.microsoft.com/en-us/library/bb239969.aspx).
        */
        wxExcelColorScaleCriteria GetColorScaleCriteria();

        /**
        Returns a String representing a formula which determines the values the icon set will be applied to. Since Excel 2007.

        [MSDN documentation for ColorScale.Formula](http://msdn.microsoft.com/en-us/library/bb224280.aspx).
        */
        wxString GetFormula();

        /**
        Sets a String representing a formula which determines the values the icon set will be applied to. Since Excel 2007.

        [MSDN documentation for ColorScale.Formula](http://msdn.microsoft.com/en-us/library/bb224280.aspx).
        */
        void SetFormula(const wxString& formula);

        /**
        Returns the priority value of the conditional formatting rule. The priority determines the order of evaluation when multiple conditional formatting rules exist in a worksheet. Since Excel 2007.

        [MSDN documentation for ColorScale.Priority](http://msdn.microsoft.com/en-us/library/bb224290.aspx).
        */
        long GetPriority();

        /**
        Sets the priority value of the conditional formatting rule. The priority determines the order of evaluation when multiple conditional formatting rules exist in a worksheet. Since Excel 2007.

        [MSDN documentation for ColorScale.Priority](http://msdn.microsoft.com/en-us/library/bb224290.aspx).
        */
        void SetPriority(long priority);

        /**
        Returns a Boolean value indicating if the conditional format is being applied to a PivotTable chart. Since Excel 2007.

        [MSDN documentation for ColorScale.PTCondition](http://msdn.microsoft.com/en-us/library/bb224295.aspx).
        */
        bool GetPTCondition();

        /**
        Returns one of the constants of the XlPivotConditionScope enumeration, which determines the scope of the conditional format when it is applied to a PivotTable chart. Since Excel 2007.

        [MSDN documentation for ColorScale.ScopeType](http://msdn.microsoft.com/en-us/library/bb224309.aspx).
        */
        XlPivotConditionScope GetScopeType();

        /**
        Returns a Boolean value that determines if additional formatting rules on the cell should be evaluated if the current rule evaluates to True. Since Excel 2007.

        [MSDN documentation for ColorScale.StopIfTrue](http://msdn.microsoft.com/en-us/library/bb224316.aspx).
        */
        bool GetStopIfTrue();

        /**
        Returns one of the constants of the XlFormatConditionType enumeration, which specifies the type of conditional format. Since Excel 2007.

        [MSDN documentation for ColorScale.Type](http://msdn.microsoft.com/en-us/library/bb224326.aspx).
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
