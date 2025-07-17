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

        [MSDN documentation for UniqueValues.Delete](http://msdn.microsoft.com/en-us/library/bb210603.aspx).
        */
        void Delete();

        /**
        Sets the cell range to which this formatting rule applies.

        [MSDN documentation for UniqueValues.ModifyAppliesToRange](http://msdn.microsoft.com/en-us/library/bb210606.aspx).
        */
        void ModifyAppliesToRange(wxExcelRange range);

        /**
        Sets the priority value for this conditional formatting rule to "1" so that it will be evaluated before all other rules on the worksheet.

        [MSDN documentation for UniqueValues.SetFirstPriority](http://msdn.microsoft.com/en-us/library/bb210609.aspx).
        */
        void SetFirstPriority();

        /**
        Sets the evaluation order for this conditional formatting rule so it is evaluated after all other rules on the worksheet.

        [MSDN documentation for UniqueValues.SetLastPriority](http://msdn.microsoft.com/en-us/library/bb210612.aspx).
        */
        void SetLastPriority();

        // ***** PROPERTIES *****

        /**
        Returns a Range object specifying the cell range to which the formatting rule is applied. Since Excel 2007.

        [MSDN documentation for UniqueValues.AppliesTo](http://msdn.microsoft.com/en-us/library/bb148150.aspx).
        */
        wxExcelRange GetAppliesTo();

        /**
        Returns a Borders collection that specifies the formatting of cell borders if the conditional formatting rule evaluates to True. Since Excel 2007.

        [MSDN documentation for UniqueValues.Borders](http://msdn.microsoft.com/en-us/library/bb148153.aspx).
        */
        wxExcelBorders GetBorders();

        /**
        Returns of sets one of the constants of XlDupeUnique enumeration specifying if the conditional format rule is looking for unique or duplicate values. Since Excel 2007.

        [MSDN documentation for UniqueValues.DupeUnique](http://msdn.microsoft.com/en-us/library/bb148156.aspx).
        */
        XlDupeUnique GetDupeUnique();

        /**
        Returns of sets one of the constants of XlDupeUnique enumeration specifying if the conditional format rule is looking for unique or duplicate values. Since Excel 2007.

        [MSDN documentation for UniqueValues.DupeUnique](http://msdn.microsoft.com/en-us/library/bb148156.aspx).
        */
        void SetDupeUnique(XlDupeUnique dupeUnique);

        /**
        Returns a Font object that specifies the font formatting if the conditional formatting rule evaluates to True. Since Excel 2007.

        [MSDN documentation for UniqueValues.Font](http://msdn.microsoft.com/en-us/library/bb148159.aspx).
        */
        wxExcelFont GetFont();

        /**
        Returns an Interior object that specifies a cell's interior attributes for a conditional formatting rule that evaluates to True. Since Excel 2007.

        [MSDN documentation for UniqueValues.Interior](http://msdn.microsoft.com/en-us/library/bb148165.aspx).
        */
        wxExcelInterior GetInterior();

        /**
        Returns the number format applied to a cell if the conditional formatting rule evaluates to True. Read/write Variant. Since Excel 2007.

        [MSDN documentation for UniqueValues.NumberFormat](http://msdn.microsoft.com/en-us/library/bb148168.aspx).
        */
        wxString GetNumberFormat();

        /**
        Sets the number format applied to a cell if the conditional formatting rule evaluates to True. Read/write Variant. Since Excel 2007.

        [MSDN documentation for UniqueValues.NumberFormat](http://msdn.microsoft.com/en-us/library/bb148168.aspx).
        */
        void SetNumberFormat(const wxString& numberFormat);

        /**
        Returns the priority value of the conditional formatting rule. The priority determines the order of evaluation when multiple conditional formatting rules exist in a worksheet. Since Excel 2007.

        [MSDN documentation for UniqueValues.Priority](http://msdn.microsoft.com/en-us/library/bb148171.aspx).
        */
        long GetPriority();

        /**
        Sets the priority value of the conditional formatting rule. The priority determines the order of evaluation when multiple conditional formatting rules exist in a worksheet. Since Excel 2007.

        [MSDN documentation for UniqueValues.Priority](http://msdn.microsoft.com/en-us/library/bb148171.aspx).
        */
        void SetPriority(long priority);

        /**
        Returns a Boolean value indicating if the conditional format is being applied to a PivotTable chart. Since Excel 2007.

        [MSDN documentation for UniqueValues.PTCondition](http://msdn.microsoft.com/en-us/library/bb148173.aspx).
        */
        bool GetPTCondition();

        /**
        Returns one of the constants of the XlPivotConditionScope enumeration, which determines the scope of the conditional format when it is applied to a PivotTable chart. Since Excel 2007.

        [MSDN documentation for UniqueValues.ScopeType](http://msdn.microsoft.com/en-us/library/bb148176.aspx).
        */
        XlPivotConditionScope GetScopeType();

        /**
        Sets one of the constants of the XlPivotConditionScope enumeration, which determines the scope of the conditional format when it is applied to a PivotTable chart. Since Excel 2007.

        [MSDN documentation for UniqueValues.ScopeType](http://msdn.microsoft.com/en-us/library/bb148176.aspx).
        */
        void SetScopeType(XlPivotConditionScope scopeType);

        /**
        Returns a Boolean value that determines if additional formatting rules on the cell should be evaluated if the current rule evaluates to True. Since Excel 2007.

        [MSDN documentation for UniqueValues.StopIfTrue](http://msdn.microsoft.com/en-us/library/bb148179.aspx).
        */
        bool GetStopIfTrue();

        /**
        Sets a Boolean value that determines if additional formatting rules on the cell should be evaluated if the current rule evaluates to True. Since Excel 2007.

        [MSDN documentation for UniqueValues.StopIfTrue](http://msdn.microsoft.com/en-us/library/bb148179.aspx).
        */
        void SetStopIfTrue(bool stopIfTrue);

        /**
        Returns one of the constants of the XlFormatConditionType enumeration, which specifies the type of conditional format. Since Excel 2007.

        [MSDN documentation for UniqueValues.Type](http://msdn.microsoft.com/en-us/library/bb211036.aspx).
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
