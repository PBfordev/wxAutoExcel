/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// License:     MIT license
/////////////////////////////////////////////////////////////////////////////


#ifndef _WXAUTOEXCEL_FORMATCONDITIONS_H
#define _WXAUTOEXCEL_FORMATCONDITIONS_H

#include "wx/wxAutoExcel_defs.h"

#if WXAUTOEXCEL_USE_CONDFORMAT

#include "wx/wxAutoExcel_object.h"
#include "wx/wxAutoExcel_enums.h"

namespace wxAutoExcel {

    /**
    @brief Represents Microsoft Excel FormatCondition object.
    */
   class WXDLLIMPEXP_WXAUTOEXCEL wxExcelFormatCondition : public wxExcelObject
    {
    public:
        // ***** METHODS *****

        /**
        Deletes the object.

        [MSDN documentation for FormatCondition.Delete](https://learn.microsoft.com/en-us/office/vba/api/excel.formatcondition.delete).
        */
        void Delete();

        /**
        Modifies an existing conditional format.

        [MSDN documentation for FormatCondition.Modify](https://learn.microsoft.com/en-us/office/vba/api/excel.formatcondition.modify).

        See wxExcelFormatCondition::Add() for a note about adding language-agnostic formulas.
        */
        void Modify(XlFormatConditionType conditionType, XlFormatConditionOperator* conditionOperator = NULL,
                    const wxVariant& formula1 = wxNullVariant, const wxVariant& formula2 = wxNullVariant);

        /**
        Sets the cell range to which this formatting rule applies.

        [MSDN documentation for FormatCondition.ModifyAppliesToRange](https://learn.microsoft.com/en-us/office/vba/api/excel.formatcondition.modifyappliestorange).
        */
        void ModifyAppliesToRange(wxExcelRange range);

        /**
        Sets the priority value for this conditional formatting rule to "1" so that it will be evaluated before all other rules on the worksheet.

        [MSDN documentation for FormatCondition.SetFirstPriority](https://learn.microsoft.com/en-us/office/vba/api/excel.formatcondition.setfirstpriority).
        */
        void SetFirstPriority();

        /**
        Sets the evaluation order for this conditional formatting rule so it is evaluated after all other rules on the worksheet.

        [MSDN documentation for FormatCondition.SetLastPriority](https://learn.microsoft.com/en-us/office/vba/api/excel.formatcondition.setlastpriority).
        */
        void SetLastPriority();

        // ***** PROPERTIES *****

        /**
        When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application. When used with an object qualifier, this property returns an Application object that represents the creator of the specified object (you can use this property with an OLE Automation object to return the application of that object).

        [MSDN documentation for FormatCondition.Application](https://learn.microsoft.com/en-us/office/vba/api/excel.formatcondition.application).
        */
        wxExcelApplication GetApplication();

        /**
        Returns a Range object specifying the cell range to which the formatting rule is applied. Since Excel 2007.

        [MSDN documentation for FormatCondition.AppliesTo](https://learn.microsoft.com/en-us/office/vba/api/excel.formatcondition.appliesto).
        */
        wxExcelRange GetAppliesTo();


        /**
        Returns a Borders collection that represents the borders of a style or a range of cells (including a range defined as part of a conditional format).

        [MSDN documentation for FormatCondition.Borders](https://learn.microsoft.com/en-us/office/vba/api/excel.formatcondition.borders).
        */
        wxExcelBorders GetBorders();

        /**
        Returns a Font object that represents the font of the specified object.

        [MSDN documentation for FormatCondition.Font](https://learn.microsoft.com/en-us/office/vba/api/excel.formatcondition.font).
        */
        wxExcelFont GetFont();

        /**
        Returns the value or expression associated with the conditional format or data validation. Can be a constant value, a string value, a cell reference, or a formula.

        [MSDN documentation for FormatCondition.Formula1](https://learn.microsoft.com/en-us/office/vba/api/excel.formatcondition.formula1).
        */
        wxString GetFormula1();

        /**
        Returns the value or expression associated with the second part of a conditional format or data validation. Used only when the data validation conditional format Operator property is xlBetween or xlNotBetween. Can be a constant value, a string value, a cell reference, or a formula.

        [MSDN documentation for FormatCondition.Formula2](https://learn.microsoft.com/en-us/office/vba/api/excel.formatcondition.formula2).
        */
        wxString GetFormula2();

        /**
        Returns an Interior object that represents the interior of the specified object.

        [MSDN documentation for FormatCondition.Interior](https://learn.microsoft.com/en-us/office/vba/api/excel.formatcondition.interior).
        */
        wxExcelInterior GetInterior();

        /**
        Returns the number format applied to a cell if the conditional formatting rule evaluates to True. Read/write Variant. Since Excel 2007.

        The name is spelled Numberformat rather than NumberFormat as in Excel, because the Windows headers define GetNumberFormat as a macro, which would otherwise rename the method.

        [MSDN documentation for FormatCondition.NumberFormat](https://learn.microsoft.com/en-us/office/vba/api/excel.formatcondition.numberformat).
        */
        wxString GetNumberformat();

        /**
        Sets the number format applied to a cell if the conditional formatting rule evaluates to True. Read/write Variant. Since Excel 2007.

        The name is spelled Numberformat rather than NumberFormat as in Excel, because the Windows headers define GetNumberFormat as a macro, which would otherwise rename the method.

        [MSDN documentation for FormatCondition.NumberFormat](https://learn.microsoft.com/en-us/office/vba/api/excel.formatcondition.numberformat).
        */
        void SetNumberformat(const wxString& numberFormat);

        /**
        Returns a Long value that represents the operator for the conditional format.

        [MSDN documentation for FormatCondition.Operator](https://learn.microsoft.com/en-us/office/vba/api/excel.formatcondition.operator).
        */
        XlFormatConditionOperator GetOperator();

        /**
        Returns the parent object for the specified object.

        [MSDN documentation for FormatCondition.Parent](https://learn.microsoft.com/en-us/office/vba/api/excel.formatcondition.parent).
        */
        wxExcelObject GetParent();

        /**
        Returns the priority value of the conditional formatting rule. The priority determines the order of evaluation when multiple conditional formatting rules exist in a worksheet. Since Excel 2007.

        [MSDN documentation for FormatCondition.Priority](https://learn.microsoft.com/en-us/office/vba/api/excel.formatcondition.priority).
        */
        long GetPriority();

        /**
        Sets the priority value of the conditional formatting rule. The priority determines the order of evaluation when multiple conditional formatting rules exist in a worksheet. Since Excel 2007.

        [MSDN documentation for FormatCondition.Priority](https://learn.microsoft.com/en-us/office/vba/api/excel.formatcondition.priority).
        */
        void SetPriority(long priority);

        /**
        Returns a Boolean value indicating if the conditional format is being applied to a PivotTable chart. Since Excel 2007.

        [MSDN documentation for FormatCondition.PTCondition](https://learn.microsoft.com/en-us/office/vba/api/excel.formatcondition.ptcondition).
        */
        bool GetPTCondition();

        /**
        Returns one of the constants of the XlPivotConditionScope enumeration, which determines the scope of the conditional format when it is applied to a PivotTable chart. Since Excel 2007.

        [MSDN documentation for FormatCondition.ScopeType](https://learn.microsoft.com/en-us/office/vba/api/excel.formatcondition.scopetype).
        */
        XlPivotConditionScope GetScopeType();

        /**
        Sets one of the constants of the XlPivotConditionScope enumeration, which determines the scope of the conditional format when it is applied to a PivotTable chart. Since Excel 2007.

        [MSDN documentation for FormatCondition.ScopeType](https://learn.microsoft.com/en-us/office/vba/api/excel.formatcondition.scopetype).
        */
        void SetScopeType(XlPivotConditionScope scopeType);

        /**
        Returns a Boolean value that determines if additional formatting rules on the cell should be evaluated if the current rule evaluates to True. Since Excel 2007.

        [MSDN documentation for FormatCondition.StopIfTrue](https://learn.microsoft.com/en-us/office/vba/api/excel.formatcondition.stopiftrue).
        */
        bool GetStopIfTrue();

        /**
        Sets a Boolean value that determines if additional formatting rules on the cell should be evaluated if the current rule evaluates to True. Since Excel 2007.

        [MSDN documentation for FormatCondition.StopIfTrue](https://learn.microsoft.com/en-us/office/vba/api/excel.formatcondition.stopiftrue).
        */
        void SetStopIfTrue(bool stopIfTrue);

        /**
        Returns a String value specifying the text string used by the conditional formatting rule. Since Excel 2007.

        [MSDN documentation for FormatCondition.Text](https://learn.microsoft.com/en-us/office/vba/api/excel.formatcondition.text).
        */
        wxString GetText();

        /**
        Sets a String value specifying the text string used by the conditional formatting rule. Since Excel 2007.

        [MSDN documentation for FormatCondition.Text](https://learn.microsoft.com/en-us/office/vba/api/excel.formatcondition.text).
        */
        void SetText(const wxString& text);

        /**
        Returns one of the constants of the XlContainsOperator enumeration specifying the text search performed by the conditional formatting rule. Since Excel 2007.

        [MSDN documentation for FormatCondition.TextOperator](https://learn.microsoft.com/en-us/office/vba/api/excel.formatcondition.textoperator).
        */
        XlContainsOperator GetTextOperator();

        /**
        Sets one of the constants of the XlContainsOperator enumeration specifying the text search performed by the conditional formatting rule. Since Excel 2007.

        [MSDN documentation for FormatCondition.TextOperator](https://learn.microsoft.com/en-us/office/vba/api/excel.formatcondition.textoperator).
        */
        void SetTextOperator(XlContainsOperator textOperator);

        /**
        Returns a Long value, containing a xlFormatConditionType value, that represents the object type.

        [MSDN documentation for FormatCondition.Type](https://learn.microsoft.com/en-us/office/vba/api/excel.formatcondition.type).
        */
        XlFormatConditionType GetType();

        /**
        Returns "FormatCondition".
        */
        virtual wxString GetAutoExcelObjectName_() const { return wxS("FormatCondition"); }

    };



    /**
    @brief Represents Microsoft Excel FormatConditions collection.
    */
   class WXDLLIMPEXP_WXAUTOEXCEL wxExcelFormatConditions : public wxExcelObject
    {
    public:
        // ***** METHODS *****

        /**
        Adds a new conditional format.

        [MSDN documentation for FormatConditions.Add](https://learn.microsoft.com/en-us/office/vba/api/excel.formatconditions.add).

        Unfortunately it appears it is impossible to add language-agnostic formula, see [here](http://stackoverflow.com/questions/13247771/excel-macro-inserting-internationally-valid-formula-during-run-time) for workaround.
        */
        wxExcelFormatCondition Add(XlFormatConditionType conditionType, XlFormatConditionOperator* conditionOperator = NULL,
                                   const wxVariant& formula1 = wxNullVariant, const wxVariant& formula2 = wxNullVariant);

        /**
        Returns a new AboveAverage object representing a conditional formatting rule for the specified range.

        [MSDN documentation for FormatConditions.AddAboveAverage](https://learn.microsoft.com/en-us/office/vba/api/excel.formatconditions.addaboveaverage).
        */
        wxExcelAboveAverage AddAboveAverage();

        /**
        Returns a new ColorScale object representing a color scale conditional formatting rule for the selected range.

        [MSDN documentation for FormatConditions.AddColorScale](https://learn.microsoft.com/en-us/office/vba/api/excel.formatconditions.addcolorscale).
        */
        wxExcelColorScale AddColorScale(long colorScaleType);

        /**
        Returns a Databar object representing a data bar conditional formatting rule for the specified range.

        [MSDN documentation for FormatConditions.AddDatabar](https://learn.microsoft.com/en-us/office/vba/api/excel.formatconditions.adddatabar).
        */
        wxExcelDatabar AddDatabar();

        /**
        Returns a new IconSetCondition object which represents an icon set conditional formatting rule for the specified range.

        [MSDN documentation for FormatConditions.AddIconSetCondition](https://learn.microsoft.com/en-us/office/vba/api/excel.formatconditions.addiconsetcondition).
        */
        wxExcelIconSetCondition AddIconSetCondition();

        /**
        Returns a Top10 object representing a conditional formatting rule for the specified range.

        [MSDN documentation for FormatConditions.AddTop10](https://learn.microsoft.com/en-us/office/vba/api/excel.formatconditions.addtop10).
        */
        wxExcelTop10 AddTop10();

        /**
        Returns a new UniqueValues object representing a conditional formatting rule for the specified range.

        [MSDN documentation for FormatConditions.AddUniqueValues](https://learn.microsoft.com/en-us/office/vba/api/excel.formatconditions.adduniquevalues).
        */
        wxExcelUniqueValues AddUniqueValues();

        /**
        Deletes the object.

        [MSDN documentation for FormatConditions.Delete](https://learn.microsoft.com/en-us/office/vba/api/excel.formatconditions.delete).
        */
        void Delete();

        //@{
        /**
        Returns a single object from a collection.

        [MSDN documentation for FormatConditions.Item](https://learn.microsoft.com/en-us/office/vba/api/excel.formatconditions.item).
        */
        wxExcelFormatCondition Item(long index);
        wxExcelFormatCondition operator[](long index);
        //@}
        // ***** PROPERTIES *****

        /**
        Returns a Long value that represents the number of objects in the collection.

        [MSDN documentation for FormatConditions.Count](https://learn.microsoft.com/en-us/office/vba/api/excel.formatconditions.count).
        */
        long GetCount();


        /**
        The iterator over the items, see wxExcelCollectionIterator.
        */
        typedef wxExcelCollectionIterator<wxExcelFormatConditions, wxExcelFormatCondition> iterator;

        //@{
        /**
        Returns the iterator referring to the first item, or past the last item,
        which allows iterating over the items with a range-based for loop,
        see wxExcelCollectionIterator.
        */
        iterator begin() { return iterator(this, 1); }
        iterator end() { return iterator(this, GetCount() + 1); }
        //@}

        /**
        Returns "FormatConditions".
        */
        virtual wxString GetAutoExcelObjectName_() const { return wxS("FormatConditions"); }

    };


} // namespace wxAutoExcel

#endif // WXAUTOEXCEL_USE_CONDFORMAT

#endif // _WXAUTOEXCEL_FORMATCONDITIONS_H
