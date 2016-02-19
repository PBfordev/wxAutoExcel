/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Modified by:
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
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

        [MSDN documentation for FormatCondition.Delete](http://msdn.microsoft.com/en-us/library/office/ff196592(v=office.14).aspx).
        */
        void Delete();

        /**
        Modifies an existing conditional format.

        [MSDN documentation for FormatCondition.Modify](http://msdn.microsoft.com/en-us/library/office/ff837106(v=office.14).aspx).
        */
        void Modify(XlFormatConditionType conditionType, XlFormatConditionOperator* conditionOperator = NULL,
                    const wxString& formula1 = wxEmptyString, const wxString& formula2 = wxEmptyString);

        /**
        Sets the cell range to which this formatting rule applies.

        [MSDN documentation for FormatCondition.ModifyAppliesToRange](http://msdn.microsoft.com/en-us/library/office/ff837422(v=office.14).aspx).
        */
        void ModifyAppliesToRange(wxExcelRange range);

        /**
        Sets the priority value for this conditional formatting rule to "1" so that it will be evaluated before all other rules on the worksheet.

        [MSDN documentation for FormatCondition.SetFirstPriority](http://msdn.microsoft.com/en-us/library/office/ff820833(v=office.14).aspx).
        */
        void SetFirstPriority();

        /**
        Sets the evaluation order for this conditional formatting rule so it is evaluated after all other rules on the worksheet.

        [MSDN documentation for FormatCondition.SetLastPriority](http://msdn.microsoft.com/en-us/library/office/ff841221(v=office.14).aspx).
        */
        void SetLastPriority();

        // ***** PROPERTIES *****

        /**
        When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application. When used with an object qualifier, this property returns an Application object that represents the creator of the specified object (you can use this property with an OLE Automation object to return the application of that object).

        [MSDN documentation for FormatCondition.Application](http://msdn.microsoft.com/en-us/library/office/ff197842(v=office.14).aspx).
        */
        wxExcelApplication GetApplication();

        /**
        Returns a Range object specifying the cell range to which the formatting rule is applied. Since Excel 2007.

        [MSDN documentation for FormatCondition.AppliesTo](http://msdn.microsoft.com/en-us/library/office/ff839719(v=office.14).aspx).
        */
        wxExcelRange GetAppliesTo();
     

        /**
        Returns a Borders collection that represents the borders of a style or a range of cells (including a range defined as part of a conditional format).

        [MSDN documentation for FormatCondition.Borders](http://msdn.microsoft.com/en-us/library/office/ff196030(v=office.14).aspx).
        */
        wxExcelBorders GetBorders();
       
        /**
        Returns a Font object that represents the font of the specified object.

        [MSDN documentation for FormatCondition.Font](http://msdn.microsoft.com/en-us/library/office/ff193040(v=office.14).aspx).
        */
        wxExcelFont GetFont();

        /**
        Returns the value or expression associated with the conditional format or data validation. Can be a constant value, a string value, a cell reference, or a formula.

        [MSDN documentation for FormatCondition.Formula1](http://msdn.microsoft.com/en-us/library/office/ff841065(v=office.14).aspx).
        */
        wxString GetFormula1();

        /**
        Returns the value or expression associated with the second part of a conditional format or data validation. Used only when the data validation conditional format Operator property is xlBetween or xlNotBetween. Can be a constant value, a string value, a cell reference, or a formula.

        [MSDN documentation for FormatCondition.Formula2](http://msdn.microsoft.com/en-us/library/office/ff195641(v=office.14).aspx).
        */
        wxString GetFormula2();

        /**
        Returns an Interior object that represents the interior of the specified object.

        [MSDN documentation for FormatCondition.Interior](http://msdn.microsoft.com/en-us/library/office/ff196979(v=office.14).aspx).
        */
        wxExcelInterior GetInterior();

        /**
        Returns the number format applied to a cell if the conditional formatting rule evaluates to True. Read/write Variant. Since Excel 2007.

        [MSDN documentation for FormatCondition.NumberFormat](http://msdn.microsoft.com/en-us/library/office/ff820867(v=office.14).aspx).
        */
        wxString GetNumberFormat();

        /**
        Sets the number format applied to a cell if the conditional formatting rule evaluates to True. Read/write Variant. Since Excel 2007.

        [MSDN documentation for FormatCondition.NumberFormat](http://msdn.microsoft.com/en-us/library/office/ff820867(v=office.14).aspx).
        */
        void SetNumberFormat(const wxString& numberFormat);

        /**
        Returns a Long value that represents the operator for the conditional format.

        [MSDN documentation for FormatCondition.Operator](http://msdn.microsoft.com/en-us/library/office/ff836182(v=office.14).aspx).
        */
        XlFormatConditionOperator GetOperator();

        /**
        Returns the parent object for the specified object.

        [MSDN documentation for FormatCondition.Parent](http://msdn.microsoft.com/en-us/library/office/ff193291(v=office.14).aspx).
        */
        wxExcelObject GetParent();

        /**
        Returns the priority value of the conditional formatting rule. The priority determines the order of evaluation when multiple conditional formatting rules exist in a worksheet. Since Excel 2007.

        [MSDN documentation for FormatCondition.Priority](http://msdn.microsoft.com/en-us/library/office/ff195509(v=office.14).aspx).
        */
        long GetPriority();

        /**
        Sets the priority value of the conditional formatting rule. The priority determines the order of evaluation when multiple conditional formatting rules exist in a worksheet. Since Excel 2007.

        [MSDN documentation for FormatCondition.Priority](http://msdn.microsoft.com/en-us/library/office/ff195509(v=office.14).aspx).
        */
        void SetPriority(long priority);

        /**
        Returns a Boolean value indicating if the conditional format is being applied to a PivotTable chart. Since Excel 2007.

        [MSDN documentation for FormatCondition.PTCondition](http://msdn.microsoft.com/en-us/library/office/ff195159(v=office.14).aspx).
        */
        bool GetPTCondition();

        /**
        Returns one of the constants of the XlPivotConditionScope enumeration, which determines the scope of the conditional format when it is applied to a PivotTable chart. Since Excel 2007.

        [MSDN documentation for FormatCondition.ScopeType](http://msdn.microsoft.com/en-us/library/office/ff193933(v=office.14).aspx).
        */
        XlPivotConditionScope GetScopeType();

        /**
        Sets one of the constants of the XlPivotConditionScope enumeration, which determines the scope of the conditional format when it is applied to a PivotTable chart. Since Excel 2007.

        [MSDN documentation for FormatCondition.ScopeType](http://msdn.microsoft.com/en-us/library/office/ff193933(v=office.14).aspx).
        */
        void SetScopeType(XlPivotConditionScope scopeType);

        /**
        Returns a Boolean value that determines if additional formatting rules on the cell should be evaluated if the current rule evaluates to True. Since Excel 2007.

        [MSDN documentation for FormatCondition.StopIfTrue](http://msdn.microsoft.com/en-us/library/office/ff838861(v=office.14).aspx).
        */
        bool GetStopIfTrue();

        /**
        Sets a Boolean value that determines if additional formatting rules on the cell should be evaluated if the current rule evaluates to True. Since Excel 2007.

        [MSDN documentation for FormatCondition.StopIfTrue](http://msdn.microsoft.com/en-us/library/office/ff838861(v=office.14).aspx).
        */
        void SetStopIfTrue(bool stopIfTrue);

        /**
        Returns a String value specifying the text string used by the conditional formatting rule. Since Excel 2007.

        [MSDN documentation for FormatCondition.Text](http://msdn.microsoft.com/en-us/library/office/ff194461(v=office.14).aspx).
        */
        wxString GetText();

        /**
        Sets a String value specifying the text string used by the conditional formatting rule. Since Excel 2007.

        [MSDN documentation for FormatCondition.Text](http://msdn.microsoft.com/en-us/library/office/ff194461(v=office.14).aspx).
        */
        void SetText(const wxString& text);

        /**
        Returns one of the constants of the XlContainsOperator enumeration specifying the text search performed by the conditional formatting rule. Since Excel 2007.

        [MSDN documentation for FormatCondition.TextOperator](http://msdn.microsoft.com/en-us/library/office/ff197985(v=office.14).aspx).
        */
        XlContainsOperator GetTextOperator();

        /**
        Sets one of the constants of the XlContainsOperator enumeration specifying the text search performed by the conditional formatting rule. Since Excel 2007.

        [MSDN documentation for FormatCondition.TextOperator](http://msdn.microsoft.com/en-us/library/office/ff197985(v=office.14).aspx).
        */
        void SetTextOperator(XlContainsOperator textOperator);

        /**
        Returns a Long value, containing a xlFormatConditionType value, that represents the object type.

        [MSDN documentation for FormatCondition.Type](http://msdn.microsoft.com/en-us/library/office/ff840778(v=office.14).aspx).
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

        [MSDN documentation for FormatConditions.Add](http://msdn.microsoft.com/en-us/library/bb211780.aspx).
        */
        wxExcelFormatCondition Add(XlFormatConditionType conditionType, XlFormatConditionOperator* conditionOperator = NULL,
                                   const wxString& formula1 = wxEmptyString, const wxString& formula2 = wxEmptyString);

        /**
        Returns a new AboveAverage object representing a conditional formatting rule for the specified range.

        [MSDN documentation for FormatConditions.AddAboveAverage](http://msdn.microsoft.com/en-us/library/bb238829.aspx).
        */
        wxExcelAboveAverage AddAboveAverage();

        /**
        Returns a new ColorScale object representing a color scale conditional formatting rule for the selected range.

        [MSDN documentation for FormatConditions.AddColorScale](http://msdn.microsoft.com/en-us/library/bb238831.aspx).
        */
        wxExcelColorScale AddColorScale(long colorScaleType);

        /**
        Returns a Databar object representing a data bar conditional formatting rule for the specified range.

        [MSDN documentation for FormatConditions.AddDatabar](http://msdn.microsoft.com/en-us/library/bb238834.aspx).
        */
        wxExcelDatabar AddDatabar();

        /**
        Returns a new IconSetCondition object which represents an icon set conditional formatting rule for the specified range.

        [MSDN documentation for FormatConditions.AddIconSetCondition](http://msdn.microsoft.com/en-us/library/bb238836.aspx).
        */
        wxExcelIconSetCondition AddIconSetCondition();

        /**
        Returns a Top10 object representing a conditional formatting rule for the specified range.

        [MSDN documentation for FormatConditions.AddTop10](http://msdn.microsoft.com/en-us/library/bb238838.aspx).
        */
        wxExcelTop10 AddTop10();

        /**
        Returns a new UniqueValues object representing a conditional formatting rule for the specified range.

        [MSDN documentation for FormatConditions.AddUniqueValues](http://msdn.microsoft.com/en-us/library/bb238839.aspx).
        */
        wxExcelUniqueValues AddUniqueValues();

        /**
        Deletes the object.

        [MSDN documentation for FormatConditions.Delete](http://msdn.microsoft.com/en-us/library/bb211783.aspx).
        */
        void Delete();

        //@{
        /**
        Returns a single object from a collection.

        [MSDN documentation for FormatConditions.Item](http://msdn.microsoft.com/en-us/library/bb211786.aspx).
        */
        wxExcelFormatCondition Item(long index);
        wxExcelFormatCondition operator[](long index);        
        //@}
        // ***** PROPERTIES *****
   
        /**
        Returns a Long value that represents the number of objects in the collection.

        [MSDN documentation for FormatConditions.Count](http://msdn.microsoft.com/en-us/library/bb148456.aspx).
        */
        long GetCount();


        /**
        Returns "FormatConditions".
        */
        virtual wxString GetAutoExcelObjectName_() const { return wxS("FormatConditions"); }

    };


} // namespace wxAutoExcel

#endif // WXAUTOEXCEL_USE_CONDFORMAT

#endif // _WXAUTOEXCEL_FORMATCONDITIONS_H
