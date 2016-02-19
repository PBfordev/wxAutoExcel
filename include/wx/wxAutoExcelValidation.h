/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Modified by:
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
/////////////////////////////////////////////////////////////////////////////


#ifndef _WXAUTOEXCEL_VALIDATION_H
#define _WXAUTOEXCEL_VALIDATION_H

#include "wx/wxAutoExcel_defs.h"

#include "wx/wxAutoExcel_object.h"
#include "wx/wxAutoExcel_enums.h"

class WXDLLIMPEXP_FWD_CORE wxColour;

namespace wxAutoExcel {


    /**
    @brief Represents Microsoft Excel Validation object.
    */
   class WXDLLIMPEXP_WXAUTOEXCEL wxExcelValidation : public wxExcelObject
    {
    public:        
        // ***** METHODS *****

        /**
        Adds data validation to the specified range.

        [MSDN documentation for Validation.Add](http://msdn.microsoft.com/en-us/library/bb179084.aspx).
        */
        void Add(XlDVType type, 
                 XlDVAlertStyle* alertStyle = NULL, XlFormatConditionOperator* conditionOperator = NULL, 
                 const wxString& formula1 = wxEmptyString, const wxString& formula2 = wxEmptyString);

        /**
        Deletes the object.

        [MSDN documentation for Validation.Delete](http://msdn.microsoft.com/en-us/library/bb179087.aspx).
        */
        void Delete();

        /**
        Modifies data validation for a range.

        [MSDN documentation for Validation.Modify](http://msdn.microsoft.com/en-us/library/bb179090.aspx).
        */
        void Modify(XlDVType* type = NULL, 
                    XlDVAlertStyle* alertStyle = NULL, XlFormatConditionOperator* conditionOperator = NULL, 
                    const wxString& formula1 = wxEmptyString, const wxString& formula2 = wxEmptyString);

        // ***** PROPERTIES *****

        /**
        Returns the validation alert style. Read-only XlDVAlertStyle.

        [MSDN documentation for Validation.AlertStyle](http://msdn.microsoft.com/en-us/library/bb220825.aspx).
        */
        XlDVAlertStyle GetAlertStyle();

        /**
        Returns the data validation error message.

        [MSDN documentation for Validation.ErrorMessage](http://msdn.microsoft.com/en-us/library/bb208476.aspx).
        */
        wxString GetErrorMessage();

        /**
        Sets the data validation error message.

        [MSDN documentation for Validation.ErrorMessage](http://msdn.microsoft.com/en-us/library/bb208476.aspx).
        */
        void SetErrorMessage(const wxString& errorMessage);

        /**
        Returns the title of the data-validation error dialog box.

        [MSDN documentation for Validation.ErrorTitle](http://msdn.microsoft.com/en-us/library/bb208482.aspx).
        */
        wxString GetErrorTitle();

        /**
        Sets the title of the data-validation error dialog box.

        [MSDN documentation for Validation.ErrorTitle](http://msdn.microsoft.com/en-us/library/bb208482.aspx).
        */
        void SetErrorTitle(const wxString& errorTitle);

        /**
        Returns the value or expression associated with the conditional format or data validation. Can be a constant value, a string value, a cell reference, or a formula.

        [MSDN documentation for Validation.Formula1](http://msdn.microsoft.com/en-us/library/bb214166.aspx).
        */
        wxString GetFormula1();

        /**
        Returns the value or expression associated with the second part of a conditional format or data validation. Used only when the data validation conditional format Operator property is xlBetween or xlNotBetween. Can be a constant value, a string value, a cell reference, or a formula.

        [MSDN documentation for Validation.Formula2](http://msdn.microsoft.com/en-us/library/bb214171.aspx).
        */
        wxString GetFormula2();

        /**
        True if blank values are permitted by the range data validation.

        [MSDN documentation for Validation.IgnoreBlank](http://msdn.microsoft.com/en-us/library/bb177589.aspx).
        */
        bool GetIgnoreBlank();

        /**
        True if blank values are permitted by the range data validation.

        [MSDN documentation for Validation.IgnoreBlank](http://msdn.microsoft.com/en-us/library/bb177589.aspx).
        */
        void SetIgnoreBlank(bool ignoreBlank);

        /**
        Returns the description of the Japanese input rules. Can be one of the XlIMEMode constants listed in the following table.

        [MSDN documentation for Validation.IMEMode](http://msdn.microsoft.com/en-us/library/bb177603.aspx).
        */
        XlIMEMode  GetIMEMode();

        /**
        Sets the description of the Japanese input rules. Can be one of the XlIMEMode constants listed in the following table.

        [MSDN documentation for Validation.IMEMode](http://msdn.microsoft.com/en-us/library/bb177603.aspx).
        */
        void SetIMEMode(XlIMEMode  iMEMode);

        /**
        True if data validation displays a drop-down list that contains acceptable values.

        [MSDN documentation for Validation.InCellDropdown](http://msdn.microsoft.com/en-us/library/bb177609.aspx).
        */
        bool GetInCellDropdown();

        /**
        True if data validation displays a drop-down list that contains acceptable values.

        [MSDN documentation for Validation.InCellDropdown](http://msdn.microsoft.com/en-us/library/bb177609.aspx).
        */
        void SetInCellDropdown(bool inCellDropdown);

        /**
        Returns the data validation input message.

        [MSDN documentation for Validation.InputMessage](http://msdn.microsoft.com/en-us/library/bb177642.aspx).
        */
        wxString GetInputMessage();

        /**
        Sets the data validation input message.

        [MSDN documentation for Validation.InputMessage](http://msdn.microsoft.com/en-us/library/bb177642.aspx).
        */
        void SetInputMessage(const wxString& inputMessage);

        /**
        Returns the title of the data-validation input dialog box.

        [MSDN documentation for Validation.InputTitle](http://msdn.microsoft.com/en-us/library/bb177646.aspx).
        */
        wxString GetInputTitle();

        /**
        Sets the title of the data-validation input dialog box.

        [MSDN documentation for Validation.InputTitle](http://msdn.microsoft.com/en-us/library/bb177646.aspx).
        */
        void SetInputTitle(const wxString& inputTitle);

        /**
        Returns a Long value that represents the operator for the data validation.

        [MSDN documentation for Validation.Operator](http://msdn.microsoft.com/en-us/library/bb214178.aspx).
        */
        long GetOperator();

        /**
        True if the data validation error message will be displayed whenever the user enters invalid data.

        [MSDN documentation for Validation.ShowError](http://msdn.microsoft.com/en-us/library/bb221688.aspx).
        */
        bool GetShowError();

        /**
        True if the data validation error message will be displayed whenever the user enters invalid data.

        [MSDN documentation for Validation.ShowError](http://msdn.microsoft.com/en-us/library/bb221688.aspx).
        */
        void SetShowError(bool showError);

        /**
        True if the data validation input message will be displayed whenever the user selects a cell in the data validation range.

        [MSDN documentation for Validation.ShowInput](http://msdn.microsoft.com/en-us/library/bb209220.aspx).
        */
        bool GetShowInput();

        /**
        True if the data validation input message will be displayed whenever the user selects a cell in the data validation range.

        [MSDN documentation for Validation.ShowInput](http://msdn.microsoft.com/en-us/library/bb209220.aspx).
        */
        void SetShowInput(bool showInput);

        /**
        Returns a Long value, containing a XlDVType constant, that represents the data type validation for a range.

        [MSDN documentation for Validation.Type](http://msdn.microsoft.com/en-us/library/bb215164.aspx).
        */
        XlDVType GetType();

        /**
        Returns a Boolean value that indicates if all the validation criteria are met (that is, if the range contains valid data).

        [MSDN documentation for Validation.Value](http://msdn.microsoft.com/en-us/library/bb215168.aspx).
        */
        bool GetValue();

        /**
        Returns "Validation".
        */
        virtual wxString GetAutoExcelObjectName_() const { return wxS("Validation"); }
    };


} // namespace wxAutoExcel

#endif //_WXAUTOEXCEL_VALIDATION_H
