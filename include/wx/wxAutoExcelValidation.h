/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// License:     MIT license
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

        [MSDN documentation for Validation.Add](https://learn.microsoft.com/en-us/office/vba/api/excel.validation.add).
        */
        void Add(XlDVType type,
                 XlDVAlertStyle* alertStyle = NULL, XlFormatConditionOperator* conditionOperator = NULL,
                 const wxString& formula1 = wxEmptyString, const wxString& formula2 = wxEmptyString);

        /**
        Deletes the object.

        [MSDN documentation for Validation.Delete](https://learn.microsoft.com/en-us/office/vba/api/excel.validation.delete).
        */
        void Delete();

        /**
        Modifies data validation for a range.

        [MSDN documentation for Validation.Modify](https://learn.microsoft.com/en-us/office/vba/api/excel.validation.modify).
        */
        void Modify(XlDVType* type = NULL,
                    XlDVAlertStyle* alertStyle = NULL, XlFormatConditionOperator* conditionOperator = NULL,
                    const wxString& formula1 = wxEmptyString, const wxString& formula2 = wxEmptyString);

        // ***** PROPERTIES *****

        /**
        Returns the validation alert style. Read-only XlDVAlertStyle.

        [MSDN documentation for Validation.AlertStyle](https://learn.microsoft.com/en-us/office/vba/api/excel.validation.alertstyle).
        */
        XlDVAlertStyle GetAlertStyle();

        /**
        Returns the data validation error message.

        [MSDN documentation for Validation.ErrorMessage](https://learn.microsoft.com/en-us/office/vba/api/excel.validation.errormessage).
        */
        wxString GetErrorMessage();

        /**
        Sets the data validation error message.

        [MSDN documentation for Validation.ErrorMessage](https://learn.microsoft.com/en-us/office/vba/api/excel.validation.errormessage).
        */
        void SetErrorMessage(const wxString& errorMessage);

        /**
        Returns the title of the data-validation error dialog box.

        [MSDN documentation for Validation.ErrorTitle](https://learn.microsoft.com/en-us/office/vba/api/excel.validation.errortitle).
        */
        wxString GetErrorTitle();

        /**
        Sets the title of the data-validation error dialog box.

        [MSDN documentation for Validation.ErrorTitle](https://learn.microsoft.com/en-us/office/vba/api/excel.validation.errortitle).
        */
        void SetErrorTitle(const wxString& errorTitle);

        /**
        Returns the value or expression associated with the conditional format or data validation. Can be a constant value, a string value, a cell reference, or a formula.

        [MSDN documentation for Validation.Formula1](https://learn.microsoft.com/en-us/office/vba/api/excel.validation.formula1).
        */
        wxString GetFormula1();

        /**
        Returns the value or expression associated with the second part of a conditional format or data validation. Used only when the data validation conditional format Operator property is xlBetween or xlNotBetween. Can be a constant value, a string value, a cell reference, or a formula.

        [MSDN documentation for Validation.Formula2](https://learn.microsoft.com/en-us/office/vba/api/excel.validation.formula2).
        */
        wxString GetFormula2();

        /**
        True if blank values are permitted by the range data validation.

        [MSDN documentation for Validation.IgnoreBlank](https://learn.microsoft.com/en-us/office/vba/api/excel.validation.ignoreblank).
        */
        bool GetIgnoreBlank();

        /**
        True if blank values are permitted by the range data validation.

        [MSDN documentation for Validation.IgnoreBlank](https://learn.microsoft.com/en-us/office/vba/api/excel.validation.ignoreblank).
        */
        void SetIgnoreBlank(bool ignoreBlank);

        /**
        Returns the description of the Japanese input rules. Can be one of the XlIMEMode constants listed in the following table.

        [MSDN documentation for Validation.IMEMode](https://learn.microsoft.com/en-us/office/vba/api/excel.validation.imemode).
        */
        XlIMEMode  GetIMEMode();

        /**
        Sets the description of the Japanese input rules. Can be one of the XlIMEMode constants listed in the following table.

        [MSDN documentation for Validation.IMEMode](https://learn.microsoft.com/en-us/office/vba/api/excel.validation.imemode).
        */
        void SetIMEMode(XlIMEMode  iMEMode);

        /**
        True if data validation displays a drop-down list that contains acceptable values.

        [MSDN documentation for Validation.InCellDropdown](https://learn.microsoft.com/en-us/office/vba/api/excel.validation.incelldropdown).
        */
        bool GetInCellDropdown();

        /**
        True if data validation displays a drop-down list that contains acceptable values.

        [MSDN documentation for Validation.InCellDropdown](https://learn.microsoft.com/en-us/office/vba/api/excel.validation.incelldropdown).
        */
        void SetInCellDropdown(bool inCellDropdown);

        /**
        Returns the data validation input message.

        [MSDN documentation for Validation.InputMessage](https://learn.microsoft.com/en-us/office/vba/api/excel.validation.inputmessage).
        */
        wxString GetInputMessage();

        /**
        Sets the data validation input message.

        [MSDN documentation for Validation.InputMessage](https://learn.microsoft.com/en-us/office/vba/api/excel.validation.inputmessage).
        */
        void SetInputMessage(const wxString& inputMessage);

        /**
        Returns the title of the data-validation input dialog box.

        [MSDN documentation for Validation.InputTitle](https://learn.microsoft.com/en-us/office/vba/api/excel.validation.inputtitle).
        */
        wxString GetInputTitle();

        /**
        Sets the title of the data-validation input dialog box.

        [MSDN documentation for Validation.InputTitle](https://learn.microsoft.com/en-us/office/vba/api/excel.validation.inputtitle).
        */
        void SetInputTitle(const wxString& inputTitle);

        /**
        Returns a Long value that represents the operator for the data validation.

        [MSDN documentation for Validation.Operator](https://learn.microsoft.com/en-us/office/vba/api/excel.validation.operator).
        */
        long GetOperator();

        /**
        True if the data validation error message will be displayed whenever the user enters invalid data.

        [MSDN documentation for Validation.ShowError](https://learn.microsoft.com/en-us/office/vba/api/excel.validation.showerror).
        */
        bool GetShowError();

        /**
        True if the data validation error message will be displayed whenever the user enters invalid data.

        [MSDN documentation for Validation.ShowError](https://learn.microsoft.com/en-us/office/vba/api/excel.validation.showerror).
        */
        void SetShowError(bool showError);

        /**
        True if the data validation input message will be displayed whenever the user selects a cell in the data validation range.

        [MSDN documentation for Validation.ShowInput](https://learn.microsoft.com/en-us/office/vba/api/excel.validation.showinput).
        */
        bool GetShowInput();

        /**
        True if the data validation input message will be displayed whenever the user selects a cell in the data validation range.

        [MSDN documentation for Validation.ShowInput](https://learn.microsoft.com/en-us/office/vba/api/excel.validation.showinput).
        */
        void SetShowInput(bool showInput);

        /**
        Returns a Long value, containing a XlDVType constant, that represents the data type validation for a range.

        [MSDN documentation for Validation.Type](https://learn.microsoft.com/en-us/office/vba/api/excel.validation.type).
        */
        XlDVType GetType();

        /**
        Returns a Boolean value that indicates if all the validation criteria are met (that is, if the range contains valid data).

        [MSDN documentation for Validation.Value](https://learn.microsoft.com/en-us/office/vba/api/excel.validation.value).
        */
        bool GetValue();

        /**
        Returns "Validation".
        */
        virtual wxString GetAutoExcelObjectName_() const { return wxS("Validation"); }
    };


} // namespace wxAutoExcel

#endif //_WXAUTOEXCEL_VALIDATION_H
