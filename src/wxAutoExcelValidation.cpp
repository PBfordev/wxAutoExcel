/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Modified by:
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
/////////////////////////////////////////////////////////////////////////////


#include "wx/wxAutoExcel_prec.h"

#include "wx/wxAutoExcelValidation.h"

#include "wx/wxAutoExcel_private.h"

namespace wxAutoExcel {

// ***** class wxExcelValidation METHODS *****

void wxExcelValidation::Add(XlDVType type, XlDVAlertStyle* alertStyle,
                            XlFormatConditionOperator* conditionOperator,
                            const wxString& formula1, const wxString& formula2)
{
    wxVariantVector args;

    args.push_back(wxVariant((long)type, wxS("Type")));

    WXAUTOEXCEL_OPTIONALCPP_TO_OPTIONALVARIANT_NAME_VECTOR(AlertStyle, ((long*)alertStyle), args);
    WXAUTOEXCEL_OPTIONALCPP_TO_OPTIONALVARIANT_NAME_VECTOR(Operator, ((long*)conditionOperator), args);
    WXAUTOEXCEL_OPTIONALCPPSTR_TO_OPTIONALVARIANT_NAME_VECTOR(Formula1, formula1, args);
    WXAUTOEXCEL_OPTIONALCPPSTR_TO_OPTIONALVARIANT_NAME_VECTOR(Formula2, formula2, args);

    WXAUTOEXCEL_CALL_METHODARR_RET("Add", args, "null");
}

void wxExcelValidation::Delete()
{
    WXAUTOEXCEL_CALL_METHOD0_RET("Delete", "null");
}

void wxExcelValidation::Modify(XlDVType* type, XlDVAlertStyle* alertStyle,
                               XlFormatConditionOperator* conditionOperator,
                               const wxString& formula1, const wxString& formula2)
{
    wxVariantVector args;

    WXAUTOEXCEL_OPTIONALCPP_TO_OPTIONALVARIANT_NAME_VECTOR(Type, type, args);
    WXAUTOEXCEL_OPTIONALCPP_TO_OPTIONALVARIANT_NAME_VECTOR(AlertStyle, ((long*)alertStyle), args);
    WXAUTOEXCEL_OPTIONALCPP_TO_OPTIONALVARIANT_NAME_VECTOR(Operator, ((long*)conditionOperator), args);
    WXAUTOEXCEL_OPTIONALCPPSTR_TO_OPTIONALVARIANT_NAME_VECTOR(Formula1, formula1, args);
    WXAUTOEXCEL_OPTIONALCPPSTR_TO_OPTIONALVARIANT_NAME_VECTOR(Formula2, formula2, args);

    WXAUTOEXCEL_CALL_METHODARR_RET("Modify", args, "null");
}

// ***** class wxExcelValidation PROPERTIES *****

XlDVAlertStyle wxExcelValidation::GetAlertStyle()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("AlertStyle", XlDVAlertStyle, xlValidAlertInformation);
}


wxString wxExcelValidation::GetErrorMessage()
{
    WXAUTOEXCEL_PROPERTY_STRING_GET0("ErrorMessage");
}

void wxExcelValidation::SetErrorMessage(const wxString& errorMessage)
{
    InvokePutProperty(wxS("ErrorMessage"), errorMessage);
}

wxString wxExcelValidation::GetErrorTitle()
{
    WXAUTOEXCEL_PROPERTY_STRING_GET0("ErrorTitle");
}

void wxExcelValidation::SetErrorTitle(const wxString& errorTitle)
{
    InvokePutProperty(wxS("ErrorTitle"), errorTitle);
}

wxString wxExcelValidation::GetFormula1()
{
    WXAUTOEXCEL_PROPERTY_STRING_GET0("Formula1");
}

wxString wxExcelValidation::GetFormula2()
{
    WXAUTOEXCEL_PROPERTY_STRING_GET0("Formula2");
}

bool wxExcelValidation::GetIgnoreBlank()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("IgnoreBlank");
}

void wxExcelValidation::SetIgnoreBlank(bool ignoreBlank)
{
    InvokePutProperty(wxS("IgnoreBlank"), ignoreBlank);
}

XlIMEMode  wxExcelValidation::GetIMEMode()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("IMEMode", XlIMEMode , xlIMEModeOff);
}

void wxExcelValidation::SetIMEMode(XlIMEMode  iMEMode)
{
    InvokePutProperty(wxS("IMEMode"), (long)iMEMode);
}

bool wxExcelValidation::GetInCellDropdown()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("InCellDropdown");
}

void wxExcelValidation::SetInCellDropdown(bool inCellDropdown)
{
    InvokePutProperty(wxS("InCellDropdown"), inCellDropdown);
}

wxString wxExcelValidation::GetInputMessage()
{
    WXAUTOEXCEL_PROPERTY_STRING_GET0("InputMessage");
}

void wxExcelValidation::SetInputMessage(const wxString& inputMessage)
{
    InvokePutProperty(wxS("InputMessage"), inputMessage);
}

wxString wxExcelValidation::GetInputTitle()
{
    WXAUTOEXCEL_PROPERTY_STRING_GET0("InputTitle");
}

void wxExcelValidation::SetInputTitle(const wxString& inputTitle)
{
    InvokePutProperty(wxS("InputTitle"), inputTitle);
}

long wxExcelValidation::GetOperator()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("Operator");
}


bool wxExcelValidation::GetShowError()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("ShowError");
}

void wxExcelValidation::SetShowError(bool showError)
{
    InvokePutProperty(wxS("ShowError"), showError);
}

bool wxExcelValidation::GetShowInput()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("ShowInput");
}

void wxExcelValidation::SetShowInput(bool showInput)
{
    InvokePutProperty(wxS("ShowInput"), showInput);
}

XlDVType wxExcelValidation::GetType()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("Type", XlDVType, xlValidateDecimal);
}

bool wxExcelValidation::GetValue()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("Value");
}


} // namespace wxAutoExcel
