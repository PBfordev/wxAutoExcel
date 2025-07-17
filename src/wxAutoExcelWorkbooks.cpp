/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
/////////////////////////////////////////////////////////////////////////////


#include "wx/wxAutoExcel_prec.h"

#include "wx/wxAutoExcelWorkbooks.h"

#include "wx/wxAutoExcelApplication.h"
#include "wx/wxAutoExcelWorkbook.h"
#include "wx/wxAutoExcelRange.h"

#include "wx/wxAutoExcel_private.h"

namespace wxAutoExcel {


wxExcelWorkbook wxExcelWorkbooks::Add(XlWBATemplate sheetType)
{
    return DoAdd(wxVariant((long)sheetType));
}

wxExcelWorkbook wxExcelWorkbooks::Add(const wxString& templateFile)
{
    return DoAdd(wxVariant(templateFile));
}

wxExcelWorkbook wxExcelWorkbooks::DoAdd(const wxVariant& templateFile)
{
    wxExcelWorkbook workbook;
    WXAUTOEXCEL_CALL_METHOD1_OBJECT("Add", templateFile, workbook);
}


bool wxExcelWorkbooks::CanCheckOut(const wxString& fileName)
{
    WXAUTOEXCEL_CALL_METHOD1_BOOL("CanCheckOut", fileName);
}

void wxExcelWorkbooks::CheckOut(const wxString& fileName)
{
     WXAUTOEXCEL_CALL_METHOD1_RET("CheckOut", fileName, "null");
}

bool wxExcelWorkbooks::Close()
{
     WXAUTOEXCEL_CALL_METHOD0_BOOL("Close");
}

wxExcelWorkbook wxExcelWorkbooks::Open(const wxString& fileName, long* updateLinks, wxXlTribool readOnly,
                              long* format,
                              const wxString& password, const wxString& writeResPassword,
                              wxXlTribool ignoreReadOnlyRecommended, XlPlatform* origin,
                              wxXlTribool editable, wxXlTribool notify, long* converter,
                              wxXlTribool addToMru, wxXlTribool local, XlCorruptLoad* corruptLoad)
{
    wxExcelWorkbook result;

    wxVariantVector args;

    WXAUTOEXCEL_OPTIONALCPP_TO_OPTIONALVARIANT_NAME_VECTOR(UpdateLinks, updateLinks, args);
    WXAUTOEXCEL_OPTIONALCPPTBOOL_TO_OPTIONALVARIANT_NAME_VECTOR(ReadOnly, readOnly, args);
    WXAUTOEXCEL_OPTIONALCPP_TO_OPTIONALVARIANT_NAME_VECTOR(Format, format, args);
    WXAUTOEXCEL_OPTIONALCPPSTR_TO_OPTIONALVARIANT_NAME_VECTOR(Password, password, args);
    WXAUTOEXCEL_OPTIONALCPPSTR_TO_OPTIONALVARIANT_NAME_VECTOR(WriteResPassword, writeResPassword, args);
    WXAUTOEXCEL_OPTIONALCPPTBOOL_TO_OPTIONALVARIANT_NAME_VECTOR(IgnoreReadOnlyRecommended, ignoreReadOnlyRecommended, args);
    WXAUTOEXCEL_OPTIONALCPP_TO_OPTIONALVARIANT_NAME_VECTOR(Origin, ((long*)origin), args);
    WXAUTOEXCEL_OPTIONALCPPTBOOL_TO_OPTIONALVARIANT_NAME_VECTOR(Editable, editable, args);
    WXAUTOEXCEL_OPTIONALCPPTBOOL_TO_OPTIONALVARIANT_NAME_VECTOR(Notify, notify, args);
    WXAUTOEXCEL_OPTIONALCPP_TO_OPTIONALVARIANT_NAME_VECTOR(Converter, converter, args);
    WXAUTOEXCEL_OPTIONALCPPTBOOL_TO_OPTIONALVARIANT_NAME_VECTOR(AddToMru, addToMru, args);
    WXAUTOEXCEL_OPTIONALCPPTBOOL_TO_OPTIONALVARIANT_NAME_VECTOR(Local, local, args);
    WXAUTOEXCEL_OPTIONALCPP_TO_OPTIONALVARIANT_NAME_VECTOR(CorruptLoad, ((long*)corruptLoad), args);

    return Open(fileName, args);
}

wxExcelWorkbook wxExcelWorkbooks::Open(const wxString& fileName, const wxVariantVector& optionalArgs)
{
    wxASSERT(!fileName.empty());
    wxExcelWorkbook workbook;

    wxVariantVector args(optionalArgs);

    args.push_back(wxVariant(fileName, wxS("FileName")));
    WXAUTOEXCEL_CALL_METHODARR("Open", args, "void*", workbook);
    VariantToObject(vResult, &workbook);
    return workbook;
}

wxExcelWorkbook wxExcelWorkbooks::OpenDatabase(const wxString& fileName, const wxString& commandText,
                                  XlCmdType commandType, bool backgroundQuery, XlImportDataAs importDataAs)
{
    wxExcelWorkbook workbook;
    WXAUTOEXCEL_OPTIONALCPPSTR_TO_OPTIONALVARIANT(CommandText, commandText);

    WXAUTOEXCEL_CALL_METHOD5("OpenDatabase", fileName, vCommandText, (long)commandType, backgroundQuery, (long)importDataAs, "void*", workbook);
    VariantToObject(vResult, &workbook);
    return workbook;
}


// ***** class wxAutoExcelWorkbooks PROPERTIES *****


long wxExcelWorkbooks::GetCount()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("Count");
}

wxExcelWorkbook wxExcelWorkbooks::GetItem(long index)
{
    wxASSERT( index > 0 );

    wxExcelWorkbook workbook;
    WXAUTOEXCEL_PROPERTY_OBJECT_GET1("Item", index, workbook);
}

wxExcelWorkbook wxExcelWorkbooks::GetItem(const wxString& name)
{
    wxExcelWorkbook workbook;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET1("Item", name, workbook);
}

wxExcelWorkbook wxExcelWorkbooks::operator[](long index)
{
    return GetItem(index);
}

wxExcelWorkbook wxExcelWorkbooks::operator[](const wxString& name)
{
    return GetItem(name);
}


} // namespace wxAutoExcel
