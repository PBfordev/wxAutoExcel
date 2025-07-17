/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
/////////////////////////////////////////////////////////////////////////////


#include "wx/wxAutoExcel_prec.h"

#include "wx/wxAutoExcelNames.h"

#include "wx/wxAutoExcelRange.h"

#include "wx/wxAutoExcel_private.h"

namespace wxAutoExcel {

// ***** class wxExcelName METHODS *****

void wxExcelName::Delete()
{
    WXAUTOEXCEL_CALL_METHOD0_RET("Delete", "null");
}

// ***** class wxExcelName PROPERTIES *****

wxString wxExcelName::GetCategory()
{
    WXAUTOEXCEL_PROPERTY_STRING_GET0("Category");
}

void wxExcelName::SetCategory(const wxString& category)
{
    InvokePutProperty(wxS("Category"), category);
}

wxString wxExcelName::GetCategoryLocal()
{
    WXAUTOEXCEL_PROPERTY_STRING_GET0("CategoryLocal");
}

void wxExcelName::SetCategoryLocal(const wxString& categoryLocal)
{
    InvokePutProperty(wxS("CategoryLocal"), categoryLocal);
}

wxString wxExcelName::GetComment()
{
    WXAUTOEXCEL_PROPERTY_STRING_GET0("Comment");
}

void wxExcelName::SetComment(const wxString& comment)
{
    InvokePutProperty(wxS("Comment"), comment);
}

long wxExcelName::GetIndex()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("Index");
}

XlXLMMacroType wxExcelName::GetMacroType()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("MacroType", XlXLMMacroType, xlFunction);
}

void wxExcelName::SetMacroType(XlXLMMacroType macroType)
{
    InvokePutProperty(wxS("MacroType"), (long)macroType);
}

wxString wxExcelName::GetName()
{
    WXAUTOEXCEL_PROPERTY_STRING_GET0("Name");
}

void wxExcelName::SetName(const wxString& name)
{
    InvokePutProperty(wxS("Name"), name);
}

wxString wxExcelName::GetNameLocal()
{
    WXAUTOEXCEL_PROPERTY_STRING_GET0("NameLocal");
}

void wxExcelName::SetNameLocal(const wxString& nameLocal)
{
    InvokePutProperty(wxS("NameLocal"), nameLocal);
}

wxString wxExcelName::GetRefersTo()
{
    WXAUTOEXCEL_PROPERTY_STRING_GET0("RefersTo");
}

void wxExcelName::SetRefersTo(const wxString& refersTo)
{
    InvokePutProperty(wxS("RefersTo"), refersTo);
}

wxString wxExcelName::GetRefersToLocal()
{
    WXAUTOEXCEL_PROPERTY_STRING_GET0("RefersToLocal");
}

void wxExcelName::SetRefersToLocal(const wxString& refersToLocal)
{
    InvokePutProperty(wxS("RefersToLocal"), refersToLocal);
}

wxString wxExcelName::GetRefersToR1C1()
{
    WXAUTOEXCEL_PROPERTY_STRING_GET0("RefersToR1C1");
}

void wxExcelName::SetRefersToR1C1(const wxString& refersToR1C1)
{
    InvokePutProperty(wxS("RefersToR1C1"), refersToR1C1);
}

wxString wxExcelName::GetRefersToR1C1Local()
{
    WXAUTOEXCEL_PROPERTY_STRING_GET0("RefersToR1C1Local");
}

void wxExcelName::SetRefersToR1C1Local(const wxString& refersToR1C1Local)
{
    InvokePutProperty(wxS("RefersToR1C1Local"), refersToR1C1Local);
}

wxExcelRange wxExcelName::GetRefersToRange()
{
    wxExcelRange range;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("RefersToRange", range);
}

wxString wxExcelName::GetShortcutKey()
{
    WXAUTOEXCEL_PROPERTY_STRING_GET0("ShortcutKey");
}

void wxExcelName::SetShortcutKey(const wxString& shortcutKey)
{
    InvokePutProperty(wxS("ShortcutKey"), shortcutKey);
}

bool wxExcelName::GetValidWorkbookParameter()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("ValidWorkbookParameter");
}

wxString wxExcelName::GetValue()
{
    WXAUTOEXCEL_PROPERTY_STRING_GET0("Value");
}

void wxExcelName::SetValue(const wxString& value)
{
    InvokePutProperty(wxS("Value"), value);
}

bool wxExcelName::GetVisible()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("Visible");
}

void wxExcelName::SetVisible(bool visible)
{
    InvokePutProperty(wxS("Visible"), visible);
}

bool wxExcelName::GetWorkbookParameter()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("WorkbookParameter");
}

void wxExcelName::SetWorkbookParameter(bool workbookParameter)
{
    InvokePutProperty(wxS("WorkbookParameter"), workbookParameter);
}

// ***** class wxExcelNames METHODS *****

wxExcelName wxExcelNames::Add(const wxString& name, const wxString& refersTo,
                              wxXlTribool visible, long* macroType, const wxString& shortCutKey,
                              const wxString& nameLocal, const wxString& refersToLocal,
                              const wxString& categoryLocal,
                              const wxString& refersToR1C1, const wxString& refersToR1C1Local)
{
    wxASSERT( !name.empty() );

    wxVariantVector args;

    WXAUTOEXCEL_OPTIONALCPPSTR_TO_OPTIONALVARIANT_NAME_VECTOR(Name, name, args);
    WXAUTOEXCEL_OPTIONALCPPSTR_TO_OPTIONALVARIANT_NAME_VECTOR(RefersTo, refersTo, args);
    WXAUTOEXCEL_OPTIONALCPPTBOOL_TO_OPTIONALVARIANT_NAME_VECTOR(Visible, visible, args);
    WXAUTOEXCEL_OPTIONALCPP_TO_OPTIONALVARIANT_NAME_VECTOR(MacroType, macroType, args);
    WXAUTOEXCEL_OPTIONALCPPSTR_TO_OPTIONALVARIANT_NAME_VECTOR(ShortCutKey, shortCutKey, args);
    WXAUTOEXCEL_OPTIONALCPPSTR_TO_OPTIONALVARIANT_NAME_VECTOR(NameLocal, nameLocal, args);
    WXAUTOEXCEL_OPTIONALCPPSTR_TO_OPTIONALVARIANT_NAME_VECTOR(RefersToLocal, refersToLocal, args);
    WXAUTOEXCEL_OPTIONALCPPSTR_TO_OPTIONALVARIANT_NAME_VECTOR(CategoryLocal, categoryLocal, args);
    WXAUTOEXCEL_OPTIONALCPPSTR_TO_OPTIONALVARIANT_NAME_VECTOR(RefersToR1C1, refersToR1C1, args);
    WXAUTOEXCEL_OPTIONALCPPSTR_TO_OPTIONALVARIANT_NAME_VECTOR(RefersToR1C1Local, refersToR1C1Local, args);

    wxExcelName nameObj;

    WXAUTOEXCEL_CALL_METHODARR("Add", args, "void*", nameObj);
    VariantToObject(vResult, &nameObj);
    return nameObj;
}

wxExcelName wxExcelNames::Item(const wxString& index, const wxString& indexLocal,
                               const wxString& refersTo)
{
    WXAUTOEXCEL_OPTIONALCPPSTR_TO_OPTIONALVARIANT_NAME(Index, index);
    WXAUTOEXCEL_OPTIONALCPPSTR_TO_OPTIONALVARIANT_NAME(IndexLocal, indexLocal);
    WXAUTOEXCEL_OPTIONALCPPSTR_TO_OPTIONALVARIANT_NAME(RefersTo, refersTo);

    wxExcelName name;

    WXAUTOEXCEL_CALL_METHOD3("Item", vIndex, vIndexLocal, vRefersTo, "void*", name);
    VariantToObject(vResult, &name);
    return name;
}

wxExcelName wxExcelNames::Item(long index)
{
    wxASSERT( index > 0 );

    wxExcelName name;
    WXAUTOEXCEL_CALL_METHOD1_OBJECT("Item", index, name);
}

wxExcelName wxExcelNames::operator[](long index)
{
    return Item(index);
}

// ***** class wxExcelNames PROPERTIES *****

long wxExcelNames::GetCount()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("Count");
}


} // namespace wxAutoExcel
