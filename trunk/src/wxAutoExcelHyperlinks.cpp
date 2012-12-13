/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Modified by:
// RCS-ID:      $Id: $
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
/////////////////////////////////////////////////////////////////////////////


#include "wx/wxAutoExcel_prec.h"

#include "wx/wxAutoExcelHyperlinks.h"

#include "wx/wxAutoExcelShape.h"
#include "wx/wxAutoExcelRange.h"

#include "wx/wxAutoExcelPrivate.h"

namespace wxAutoExcel {

// ***** class wxExcelHyperlink METHODS *****

void wxExcelHyperlink::AddToFavorites()
{
    WXAUTOEXCEL_CALL_METHOD0_RET("AddToFavorites", "null");
}

void wxExcelHyperlink::CreateNewDocument(const wxString& fileName, bool editNow, bool overwrite)
{
    WXAUTOEXCEL_CALL_METHOD3_RET("CreateNewDocument", fileName, editNow, overwrite, "null");
}

void wxExcelHyperlink::Delete()
{

    WXAUTOEXCEL_CALL_METHOD0_RET("Delete", "null");
}

void wxExcelHyperlink::Follow(wxXlTribool newWindow, wxXlTribool addHistory,
                              MsoExtraInfoMethod* method, const wxString& headerInfo)
{
    WXAUTOEXCEL_OPTIONALCPPTBOOL_TO_OPTIONALVARIANT_NAME(NewWindow, newWindow);
    WXAUTOEXCEL_OPTIONALCPPTBOOL_TO_OPTIONALVARIANT_NAME(AddHistory, addHistory);
    WXAUTOEXCEL_OPTIONALCPP_TO_OPTIONALVARIANT_NAME(Method, ((long*)method));
    WXAUTOEXCEL_OPTIONALCPPSTR_TO_OPTIONALVARIANT_NAME(HeaderInfo, headerInfo);

    WXAUTOEXCEL_CALL_METHOD4_RET("Follow", vNewWindow, vAddHistory, vMethod, vHeaderInfo, "null");
}

// ***** class wxExcelHyperlink PROPERTIES *****

wxString wxExcelHyperlink::GetAddress()
{
    WXAUTOEXCEL_PROPERTY_STRING_GET0("Address");
}

void wxExcelHyperlink::SetAddress(const wxString& address)
{
    InvokePutProperty(wxS("Address"), address);
}

wxString wxExcelHyperlink::GetEmailSubject()
{
    WXAUTOEXCEL_PROPERTY_STRING_GET0("EmailSubject");
}

void wxExcelHyperlink::SetEmailSubject(const wxString& emailSubject)
{
    InvokePutProperty(wxS("EmailSubject"), emailSubject);
}

wxString wxExcelHyperlink::GetName()
{
    WXAUTOEXCEL_PROPERTY_STRING_GET0("Name");
}


wxExcelRange wxExcelHyperlink::GetRange()
{
    wxExcelRange range;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("Range", range);
}

wxString wxExcelHyperlink::GetScreenTip()
{
    WXAUTOEXCEL_PROPERTY_STRING_GET0("ScreenTip");
}

void wxExcelHyperlink::SetScreenTip(const wxString& screenTip)
{
    InvokePutProperty(wxS("ScreenTip"), screenTip);
}

#if WXAUTOEXCEL_USE_SHAPES

wxExcelShape wxExcelHyperlink::GetShape()
{
    wxExcelShape shape;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("Shape", shape);
}

#endif // #if WXAUTOEXCEL_USE_SHAPES

wxString wxExcelHyperlink::GetSubAddress()
{
    WXAUTOEXCEL_PROPERTY_STRING_GET0("SubAddress");
}

void wxExcelHyperlink::SetSubAddress(const wxString& subAddress)
{
    InvokePutProperty(wxS("SubAddress"), subAddress);
}

wxString wxExcelHyperlink::GetTextToDisplay()
{
    WXAUTOEXCEL_PROPERTY_STRING_GET0("TextToDisplay");
}

void wxExcelHyperlink::SetTextToDisplay(const wxString& textToDisplay)
{
    InvokePutProperty(wxS("TextToDisplay"), textToDisplay);
}

MsoHyperlinkType wxExcelHyperlink::GetType()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("Type", MsoHyperlinkType, msoHyperlinkRange);
}

// ***** class wxExcelHyperlinks METHODS *****

wxExcelHyperlink wxExcelHyperlinks::Add(wxExcelObject* anchor, const wxString& address, const wxString& subAddress,
                                        const wxString& screenTip, const wxString& textToDisplay)
{
    wxExcelHyperlink hyperlink;

    wxCHECK( anchor && anchor->IsOk_(), hyperlink );

    wxVariant vAnchor;
    if ( !wxExcelObject::ObjectToVariant(anchor, vAnchor, wxS("Anchor")) )
        return hyperlink;    

    wxVariant vAddress(address, wxS("Address"));    
    
    WXAUTOEXCEL_OPTIONALCPPSTR_TO_OPTIONALVARIANT_NAME(SubAddress, subAddress);
    WXAUTOEXCEL_OPTIONALCPPSTR_TO_OPTIONALVARIANT_NAME(ScreenTip, screenTip);
    WXAUTOEXCEL_OPTIONALCPPSTR_TO_OPTIONALVARIANT_NAME(TextToDisplay, textToDisplay);

    WXAUTOEXCEL_CALL_METHOD4("Add", vAnchor, vAddress, vSubAddress, vTextToDisplay, "void*", hyperlink);
    
    VariantToObject(vResult, &hyperlink);
    return hyperlink;
}

void wxExcelHyperlinks::Delete()
{
    WXAUTOEXCEL_CALL_METHOD0_RET("Delete", "null");
}

// ***** class wxExcelHyperlinks PROPERTIES *****


long wxExcelHyperlinks::GetCount()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("Count");
}

wxExcelHyperlink wxExcelHyperlinks::GetItem(long index)
{
    wxASSERT( index > 0 );
    
    wxExcelHyperlink hyperlink;
    WXAUTOEXCEL_PROPERTY_OBJECT_GET1("Item", index, hyperlink);
}

wxExcelHyperlink wxExcelHyperlinks::operator[](long index)
{
    return GetItem(index);
}
} // namespace wxAutoExcel
