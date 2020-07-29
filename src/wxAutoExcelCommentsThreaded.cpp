/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Modified by:
// Copyright:   (c) 2020 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
/////////////////////////////////////////////////////////////////////////////

#include "wx/wxAutoExcel_prec.h"

#include "wx/wxAutoExcelCommentsThreaded.h"

#include "wx/wxAutoExcelAuthor.h"

#include "wx/wxAutoExcel_private.h"

namespace wxAutoExcel {

// ***** class wxExcelCommentThreaded METHODS *****

wxExcelCommentThreaded wxExcelCommentThreaded::AddReply(const wxString& text)
{
    wxExcelCommentThreaded object;

    WXAUTOEXCEL_OPTIONALCPPSTR_TO_OPTIONALVARIANT_NAME(Text, text);

    WXAUTOEXCEL_CALL_METHOD1("AddReply", vText, "void*", object);
    VariantToObject(vResult, &object);
    return object;
}

void wxExcelCommentThreaded::Delete()
{
    WXAUTOEXCEL_CALL_METHOD0_RET("Delete", "null");
}

wxExcelCommentThreaded wxExcelCommentThreaded::Next()
{
    wxExcelCommentThreaded object;

    WXAUTOEXCEL_CALL_METHOD0("Next", "void*", object);
    VariantToObject(vResult, &object);
    return object;    
}

wxExcelCommentThreaded wxExcelCommentThreaded::Previous()
{
    wxExcelCommentThreaded object;

    WXAUTOEXCEL_CALL_METHOD0("Previous", "void*", object);
    VariantToObject(vResult, &object);
    return object;    
}

wxString wxExcelCommentThreaded::Text(const wxString& text, long* start, wxXlTribool overwrite)
{    
    WXAUTOEXCEL_OPTIONALCPPSTR_TO_OPTIONALVARIANT_NAME(Text, text);
    WXAUTOEXCEL_OPTIONALCPP_TO_OPTIONALVARIANT_NAME(Start, start);
    WXAUTOEXCEL_OPTIONALCPPTBOOL_TO_OPTIONALVARIANT_NAME(Overwrite, overwrite);

    WXAUTOEXCEL_CALL_METHOD3("Text", vText, vStart, vOverwrite, "string", "");    
    return vResult.GetString();
}

// ***** class wxExcelCommentThreaded PROPERTIES *****

wxExcelAuthor wxExcelCommentThreaded::GetAuthor()
{
    wxExcelAuthor author;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("Author", author);
}

wxVariant wxExcelCommentThreaded::GetDate()
{
    wxVariant result;

    InvokeGetProperty(wxS("Date"), result);
    return result;
}

wxExcelCommentsThreaded wxExcelCommentThreaded::GetReplies()
{
    wxExcelCommentsThreaded commentsThreaded;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("Replies", commentsThreaded);
}



// ***** class wxExcelCommentsThreaded METHODS *****

wxExcelCommentThreaded wxExcelCommentsThreaded::GetItem(long index)
{
    wxASSERT( index > 0 );

    wxExcelCommentThreaded object;

    WXAUTOEXCEL_CALL_METHOD1("Item", index, "void*", object);
    VariantToObject(vResult, &object);
    return object; 
}

wxExcelCommentThreaded wxExcelCommentsThreaded::operator[](long index)
{
    return GetItem(index);
}


// ***** class wxExcelCommentThreadeds PROPERTIES *****

long wxExcelCommentsThreaded::GetCount()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("Count");
}


} // namespace wxAutoExcel 
