/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// License:     MIT license
/////////////////////////////////////////////////////////////////////////////


#include "wx/wxAutoExcel_prec.h"

#include "wx/wxAutoExcelComments.h"
#include "wx/wxAutoExcelShape.h"

#include "wx/wxAutoExcel_private.h"

namespace wxAutoExcel {

// ***** class wxExcelComment METHODS *****

void wxExcelComment::Delete()
{
    WXAUTOEXCEL_CALL_METHOD0_RET("Delete", "null");
}

wxExcelComment wxExcelComment::Next()
{
    wxExcelComment comment;

    WXAUTOEXCEL_CALL_METHOD0("Next", "void*", comment);
    VariantToObject(vResult, &comment);
    return comment;
}

wxExcelComment wxExcelComment::Previous()
{
    wxExcelComment comment;

    WXAUTOEXCEL_CALL_METHOD0("Previous", "void*", comment);
    VariantToObject(vResult, &comment);
    return comment;
}


wxString wxExcelComment::Text(const wxString& text, long* start, wxXlTribool overwrite)
{
    WXAUTOEXCEL_OPTIONALCPPSTR_TO_OPTIONALVARIANT_NAME(Text, text);
    WXAUTOEXCEL_OPTIONALCPP_TO_OPTIONALVARIANT_NAME(Start, start);
    WXAUTOEXCEL_OPTIONALCPPTBOOL_TO_OPTIONALVARIANT_NAME(Overwrite, overwrite);

    WXAUTOEXCEL_CALL_METHOD3("Text", vText, vStart, vOverwrite, "string", wxS(""));
    return vResult.GetString();
}

// ***** class wxExcelComment PROPERTIES *****

wxString wxExcelComment::GetAuthor()
{
    WXAUTOEXCEL_PROPERTY_STRING_GET0("Author");
}

void wxExcelComment::SetAuthor(const wxString& author)
{
    InvokePutProperty(wxS("Author"), author);
}

#if WXAUTOEXCEL_USE_SHAPES

wxExcelShape wxExcelComment::GetShape()
{
    wxExcelShape shape;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("Shape", shape);
}

#endif // #if WXAUTOEXCEL_USE_SHAPES

bool wxExcelComment::GetVisible()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("Visible");
}

void wxExcelComment::SetVisible(bool visible)
{
    InvokePutProperty(wxS("Visible"), visible);
}

// ***** class wxExcelComments METHODS *****

wxExcelComment wxExcelComments::GetItem(long index)
{
    wxASSERT( index > 0 );

    wxExcelComment comment;

    WXAUTOEXCEL_CALL_METHOD1_OBJECT("Item", index, comment);
}

wxExcelComment wxExcelComments::operator[](long index)
{
    return GetItem(index);
}

// ***** class wxExcelComments PROPERTIES *****

long wxExcelComments::GetCount()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("Count");
}

} // namespace wxAutoExcel