/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Modified by:
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
/////////////////////////////////////////////////////////////////////////////


#include "wx/wxAutoExcel_prec.h"

#include "wx/wxAutoExcelBorders.h"

#include "wx/wxAutoExcelFont.h"
#include "wx/wxAutoExcelInterior.h"
#include "wx/wxAutoExcelStyles.h"
#include "wx/wxAutoExcelWorkbook.h"

#include "wx/wxAutoExcel_private.h"

namespace wxAutoExcel {

bool wxExcelStyle::Delete()
{
    WXAUTOEXCEL_CALL_METHOD0_BOOL("Delete");
}

// ***** class wxExcelStyle PROPERTIES *****

bool wxExcelStyle::GetAddIndent()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("AddIndent");
}

void wxExcelStyle::SetAddIndent(bool addIndent)
{
    InvokePutProperty(wxS("AddIndent"), addIndent);
}

wxExcelBorders wxExcelStyle::GetBorders()
{
    wxExcelBorders borders;
    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("Borders", borders);
}

bool wxExcelStyle::GetBuiltIn()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("BuiltIn");
}

wxExcelFont wxExcelStyle::GetFont()
{
    wxExcelFont font;
    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("Font", font);
}

bool wxExcelStyle::GetFormulaHidden()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("FormulaHidden");
}

void wxExcelStyle::SetFormulaHidden(bool formulaHidden)
{
    InvokePutProperty(wxS("FormulaHidden"), formulaHidden);
}

XlHAlign wxExcelStyle::GetHorizontalAlignment()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("HorizontalAlignment", XlHAlign, xlHAlignGeneral);
}

void wxExcelStyle::SetHorizontalAlignment(XlHAlign horizontalAlignment)
{
    InvokePutProperty(wxS("HorizontalAlignment"), (long)horizontalAlignment);
}

bool wxExcelStyle::GetIncludeAlignment()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("IncludeAlignment");
}

void wxExcelStyle::SetIncludeAlignment(bool includeAlignment)
{
    InvokePutProperty(wxS("IncludeAlignment"), includeAlignment);
}

bool wxExcelStyle::GetIncludeBorder()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("IncludeBorder");
}

void wxExcelStyle::SetIncludeBorder(bool includeBorder)
{
    InvokePutProperty(wxS("IncludeBorder"), includeBorder);
}

bool wxExcelStyle::GetIncludeFont()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("IncludeFont");
}

void wxExcelStyle::SetIncludeFont(bool includeFont)
{
    InvokePutProperty(wxS("IncludeFont"), includeFont);
}

bool wxExcelStyle::GetIncludeNumber()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("IncludeNumber");
}

void wxExcelStyle::SetIncludeNumber(bool includeNumber)
{
    InvokePutProperty(wxS("IncludeNumber"), includeNumber);
}

bool wxExcelStyle::GetIncludePatterns()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("IncludePatterns");
}

void wxExcelStyle::SetIncludePatterns(bool includePatterns)
{
    InvokePutProperty(wxS("IncludePatterns"), includePatterns);
}

bool wxExcelStyle::GetIncludeProtection()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("IncludeProtection");
}

void wxExcelStyle::SetIncludeProtection(bool includeProtection)
{
    InvokePutProperty(wxS("IncludeProtection"), includeProtection);
}

long wxExcelStyle::GetIndentLevel()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("IndentLevel");
}

void wxExcelStyle::SetIndentLevel(long indentLevel)
{
    InvokePutProperty(wxS("IndentLevel"), indentLevel);
}

wxExcelInterior wxExcelStyle::GetInterior()
{
    wxExcelInterior interior;
    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("Interior", interior);
}

bool wxExcelStyle::GetLocked()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("Locked");
}

void wxExcelStyle::SetLocked(bool locked)
{
    InvokePutProperty(wxS("Locked"), locked);
}

bool wxExcelStyle::GetMergeCells()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("MergeCells");
}

wxString wxExcelStyle::GetName()
{
    WXAUTOEXCEL_PROPERTY_STRING_GET0("Name");
}

wxString wxExcelStyle::GetNameLocal()
{
    WXAUTOEXCEL_PROPERTY_STRING_GET0("NameLocal");
}

wxString wxExcelStyle::GetNumberFormat()
{
    WXAUTOEXCEL_PROPERTY_STRING_GET0("NumberFormat");
}

void wxExcelStyle::SetNumberFormat(const wxString& numberFormat)
{
    InvokePutProperty(wxS("NumberFormat"), numberFormat);
}

wxString wxExcelStyle::GetNumberFormatLocal()
{
    WXAUTOEXCEL_PROPERTY_STRING_GET0("NumberFormatLocal");
}

void wxExcelStyle::SetNumberFormatLocal(const wxString& numberFormatLocal)
{
    InvokePutProperty(wxS("NumberFormatLocal"), numberFormatLocal);
}

XlOrientation wxExcelStyle::GetOrientation()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("Orientation", XlOrientation, xlHorizontal);
}

void wxExcelStyle::SetOrientation(XlOrientation orientation)
{
    InvokePutProperty(wxS("Orientation"), (long)orientation);
}

long wxExcelStyle::GetReadingOrder()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("ReadingOrder");
}

void wxExcelStyle::SetReadingOrder(long readingOrder)
{
    InvokePutProperty(wxS("ReadingOrder"), readingOrder);
}

bool wxExcelStyle::GetShrinkToFit()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("ShrinkToFit");
}

void wxExcelStyle::SetShrinkToFit(bool shrinkToFit)
{
    InvokePutProperty(wxS("ShrinkToFit"), shrinkToFit);
}

XlVAlign wxExcelStyle::GetVerticalAlignment()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("VerticalAlignment", XlVAlign, xlVAlignTop);
}

void wxExcelStyle::SetVerticalAlignment(XlVAlign verticalAlignment)
{
    InvokePutProperty(wxS("VerticalAlignment"), (long)verticalAlignment);
}

bool wxExcelStyle::GetWrapText()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("WrapText");
}

void wxExcelStyle::SetWrapText(bool wrapText)
{
    InvokePutProperty(wxS("WrapText"), wrapText);
}


// ***** class wxExcelStyles METHODS *****

wxExcelStyle wxExcelStyles::Add(const wxString& name, wxExcelStyle* basedOn)
{
    wxExcelStyle style;

    wxVariant vBasedOn;
    if ( basedOn != NULL )
    {
        if ( !ObjectToVariant(basedOn, vBasedOn) )
            return style;
    }

    WXAUTOEXCEL_CALL_METHOD2("Add", name, vBasedOn, "void*", style);
    VariantToObject(vResult, &style);
    return style;
}

void wxExcelStyles::Merge(wxExcelWorkbook workbook)
{
    wxVariant vWorkbook;

    if ( !ObjectToVariant(&workbook, vWorkbook) )
        return;

    WXAUTOEXCEL_CALL_METHOD1_RET("Merge", vWorkbook, "null");
}

// ***** class wxExcelStyles PROPERTIES *****

long wxExcelStyles::GetCount()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("Count");
}

wxExcelStyle wxExcelStyles::GetItem(long index)
{
    wxASSERT( index > 0 );
    
    wxExcelStyle style;
    WXAUTOEXCEL_PROPERTY_OBJECT_GET1("Item", index, style);
}

wxExcelStyle wxExcelStyles::operator[](long index)
{
    return GetItem(index);
}

wxExcelStyle wxExcelStyles::GetItem(const wxString& name)
{        
    wxExcelStyle style;
    WXAUTOEXCEL_PROPERTY_OBJECT_GET1("Item", name, style);
}

wxExcelStyle wxExcelStyles::operator[](const wxString& name)
{
    return GetItem(name);
}

} // namespace wxAutoExcel
