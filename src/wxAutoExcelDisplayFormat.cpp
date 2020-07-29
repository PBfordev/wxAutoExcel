/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Modified by:
// Copyright:   (c) 2020 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
/////////////////////////////////////////////////////////////////////////////

#include "wx/wxAutoExcel_prec.h"

#include "wx/wxAutoExcelDisplayFormat.h"

#include "wx/wxAutoExcelBorders.h"
#include "wx/wxAutoExcelCharacters.h"
#include "wx/wxAutoExcelFont.h"
#include "wx/wxAutoExcelInterior.h"
#include "wx/wxAutoExcelStyles.h"

#include "wx/wxAutoExcel_private.h"

namespace wxAutoExcel {

// ***** class wxExcelDisplayFormat PROPERTIES *****

bool wxExcelDisplayFormat::GetAddIndent()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("AddIndent");
}

wxExcelBorders wxExcelDisplayFormat::GetBorders()
{
    wxExcelBorders borders;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("Borders", borders);
}

wxExcelCharacters wxExcelDisplayFormat::GetCharacters(long start, long* length)
{
   wxExcelCharacters characters;
   WXAUTOEXCEL_OPTIONALCPP_TO_OPTIONALVARIANT(Length, length);
   WXAUTOEXCEL_PROPERTY_OBJECT_GET2("Characters", start, vLength, characters);
}

wxExcelFont wxExcelDisplayFormat::GetFont()
{
    wxExcelFont font;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("Font", font);
}

wxXlTribool wxExcelDisplayFormat::GetFormulaHidden()
{
    wxXlTribool tb;
    wxVariant vResult;

    if ( InvokeGetProperty(wxS("FormulaHidden"), vResult) && vResult.GetType() == wxS("bool") )
    {
        tb = vResult.GetBool();
    }
    return tb;
}

long wxExcelDisplayFormat::GetHorizontalAlignment()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("HorizontalAlignment");
}

long wxExcelDisplayFormat::GetIndentLevel()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("IndentLevel");
}

wxExcelInterior wxExcelDisplayFormat::GetInterior()
{
    wxExcelInterior interior;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("Interior", interior);
}

wxXlTribool wxExcelDisplayFormat::GetLocked()
{
    wxXlTribool tb;
    wxVariant vResult;

    if ( InvokeGetProperty(wxS("Locked"), vResult) && vResult.GetType() == wxS("bool") )
    {
        tb = vResult.GetBool();
    }
    return tb;
}

bool wxExcelDisplayFormat::GetMergeCells()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("MergeCells");
}

wxString wxExcelDisplayFormat::GetNumberFormat()
{
    wxVariant vResult;

    // NumberFormat returns either the number format string
    // if all the cells in the range have the same number format;
    // or null if they have not.
    if ( InvokeGetProperty(wxS("NumberFormat"), vResult) )
    {
        if ( vResult.IsType(wxS("string")) )
            return vResult.GetString();
    }

    return wxEmptyString;
}

wxString wxExcelDisplayFormat::GetNumberFormatLocal()
{
     wxVariant vResult;

    // NumberFormatLocal returns either the number format string
    // if all the cells in the range have the same number format;
    // or null if they have not.
    if ( InvokeGetProperty(wxS("NumberFormatLocal"), vResult) )
    {
        if ( vResult.IsType(wxS("string")) )
            return vResult.GetString();;
    }

    return wxEmptyString;
}

long wxExcelDisplayFormat::GetOrientation()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("Orientation");
}

long wxExcelDisplayFormat::GetReadingOrder()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("ReadingOrder");
}

wxXlTribool wxExcelDisplayFormat::GetShrinkToFit()
{
    wxXlTribool tb;
    wxVariant vResult;

    if ( InvokeGetProperty(wxS("ShrinkToFit"), vResult) && vResult.GetType() == wxS("bool") )
    {
        tb = vResult.GetBool();
    }
    return tb;
}

wxExcelStyle wxExcelDisplayFormat::GetStyle()
{
    wxExcelStyle style;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("Style", style);
}

long wxExcelDisplayFormat::GetVerticalAlignment()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("VerticalAlignment");
}

wxXlTribool wxExcelDisplayFormat::GetWrapText()
{
    wxXlTribool tb;
    wxVariant vResult;

    if ( InvokeGetProperty(wxS("WrapText"), vResult) && vResult.GetType() == wxS("bool"))
    {
        tb = vResult.GetBool();
    }
    return tb;
}

} // namespace wxAutoExcel 
