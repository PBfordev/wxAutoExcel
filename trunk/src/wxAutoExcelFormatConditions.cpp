/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Modified by:
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
/////////////////////////////////////////////////////////////////////////////


#include "wx/wxAutoExcel_prec.h"

#include "wx/wxAutoExcelFormatConditions.h"

#if WXAUTOEXCEL_USE_CONDFORMAT

#include "wx/wxAutoExcelPrivate.h"

#include "wx/wxAutoExcelRange.h"
#include "wx/wxAutoExcelBorders.h"
#include "wx/wxAutoExcelFont.h"
#include "wx/wxAutoExcelInterior.h"
#include "wx/wxAutoExcelAboveAverage.h"
#include "wx/wxAutoExcelColorScale.h"
#include "wx/wxAutoExcelDatabar.h"
#include "wx/wxAutoExcelIconSetCondition.h"
#include "wx/wxAutoExcelTop10.h"
#include "wx/wxAutoExcelUniqueValues.h"

namespace wxAutoExcel {

// ***** class wxExcelFormatCondition METHODS *****

void wxExcelFormatCondition::Delete()
{
    WXAUTOEXCEL_CALL_METHOD0_RET("Delete", "null");
}

void wxExcelFormatCondition::Modify(XlFormatConditionType conditionType, XlFormatConditionOperator* conditionOperator,
                                    const wxString& formula1, const wxString& formula2)
{
    wxVariantVector args;

    args.push_back(wxVariant((long)conditionType, wxS("Type")));

    WXAUTOEXCEL_OPTIONALCPP_TO_OPTIONALVARIANT_NAME_VECTOR(Operator, ((long*)conditionOperator), args);
    WXAUTOEXCEL_OPTIONALCPPSTR_TO_OPTIONALVARIANT_NAME_VECTOR(Formula1, formula1, args);
    WXAUTOEXCEL_OPTIONALCPPSTR_TO_OPTIONALVARIANT_NAME_VECTOR(Formula2, formula2, args);
    
    WXAUTOEXCEL_CALL_METHODARR_RET("Modify", args, "null");      
}

void wxExcelFormatCondition::ModifyAppliesToRange(wxExcelRange range)
{
    wxVariant vRange;

    if ( ObjectToVariant(&range, vRange, wxS("Range")) )
    {
        WXAUTOEXCEL_CALL_METHOD1_RET("ModifyAppliesToRange", vRange, "null");
    }        
}

void wxExcelFormatCondition::SetFirstPriority()
{
    WXAUTOEXCEL_CALL_METHOD0_RET("SetFirstPriority", "null");
}

void wxExcelFormatCondition::SetLastPriority()
{
    WXAUTOEXCEL_CALL_METHOD0_RET("SetLastPriority", "null");
}

// ***** class wxExcelFormatCondition PROPERTIES *****

wxExcelRange wxExcelFormatCondition::GetAppliesTo()
{
    wxExcelRange range;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("AppliesTo", range);
}


wxExcelBorders wxExcelFormatCondition::GetBorders()
{
    wxExcelBorders borders;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("Borders", borders);
}


wxExcelFont wxExcelFormatCondition::GetFont()
{
    wxExcelFont font;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("Font", font);
}

wxString wxExcelFormatCondition::GetFormula1()
{
    WXAUTOEXCEL_PROPERTY_STRING_GET0("Formula1");
}

wxString wxExcelFormatCondition::GetFormula2()
{
    WXAUTOEXCEL_PROPERTY_STRING_GET0("Formula2");
}

wxExcelInterior wxExcelFormatCondition::GetInterior()
{
    wxExcelInterior interior;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("Interior", interior);
}

wxString wxExcelFormatCondition::GetNumberFormat()
{
    WXAUTOEXCEL_PROPERTY_STRING_GET0("NumberFormat");
}

void wxExcelFormatCondition::SetNumberFormat(const wxString& numberFormat)
{
    InvokePutProperty(wxS("NumberFormat"), numberFormat);
}

XlFormatConditionOperator wxExcelFormatCondition::GetOperator()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("Operator", XlFormatConditionOperator, xlBetween);
}


long wxExcelFormatCondition::GetPriority()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("Priority");
}

void wxExcelFormatCondition::SetPriority(long priority)
{
    InvokePutProperty(wxS("Priority"), priority);
}

bool wxExcelFormatCondition::GetPTCondition()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("PTCondition");
}

XlPivotConditionScope wxExcelFormatCondition::GetScopeType()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("ScopeType", XlPivotConditionScope, xlSelectionScope);
}

void wxExcelFormatCondition::SetScopeType(XlPivotConditionScope scopeType)
{
    InvokePutProperty(wxS("ScopeType"), (long)scopeType);
}

bool wxExcelFormatCondition::GetStopIfTrue()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("StopIfTrue");
}

void wxExcelFormatCondition::SetStopIfTrue(bool stopIfTrue)
{
    InvokePutProperty(wxS("StopIfTrue"), stopIfTrue);
}

wxString wxExcelFormatCondition::GetText()
{
    WXAUTOEXCEL_PROPERTY_STRING_GET0("Text");
}

void wxExcelFormatCondition::SetText(const wxString& text)
{
    InvokePutProperty(wxS("Text"), text);
}

XlContainsOperator wxExcelFormatCondition::GetTextOperator()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("TextOperator", XlContainsOperator, xlContains);
}

void wxExcelFormatCondition::SetTextOperator(XlContainsOperator textOperator)
{
    InvokePutProperty(wxS("TextOperator"), (long)textOperator);
}

XlFormatConditionType wxExcelFormatCondition::GetType()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("Type", XlFormatConditionType, xlCellValue);
}


// ***** class wxExcelFormatConditions METHODS *****

wxExcelFormatCondition wxExcelFormatConditions::Add(XlFormatConditionType conditionType, XlFormatConditionOperator* conditionOperator,
                                                    const wxString& formula1, const wxString& formula2)
{
    wxVariantVector args;
    wxExcelFormatCondition condition;

    args.push_back(wxVariant((long)conditionType, wxS("Type")));

    WXAUTOEXCEL_OPTIONALCPP_TO_OPTIONALVARIANT_NAME_VECTOR(Operator, ((long*)conditionOperator), args);
    WXAUTOEXCEL_OPTIONALCPPSTR_TO_OPTIONALVARIANT_NAME_VECTOR(Formula1, formula1, args);
    WXAUTOEXCEL_OPTIONALCPPSTR_TO_OPTIONALVARIANT_NAME_VECTOR(Formula2, formula2, args);
        
    WXAUTOEXCEL_CALL_METHODARR("Modify", args, "void*", condition);    
    VariantToObject(vResult, &condition);
    return condition;
}


wxExcelAboveAverage wxExcelFormatConditions::AddAboveAverage()
{        
    wxExcelAboveAverage aboveAverage;
    WXAUTOEXCEL_CALL_METHOD0("AddAboveAverage", "void*", aboveAverage);
    VariantToObject(vResult, &aboveAverage);
    return aboveAverage;
}

wxExcelColorScale wxExcelFormatConditions::AddColorScale(long colorScaleType)
{
    wxExcelColorScale colorScale;
    WXAUTOEXCEL_CALL_METHOD1("AddCcolorScale", colorScaleType, "void*", colorScale);
    VariantToObject(vResult, &colorScale);
    return colorScale;
}

wxExcelDatabar wxExcelFormatConditions::AddDatabar()
{
    wxExcelDatabar databar;
    
    WXAUTOEXCEL_CALL_METHOD0("AddDataBar", "void*", databar);
    VariantToObject(vResult, &databar);
    return databar;
}

wxExcelIconSetCondition wxExcelFormatConditions::AddIconSetCondition()
{
    wxExcelIconSetCondition isc;
    
    WXAUTOEXCEL_CALL_METHOD0("AddIconSetCondition", "void*", isc);
    VariantToObject(vResult, &isc);
    return isc;
}

wxExcelTop10 wxExcelFormatConditions::AddTop10()
{
    wxExcelTop10 top10;
    
    WXAUTOEXCEL_CALL_METHOD0("AddTop10", "void*", top10);
    VariantToObject(vResult, &top10);
    return top10;
}

wxExcelUniqueValues wxExcelFormatConditions::AddUniqueValues()
{
    wxExcelUniqueValues uv;
    
    WXAUTOEXCEL_CALL_METHOD0("AddUniqueValues", "void*", uv);
    VariantToObject(vResult, &uv);
    return uv;    
}

void wxExcelFormatConditions::Delete()
{
    WXAUTOEXCEL_CALL_METHOD0_RET("Delete", "null");
}

wxExcelFormatCondition wxExcelFormatConditions::Item(long index)
{
    wxASSERT(index > 0);

    wxExcelFormatCondition condition;
    
    WXAUTOEXCEL_CALL_METHOD1("Item", index, "void*", condition);
    VariantToObject(vResult, &condition);
    return condition; 
}

wxExcelFormatCondition wxExcelFormatConditions::operator[](long index)
{
    return Item(index);
}

// ***** class wxExcelFormatConditions PROPERTIES *****

long wxExcelFormatConditions::GetCount()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("Count");
}

} // namespace wxAutoExcel


#endif // WXAUTOEXCEL_USE_CONDFORMAT