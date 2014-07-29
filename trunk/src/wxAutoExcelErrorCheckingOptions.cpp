/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Modified by:
// RCS-ID:      $Id: $
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
/////////////////////////////////////////////////////////////////////////////


#include "wx/wxAutoExcel_prec.h"

#include "wx/wxAutoExcelErrorCheckingOptions.h"

#include "wx/wxAutoExcelPrivate.h"


namespace wxAutoExcel {

// ***** class wxExcelErrorCheckingOptions PROPERTIES *****


bool wxExcelErrorCheckingOptions::GetBackgroundChecking()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("BackgroundChecking");
}

void wxExcelErrorCheckingOptions::SetBackgroundChecking(bool backgroundChecking)
{
    InvokePutProperty(wxS("BackgroundChecking"), backgroundChecking);
}

bool wxExcelErrorCheckingOptions::GetEmptyCellReferences()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("EmptyCellReferences");
}

void wxExcelErrorCheckingOptions::SetEmptyCellReferences(bool emptyCellReferences)
{
    InvokePutProperty(wxS("EmptyCellReferences"), emptyCellReferences);
}

bool wxExcelErrorCheckingOptions::GetEvaluateToError()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("EvaluateToError");
}

void wxExcelErrorCheckingOptions::SetEvaluateToError(bool evaluateToError)
{
    InvokePutProperty(wxS("EvaluateToError"), evaluateToError);
}

bool wxExcelErrorCheckingOptions::GetInconsistentFormula()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("InconsistentFormula");
}

void wxExcelErrorCheckingOptions::SetInconsistentFormula(bool inconsistentFormula)
{
    InvokePutProperty(wxS("InconsistentFormula"), inconsistentFormula);
}

bool wxExcelErrorCheckingOptions::GetInconsistentTableFormula()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("InconsistentTableFormula");
}

void wxExcelErrorCheckingOptions::SetInconsistentTableFormula(bool inconsistentTableFormula)
{
    InvokePutProperty(wxS("InconsistentTableFormula"), inconsistentTableFormula);
}

XlColorIndex wxExcelErrorCheckingOptions::GetIndicatorColorIndex()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("IndicatorColorIndex", XlColorIndex, xlColorIndexNone);
}

void wxExcelErrorCheckingOptions::SetIndicatorColorIndex(XlColorIndex indicatorColorIndex)
{
    InvokePutProperty(wxS("IndicatorColorIndex"), (long)indicatorColorIndex);
}

bool wxExcelErrorCheckingOptions::GetListDataValidation()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("ListDataValidation");
}

void wxExcelErrorCheckingOptions::SetListDataValidation(bool listDataValidation)
{
    InvokePutProperty(wxS("ListDataValidation"), listDataValidation);
}

bool wxExcelErrorCheckingOptions::GetNumberAsText()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("NumberAsText");
}

void wxExcelErrorCheckingOptions::SetNumberAsText(bool numberAsText)
{
    InvokePutProperty(wxS("NumberAsText"), numberAsText);
}

bool wxExcelErrorCheckingOptions::GetOmittedCells()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("OmittedCells");
}

void wxExcelErrorCheckingOptions::SetOmittedCells(bool omittedCells)
{
    InvokePutProperty(wxS("OmittedCells"), omittedCells);
}

bool wxExcelErrorCheckingOptions::GetTextDate()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("TextDate");
}

void wxExcelErrorCheckingOptions::SetTextDate(bool textDate)
{
    InvokePutProperty(wxS("TextDate"), textDate);
}

bool wxExcelErrorCheckingOptions::GetUnlockedFormulaCells()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("UnlockedFormulaCells");
}

void wxExcelErrorCheckingOptions::SetUnlockedFormulaCells(bool unlockedFormulaCells)
{
    InvokePutProperty(wxS("UnlockedFormulaCells"), unlockedFormulaCells);
}




} // namespace wxAutoExcel
