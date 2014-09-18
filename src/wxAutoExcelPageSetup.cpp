/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Modified by:
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
/////////////////////////////////////////////////////////////////////////////


#include "wx/wxAutoExcel_prec.h"

#include "wx/wxAutoExcelPageSetup.h"

#include "wx/wxAutoExcelPages.h"
#include "wx/wxAutoExcelGraphic.h"

#include "wx/wxAutoExcelPrivate.h"

namespace wxAutoExcel {

// ***** class wxExcelPageSetup PROPERTIES *****

bool wxExcelPageSetup::GetAlignMarginsHeaderFooter()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("AlignMarginsHeaderFooter");
}

void wxExcelPageSetup::SetAlignMarginsHeaderFooter(bool alignMarginsHeaderFooter)
{
    InvokePutProperty(wxS("AlignMarginsHeaderFooter"), alignMarginsHeaderFooter);
}

bool wxExcelPageSetup::GetBlackAndWhite()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("BlackAndWhite");
}

void wxExcelPageSetup::SetBlackAndWhite(bool blackAndWhite)
{
    InvokePutProperty(wxS("BlackAndWhite"), blackAndWhite);
}

double wxExcelPageSetup::GetBottomMargin()
{
    WXAUTOEXCEL_PROPERTY_DOUBLE_GET0("BottomMargin");
}

void wxExcelPageSetup::SetBottomMargin(double bottomMargin)
{
    InvokePutProperty(wxS("BottomMargin"), bottomMargin);
}

wxString wxExcelPageSetup::GetCenterFooter()
{
    WXAUTOEXCEL_PROPERTY_STRING_GET0("CenterFooter");
}

void wxExcelPageSetup::SetCenterFooter(const wxString& centerFooter)
{
    InvokePutProperty(wxS("CenterFooter"), centerFooter);
}

wxExcelGraphic wxExcelPageSetup::GetCenterFooterPicture()
{
    wxExcelGraphic graphic;
    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("CenterFooterPicture", graphic);
}

wxString wxExcelPageSetup::GetCenterHeader()
{
    WXAUTOEXCEL_PROPERTY_STRING_GET0("CenterHeader");
}

wxExcelGraphic wxExcelPageSetup::GetCenterHeaderPicture()
{
    wxExcelGraphic graphic;
    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("CenterHeaderPicture", graphic);
}

void wxExcelPageSetup::SetCenterHeader(const wxString& centerHeader)
{
    InvokePutProperty(wxS("CenterHeader"), centerHeader);
}

bool wxExcelPageSetup::GetCenterHorizontally()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("CenterHorizontally");
}

void wxExcelPageSetup::SetCenterHorizontally(bool centerHorizontally)
{
    InvokePutProperty(wxS("CenterHorizontally"), centerHorizontally);
}

bool wxExcelPageSetup::GetCenterVertically()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("CenterVertically");
}

void wxExcelPageSetup::SetCenterVertically(bool centerVertically)
{
    InvokePutProperty(wxS("CenterVertically"), centerVertically);
}

bool wxExcelPageSetup::GetDifferentFirstPageHeaderFooter()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("DifferentFirstPageHeaderFooter");
}

void wxExcelPageSetup::SetDifferentFirstPageHeaderFooter(bool differentFirstPageHeaderFooter)
{
    InvokePutProperty(wxS("DifferentFirstPageHeaderFooter"), differentFirstPageHeaderFooter);
}

bool wxExcelPageSetup::GetDraft()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("Draft");
}

void wxExcelPageSetup::SetDraft(bool draft)
{
    InvokePutProperty(wxS("Draft"), draft);
}

wxExcelPage wxExcelPageSetup::GetEvenPage()
{
    wxExcelPage page;
    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("EvenPage", page);
}

wxExcelPage wxExcelPageSetup::GetFirstPage()
{
    wxExcelPage page;
    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("FirstPage", page);
}

long wxExcelPageSetup::GetFirstPageNumber()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("FirstPageNumber");
}

void wxExcelPageSetup::SetFirstPageNumber(long firstPageNumber)
{
    InvokePutProperty(wxS("FirstPageNumber"), firstPageNumber);
}

long wxExcelPageSetup::GetFitToPagesTall()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("FitToPagesTall");
}

void wxExcelPageSetup::SetFitToPagesTall(long fitToPagesTall)
{
    InvokePutProperty(wxS("FitToPagesTall"), fitToPagesTall);
}

long wxExcelPageSetup::GetFitToPagesWide()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("FitToPagesWide");
}

void wxExcelPageSetup::SetFitToPagesWide(long fitToPagesWide)
{
    InvokePutProperty(wxS("FitToPagesWide"), fitToPagesWide);
}

double wxExcelPageSetup::GetFooterMargin()
{
    WXAUTOEXCEL_PROPERTY_DOUBLE_GET0("FooterMargin");
}

void wxExcelPageSetup::SetFooterMargin(double footerMargin)
{
    InvokePutProperty(wxS("FooterMargin"), footerMargin);
}

double wxExcelPageSetup::GetHeaderMargin()
{
    WXAUTOEXCEL_PROPERTY_DOUBLE_GET0("HeaderMargin");
}

void wxExcelPageSetup::SetHeaderMargin(double headerMargin)
{
    InvokePutProperty(wxS("HeaderMargin"), headerMargin);
}

wxString wxExcelPageSetup::GetLeftFooter()
{
    WXAUTOEXCEL_PROPERTY_STRING_GET0("LeftFooter");
}

void wxExcelPageSetup::SetLeftFooter(const wxString& leftFooter)
{
    InvokePutProperty(wxS("LeftFooter"), leftFooter);
}

wxExcelGraphic wxExcelPageSetup::GetLeftFooterPicture()
{
    wxExcelGraphic graphic;
    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("LeftFooterPicture", graphic);
}


wxString wxExcelPageSetup::GetLeftHeader()
{
    WXAUTOEXCEL_PROPERTY_STRING_GET0("LeftHeader");
}

void wxExcelPageSetup::SetLeftHeader(const wxString& leftHeader)
{
    InvokePutProperty(wxS("LeftHeader"), leftHeader);
}

wxExcelGraphic wxExcelPageSetup::GetLeftHeaderPicture()
{
    wxExcelGraphic graphic;
    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("LeftHeaderPicture", graphic);
}


double wxExcelPageSetup::GetLeftMargin()
{
    WXAUTOEXCEL_PROPERTY_DOUBLE_GET0("LeftMargin");
}

void wxExcelPageSetup::SetLeftMargin(double leftMargin)
{
    InvokePutProperty(wxS("LeftMargin"), leftMargin);
}

bool wxExcelPageSetup::GetOddAndEvenPagesHeaderFooter()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("OddAndEvenPagesHeaderFooter");
}

void wxExcelPageSetup::SetOddAndEvenPagesHeaderFooter(bool oddAndEvenPagesHeaderFooter)
{
    InvokePutProperty(wxS("OddAndEvenPagesHeaderFooter"), oddAndEvenPagesHeaderFooter);
}

// contrary to the documentation, the result is returned as a double?
XlOrder wxExcelPageSetup::GetOrder()
{
    wxVariant vResult;
    XlOrder result = xlDownThenOver;
    
    if ( InvokeGetProperty(wxS("Order"), vResult) )
    {        
        unsigned long l;
        if ( vResult.GetString().ToCULong(&l) )
        {
            result = (XlOrder)l;
        }        
    }
    return result;
}

void wxExcelPageSetup::SetOrder(XlOrder order)
{
    InvokePutProperty(wxS("Order"), (long)order);
}

// contrary to the documentation, the result is returned as a double?
XlPageOrientation wxExcelPageSetup::GetOrientation()
{    
    wxVariant vResult;
    XlPageOrientation result = xlPortrait;
    
    if ( InvokeGetProperty(wxS("Orientation"), vResult) )
    {        
        unsigned long l;
        if ( vResult.GetString().ToCULong(&l) )
        {
            result = (XlPageOrientation)l;
        }        
    }
    return result;
}

void wxExcelPageSetup::SetOrientation(XlPageOrientation orientation)
{
    InvokePutProperty(wxS("Orientation"), (long)orientation);
}

wxExcelPages wxExcelPageSetup::GetPages()
{
    wxExcelPages pages;
    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("Pages", pages);
}

// contrary to the documentation, the result is returned as a double?
XlPaperSize wxExcelPageSetup::GetPaperSize()
{    
    wxVariant vResult;
    XlPaperSize result = xlPaperA4;
    
    if ( InvokeGetProperty(wxS("PaperSize"), vResult) )
    {        
        unsigned long l;
        if ( vResult.GetString().ToCULong(&l) )
        {
            result = (XlPaperSize)l;
        }        
    }
    return result;
}

void wxExcelPageSetup::SetPaperSize(XlPaperSize paperSize)
{
    InvokePutProperty(wxS("PaperSize"), (long)paperSize);
}

wxString wxExcelPageSetup::GetPrintArea()
{
    WXAUTOEXCEL_PROPERTY_STRING_GET0("PrintArea");
}

void wxExcelPageSetup::SetPrintArea(const wxString& printArea)
{
    InvokePutProperty(wxS("PrintArea"), printArea);
}

XlPrintLocation wxExcelPageSetup::GetPrintComments()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("PrintComments", XlPrintLocation, xlPrintSheetEnd);
}

void wxExcelPageSetup::SetPrintComments(XlPrintLocation printComments)
{
    InvokePutProperty(wxS("PrintComments"), (long)printComments);
}

XlPrintErrors wxExcelPageSetup::GetPrintErrors()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("PrintErrors", XlPrintErrors, xlPrintErrorsDisplayed);
}

void wxExcelPageSetup::SetPrintErrors(XlPrintErrors printErrors)
{
    InvokePutProperty(wxS("PrintErrors"), (long)printErrors);
}

bool wxExcelPageSetup::GetPrintGridlines()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("PrintGridlines");
}

void wxExcelPageSetup::SetPrintGridlines(bool printGridlines)
{
    InvokePutProperty(wxS("PrintGridlines"), printGridlines);
}

bool wxExcelPageSetup::GetPrintHeadings()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("PrintHeadings");
}

void wxExcelPageSetup::SetPrintHeadings(bool printHeadings)
{
    InvokePutProperty(wxS("PrintHeadings"), printHeadings);
}

bool wxExcelPageSetup::GetPrintNotes()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("PrintNotes");
}

void wxExcelPageSetup::SetPrintNotes(bool printNotes)
{
    InvokePutProperty(wxS("PrintNotes"), printNotes);
}

long wxExcelPageSetup::GetPrintQuality(int index)
{
    WXAUTOEXCEL_PROPERTY_LONG_GET1("PrintQuality", index);
}

void wxExcelPageSetup::SetPrintQuality(int index, long quality)
{
    InvokePutProperty(wxS("PrintQuality"), index, quality);
}

wxString wxExcelPageSetup::GetPrintTitleColumns()
{
    WXAUTOEXCEL_PROPERTY_STRING_GET0("PrintTitleColumns");
}

void wxExcelPageSetup::SetPrintTitleColumns(const wxString& printTitleColumns)
{
    InvokePutProperty(wxS("PrintTitleColumns"), printTitleColumns);
}

wxString wxExcelPageSetup::GetPrintTitleRows()
{
    WXAUTOEXCEL_PROPERTY_STRING_GET0("PrintTitleRows");
}

void wxExcelPageSetup::SetPrintTitleRows(const wxString& printTitleRows)
{
    InvokePutProperty(wxS("PrintTitleRows"), printTitleRows);
}

wxString wxExcelPageSetup::GetRightFooter()
{
    WXAUTOEXCEL_PROPERTY_STRING_GET0("RightFooter");
}

void wxExcelPageSetup::SetRightFooter(const wxString& rightFooter)
{
    InvokePutProperty(wxS("RightFooter"), rightFooter);
}

wxExcelGraphic wxExcelPageSetup::GetRightFooterPicture()
{
    wxExcelGraphic graphic;
    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("RightFooterPicture", graphic);
}


wxString wxExcelPageSetup::GetRightHeader()
{
    WXAUTOEXCEL_PROPERTY_STRING_GET0("RightHeader");
}

void wxExcelPageSetup::SetRightHeader(const wxString& rightHeader)
{
    InvokePutProperty(wxS("RightHeader"), rightHeader);
}

wxExcelGraphic wxExcelPageSetup::GetRightHeaderPicture()
{
    wxExcelGraphic graphic;
    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("RightHeaderPicture", graphic);
}

double wxExcelPageSetup::GetRightMargin()
{
    WXAUTOEXCEL_PROPERTY_DOUBLE_GET0("RightMargin");
}

void wxExcelPageSetup::SetRightMargin(double rightMargin)
{
    InvokePutProperty(wxS("RightMargin"), rightMargin);
}

bool wxExcelPageSetup::GetScaleWithDocHeaderFooter()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("ScaleWithDocHeaderFooter");
}

void wxExcelPageSetup::SetScaleWithDocHeaderFooter(bool scaleWithDocHeaderFooter)
{
    InvokePutProperty(wxS("ScaleWithDocHeaderFooter"), scaleWithDocHeaderFooter);
}

double wxExcelPageSetup::GetTopMargin()
{
    WXAUTOEXCEL_PROPERTY_DOUBLE_GET0("TopMargin");
}

void wxExcelPageSetup::SetTopMargin(double topMargin)
{
    InvokePutProperty(wxS("TopMargin"), topMargin);
}

long wxExcelPageSetup::GetZoom()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("Zoom");
}

void wxExcelPageSetup::SetZoom(long zoom)
{
    InvokePutProperty(wxS("Zoom"), zoom);
}

} // namespace wxAutoExcel
