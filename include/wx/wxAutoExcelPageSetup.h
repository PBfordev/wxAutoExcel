/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// License:     MIT license
/////////////////////////////////////////////////////////////////////////////


#ifndef _WXAUTOEXCEL_PAGESETUP_H
#define _WXAUTOEXCEL_PAGESETUP_H

#include "wx/wxAutoExcel_defs.h"

#include "wx/wxAutoExcel_object.h"
#include "wx/wxAutoExcel_enums.h"

namespace wxAutoExcel {

    /**
    @brief Represents Microsoft Excel PageSetup object.
    */
   class WXDLLIMPEXP_WXAUTOEXCEL wxExcelPageSetup : public wxExcelObject
    {
    public:
        /**
        True if the header and the footer are aligned. with the margins set in the page setup options. Since Excel 2007.

        [MSDN documentation for PageSetup.AlignMarginsHeaderFooter](https://learn.microsoft.com/en-us/office/vba/api/excel.pagesetup.alignmarginsheaderfooter).
        */
        bool GetAlignMarginsHeaderFooter();

        /**
        True if the header and the footer are aligned. with the margins set in the page setup options. Since Excel 2007.

        [MSDN documentation for PageSetup.AlignMarginsHeaderFooter](https://learn.microsoft.com/en-us/office/vba/api/excel.pagesetup.alignmarginsheaderfooter).
        */
        void SetAlignMarginsHeaderFooter(bool alignMarginsHeaderFooter);

        /**
        True if elements of the document will be printed in black and white.

        [MSDN documentation for PageSetup.BlackAndWhite](https://learn.microsoft.com/en-us/office/vba/api/excel.pagesetup.blackandwhite).
        */
        bool GetBlackAndWhite();

        /**
        True if elements of the document will be printed in black and white.

        [MSDN documentation for PageSetup.BlackAndWhite](https://learn.microsoft.com/en-us/office/vba/api/excel.pagesetup.blackandwhite).
        */
        void SetBlackAndWhite(bool blackAndWhite);

        /**
        Returns the size of the bottom margin, in points.

        [MSDN documentation for PageSetup.BottomMargin](https://learn.microsoft.com/en-us/office/vba/api/excel.pagesetup.bottommargin).
        */
        double GetBottomMargin();

        /**
        Sets the size of the bottom margin, in points.

        [MSDN documentation for PageSetup.BottomMargin](https://learn.microsoft.com/en-us/office/vba/api/excel.pagesetup.bottommargin).
        */
        void SetBottomMargin(double bottomMargin);

        /**
        Returns the center footer.

        [MSDN documentation for PageSetup.CenterFooter](https://learn.microsoft.com/en-us/office/vba/api/excel.pagesetup.centerfooter).
        */
        wxString GetCenterFooter();

        /**
        Sets the center footer.

        [MSDN documentation for PageSetup.CenterFooter](https://learn.microsoft.com/en-us/office/vba/api/excel.pagesetup.centerfooter).
        */
        void SetCenterFooter(const wxString& centerFooter);

        /**
        Returns a Graphic Represents the picture for the center section of the footer. Used to set attributes about the picture.

        [MSDN documentation for PageSetup.CenterFooterPicture](https://learn.microsoft.com/en-us/office/vba/api/excel.pagesetup.centerfooterpicture).
        */
        wxExcelGraphic GetCenterFooterPicture();

        /**
        Returns the center heaer.

        [MSDN documentation for PageSetup.CenterHeader](https://learn.microsoft.com/en-us/office/vba/api/excel.pagesetup.centerheader).
        */
        wxString GetCenterHeader();

        /**
        Sets the center heaer.

        [MSDN documentation for PageSetup.CenterHeader](https://learn.microsoft.com/en-us/office/vba/api/excel.pagesetup.centerheader).
        */
        void SetCenterHeader(const wxString& centerHeader);

        /**
        Returns a Graphic Represents the picture for the center section of the header. Used to set attributes about the picture.

        [MSDN documentation for PageSetup.CenterHeaderPicture](https://learn.microsoft.com/en-us/office/vba/api/excel.pagesetup.centerheaderpicture).
        */
        wxExcelGraphic GetCenterHeaderPicture();


        /**
        True if the sheet is centered horizontally on the page when it's printed.

        [MSDN documentation for PageSetup.CenterHorizontally](https://learn.microsoft.com/en-us/office/vba/api/excel.pagesetup.centerhorizontally).
        */
        bool GetCenterHorizontally();

        /**
        True if the sheet is centered horizontally on the page when it's printed.

        [MSDN documentation for PageSetup.CenterHorizontally](https://learn.microsoft.com/en-us/office/vba/api/excel.pagesetup.centerhorizontally).
        */
        void SetCenterHorizontally(bool centerHorizontally);

        /**
        True if the sheet is centered vertically on the page when it's printed.

        [MSDN documentation for PageSetup.CenterVertically](https://learn.microsoft.com/en-us/office/vba/api/excel.pagesetup.centervertically).
        */
        bool GetCenterVertically();

        /**
        True if the sheet is centered vertically on the page when it's printed.

        [MSDN documentation for PageSetup.CenterVertically](https://learn.microsoft.com/en-us/office/vba/api/excel.pagesetup.centervertically).
        */
        void SetCenterVertically(bool centerVertically);

        /**
        True if a different header or footer is used on the first page.

        [MSDN documentation for PageSetup.DifferentFirstPageHeaderFooter](https://learn.microsoft.com/en-us/office/vba/api/excel.pagesetup.differentfirstpageheaderfooter).
        */
        bool GetDifferentFirstPageHeaderFooter();

        /**
        True if a different header or footer is used on the first page.

        [MSDN documentation for PageSetup.DifferentFirstPageHeaderFooter](https://learn.microsoft.com/en-us/office/vba/api/excel.pagesetup.differentfirstpageheaderfooter).
        */
        void SetDifferentFirstPageHeaderFooter(bool differentFirstPageHeaderFooter);

        /**
        True if the sheet will be printed without graphics.

        [MSDN documentation for PageSetup.Draft](https://learn.microsoft.com/en-us/office/vba/api/excel.pagesetup.draft).
        */
        bool GetDraft();

        /**
        True if the sheet will be printed without graphics.

        [MSDN documentation for PageSetup.Draft](https://learn.microsoft.com/en-us/office/vba/api/excel.pagesetup.draft).
        */
        void SetDraft(bool draft);

        /**
        Returns the page settings for the even pages. Since Excel 2007.

        [MSDN documentation for PageSetup.EvenPage](https://learn.microsoft.com/en-us/office/vba/api/excel.pagesetup.evenpage).
        */
        wxExcelPage GetEvenPage();

        /**
        Returns the page settings for the first page. Since Excel 2007.

        [MSDN documentation for PageSetup.FirstPage](https://learn.microsoft.com/en-us/office/vba/api/excel.pagesetup.firstpage).
        */
        wxExcelPage GetFirstPage();

        /**
        Returns the first page number that will be used when this sheet is printed. If xlAutomatic, Microsoft Excel chooses the first page number. The default is xlAutomatic.

        [MSDN documentation for PageSetup.FirstPageNumber](https://learn.microsoft.com/en-us/office/vba/api/excel.pagesetup.firstpagenumber).
        */
        long GetFirstPageNumber();

        /**
        Sets the first page number that will be used when this sheet is printed. If xlAutomatic, Microsoft Excel chooses the first page number. The default is xlAutomatic.

        [MSDN documentation for PageSetup.FirstPageNumber](https://learn.microsoft.com/en-us/office/vba/api/excel.pagesetup.firstpagenumber).
        */
        void SetFirstPageNumber(long firstPageNumber);

        /**
        Returns the number of pages tall the worksheet will be scaled to when it's printed. Applies only to worksheets.

        [MSDN documentation for PageSetup.FitToPagesTall](https://learn.microsoft.com/en-us/office/vba/api/excel.pagesetup.fittopagestall).
        */
        long GetFitToPagesTall();

        /**
        Sets the number of pages tall the worksheet will be scaled to when it's printed. Applies only to worksheets.

        [MSDN documentation for PageSetup.FitToPagesTall](https://learn.microsoft.com/en-us/office/vba/api/excel.pagesetup.fittopagestall).
        */
        void SetFitToPagesTall(long fitToPagesTall);

        /**
        Returns the number of pages wide the worksheet will be scaled to when it's printed. Applies only to worksheets.

        [MSDN documentation for PageSetup.FitToPagesWide](https://learn.microsoft.com/en-us/office/vba/api/excel.pagesetup.fittopageswide).
        */
        long GetFitToPagesWide();

        /**
        Sets the number of pages wide the worksheet will be scaled to when it's printed. Applies only to worksheets.

        [MSDN documentation for PageSetup.FitToPagesWide](https://learn.microsoft.com/en-us/office/vba/api/excel.pagesetup.fittopageswide).
        */
        void SetFitToPagesWide(long fitToPagesWide);

        /**
        Returns the distance from the bottom of the page to the footer, in points.

        [MSDN documentation for PageSetup.FooterMargin](https://learn.microsoft.com/en-us/office/vba/api/excel.pagesetup.footermargin).
        */
        double GetFooterMargin();

        /**
        Sets the distance from the bottom of the page to the footer, in points.

        [MSDN documentation for PageSetup.FooterMargin](https://learn.microsoft.com/en-us/office/vba/api/excel.pagesetup.footermargin).
        */
        void SetFooterMargin(double footerMargin);

        /**
        Returns the distance from the top of the page to the header, in points.

        [MSDN documentation for PageSetup.HeaderMargin](https://learn.microsoft.com/en-us/office/vba/api/excel.pagesetup.headermargin).
        */
        double GetHeaderMargin();

        /**
        Sets the distance from the top of the page to the header, in points.

        [MSDN documentation for PageSetup.HeaderMargin](https://learn.microsoft.com/en-us/office/vba/api/excel.pagesetup.headermargin).
        */
        void SetHeaderMargin(double headerMargin);

        /**
        Returns the left footer.

        [MSDN documentation for PageSetup.LeftFooter](https://learn.microsoft.com/en-us/office/vba/api/excel.pagesetup.leftfooter).
        */
        wxString GetLeftFooter();

        /**
        Sets the left footer.

        [MSDN documentation for PageSetup.LeftFooter](https://learn.microsoft.com/en-us/office/vba/api/excel.pagesetup.leftfooter).
        */
        void SetLeftFooter(const wxString& leftFooter);

        /**
        Returns a Graphic Represents the picture for the left section of the footer. Used to set attributes about the picture.

        [MSDN documentation for PageSetup.LeftFooterPicture](https://learn.microsoft.com/en-us/office/vba/api/excel.pagesetup.leftfooterpicture).
        */
        wxExcelGraphic GetLeftFooterPicture();

        /**
        Returns the left header.

        [MSDN documentation for PageSetup.LeftHeader](https://learn.microsoft.com/en-us/office/vba/api/excel.pagesetup.leftheader).
        */
        wxString GetLeftHeader();

        /**
        Sets the left header.

        [MSDN documentation for PageSetup.LeftHeader](https://learn.microsoft.com/en-us/office/vba/api/excel.pagesetup.leftheader).
        */
        void SetLeftHeader(const wxString& leftHeader);

        /**
        Returns a Graphic Represents the picture for the left section of the header. Used to set attributes about the picture.

        [MSDN documentation for PageSetup.LeftHeaderPicture](https://learn.microsoft.com/en-us/office/vba/api/excel.pagesetup.leftheaderpicture).
        */
        wxExcelGraphic GetLeftHeaderPicture();

        /**
        Returns the size of the left margin, in points.

        [MSDN documentation for PageSetup.LeftMargin](https://learn.microsoft.com/en-us/office/vba/api/excel.pagesetup.leftmargin).
        */
        double GetLeftMargin();

        /**
        Sets the size of the left margin, in points.

        [MSDN documentation for PageSetup.LeftMargin](https://learn.microsoft.com/en-us/office/vba/api/excel.pagesetup.leftmargin).
        */
        void SetLeftMargin(double leftMargin);

        /**
        True if different headers and footers for odd-numbered and even-numbered pages. Since Excel 2007.

        [MSDN documentation for PageSetup.OddAndEvenPagesHeaderFooter](https://learn.microsoft.com/en-us/office/vba/api/excel.pagesetup.oddandevenpagesheaderfooter).
        */
        bool GetOddAndEvenPagesHeaderFooter();

        /**
        True if different headers and footers for odd-numbered and even-numbered pages. Since Excel 2007.

        [MSDN documentation for PageSetup.OddAndEvenPagesHeaderFooter](https://learn.microsoft.com/en-us/office/vba/api/excel.pagesetup.oddandevenpagesheaderfooter).
        */
        void SetOddAndEvenPagesHeaderFooter(bool oddAndEvenPagesHeaderFooter);

        /**
        Returns a XlOrder value that represents the order that Microsoft Excel uses to number pages when printing a large worksheet.

        [MSDN documentation for PageSetup.Order](https://learn.microsoft.com/en-us/office/vba/api/excel.pagesetup.order).
        */
        XlOrder GetOrder();

        /**
        Sets a XlOrder value that represents the order that Microsoft Excel uses to number pages when printing a large worksheet.

        [MSDN documentation for PageSetup.Order](https://learn.microsoft.com/en-us/office/vba/api/excel.pagesetup.order).
        */
        void SetOrder(XlOrder order);

        /**
        Returns a XlPageOrientation value that represents the portrait or landscape printing mode.

        [MSDN documentation for PageSetup.Orientation](https://learn.microsoft.com/en-us/office/vba/api/excel.pagesetup.orientation).
        */
        XlPageOrientation GetOrientation();

        /**
        Sets a XlPageOrientation value that represents the portrait or landscape printing mode.

        [MSDN documentation for PageSetup.Orientation](https://learn.microsoft.com/en-us/office/vba/api/excel.pagesetup.orientation).
        */
        void SetOrientation(XlPageOrientation orientation);

        /**
        Returns Pages collection. Since Excel 2007.

        [MSDN documentation for PageSetup.Pages](https://learn.microsoft.com/en-us/office/vba/api/excel.pagesetup.pages).
        */
        wxExcelPages GetPages();

        /**
        Returns the size of the paper. Read/write XlPaperSize.

        [MSDN documentation for PageSetup.PaperSize](https://learn.microsoft.com/en-us/office/vba/api/excel.pagesetup.papersize).
        */
        XlPaperSize GetPaperSize();

        /**
        Sets the size of the paper. Read/write XlPaperSize.

        [MSDN documentation for PageSetup.PaperSize](https://learn.microsoft.com/en-us/office/vba/api/excel.pagesetup.papersize).
        */
        void SetPaperSize(XlPaperSize paperSize);

        /**
        Returns the range to be printed, as a string using A1-style references in the language of the macro.

        [MSDN documentation for PageSetup.PrintArea](https://learn.microsoft.com/en-us/office/vba/api/excel.pagesetup.printarea).
        */
        wxString GetPrintArea();

        /**
        Sets the range to be printed, as a string using A1-style references in the language of the macro.

        [MSDN documentation for PageSetup.PrintArea](https://learn.microsoft.com/en-us/office/vba/api/excel.pagesetup.printarea).
        */
        void SetPrintArea(const wxString& printArea);

        /**
        Returns the way comments are printed with the sheet. Read/write XlPrintLocation.

        [MSDN documentation for PageSetup.PrintComments](https://learn.microsoft.com/en-us/office/vba/api/excel.pagesetup.printcomments).
        */
        XlPrintLocation GetPrintComments();

        /**
        Sets the way comments are printed with the sheet. Read/write XlPrintLocation.

        [MSDN documentation for PageSetup.PrintComments](https://learn.microsoft.com/en-us/office/vba/api/excel.pagesetup.printcomments).
        */
        void SetPrintComments(XlPrintLocation printComments);

        /**
        Returns an XlPrintErrors contstant specifying the type of print error displayed. This feature allows users to suppress the display of error values when printing a worksheet.

        [MSDN documentation for PageSetup.PrintErrors](https://learn.microsoft.com/en-us/office/vba/api/excel.pagesetup.printerrors).
        */
        XlPrintErrors GetPrintErrors();

        /**
        Sets an XlPrintErrors contstant specifying the type of print error displayed. This feature allows users to suppress the display of error values when printing a worksheet.

        [MSDN documentation for PageSetup.PrintErrors](https://learn.microsoft.com/en-us/office/vba/api/excel.pagesetup.printerrors).
        */
        void SetPrintErrors(XlPrintErrors printErrors);

        /**
        True if cell gridlines are printed on the page. Applies only to worksheets.

        [MSDN documentation for PageSetup.PrintGridlines](https://learn.microsoft.com/en-us/office/vba/api/excel.pagesetup.printgridlines).
        */
        bool GetPrintGridlines();

        /**
        True if cell gridlines are printed on the page. Applies only to worksheets.

        [MSDN documentation for PageSetup.PrintGridlines](https://learn.microsoft.com/en-us/office/vba/api/excel.pagesetup.printgridlines).
        */
        void SetPrintGridlines(bool printGridlines);

        /**
        True if row and column headings are printed with this page. Applies only to worksheets.

        [MSDN documentation for PageSetup.PrintHeadings](https://learn.microsoft.com/en-us/office/vba/api/excel.pagesetup.printheadings).
        */
        bool GetPrintHeadings();

        /**
        True if row and column headings are printed with this page. Applies only to worksheets.

        [MSDN documentation for PageSetup.PrintHeadings](https://learn.microsoft.com/en-us/office/vba/api/excel.pagesetup.printheadings).
        */
        void SetPrintHeadings(bool printHeadings);

        /**
        True if cell notes are printed as end notes with the sheet. Applies only to worksheets.

        [MSDN documentation for PageSetup.PrintNotes](https://learn.microsoft.com/en-us/office/vba/api/excel.pagesetup.printnotes).
        */
        bool GetPrintNotes();

        /**
        True if cell notes are printed as end notes with the sheet. Applies only to worksheets.

        [MSDN documentation for PageSetup.PrintNotes](https://learn.microsoft.com/en-us/office/vba/api/excel.pagesetup.printnotes).
        */
        void SetPrintNotes(bool printNotes);

        /**
        Returns the print quality.  Index can be either 1 for horizontal quality or 2 for vertical quality, which a printer may not support

        [MSDN documentation for PageSetup.PrintQuality](https://learn.microsoft.com/en-us/office/vba/api/excel.pagesetup.printquality).
        */
        long GetPrintQuality(int index);

        /**
        Sets the print quality. Index can be either 1 for horizontal quality or 2 for vertical quality, which a printer may not support

        [MSDN documentation for PageSetup.PrintQuality](https://learn.microsoft.com/en-us/office/vba/api/excel.pagesetup.printquality).
        */
        void SetPrintQuality(int index, long quality);

        /**
        Returns the columns that contain the cells to be repeated on the left side of each page, as a string in A1-style notation in the language of the macro.

        [MSDN documentation for PageSetup.PrintTitleColumns](https://learn.microsoft.com/en-us/office/vba/api/excel.pagesetup.printtitlecolumns).
        */
        wxString GetPrintTitleColumns();

        /**
        Sets the columns that contain the cells to be repeated on the left side of each page, as a string in A1-style notation in the language of the macro.

        [MSDN documentation for PageSetup.PrintTitleColumns](https://learn.microsoft.com/en-us/office/vba/api/excel.pagesetup.printtitlecolumns).
        */
        void SetPrintTitleColumns(const wxString& printTitleColumns);

        /**
        Returns the rows that contain the cells to be repeated at the top of each page, as a string in A1-style notation in the language of the macro.

        [MSDN documentation for PageSetup.PrintTitleRows](https://learn.microsoft.com/en-us/office/vba/api/excel.pagesetup.printtitlerows).
        */
        wxString GetPrintTitleRows();

        /**
        Sets the rows that contain the cells to be repeated at the top of each page, as a string in A1-style notation in the language of the macro.

        [MSDN documentation for PageSetup.PrintTitleRows](https://learn.microsoft.com/en-us/office/vba/api/excel.pagesetup.printtitlerows).
        */
        void SetPrintTitleRows(const wxString& printTitleRows);

        /**
        Returns the right part of the footer. Since Excel 2007.

        [MSDN documentation for PageSetup.RightFooter](https://learn.microsoft.com/en-us/office/vba/api/excel.pagesetup.rightfooter).
        */
        wxString GetRightFooter();

        /**
        Sets the right part of the footer.

        [MSDN documentation for PageSetup.RightFooter](https://learn.microsoft.com/en-us/office/vba/api/excel.pagesetup.rightfooter).
        */
        void SetRightFooter(const wxString& rightFooter);

        /**
        Returns a Graphic Represents the picture for the right section of the footer. Used to set attributes about the picture.

        [MSDN documentation for PageSetup.RightFooterPicture](https://learn.microsoft.com/en-us/office/vba/api/excel.pagesetup.rightfooterpicture).
        */
        wxExcelGraphic GetRightFooterPicture();

        /**
        Returns the right part of the header.

        [MSDN documentation for PageSetup.RightHeader](https://learn.microsoft.com/en-us/office/vba/api/excel.pagesetup.rightheader).
        */
        wxString GetRightHeader();

        /**
        Sets the right part of the header.

        [MSDN documentation for PageSetup.RightHeader](https://learn.microsoft.com/en-us/office/vba/api/excel.pagesetup.rightheader).
        */
        void SetRightHeader(const wxString& rightHeader);

        /**
        Returns a Graphic Represents the picture for the right section of the header. Used to set attributes about the picture.

        [MSDN documentation for PageSetup.RightHeaderPicture](https://learn.microsoft.com/en-us/office/vba/api/excel.pagesetup.rightheaderpicture).
        */
        wxExcelGraphic GetRightHeaderPicture();

        /**
        Returns the size of the right margin, in points.

        [MSDN documentation for PageSetup.RightMargin](https://learn.microsoft.com/en-us/office/vba/api/excel.pagesetup.rightmargin).
        */
        double GetRightMargin();

        /**
        Sets the size of the right margin, in points.

        [MSDN documentation for PageSetup.RightMargin](https://learn.microsoft.com/en-us/office/vba/api/excel.pagesetup.rightmargin).
        */
        void SetRightMargin(double rightMargin);

        /**
        True if the header and footer should be scaled with the document when the size of the document changes. Since Excel 2007.

        [MSDN documentation for PageSetup.ScaleWithDocHeaderFooter](https://learn.microsoft.com/en-us/office/vba/api/excel.pagesetup.scalewithdocheaderfooter).
        */
        bool GetScaleWithDocHeaderFooter();

        /**
        True if the header and footer should be scaled with the document when the size of the document changes. Since Excel 2007.

        [MSDN documentation for PageSetup.ScaleWithDocHeaderFooter](https://learn.microsoft.com/en-us/office/vba/api/excel.pagesetup.scalewithdocheaderfooter).
        */
        void SetScaleWithDocHeaderFooter(bool scaleWithDocHeaderFooter);

        /**
        Returns the size of the top margin, in points.

        [MSDN documentation for PageSetup.TopMargin](https://learn.microsoft.com/en-us/office/vba/api/excel.pagesetup.topmargin).
        */
        double GetTopMargin();

        /**
        Sets the size of the top margin, in points.

        [MSDN documentation for PageSetup.TopMargin](https://learn.microsoft.com/en-us/office/vba/api/excel.pagesetup.topmargin).
        */
        void SetTopMargin(double topMargin);

        /**
        Returns a value that represents a percentage (between 10 and 400 percent) by which Microsoft Excel will scale the worksheet for printing.

        [MSDN documentation for PageSetup.Zoom](https://learn.microsoft.com/en-us/office/vba/api/excel.pagesetup.zoom).
        */
        long GetZoom();

        /**
        Sets a value that represents a percentage (between 10 and 400 percent) by which Microsoft Excel will scale the worksheet for printing.

        [MSDN documentation for PageSetup.Zoom](https://learn.microsoft.com/en-us/office/vba/api/excel.pagesetup.zoom).
        */
        void SetZoom(long zoom);

        /**
        Returns "PageSetup".
        */
        virtual wxString GetAutoExcelObjectName_() const { return wxS("PageSetup"); }
    };


} // namespace wxAutoExcel

#endif //_WXAUTOEXCEL_PAGESETUP_H
