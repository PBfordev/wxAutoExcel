/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Modified by:
// RCS-ID:      $Id: $
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
/////////////////////////////////////////////////////////////////////////////


#ifndef _WXAUTOEXCEL_PAGESETUP_H
#define _WXAUTOEXCEL_PAGESETUP_H

#include "wx/wxAutoExcel_defs.h"

#include "wx/wxAutoExcelObject.h"
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

        [MSDN documentation for PageSetup.AlignMarginsHeaderFooter](http://msdn.microsoft.com/en-us/library/bb225112).
        */
        bool GetAlignMarginsHeaderFooter();

        /**
        True if the header and the footer are aligned. with the margins set in the page setup options. Since Excel 2007.

        [MSDN documentation for PageSetup.AlignMarginsHeaderFooter](http://msdn.microsoft.com/en-us/library/bb225112).
        */
        void SetAlignMarginsHeaderFooter(bool alignMarginsHeaderFooter);

        /**
        True if elements of the document will be printed in black and white.

        [MSDN documentation for PageSetup.BlackAndWhite](http://msdn.microsoft.com/en-us/library/bb220890).
        */
        bool GetBlackAndWhite();

        /**
        True if elements of the document will be printed in black and white.

        [MSDN documentation for PageSetup.BlackAndWhite](http://msdn.microsoft.com/en-us/library/bb220890).
        */
        void SetBlackAndWhite(bool blackAndWhite);

        /**
        Returns the size of the bottom margin, in points.

        [MSDN documentation for PageSetup.BottomMargin](http://msdn.microsoft.com/en-us/library/bb220892).
        */
        double GetBottomMargin();

        /**
        Sets the size of the bottom margin, in points. 

        [MSDN documentation for PageSetup.BottomMargin](http://msdn.microsoft.com/en-us/library/bb220892).
        */
        void SetBottomMargin(double bottomMargin);

        /**
        Returns the center footer.

        [MSDN documentation for PageSetup.CenterFooter](http://msdn.microsoft.com/en-us/library/bb225118).
        */
        wxString GetCenterFooter();

        /**
        Sets the center footer.

        [MSDN documentation for PageSetup.CenterFooter](http://msdn.microsoft.com/en-us/library/bb225118).
        */
        void SetCenterFooter(const wxString& centerFooter);

        /**
        Returns a Graphic Represents the picture for the center section of the footer. Used to set attributes about the picture.

        [MSDN documentation for PageSetup.CenterFooterPicture](http://msdn.microsoft.com/en-us/library/bb220912).
        */
        wxExcelGraphic GetCenterFooterPicture();

        /**
        Returns the center heaer.

        [MSDN documentation for PageSetup.CenterHeader](http://msdn.microsoft.com/en-us/library/bb225121).
        */
        wxString GetCenterHeader();

        /**
        Sets the center heaer.

        [MSDN documentation for PageSetup.CenterHeader](http://msdn.microsoft.com/en-us/library/bb225121).
        */
        void SetCenterHeader(const wxString& centerHeader);

        /**
        Returns a Graphic Represents the picture for the center section of the header. Used to set attributes about the picture.

        [MSDN documentation for PageSetup.CenterHeaderPicture](http://msdn.microsoft.com/en-us/library/bb220913).
        */
        wxExcelGraphic GetCenterHeaderPicture();


        /**
        True if the sheet is centered horizontally on the page when it's printed.

        [MSDN documentation for PageSetup.CenterHorizontally](http://msdn.microsoft.com/en-us/library/bb220914).
        */
        bool GetCenterHorizontally();

        /**
        True if the sheet is centered horizontally on the page when it's printed.

        [MSDN documentation for PageSetup.CenterHorizontally](http://msdn.microsoft.com/en-us/library/bb220914).
        */
        void SetCenterHorizontally(bool centerHorizontally);

        /**
        True if the sheet is centered vertically on the page when it's printed.

        [MSDN documentation for PageSetup.CenterVertically](http://msdn.microsoft.com/en-us/library/bb220915).
        */
        bool GetCenterVertically();

        /**
        True if the sheet is centered vertically on the page when it's printed.

        [MSDN documentation for PageSetup.CenterVertically](http://msdn.microsoft.com/en-us/library/bb220915).
        */
        void SetCenterVertically(bool centerVertically);

        /**
        True if a different header or footer is used on the first page.

        [MSDN documentation for PageSetup.DifferentFirstPageHeaderFooter](http://msdn.microsoft.com/en-us/library/bb225124).
        */
        bool GetDifferentFirstPageHeaderFooter();

        /**
        True if a different header or footer is used on the first page.

        [MSDN documentation for PageSetup.DifferentFirstPageHeaderFooter](http://msdn.microsoft.com/en-us/library/bb225124).
        */
        void SetDifferentFirstPageHeaderFooter(bool differentFirstPageHeaderFooter);

        /**
        True if the sheet will be printed without graphics.

        [MSDN documentation for PageSetup.Draft](http://msdn.microsoft.com/en-us/library/bb221052).
        */
        bool GetDraft();

        /**
        True if the sheet will be printed without graphics.

        [MSDN documentation for PageSetup.Draft](http://msdn.microsoft.com/en-us/library/bb221052).
        */
        void SetDraft(bool draft);

        /**
        Returns the page settings for the even pages. Since Excel 2007.

        [MSDN documentation for PageSetup.EvenPage](http://msdn.microsoft.com/en-us/library/bb225126).
        */
        wxExcelPage GetEvenPage();

        /**
        Returns the page settings for the first page. Since Excel 2007.

        [MSDN documentation for PageSetup.FirstPage](http://msdn.microsoft.com/en-us/library/bb225128).
        */
        wxExcelPage GetFirstPage();

        /**
        Returns the first page number that will be used when this sheet is printed. If xlAutomatic, Microsoft Excel chooses the first page number. The default is xlAutomatic.

        [MSDN documentation for PageSetup.FirstPageNumber](http://msdn.microsoft.com/en-us/library/bb208512).
        */
        long GetFirstPageNumber();

        /**
        Sets the first page number that will be used when this sheet is printed. If xlAutomatic, Microsoft Excel chooses the first page number. The default is xlAutomatic.

        [MSDN documentation for PageSetup.FirstPageNumber](http://msdn.microsoft.com/en-us/library/bb208512).
        */
        void SetFirstPageNumber(long firstPageNumber);

        /**
        Returns the number of pages tall the worksheet will be scaled to when it's printed. Applies only to worksheets.

        [MSDN documentation for PageSetup.FitToPagesTall](http://msdn.microsoft.com/en-us/library/bb208514).
        */
        long GetFitToPagesTall();

        /**
        Sets the number of pages tall the worksheet will be scaled to when it's printed. Applies only to worksheets.

        [MSDN documentation for PageSetup.FitToPagesTall](http://msdn.microsoft.com/en-us/library/bb208514).
        */
        void SetFitToPagesTall(long fitToPagesTall);

        /**
        Returns the number of pages wide the worksheet will be scaled to when it's printed. Applies only to worksheets.

        [MSDN documentation for PageSetup.FitToPagesWide](http://msdn.microsoft.com/en-us/library/bb208515).
        */
        long GetFitToPagesWide();

        /**
        Sets the number of pages wide the worksheet will be scaled to when it's printed. Applies only to worksheets. 

        [MSDN documentation for PageSetup.FitToPagesWide](http://msdn.microsoft.com/en-us/library/bb208515).
        */
        void SetFitToPagesWide(long fitToPagesWide);

        /**
        Returns the distance from the bottom of the page to the footer, in points. 

        [MSDN documentation for PageSetup.FooterMargin](http://msdn.microsoft.com/en-us/library/bb208526).
        */
        double GetFooterMargin();

        /**
        Sets the distance from the bottom of the page to the footer, in points. 

        [MSDN documentation for PageSetup.FooterMargin](http://msdn.microsoft.com/en-us/library/bb208526).
        */
        void SetFooterMargin(double footerMargin);

        /**
        Returns the distance from the top of the page to the header, in points. 

        [MSDN documentation for PageSetup.HeaderMargin](http://msdn.microsoft.com/en-us/library/bb208664).
        */
        double GetHeaderMargin();

        /**
        Sets the distance from the top of the page to the header, in points. 

        [MSDN documentation for PageSetup.HeaderMargin](http://msdn.microsoft.com/en-us/library/bb208664).
        */
        void SetHeaderMargin(double headerMargin);

        /**
        Returns the left footer.

        [MSDN documentation for PageSetup.LeftFooter](http://msdn.microsoft.com/en-us/library/bb225132).
        */
        wxString GetLeftFooter();

        /**
        Sets the left footer.

        [MSDN documentation for PageSetup.LeftFooter](http://msdn.microsoft.com/en-us/library/bb225132).
        */
        void SetLeftFooter(const wxString& leftFooter);

        /**
        Returns a Graphic Represents the picture for the left section of the footer. Used to set attributes about the picture.

        [MSDN documentation for PageSetup.LeftFooterPicture](http://msdn.microsoft.com/en-us/library/bb177847).
        */
        wxExcelGraphic GetLeftFooterPicture();

        /**
        Returns the left header.

        [MSDN documentation for PageSetup.LeftHeader](http://msdn.microsoft.com/en-us/library/bb225135).
        */
        wxString GetLeftHeader();

        /**
        Sets the left header.

        [MSDN documentation for PageSetup.LeftHeader](http://msdn.microsoft.com/en-us/library/bb225135).
        */
        void SetLeftHeader(const wxString& leftHeader);

        /**
        Returns a Graphic Represents the picture for the left section of the header. Used to set attributes about the picture.

        [MSDN documentation for PageSetup.LeftHeaderPicture](http://msdn.microsoft.com/en-us/library/bb177850).
        */
        wxExcelGraphic GetLeftHeaderPicture();

        /**
        Returns the size of the left margin, in points.

        [MSDN documentation for PageSetup.LeftMargin](http://msdn.microsoft.com/en-us/library/bb177853).
        */
        double GetLeftMargin();

        /**
        Sets the size of the left margin, in points.

        [MSDN documentation for PageSetup.LeftMargin](http://msdn.microsoft.com/en-us/library/bb177853).
        */
        void SetLeftMargin(double leftMargin);

        /**
        True if different headers and footers for odd-numbered and even-numbered pages. Since Excel 2007.

        [MSDN documentation for PageSetup.OddAndEvenPagesHeaderFooter](http://msdn.microsoft.com/en-us/library/bb242106).
        */
        bool GetOddAndEvenPagesHeaderFooter();

        /**
        True if different headers and footers for odd-numbered and even-numbered pages. Since Excel 2007.

        [MSDN documentation for PageSetup.OddAndEvenPagesHeaderFooter](http://msdn.microsoft.com/en-us/library/bb242106).
        */
        void SetOddAndEvenPagesHeaderFooter(bool oddAndEvenPagesHeaderFooter);

        /**
        Returns a XlOrder value that represents the order that Microsoft Excel uses to number pages when printing a large worksheet.

        [MSDN documentation for PageSetup.Order](http://msdn.microsoft.com/en-us/library/bb213218).
        */
        XlOrder GetOrder();

        /**
        Sets a XlOrder value that represents the order that Microsoft Excel uses to number pages when printing a large worksheet.

        [MSDN documentation for PageSetup.Order](http://msdn.microsoft.com/en-us/library/bb213218).
        */
        void SetOrder(XlOrder order);

        /**
        Returns a XlPageOrientation value that represents the portrait or landscape printing mode.

        [MSDN documentation for PageSetup.Orientation](http://msdn.microsoft.com/en-us/library/bb213219).
        */
        XlPageOrientation GetOrientation();

        /**
        Sets a XlPageOrientation value that represents the portrait or landscape printing mode.

        [MSDN documentation for PageSetup.Orientation](http://msdn.microsoft.com/en-us/library/bb213219).
        */
        void SetOrientation(XlPageOrientation orientation);

        /**
        Returns Pages collection. Since Excel 2007.

        [MSDN documentation for PageSetup.Pages](http://msdn.microsoft.com/en-us/library/bb242110).
        */
        wxExcelPages GetPages();

        /**
        Returns the size of the paper. Read/write XlPaperSize.

        [MSDN documentation for PageSetup.PaperSize](http://msdn.microsoft.com/en-us/library/bb208944).
        */
        XlPaperSize GetPaperSize();

        /**
        Sets the size of the paper. Read/write XlPaperSize.

        [MSDN documentation for PageSetup.PaperSize](http://msdn.microsoft.com/en-us/library/bb208944).
        */
        void SetPaperSize(XlPaperSize paperSize);

        /**
        Returns the range to be printed, as a string using A1-style references in the language of the macro.

        [MSDN documentation for PageSetup.PrintArea](http://msdn.microsoft.com/en-us/library/bb208995).
        */
        wxString GetPrintArea();

        /**
        Sets the range to be printed, as a string using A1-style references in the language of the macro.

        [MSDN documentation for PageSetup.PrintArea](http://msdn.microsoft.com/en-us/library/bb208995).
        */
        void SetPrintArea(const wxString& printArea);

        /**
        Returns the way comments are printed with the sheet. Read/write XlPrintLocation.

        [MSDN documentation for PageSetup.PrintComments](http://msdn.microsoft.com/en-us/library/bb208997).
        */
        XlPrintLocation GetPrintComments();

        /**
        Sets the way comments are printed with the sheet. Read/write XlPrintLocation.

        [MSDN documentation for PageSetup.PrintComments](http://msdn.microsoft.com/en-us/library/bb208997).
        */
        void SetPrintComments(XlPrintLocation printComments);

        /**
        Returns an XlPrintErrors contstant specifying the type of print error displayed. This feature allows users to suppress the display of error values when printing a worksheet.

        [MSDN documentation for PageSetup.PrintErrors](http://msdn.microsoft.com/en-us/library/bb209002).
        */
        XlPrintErrors GetPrintErrors();

        /**
        Sets an XlPrintErrors contstant specifying the type of print error displayed. This feature allows users to suppress the display of error values when printing a worksheet.

        [MSDN documentation for PageSetup.PrintErrors](http://msdn.microsoft.com/en-us/library/bb209002).
        */
        void SetPrintErrors(XlPrintErrors printErrors);

        /**
        True if cell gridlines are printed on the page. Applies only to worksheets.

        [MSDN documentation for PageSetup.PrintGridlines](http://msdn.microsoft.com/en-us/library/bb209004).
        */
        bool GetPrintGridlines();

        /**
        True if cell gridlines are printed on the page. Applies only to worksheets.

        [MSDN documentation for PageSetup.PrintGridlines](http://msdn.microsoft.com/en-us/library/bb209004).
        */
        void SetPrintGridlines(bool printGridlines);

        /**
        True if row and column headings are printed with this page. Applies only to worksheets.

        [MSDN documentation for PageSetup.PrintHeadings](http://msdn.microsoft.com/en-us/library/bb209007).
        */
        bool GetPrintHeadings();

        /**
        True if row and column headings are printed with this page. Applies only to worksheets.

        [MSDN documentation for PageSetup.PrintHeadings](http://msdn.microsoft.com/en-us/library/bb209007).
        */
        void SetPrintHeadings(bool printHeadings);

        /**
        True if cell notes are printed as end notes with the sheet. Applies only to worksheets.

        [MSDN documentation for PageSetup.PrintNotes](http://msdn.microsoft.com/en-us/library/bb209008).
        */
        bool GetPrintNotes();

        /**
        True if cell notes are printed as end notes with the sheet. Applies only to worksheets.

        [MSDN documentation for PageSetup.PrintNotes](http://msdn.microsoft.com/en-us/library/bb209008).
        */
        void SetPrintNotes(bool printNotes);

        /**
        Returns the print quality.  Index can be either 1 for horizontal quality or 2 for vertical quality, which a printer may not support

        [MSDN documentation for PageSetup.PrintQuality](http://msdn.microsoft.com/en-us/library/bb209010).
        */
        long GetPrintQuality(int index);

        /**
        Sets the print quality. Index can be either 1 for horizontal quality or 2 for vertical quality, which a printer may not support

        [MSDN documentation for PageSetup.PrintQuality](http://msdn.microsoft.com/en-us/library/bb209010).
        */
        void SetPrintQuality(int index, long quality);

        /**
        Returns the columns that contain the cells to be repeated on the left side of each page, as a string in A1-style notation in the language of the macro.

        [MSDN documentation for PageSetup.PrintTitleColumns](http://msdn.microsoft.com/en-us/library/bb209013).
        */
        wxString GetPrintTitleColumns();

        /**
        Sets the columns that contain the cells to be repeated on the left side of each page, as a string in A1-style notation in the language of the macro.

        [MSDN documentation for PageSetup.PrintTitleColumns](http://msdn.microsoft.com/en-us/library/bb209013).
        */
        void SetPrintTitleColumns(const wxString& printTitleColumns);

        /**
        Returns the rows that contain the cells to be repeated at the top of each page, as a string in A1-style notation in the language of the macro.

        [MSDN documentation for PageSetup.PrintTitleRows](http://msdn.microsoft.com/en-us/library/bb209015).
        */
        wxString GetPrintTitleRows();

        /**
        Sets the rows that contain the cells to be repeated at the top of each page, as a string in A1-style notation in the language of the macro.

        [MSDN documentation for PageSetup.PrintTitleRows](http://msdn.microsoft.com/en-us/library/bb209015).
        */
        void SetPrintTitleRows(const wxString& printTitleRows);

        /**
        Returns the right part of the footer. Since Excel 2007.

        [MSDN documentation for PageSetup.RightFooter](http://msdn.microsoft.com/en-us/library/bb225140).
        */
        wxString GetRightFooter();

        /**
        Sets the right part of the footer.

        [MSDN documentation for PageSetup.RightFooter](http://msdn.microsoft.com/en-us/library/bb225140).
        */
        void SetRightFooter(const wxString& rightFooter);

        /**
        Returns a Graphic Represents the picture for the right section of the footer. Used to set attributes about the picture.

        [MSDN documentation for PageSetup.RightFooterPicture](http://msdn.microsoft.com/en-us/library/bb209165).
        */
        wxExcelGraphic GetRightFooterPicture();

        /**
        Returns the right part of the header.

        [MSDN documentation for PageSetup.RightHeader](http://msdn.microsoft.com/en-us/library/bb209171).
        */
        wxString GetRightHeader();

        /**
        Sets the right part of the header.

        [MSDN documentation for PageSetup.RightHeader](http://msdn.microsoft.com/en-us/library/bb209171).
        */
        void SetRightHeader(const wxString& rightHeader);

        /**
        Returns a Graphic Represents the picture for the right section of the header. Used to set attributes about the picture.

        [MSDN documentation for PageSetup.RightHeaderPicture](http://msdn.microsoft.com/en-us/library/bb225143).
        */
        wxExcelGraphic GetRightHeaderPicture();

        /**
        Returns the size of the right margin, in points. 

        [MSDN documentation for PageSetup.RightMargin](http://msdn.microsoft.com/en-us/library/bb209175).
        */
        double GetRightMargin();

        /**
        Sets the size of the right margin, in points.

        [MSDN documentation for PageSetup.RightMargin](http://msdn.microsoft.com/en-us/library/bb209175).
        */
        void SetRightMargin(double rightMargin);

        /**
        True if the header and footer should be scaled with the document when the size of the document changes. Since Excel 2007.

        [MSDN documentation for PageSetup.ScaleWithDocHeaderFooter](http://msdn.microsoft.com/en-us/library/bb225146).
        */
        bool GetScaleWithDocHeaderFooter();

        /**
        True if the header and footer should be scaled with the document when the size of the document changes. Since Excel 2007.

        [MSDN documentation for PageSetup.ScaleWithDocHeaderFooter](http://msdn.microsoft.com/en-us/library/bb225146).
        */
        void SetScaleWithDocHeaderFooter(bool scaleWithDocHeaderFooter);

        /**
        Returns the size of the top margin, in points.

        [MSDN documentation for PageSetup.TopMargin](http://msdn.microsoft.com/en-us/library/bb221846).
        */
        double GetTopMargin();

        /**
        Sets the size of the top margin, in points.

        [MSDN documentation for PageSetup.TopMargin](http://msdn.microsoft.com/en-us/library/bb221846).
        */
        void SetTopMargin(double topMargin);

        /**
        Returns a value that represents a percentage (between 10 and 400 percent) by which Microsoft Excel will scale the worksheet for printing.

        [MSDN documentation for PageSetup.Zoom](http://msdn.microsoft.com/en-us/library/bb214929).
        */
        long GetZoom();

        /**
        Sets a value that represents a percentage (between 10 and 400 percent) by which Microsoft Excel will scale the worksheet for printing.

        [MSDN documentation for PageSetup.Zoom](http://msdn.microsoft.com/en-us/library/bb214929).
        */
        void SetZoom(long zoom);

        /**
        Returns "PageSetup".
        */
        virtual wxString GetAutoExcelObjectName_() const { return wxS("PageSetup"); }
    };


} // namespace wxAutoExcel

#endif //_WXAUTOEXCEL_PAGESETUP_H
