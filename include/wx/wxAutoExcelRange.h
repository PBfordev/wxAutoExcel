/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// License:     MIT license
/////////////////////////////////////////////////////////////////////////////


#ifndef _WXAUTOEXCEL_RANGE_H
#define _WXAUTOEXCEL_RANGE_H

#include "wx/wxAutoExcelRangeOwner.h"

#include "wx/wxAutoExcel_enums.h"

namespace wxAutoExcel {

    typedef wxVector<wxExcelRange> wxExcelRangeVector;

    /**
    @brief Represents Microsoft Excel Range, i.e. a collection of cells.
    */
   class WXDLLIMPEXP_WXAUTOEXCEL wxExcelRange : public wxExcelRangeOwner
    {
    public:
        // ***** METHODS *****

        /**
        Activates a single cell, which must be inside the current selection. To select a range of cells, use the Select method.

        [MSDN documentation for Range.Activate](https://learn.microsoft.com/en-us/office/vba/api/excel.range.activate).
        */
        bool Activate();

        /**
        Adds a comment to the range.

        [MSDN documentation for Range.AddComment](https://learn.microsoft.com/en-us/office/vba/api/excel.range.addcomment).
        */
        wxExcelComment AddComment(const wxString& text = wxEmptyString);

        /**
        Adds a threaded comment to the range. Since Excel for Microsoft 365.

        [MSDN documentation for Range.AddCommentThreaded](https://learn.microsoft.com/en-us/office/vba/api/excel.range.addcommentthreaded).
        */
        wxExcelCommentThreaded AddCommentThreaded(const wxString& text);

        /**
        Filters or copies data from a list based on a criteria range. If the initial selection is a single cell, that cell's current region is used.

        [MSDN documentation for Range.AdvancedFilter](https://learn.microsoft.com/en-us/office/vba/api/excel.range.advancedfilter).
        */
        bool AdvancedFilter(enum XlFilterAction action,
                            wxExcelRange* criteriaRange = NULL, wxExcelRange* copyToRange = NULL,
                            wxXlTribool unique = wxDefaultXlTribool);

        /**
        Performs a writeback operation for all edited cells in a range based on an OLAP data source. Since Excel 2010.

        [MSDN documentation for Range.AllocateChanges](https://learn.microsoft.com/en-us/office/vba/api/excel.range.allocatechanges).
        */
        void AllocateChanges();

        /**
        Applies names to the cells in the specified range.

        [MSDN documentation for Range.ApplyNames](https://learn.microsoft.com/en-us/office/vba/api/excel.range.applynames).
        */
        void ApplyNames(wxArrayString* names = NULL,
                        wxXlTribool ignoreRelativeAbsolute =  NULL, wxXlTribool useRowColumnNames = wxDefaultXlTribool,
                        wxXlTribool omitColumn = wxDefaultXlTribool, wxXlTribool omitRow = wxDefaultXlTribool,
                        XlApplyNamesOrder* order = NULL, wxXlTribool appendLast = wxDefaultXlTribool);

        /**
        Applies outlining styles to the specified range.

        [MSDN documentation for Range.ApplyOutlineStyles](https://learn.microsoft.com/en-us/office/vba/api/excel.range.applyoutlinestyles).
        */
        void ApplyOutlineStyles();

        /**
        Returns an AutoComplete match from the list. If there's no AutoComplete match or if more than one entry in the list matches the string to complete, this method returns an empty string.

        [MSDN documentation for Range.AutoComplete](https://learn.microsoft.com/en-us/office/vba/api/excel.range.autocomplete).
        */
        wxString AutoComplete(const wxString& str);

        /**
        Performs an autofill on the cells in the specified range.

        [MSDN documentation for Range.AutoFill](https://learn.microsoft.com/en-us/office/vba/api/excel.range.autofill).
        */
        bool AutoFill(wxExcelRange destination, XlAutoFillType* type = NULL);

        /**
        Filters a list using the AutoFilter.

        [MSDN documentation for Range.AutoFilter](https://learn.microsoft.com/en-us/office/vba/api/excel.range.autofilter).
        */
        bool AutoFilter(long* field = NULL, const wxString& criteria1 = wxEmptyString,
                        XlAutoFilterOperator* oper = NULL, const wxString& criteria2 = wxEmptyString,
                        wxXlTribool visibleDropDown = wxDefaultXlTribool);

        /**
        Changes the width of the columns in the range or the height of the rows in the range to achieve the best fit.

        [MSDN documentation for Range.AutoFit](https://learn.microsoft.com/en-us/office/vba/api/excel.range.autofit).
        */
        bool AutoFit();

        /**
        Automatically creates an outline for the specified range. If the range is a single cell, Microsoft Excel creates an outline for the entire sheet. The new outline replaces any existing outline.

        [MSDN documentation for Range.AutoOutline](https://learn.microsoft.com/en-us/office/vba/api/excel.range.autooutline).
        */
        void AutoOutline();

        /**
        Adds a border to a range and sets the Color, LineStyle, and Weight properties for the new border.

        [MSDN documentation for Range.BorderAround](https://learn.microsoft.com/en-us/office/vba/api/excel.range.borderaround).
        */
        bool BorderAround(XlLineStyle* lineStyle = NULL, XlBorderWeight* weight = NULL,
                          long* colorIndex = NULL, const wxColour* color = NULL);

        /**
        Calculates the range.

        [MSDN documentation for Range.Calculate](https://learn.microsoft.com/en-us/office/vba/api/excel.range.calculate).
        */
        void Calculate();

        /**
        Calculates the range the same way Excel 2000 and earlier did.

        [MSDN documentation for Range.CalculateRowMajorOrder](https://learn.microsoft.com/en-us/office/vba/api/excel.range.calculaterowmajororder).
        */
        void CalculateRowMajorOrder();

        /**
        Checks the spelling of an object.

        [MSDN documentation for Range.CheckSpelling](https://learn.microsoft.com/en-us/office/vba/api/excel.range.checkspelling).
        */
        void CheckSpelling(const wxString& customDictionary = wxEmptyString, wxXlTribool ignoreUpperCase = wxDefaultXlTribool,
                           wxXlTribool alwaysSuggest = wxDefaultXlTribool, MsoLanguageID* spellLang = NULL);

        /**
        Clears the entire object.

        [MSDN documentation for Range.Clear](https://learn.microsoft.com/en-us/office/vba/api/excel.range.clear).
        */
        bool Clear();

        /**
        Clears all cell comments from the specified range.

        [MSDN documentation for Range.ClearComments](https://learn.microsoft.com/en-us/office/vba/api/excel.range.clearcomments).
        */
        bool ClearComments();

        /**
        Clears the formulas from the range.

        [MSDN documentation for Range.ClearContents](https://learn.microsoft.com/en-us/office/vba/api/excel.range.clearcontents).
        */
        bool ClearContents();

        /**
        Clears the formatting of the object.

        [MSDN documentation for Range.ClearFormats](https://learn.microsoft.com/en-us/office/vba/api/excel.range.clearformats).
        */
        bool ClearFormats();

         /**
        Removes all hyperlinks from the specified range. Since Excel 2010.

        [MSDN documentation for Range.ClearHyperlinks](https://learn.microsoft.com/en-us/office/vba/api/excel.range.clearhyperlinks).
        */
        bool ClearHyperlinks();

        /**
        Clears notes and sound notes from all the cells in the specified range.

        [MSDN documentation for Range.ClearNotes](https://learn.microsoft.com/en-us/office/vba/api/excel.range.clearnotes).
        */
        bool ClearNotes();

        /**
        Clears the outline for the specified range.

        [MSDN documentation for Range.ClearOutline](https://learn.microsoft.com/en-us/office/vba/api/excel.range.clearoutline).
        */
        bool ClearOutline();

        /**
        Returns a Range Represents all the cells whose contents are different from the comparison cell in each column.

        [MSDN documentation for Range.ColumnDifferences](https://learn.microsoft.com/en-us/office/vba/api/excel.range.columndifferences).
        */
        wxExcelRange ColumnDifferences(wxExcelRange comparison);

        /**
        Copies the range to the specified range or to the Clipboard.

        [MSDN documentation for Range.Copy](https://learn.microsoft.com/en-us/office/vba/api/excel.range.copy).
        */
        bool Copy(const wxExcelRange* destination = NULL);

        /**
        Copies the selected object to the Clipboard as a picture.

        [MSDN documentation for Range.CopyPicture](https://learn.microsoft.com/en-us/office/vba/api/excel.range.copypicture).
        */
        bool CopyPicture(XlPictureAppearance* appearance = NULL, XlCopyPictureFormat* format = NULL);

        /**
        Creates names in the specified range, based on text labels in the sheet.

        [MSDN documentation for Range.CreateNames](https://learn.microsoft.com/en-us/office/vba/api/excel.range.createnames).
        */
        bool CreateNames(wxXlTribool top = wxDefaultXlTribool, wxXlTribool left = wxDefaultXlTribool, wxXlTribool bottom = wxDefaultXlTribool, wxXlTribool right = wxDefaultXlTribool);

        /**
        Cuts the object to the Clipboard or pastes it into a specified destination.

        [MSDN documentation for Range.Cut](https://learn.microsoft.com/en-us/office/vba/api/excel.range.cut).
        */
        bool Cut(const wxExcelRange* destination = NULL);

        /**
        Creates a data series in the specified range.

        [MSDN documentation for Range.DataSeries](https://learn.microsoft.com/en-us/office/vba/api/excel.range.dataseries).
        */
        bool DataSeries(XlRowCol* rowCol = NULL, XlDataSeriesType* type = NULL, XlDataSeriesDate* date = NULL,
                                long* step = NULL, const wxVariant& stop = wxNullVariant, wxXlTribool trend = wxDefaultXlTribool);

        /**
        Deletes the object.

        [MSDN documentation for Range.Delete](https://learn.microsoft.com/en-us/office/vba/api/excel.range.delete).
        */
        bool Delete(XlDeleteShiftDirection* shift = NULL);


        /**
        Designates a range to be recalculated when the next recalculation occurs.

        [MSDN documentation for Range.Dirty](https://learn.microsoft.com/en-us/office/vba/api/excel.range.dirty).
        */
        void Dirty();

        /**
        Discards all changes in the edited cells of the range. Since Excel 2010.

        [MSDN documentation for Range.DiscardChanges](https://learn.microsoft.com/en-us/office/vba/api/excel.range.discardchanges).
        */
        void DiscardChanges();


        //@{
        /**
        Exports to an XPS or PDF file.

        [MSDN documentation for Range.ExportAsFixedFormat](https://learn.microsoft.com/en-us/office/vba/api/excel.range.exportasfixedformat).
        */
        void ExportAsFixedFormat(XlFixedFormatType type, const wxString& fileName = wxEmptyString,
                                 XlFixedFormatQuality* quality = NULL, wxXlTribool includeDocProperties = wxDefaultXlTribool,
                                 wxXlTribool ignorePrintAreas = wxDefaultXlTribool,
                                 long* from = NULL, long* to = NULL, wxXlTribool openAfterPublish = wxDefaultXlTribool);

        void ExportAsFixedFormat(XlFixedFormatType type, const wxVariantVector& optionalArgs);
        //@}

        /**
        Fills down from the top cell or cells in the specified range to the bottom of the range. The contents and formatting of the cell or cells in the top row of a range are copied into the rest of the rows in the range.

        [MSDN documentation for Range.FillDown](https://learn.microsoft.com/en-us/office/vba/api/excel.range.filldown).
        */
        bool FillDown();

        /**
        Fills left from the rightmost cell or cells in the specified range. The contents and formatting of the cell or cells in the rightmost column of a range are copied into the rest of the columns in the range.

        [MSDN documentation for Range.FillLeft](https://learn.microsoft.com/en-us/office/vba/api/excel.range.fillleft).
        */
        bool FillLeft();

        /**
        Fills right from the leftmost cell or cells in the specified range. The contents and formatting of the cell or cells in the leftmost column of a range are copied into the rest of the columns in the range.

        [MSDN documentation for Range.FillRight](https://learn.microsoft.com/en-us/office/vba/api/excel.range.fillright).
        */
        bool FillRight();

        /**
        Fills up from the bottom cell or cells in the specified range to the top of the range. The contents and formatting of the cell or cells in the bottom row of a range are copied into the rest of the rows in the range.

        [MSDN documentation for Range.FillUp](https://learn.microsoft.com/en-us/office/vba/api/excel.range.fillup).
        */
        bool FillUp();

         /**
        Flash Fill automatically fills the data when it senses a pattern.

        [Excel VBA documentation for Range.FlashFill](https://learn.microsoft.com/en-us/office/vba/api/excel.range.flashfill)
        */
        void FlashFill();


        /**
        Starts the Function Wizard for the upper-left cell of the range.

        [MSDN documentation for Range.FunctionWizard](https://learn.microsoft.com/en-us/office/vba/api/excel.range.functionwizard).
        */
        void FunctionWizard();

        /**
        Inserts a cell or a range of cells into the worksheet or macro sheet and shifts other cells away to make space.

        [MSDN documentation for Range.Insert](https://learn.microsoft.com/en-us/office/vba/api/excel.range.insert).
        */
        bool Insert(XlInsertShiftDirection* shift);

        /**
        Adds an indent to the specified range.

        [MSDN documentation for Range.InsertIndent](https://learn.microsoft.com/en-us/office/vba/api/excel.range.insertindent).
        */
        void InsertIndent(long insertAmount);

        /**
        Rearranges the text in a range so that it fills the range evenly.

        [MSDN documentation for Range.Justify](https://learn.microsoft.com/en-us/office/vba/api/excel.range.justify).
        */
        bool Justify();

        /**
        Pastes a list of all nonhidden names onto the worksheet, beginning with the first cell in the range.

        [MSDN documentation for Range.ListNames](https://learn.microsoft.com/en-us/office/vba/api/excel.range.listnames).
        */
        bool ListNames();

        /**
        Creates a merged cell from the specified Range object.

        [MSDN documentation for Range.Merge](https://learn.microsoft.com/en-us/office/vba/api/excel.range.merge).
        */
        void Merge(wxXlTribool across = wxDefaultXlTribool);

        /**
        Navigates a tracer arrow for the specified range to the precedent, dependent, or error-causing cell or cells. Selects the precedent, dependent, or error cells and returns a Range Represents the new selection. This method causes an error if it's applied to a cell without visible tracer arrows.

        [MSDN documentation for Range.NavigateArrow](https://learn.microsoft.com/en-us/office/vba/api/excel.range.navigatearrow).
        */
        bool NavigateArrow(wxXlTribool towardPrecedent = wxDefaultXlTribool, long* arrowNumber = NULL, long* linkNumber = NULL);

        /**
        Pastes a Range from the Clipboard into the specified range.

        [MSDN documentation for Range.PasteSpecial](https://learn.microsoft.com/en-us/office/vba/api/excel.range.pastespecial).
        */
        bool PasteSpecial(XlPasteType* paste = NULL, XlPasteSpecialOperation* operation = NULL,
                          wxXlTribool skipBlanks = wxDefaultXlTribool, wxXlTribool transpose = wxDefaultXlTribool);

        //@{
        /**
        Prints the object.

        [MSDN documentation for Range.PrintOut](https://learn.microsoft.com/en-us/office/vba/api/excel.range.printout).
        */
        bool PrintOut(long* from = NULL, long* to = NULL, long* copies = NULL, wxXlTribool preview = wxDefaultXlTribool,
                      const wxString& activePrinter = wxEmptyString, wxXlTribool printToFile = wxDefaultXlTribool,
                      wxXlTribool collate = wxDefaultXlTribool, const wxString& prToFileName= wxEmptyString);
        bool PrintOut(const wxVariantVector& args);
        //@}

        /**
        Shows a preview of the object as it would look when printed.

        [MSDN documentation for Range.PrintPreview](https://learn.microsoft.com/en-us/office/vba/api/excel.range.printpreview).
        */
        bool PrintPreview(wxXlTribool enableChanges = wxDefaultXlTribool);

        /**
        Removes duplicate values from a range of values.

        [MSDN documentation for Range.RemoveDuplicates](https://learn.microsoft.com/en-us/office/vba/api/excel.range.removeduplicates).
        */
        void RemoveDuplicates(const wxArrayLong& columns, XlYesNoGuess* header = NULL);

        /**
        Removes subtotals from a list.

        [MSDN documentation for Range.RemoveSubtotal](https://learn.microsoft.com/en-us/office/vba/api/excel.range.removesubtotal).
        */
        void RemoveSubtotal();

        /**
        Returns a Range Represents all the cells whose contents are different from those of the comparison cell in each row.

        [MSDN documentation for Range.RowDifferences](https://learn.microsoft.com/en-us/office/vba/api/excel.range.rowdifferences).
        */
        wxExcelRange RowDifferences(wxExcelRange comparison);

        /**
        Selects the object.

        [MSDN documentation for Range.Select](https://learn.microsoft.com/en-us/office/vba/api/excel.range.select).
        */
        bool Select();


        /**
        Scrolls through the contents of the active window to move the range into view. The range must consist of a single cell in the active document.

        [MSDN documentation for Range.Show](https://learn.microsoft.com/en-us/office/vba/api/excel.range.show).
        */
        bool Show();

        /**
        Draws tracer arrows to the direct dependents of the range.

        [MSDN documentation for Range.ShowDependents](https://learn.microsoft.com/en-us/office/vba/api/excel.range.showdependents).
        */
        bool ShowDependents(wxXlTribool remove = wxDefaultXlTribool);

        /**
        Draws tracer arrows through the precedents tree to the cell that's the source of the error, and returns the range that contains that cell.

        [MSDN documentation for Range.ShowErrors](https://learn.microsoft.com/en-us/office/vba/api/excel.range.showerrors).
        */
        wxExcelRange ShowErrors();

        /**
        Draws tracer arrows to the direct precedents of the range.

        [MSDN documentation for Range.ShowPrecedents](https://learn.microsoft.com/en-us/office/vba/api/excel.range.showprecedents).
        */
        bool ShowPrecedents(wxXlTribool remove = wxDefaultXlTribool);

        /**
        Returns a Range Represents all the cells that match the specified type and value.

        [MSDN documentation for Range.SpecialCells](https://learn.microsoft.com/en-us/office/vba/api/excel.range.specialcells).
        */
        wxExcelRange SpecialCells(XlCellType type, const wxVariant& value = wxNullVariant);


        /**
        Creates subtotals for the range (or the current region, if the range is a single cell).

        [MSDN documentation for Range.Subtotal](https://learn.microsoft.com/en-us/office/vba/api/excel.range.subtotal).
        */
        bool Subtotal(XlConsolidationFunction groupBy, XlConsolidationFunction function,
                      wxArrayLong& totalList, wxXlTribool replace = wxDefaultXlTribool, wxXlTribool pageBreaks = wxDefaultXlTribool,
                      XlSummaryRow* summaryBelowData = NULL);

        /**
        Creates a data table based on input values and formulas that you define on a worksheet.

        [MSDN documentation for Range.Table](https://learn.microsoft.com/en-us/office/vba/api/excel.range.table).
        */
        void Table(wxExcelRange* rowInput = NULL, wxExcelRange* columnInput = NULL);

        /**
        Promotes a range in an outline (that is, decreases its outline level). The specified range must be a row or column, or a range of rows or columns. If the range is in a PivotTable report, this method ungroups the items contained in the range.

        [MSDN documentation for Range.Ungroup](https://learn.microsoft.com/en-us/office/vba/api/excel.range.ungroup).
        */
        bool Ungroup();

        /**
        Separates a merged area into individual cells.

        [MSDN documentation for Range.UnMerge](https://learn.microsoft.com/en-us/office/vba/api/excel.range.unmerge).
        */
        void UnMerge();

        // ***** PROPERTIES *****

        /**
        True if text is automatically indented when the text alignment in a cell is set to equal distribution (either horizontally or vertically.);
        [MSDN documentation for Range.AddIndent](https://learn.microsoft.com/en-us/office/vba/api/excel.range.addindent).
        */
        bool GetAddIndent();
        /**
        True if text is automatically indented when the text alignment in a cell is set to equal distribution (either horizontally or vertically.);
        [MSDN documentation for Range.AddIndent](https://learn.microsoft.com/en-us/office/vba/api/excel.range.addindent).
        */
        void SetAddIndent(bool addIndent);

        /**
        Returns a value that represents the range reference in the language of the macro.

        For a range with several areas, the separator of the areas is either a comma or the list
        separator of the Windows regional settings, depending on how the range was obtained.
        Use GetAddressLocal() to get an address which can be passed to GetRange().

        [MSDN documentation for Range.Address](https://learn.microsoft.com/en-us/office/vba/api/excel.range.address).
        */
        wxString GetAddress(wxXlTribool rowAbsolute = wxDefaultXlTribool, wxXlTribool columnAbsolute = wxDefaultXlTribool,
                            XlReferenceStyle* referenceStyle = NULL, wxXlTribool external = wxDefaultXlTribool,
                            wxExcelRange* relativeTo = NULL);
        /**
        Returns the range reference for the specified range in the language of the user.

        For a range with several areas, the areas are separated with the list separator of the
        Windows regional settings, which GetRange() expects, see wxExcelApplication::GetListSeparator_().

        [MSDN documentation for Range.AddressLocal](https://learn.microsoft.com/en-us/office/vba/api/excel.range.addresslocal).
        */
        wxString GetAddressLocal(wxXlTribool rowAbsolute = wxDefaultXlTribool, wxXlTribool columnAbsolute = wxDefaultXlTribool,
                                XlReferenceStyle* referenceStyle = NULL, wxXlTribool external = wxDefaultXlTribool,
                                wxExcelRange* relativeTo = NULL);
        /**
        True if the range can be edited on a protected worksheet.

        [MSDN documentation for Range.AllowEdit](https://learn.microsoft.com/en-us/office/vba/api/excel.range.allowedit).
        */
        bool GetAllowEdit();

        /**
        Returns an Areas collection that represents all the ranges in a multiple-area selection.

        [MSDN documentation for Range.Areas](https://learn.microsoft.com/en-us/office/vba/api/excel.range.areas).
        */
        wxExcelAreas GetAreas();

        /**
        Returns a Borders collection that represents the borders of a style or a range of cells (including a range defined as part of a conditional format).

        [MSDN documentation for Range.Borders](https://learn.microsoft.com/en-us/office/vba/api/excel.range.borders).
        */
        wxExcelBorders GetBorders();

        /**
        Returns a Characters Represents a range of characters within the object text. You can use the Characters object to format characters within a text string.

        [MSDN documentation for Range.Characters](https://learn.microsoft.com/en-us/office/vba/api/excel.range.characters).
        */
        wxExcelCharacters GetCharacters(long start = 1, long* length = NULL);

        /**
        Returns the number of the first column in the first area in the specified range.

        [MSDN documentation for Range.Column](https://learn.microsoft.com/en-us/office/vba/api/excel.range.column).
        */
        long GetColumn();

        /**
        If all columns in the range have same width returns the width in points else returns -1. Returns -2 on error.

        [MSDN documentation for Range.ColumnWidth](https://learn.microsoft.com/en-us/office/vba/api/excel.range.columnwidth).
        */
        double GetColumnWidth();

         /**
        Sets the column width for all the columns in the range.

        [MSDN documentation for Range.ColumnWidth](https://learn.microsoft.com/en-us/office/vba/api/excel.range.columnwidth).
        */
        void SetColumnWidth(double colWidth);

        /**
        Returns a comment associated with the cell in the upper-left corner of the range.

        [MSDN documentation for Range.Comment](https://learn.microsoft.com/en-us/office/vba/api/excel.range.comment).
        */
        wxExcelComment GetComment();

        /**
        Returns a CommentThreaded object that represents the threaded comment associated with the cell in the upper-left corner of the range.

        [MSDN documentation for Range.CommentThreaded](https://learn.microsoft.com/en-us/office/vba/api/excel.range.commentthreaded).
        */
        wxExcelCommentThreaded GetCommentThreaded();

        /**
        Returns a value that represents the number of cells in the range.

        [MSDN documentation for Range.Count](https://learn.microsoft.com/en-us/office/vba/api/excel.range.count).
        */
        long GetCount();

        /**
        Returns a value that represents the number of cells in the range. Since Excel 2007.

        [MSDN documentation for Range.CountLarge](https://learn.microsoft.com/en-us/office/vba/api/excel.range.countlarge).
        */
        wxLongLong GetCountLarge();

        /**
        If the specified cell is part of an array, returns a Range Represents the entire array.

        [MSDN documentation for Range.CurrentArray](https://learn.microsoft.com/en-us/office/vba/api/excel.range.currentarray).
        */
        wxExcelRange GetCurrentArray();

        /**
        Returns a Range Represents the current region. The current region is a range bounded by any combination of blank rows and blank columns.

        [MSDN documentation for Range.CurrentRegion](https://learn.microsoft.com/en-us/office/vba/api/excel.range.currentregion).
        */
        wxExcelRange GetCurrentRegion();

        /**
        Returns a Range Represents the range containing all the dependents of a cell. This can be a multiple selection (a union of Range objects) if there's more than one dependent.

        [MSDN documentation for Range.Dependents](https://learn.microsoft.com/en-us/office/vba/api/excel.range.dependents).
        */
        wxExcelRange GetDependents();

        /**
        Returns a Range Represents the range containing all the direct dependents of a cell. This can be a multiple selection (a union of Range objects) if there's more than one dependent.

        [MSDN documentation for Range.DirectDependents](https://learn.microsoft.com/en-us/office/vba/api/excel.range.directdependents).
        */
        wxExcelRange GetDirectDependents();

        /**
        Returns a Range Represents the range containing all the direct precedents of a cell. This can be a multiple selection (a union of Range objects) if there's more than one precedent.

        [MSDN documentation for Range.DirectPrecedents](https://learn.microsoft.com/en-us/office/vba/api/excel.range.directprecedents).
        */
        wxExcelRange GetDirectPrecedents();

        /**
        Returns a DisplayFormat object that represents the display settings for the range,
        including the formatting applied by conditional formatting. Since Excel 2010.

        [MSDN documentation for Range.DisplayFormat](https://learn.microsoft.com/en-us/office/vba/api/excel.range.displayformat).
        */
        wxExcelDisplayFormat GetDisplayFormat();

        /**
        Returns a Range Represents the cell at the end of the region that contains the source range. Equivalent to pressing END+UP ARROW, END+DOWN ARROW, END+LEFT ARROW, or END+RIGHT ARROW.

        [MSDN documentation for Range.End](https://learn.microsoft.com/en-us/office/vba/api/excel.range.end).
        */
        wxExcelRange GetEnd(XlDirection direction);

        /**
        Returns a Range Represents the entire column (or columns) that contains the specified range.

        [MSDN documentation for Range.EntireColumn](https://learn.microsoft.com/en-us/office/vba/api/excel.range.entirecolumn).
        */
        wxExcelRange GetEntireColumn();

        /**
        Returns a Range Represents the entire row (or rows) that contains the specified range.

        [MSDN documentation for Range.EntireRow](https://learn.microsoft.com/en-us/office/vba/api/excel.range.entirerow).
        */
        wxExcelRange GetEntireRow();


        /**
        Allows the user to to access error checking options.

        [MSDN documentation for Range.Errors](https://learn.microsoft.com/en-us/office/vba/api/excel.range.errors).
        */
        wxExcelErrors GetErrors();

        /**
        Returns a Font Represents the font of the specified object.

        [MSDN documentation for Range.Font](https://learn.microsoft.com/en-us/office/vba/api/excel.range.font).
        */
        wxExcelFont GetFont();

#if WXAUTOEXCEL_USE_CONDFORMAT
        /**
        Returns a FormatConditions collection that represents all the conditional formats for the specified range.

        [MSDN documentation for Range.FormatConditions](https://learn.microsoft.com/en-us/office/vba/api/excel.range.formatconditions).
        */

        wxExcelFormatConditions GetFormatConditions();

#endif  // WXAUTOEXCEL_USE_CONDFORMAT



        /**
        Returns the formula of a single cell, in A1-style notation and in the language of the macro.

        For a range with more than one cell, Excel returns the formulas of all
        its cells, which this method cannot return, so it fails; use
        GetFormula(wxVariant&) instead.

        [MSDN documentation for Range.Formula](https://learn.microsoft.com/en-us/office/vba/api/excel.range.formula).
        */
        wxString GetFormula();
        /**
        Returns the formulas of all cells in the range, in A1-style notation and in the language of the macro.

        For a single cell, @a formulas contains a string. For more cells, it
        contains the formulas in the same way GetValue() returns the values,
        i.e. a list of strings ordered by columns, or a two-dimensional
        SAFEARRAY when @c wxOleConvertVariant_ReturnSafeArrays is set with
        SetConvertVariantFlags_(). Cells without a formula have their constant
        value as a string, empty cells an empty string.

        Returns @c true if the formulas were obtained.

        [MSDN documentation for Range.Formula](https://learn.microsoft.com/en-us/office/vba/api/excel.range.formula).
        */
        bool GetFormula(wxVariant& formulas);
        /**
        Sets the formula of all cells in the range, in A1-style notation and in the language of the macro.

        A string sets the same formula for all cells, with its relative references
        adjusted for each cell. To set a different formula for each cell, pass an
        array with the same dimensions as the range, e.g. a two-dimensional
        SAFEARRAY created with wxSafeArray, see the bulkdata sample; a list
        of strings can be used for a range with a single row.
        A null wxVariant clears the cells.

        [MSDN documentation for Range.Formula](https://learn.microsoft.com/en-us/office/vba/api/excel.range.formula).
        */
        void SetFormula(const wxVariant& formula);

        /**
        Returns the array formula of a range. Returns (or can be set to) a single formula or a Visual Basic array. If the specified range doesn't contain an array formula, this property returns null.

        [MSDN documentation for Range.FormulaArray](https://learn.microsoft.com/en-us/office/vba/api/excel.range.formulaarray).
        */
        wxString GetFormulaArray();
        /**
        Sets the array formula of a range.

        [MSDN documentation for Range.FormulaArray](https://learn.microsoft.com/en-us/office/vba/api/excel.range.formulaarray).
        */
        void SetFormulaArray(const wxString& formula);

        /**
        Whether the formula will be hidden when the worksheet is protected.
        Returns @c tb_true if all cells in the range have this property set to true,
        @c tb_false if all cells in the range have this property set to false,
        and @c tb_default if the value of this property is not same for all cells in the range.

        [MSDN documentation for Range.FormulaHidden](https://learn.microsoft.com/en-us/office/vba/api/excel.range.formulahidden).
        */
        wxXlTribool GetFormulaHidden();

        /**
        True if the formula will be hidden when the worksheet is protected.

        [MSDN documentation for Range.FormulaHidden](https://learn.microsoft.com/en-us/office/vba/api/excel.range.formulahidden).
        */
        void SetFormulaHidden(const bool hidden);

        /**
        Returns the formula of a single cell, using A1-style references in the language of the user.

        For a range with more than one cell, Excel returns the formulas of all
        its cells, which this method cannot return, so it fails; use
        GetFormulaLocal(wxVariant&) instead.

        [MSDN documentation for Range.FormulaLocal](https://learn.microsoft.com/en-us/office/vba/api/excel.range.formulalocal).
        */
        wxString GetFormulaLocal();
        /**
        Returns the formulas of all cells in the range, using A1-style references in the language of the user.

        For a single cell, @a formulas contains a string. For more cells, it
        contains the formulas in the same way GetValue() returns the values,
        i.e. a list of strings ordered by columns, or a two-dimensional
        SAFEARRAY when @c wxOleConvertVariant_ReturnSafeArrays is set with
        SetConvertVariantFlags_(). Cells without a formula have their constant
        value as a string, empty cells an empty string.

        Returns @c true if the formulas were obtained.

        [MSDN documentation for Range.FormulaLocal](https://learn.microsoft.com/en-us/office/vba/api/excel.range.formulalocal).
        */
        bool GetFormulaLocal(wxVariant& formulas);
        /**
        Sets the formula of all cells in the range, using A1-style references in the language of the user.

        A string sets the same formula for all cells, with its relative references
        adjusted for each cell. To set a different formula for each cell, pass an
        array with the same dimensions as the range, e.g. a two-dimensional
        SAFEARRAY created with wxSafeArray, see the bulkdata sample; a list
        of strings can be used for a range with a single row.
        A null wxVariant clears the cells.

        [MSDN documentation for Range.FormulaLocal](https://learn.microsoft.com/en-us/office/vba/api/excel.range.formulalocal).
        */
        void SetFormulaLocal(const wxVariant& formulaLocal);

        /**
        Returns the formula of a single cell, using R1C1-style notation in the language of the macro.

        For a range with more than one cell, Excel returns the formulas of all
        its cells, which this method cannot return, so it fails; use
        GetFormulaR1C1(wxVariant&) instead.

        [MSDN documentation for Range.FormulaR1C1](https://learn.microsoft.com/en-us/office/vba/api/excel.range.formular1c1).
        */
        wxString GetFormulaR1C1();
        /**
        Returns the formulas of all cells in the range, using R1C1-style notation in the language of the macro.

        For a single cell, @a formulas contains a string. For more cells, it
        contains the formulas in the same way GetValue() returns the values,
        i.e. a list of strings ordered by columns, or a two-dimensional
        SAFEARRAY when @c wxOleConvertVariant_ReturnSafeArrays is set with
        SetConvertVariantFlags_(). Cells without a formula have their constant
        value as a string, empty cells an empty string.

        Returns @c true if the formulas were obtained.

        [MSDN documentation for Range.FormulaR1C1](https://learn.microsoft.com/en-us/office/vba/api/excel.range.formular1c1).
        */
        bool GetFormulaR1C1(wxVariant& formulas);
        /**
        Sets the formula of all cells in the range, using R1C1-style notation in the language of the macro.

        A string sets the same formula for all cells, with its relative references
        adjusted for each cell. To set a different formula for each cell, pass an
        array with the same dimensions as the range, e.g. a two-dimensional
        SAFEARRAY created with wxSafeArray, see the bulkdata sample; a list
        of strings can be used for a range with a single row.
        A null wxVariant clears the cells.

        [MSDN documentation for Range.FormulaR1C1](https://learn.microsoft.com/en-us/office/vba/api/excel.range.formular1c1).
        */
        void SetFormulaR1C1(const wxVariant& formulaR1C1);

        /**
        Returns the formula of a single cell, using R1C1-style notation in the language of the user.

        For a range with more than one cell, Excel returns the formulas of all
        its cells, which this method cannot return, so it fails; use
        GetFormulaR1C1Local(wxVariant&) instead.

        [MSDN documentation for Range.FormulaR1C1Local](https://learn.microsoft.com/en-us/office/vba/api/excel.range.formular1c1local).
        */
        wxString GetFormulaR1C1Local();
        /**
        Returns the formulas of all cells in the range, using R1C1-style notation in the language of the user.

        For a single cell, @a formulas contains a string. For more cells, it
        contains the formulas in the same way GetValue() returns the values,
        i.e. a list of strings ordered by columns, or a two-dimensional
        SAFEARRAY when @c wxOleConvertVariant_ReturnSafeArrays is set with
        SetConvertVariantFlags_(). Cells without a formula have their constant
        value as a string, empty cells an empty string.

        Returns @c true if the formulas were obtained.

        [MSDN documentation for Range.FormulaR1C1Local](https://learn.microsoft.com/en-us/office/vba/api/excel.range.formular1c1local).
        */
        bool GetFormulaR1C1Local(wxVariant& formulas);
        /**
        Sets the formula of all cells in the range, using R1C1-style notation in the language of the user.

        A string sets the same formula for all cells, with its relative references
        adjusted for each cell. To set a different formula for each cell, pass an
        array with the same dimensions as the range, e.g. a two-dimensional
        SAFEARRAY created with wxSafeArray, see the bulkdata sample; a list
        of strings can be used for a range with a single row.
        A null wxVariant clears the cells.

        [MSDN documentation for Range.FormulaR1C1Local](https://learn.microsoft.com/en-us/office/vba/api/excel.range.formular1c1local).
        */
        void SetFormulaR1C1Local(const wxVariant& formulaR1C1Local);

        /**
        True if the specified cell is part of an array formula.

        [MSDN documentation for Range.HasArray](https://learn.microsoft.com/en-us/office/vba/api/excel.range.hasarray).
        */
        wxXlTribool GetHasArray();
        /**
        Returns @c tb_true if all cells in the range contain formulas, @c tb_false if none of the cells in the range contains a formula,
        and @c tb_default otherwise.

        [MSDN documentation for Range.HasFormula](https://learn.microsoft.com/en-us/office/vba/api/excel.range.hasformula).
        */
        wxXlTribool GetHasFormula();
        /**
        The range height in points.

        [MSDN documentation for Range.Height](https://learn.microsoft.com/en-us/office/vba/api/excel.range.height).
        */
        double GetHeight();
        /**
        If true the rows or columns are hidden. The range must span entire row or column.

        [MSDN documentation for Range.Hidden](https://learn.microsoft.com/en-us/office/vba/api/excel.range.hidden).
        */
        bool GetHidden();
        /**
        If true the rows or columns are hidden. The range must span entire row or column.

        [MSDN documentation for Range.Hidden](https://learn.microsoft.com/en-us/office/vba/api/excel.range.hidden).
        */
        void SetHidden(bool hidden);

        /**
        Horizontal alignment for the range.

        [MSDN documentation for Range.HorizontalAlignment](https://learn.microsoft.com/en-us/office/vba/api/excel.range.horizontalalignment).
        */
        long GetHorizontalAlignment();
        /**
        Horizontal alignment for the range.

        [MSDN documentation for Range.HorizontalAlignment](https://learn.microsoft.com/en-us/office/vba/api/excel.range.horizontalalignment).
        */
        void SetHorizontalAlignment(const long alignment);

        /**
        Returns a Hyperlinks collection that represents the hyperlinks for the range.

        [MSDN documentation for Range.Hyperlinks](https://learn.microsoft.com/en-us/office/vba/api/excel.range.hyperlinks).
        */
        wxExcelHyperlinks GetHyperlinks();

        /**
        The identifying label for the specified cell when the page is saved as a Web page.

        [MSDN documentation for Range.ID](https://learn.microsoft.com/en-us/office/vba/api/excel.range.id).
        */
        wxString GetID();
        /**
        The identifying label for the specified cell when the page is saved as a Web page.

        [MSDN documentation for Range.ID](https://learn.microsoft.com/en-us/office/vba/api/excel.range.id).
        */
        void SetID(const wxString& ID);

        /**
        The indent level for the range. Can be an integer from 0 to 15.

        [MSDN documentation for Range.IndentLevel](https://learn.microsoft.com/en-us/office/vba/api/excel.range.indentlevel).
        */
        long GetIndentLevel();
        /**
        The indent level for the range. Can be an integer from 0 to 15.

        [MSDN documentation for Range.IndentLevel](https://learn.microsoft.com/en-us/office/vba/api/excel.range.indentlevel).
        */
        void SetIndentLevel(long indentLevel);

        /**
        Returns an Interior Represents the interior of the range.

        [MSDN documentation for Range.Interior](https://learn.microsoft.com/en-us/office/vba/api/excel.range.interior).
        */
        wxExcelInterior GetInterior();

        //@{
        /**
        Returns a Range Represents a range at an offset to the specified range.

        [MSDN documentation for Range.Item](https://learn.microsoft.com/en-us/office/vba/api/excel.range.item).
        */
        wxExcelRange GetItem(long rowIndex, long* columnIndex = NULL);
        wxExcelRange GetItem(long rowIndex, const wxString& columnIndex);
        wxExcelRange GetItem(const wxString& cell);
        //@}

        /**
        Returns the distance, in points, from the left edge of column A to the left edge of the range.

        [MSDN documentation for Range.Left](https://learn.microsoft.com/en-us/office/vba/api/excel.range.left).
        */
        double GetLeft();
        /**
        Returns the number of header rows for the specified range.

        [MSDN documentation for Range.ListHeaderRows](https://learn.microsoft.com/en-us/office/vba/api/excel.range.listheaderrows).
        */
        long GetListHeaderRows();

        /**
        Returns a ListObject for Range.

        [MSDN documentation for Range.ListObject](https://learn.microsoft.com/en-us/office/vba/api/excel.range.listobject).
        */
        wxExcelListObject GetListObject();

        /**
        Returns a constant that describes the part of the PivotTable report that contains the upper-left corner of the specified range. Can be one of the following XlLocationInTable. constants.

        [MSDN documentation for Range.LocationInTable](https://learn.microsoft.com/en-us/office/vba/api/excel.range.locationintable).
        */
        XlLocationInTable GetLocationInTable();

        /**
        Returns @c tb_true if all cells in the range are locked, @c tb_true if all cells in the range are unlocked,
        and @c tb_default is some cells are locked and some are not.

        [MSDN documentation for Range.Locked](https://learn.microsoft.com/en-us/office/vba/api/excel.range.locked).
        */
        wxXlTribool GetLocked();
        /**
        Sets a value that indicates if the object is locked.

        [MSDN documentation for Range.Locked](https://learn.microsoft.com/en-us/office/vba/api/excel.range.locked).
        */
        void SetLocked(bool locked);

        /**
        Returns the MDX name for the specified Range object. Since Excel 2007.

        [MSDN documentation for Range.MDX](https://learn.microsoft.com/en-us/office/vba/api/excel.range.mdx).
        */
        wxString GetMDX();
        /**
        Returns a Range Represents the merged range containing the specified cell. If the specified cell isn't in a merged range, this property returns the specified cell.

        [MSDN documentation for Range.MergeArea](https://learn.microsoft.com/en-us/office/vba/api/excel.range.mergearea).
        */
        wxExcelRange GetMergeArea();
        /**
        True if the range contains merged cells.

        [MSDN documentation for Range.MergeCells](https://learn.microsoft.com/en-us/office/vba/api/excel.range.mergecells).
        */
        bool GetMergeCells();

        /**
        Returns the name of the object.

        [MSDN documentation for Range.Name](https://learn.microsoft.com/en-us/office/vba/api/excel.range.name).
        */
        wxExcelName GetName();

        /**
        Sets the name of the object.

        [MSDN documentation for Range.Name](https://learn.microsoft.com/en-us/office/vba/api/excel.range.name).
        */
        void SetName(const wxString& name);

        /**
        Returns a range representing the next cell.

        [MSDN documentation for Range.Next](https://learn.microsoft.com/en-us/office/vba/api/excel.range.next).
        */
        wxExcelRange GetNext();

        /**
        Returns the format code for the range or an empty string if all the cells in range don't share the same format.

        The name is spelled Numberformat rather than NumberFormat as in Excel, because the Windows headers define GetNumberFormat as a macro, which would otherwise rename the method.

        [MSDN documentation for Range.NumberFormat](https://learn.microsoft.com/en-us/office/vba/api/excel.range.numberformat).
        */
        wxString GetNumberformat();

        /**
        Sets the format code for the range.

        The name is spelled Numberformat rather than NumberFormat as in Excel, because the Windows headers define GetNumberFormat as a macro, which would otherwise rename the method.

        [MSDN documentation for Range.NumberFormat](https://learn.microsoft.com/en-us/office/vba/api/excel.range.numberformat).
        */
        void SetNumberformat(const wxString& numberFormat);

        /**
        Returns the format code for the range as a string in the language of the user or an empty string if all the cells in range don't share the same format.

        The name is spelled Numberformat rather than NumberFormat as in Excel, because the Windows headers define GetNumberFormat as a macro, which would otherwise rename the method.

        [MSDN documentation for Range.NumberFormatLocal](https://learn.microsoft.com/en-us/office/vba/api/excel.range.numberformatlocal).
        */
        wxString GetNumberformatLocal();

        /**
        Sets the format code for the range as a string in the language of the user.

        The name is spelled Numberformat rather than NumberFormat as in Excel, because the Windows headers define GetNumberFormat as a macro, which would otherwise rename the method.

        [MSDN documentation for Range.NumberFormatLocal](https://learn.microsoft.com/en-us/office/vba/api/excel.range.numberformatlocal).
        */
        void SetNumberformatLocal(const wxString& numberFormatLocal);

        /**
        Returns a Range Represents a range that's offset from the specified range.

        [MSDN documentation for Range.Offset](https://learn.microsoft.com/en-us/office/vba/api/excel.range.offset).
        */
        wxExcelRange GetOffset(long rowOffset = 0, long columnOffset = 0);

        /**
        Returns the text orientation. See XlOrientation

        [MSDN documentation for Range.Orientation](https://learn.microsoft.com/en-us/office/vba/api/excel.range.orientation).
        */
        long GetOrientation();

        /**
        Sets the text orientation. See XlOrientation

        [MSDN documentation for Range.Orientation](https://learn.microsoft.com/en-us/office/vba/api/excel.range.orientation).
        */
        void SetOrientation(long orientation);

        /**
        Returns the current outline level of the specified row or column.

        [MSDN documentation for Range.OutlineLevel](https://learn.microsoft.com/en-us/office/vba/api/excel.range.outlinelevel).
        */
        long GetOutlineLevel();

        /**
        Sets the current outline level of the specified row or column.

        [MSDN documentation for Range.OutlineLevel](https://learn.microsoft.com/en-us/office/vba/api/excel.range.outlinelevel).
        */
        void SetOutlineLevel(long outlineLevel);

        /**
        Returns the location of a page break. Can be one of the following XlPageBreak constants: xlPageBreakAutomatic, xlPageBreakManual, or xlPageBreakNone.

        [MSDN documentation for Range.PageBreak](https://learn.microsoft.com/en-us/office/vba/api/excel.range.pagebreak).
        */
        XlPageBreak GetPageBreak();

        /**
        Sets the location of a page break. Can be one of the following XlPageBreak constants: xlPageBreakAutomatic, xlPageBreakManual, or xlPageBreakNone.

        [MSDN documentation for Range.PageBreak](https://learn.microsoft.com/en-us/office/vba/api/excel.range.pagebreak).
        */
        void SetPageBreak(XlPageBreak pageBreak);

        /**

        [MSDN documentation for Range.Precedents](https://learn.microsoft.com/en-us/office/vba/api/excel.range.precedents).
        */
        wxExcelRange GetPrecedents();

        /**
        Returns the prefix character for the cell.

        [MSDN documentation for Range.PrefixCharacter](https://learn.microsoft.com/en-us/office/vba/api/excel.range.prefixcharacter).
        */
        wxString GetPrefixCharacter();

        /**
        Returns a range represnting the next cell.

        [MSDN documentation for Range.Previous](https://learn.microsoft.com/en-us/office/vba/api/excel.range.previous).
        */
        wxExcelRange GetPrevious();

        /**
        Returns the reading order for the specified object. Can be one of the following constants: xlRTL (right-to-left), xlLTR (left-to-right), or xlContext.

        [MSDN documentation for Range.ReadingOrder](https://learn.microsoft.com/en-us/office/vba/api/excel.range.readingorder).
        */
        long GetReadingOrder();

        /**
        Resizes the specified range. Returns a Range Represents the resized range.

        [MSDN documentation for Range.Resize](https://learn.microsoft.com/en-us/office/vba/api/excel.range.resize).
        */
        wxExcelRange GetResize(long* rowSize= NULL, long* columnSize = NULL);

        /**
        Returns the number of the first row of the first area in the range.

        [MSDN documentation for Range.Row](https://learn.microsoft.com/en-us/office/vba/api/excel.range.row).
        */
        long GetRow();

        /**
        If all rows in the range have same height returns the height in points else returns -1. Returns -2 on error.

        [MSDN documentation for Range.RowHeight](https://learn.microsoft.com/en-us/office/vba/api/excel.range.rowheight).
        */
        double GetRowHeight();

        /**
        Sets the height of all the rows in the range specified, measured in points.

        [MSDN documentation for Range.RowHeight](https://learn.microsoft.com/en-us/office/vba/api/excel.range.rowheight).
        */
        void SetRowHeight(double rowHeight);

        /**
        True if the outline is expanded for the specified range (so that the detail of the column or row is visible). The specified range must be a single summary column or row in an outline. For the PivotItem object (or the Range object if the range is in a PivotTable report), this property is set to True if the item is showing detail.

        [MSDN documentation for Range.ShowDetail](https://learn.microsoft.com/en-us/office/vba/api/excel.range.showdetail).
        */
        bool GetShowDetail();

        /**
        Returns @c tb_true if the text shrinks to fit into the cell in all cells in the range.
        Returns @c tb_default if this property is not the same in all cells in the range.

        [MSDN documentation for Range.ShrinkToFit](https://learn.microsoft.com/en-us/office/vba/api/excel.range.shrinktofit).
        */
        wxXlTribool GetShrinkToFit();

        /**
        Set true if the text is to fit into the cell.

        [MSDN documentation for Range.ShrinkToFit](https://learn.microsoft.com/en-us/office/vba/api/excel.range.shrinktofit).
        */
        void SetShrinkToFit(bool shrinkToFit);

#if WXAUTOEXCEL_USE_CHARTS
        /**
        Returns a SparklineGroups object that represents an existing group of sparklines from the specified range. Since Excel 2010.

        [MSDN documentation for Range.SparklineGroups](https://learn.microsoft.com/en-us/office/vba/api/excel.range.sparklinegroups).
        */

        wxExcelSparklineGroups GetSparklineGroups();

#endif  // WXAUTOEXCEL_USE_CHARTS

        /**
        Returns the style of the range.

        [MSDN documentation for Range.Style](https://learn.microsoft.com/en-us/office/vba/api/excel.range.style).
        */
        wxExcelStyle GetStyle();

        /**
        Sets the style of the range.

        [MSDN documentation for Range.Style](https://learn.microsoft.com/en-us/office/vba/api/excel.range.style).
        */
        void SetStyle(wxExcelStyle style);

        /**
        True if the range is an outlining summary row or column. The range should be a row or a column.

        [MSDN documentation for Range.Summary](https://learn.microsoft.com/en-us/office/vba/api/excel.range.summary).
        */
        bool GetSummary();

        /**
        Returns the text value as seen by user in MS Excel. The range must comprise of a single cell, else an empty string is returned.

        [MSDN documentation for Range.Text](https://learn.microsoft.com/en-us/office/vba/api/excel.range.text).
        */
        wxString GetText();

        /**
        Returns the distance, in points, from the top edge of row 1 to the top edge of the range.

        [MSDN documentation for Range.Top](https://learn.microsoft.com/en-us/office/vba/api/excel.range.top).
        */
        double GetTop();

        /**
        Returns @c tb_true if the row height of the Range object equals the standard height of the sheet.
        Returns @c tb_default if the range contains more than one row and the rows are not all the same height.

        [MSDN documentation for Range.UseStandardHeight](https://learn.microsoft.com/en-us/office/vba/api/excel.range.usestandardheight).
        */
        wxXlTribool GetUseStandardHeight();

        /**
        Whether all cells in the range should use standard row height of the sheet.

        [MSDN documentation for Range.UseStandardHeight](https://learn.microsoft.com/en-us/office/vba/api/excel.range.usestandardheight).
        */
        void SetUseStandardHeight(bool useStandardHeight);

        /**
        Returns @c tb_true if the column width of the Range object equals the standard width of the sheet.
        Returns @c tb_default if the range contains more than one column and the columns are not all the same width.

        [MSDN documentation for Range.UseStandardWidth](https://learn.microsoft.com/en-us/office/vba/api/excel.range.usestandardwidth).
        */
        wxXlTribool GetUseStandardWidth();

        /**
        Whether all cells in the range should use standard column width of the sheet.

        [MSDN documentation for Range.UseStandardWidth](https://learn.microsoft.com/en-us/office/vba/api/excel.range.usestandardwidth).
        */
        void SetUseStandardWidth(bool useStandardWidth);

        /**

        Returns the Validation object that represents data validation for the specified range.

        [MSDN documentation for Range.Validation](https://learn.microsoft.com/en-us/office/vba/api/excel.range.validation).
        */
        wxExcelValidation GetValidation();

        /**
        Returns the value(s) of the specified range.

        [MSDN documentation for Range.Value](https://learn.microsoft.com/en-us/office/vba/api/excel.range.value).

        Due to wxWidgets internals, GetValue and SetValue work asymetrically by default. For example, let's have
        a range with address A1:C2, i.e. three columns and two rows. When you ask Excel for values,
        it puts them into a two-dimensional array. By default, wxVariant does not support multidimensional arrays,
        so you get values in wxVariant as a single list ordered by columns, in our example it will
        contain values of cells in this order: A1, A2, B1, B2, C1, C2.

        See bulkdata sample to see how efficiently transfer large numbers of values
        from/to MS Excel as a two-dimensional wxSafeArray.
        @see SetConvertVariantFlags_()

        */
        wxVariant GetValue();

        /**
        Same as calling GetValue();
        */
        operator wxVariant() { return GetValue(); }

        /**
        Sets the value for a Range. Passing a null wxVariant clears the cells.

        [MSDN documentation for Range.Value](https://learn.microsoft.com/en-us/office/vba/api/excel.range.value).

        Be aware that the automation LCID (see wxExcelObject::SetAutomationLCID_()) also affects how MS Excel may
        interpret the values passed as strings. E.g. "1,234" may be converted to an integer with value 1234 if the locale uses
        comma for the thousand separator and the decimal period but the same string may be converted to a float with value 1.234
        if the locale has decimal comma.

        Check bulkdata sample for an example how to efficiently transfer data to large two-dimensional Ranges.
        @see SetConvertVariantFlags_()
        */
        void SetValue(const wxVariant& value);

        /**
        Same as calling SetValue();
        */
        void operator=(const wxVariant& value) { SetValue(value); }

        /**
        Works almost like GetValue(), except that it returns DateTime and Currency as doubles and not their respective types.

        Also see GetValue() for some oddities of this function.

        [MSDN documentation for Range.Value2](https://learn.microsoft.com/en-us/office/vba/api/excel.range.value2).
        */
        wxVariant GetValue2();

        /**
        Sets the cell value.

        Also see SetValue() for some oddities of this function.

        [MSDN documentation for Range.Value2](https://learn.microsoft.com/en-us/office/vba/api/excel.range.value2).
        */
        void SetValue2(const wxVariant& value);

        /**
        Returns the vertical alignment of the specified object.

        [MSDN documentation for Range.VerticalAlignment](https://learn.microsoft.com/en-us/office/vba/api/excel.range.verticalalignment).
        */
        long GetVerticalAlignment();

        /**
        Sets the vertical alignment of the specified object.

        [MSDN documentation for Range.VerticalAlignment](https://learn.microsoft.com/en-us/office/vba/api/excel.range.verticalalignment).
        */
        void SetVerticalAlignment(long verticalAlignment);

        /**
        Returns the width, in units, of the range.

        [MSDN documentation for Range.Width](https://learn.microsoft.com/en-us/office/vba/api/excel.range.width).
        */
        double GetWidth();

        /**
        Returns the worksheet containing the specified range.

        [MSDN documentation for Range.Worksheet](https://learn.microsoft.com/en-us/office/vba/api/excel.range.worksheet).
        */
        wxExcelWorksheet GetWorksheet();

        /**
        Returns @c tb_true if all cells in the range wrap the text, @c tb_false if all cells in the range
        do not wrap the text, and @c tb_default if some cells wrap the text and some do not.

        [MSDN documentation for Range.WrapText](https://learn.microsoft.com/en-us/office/vba/api/excel.range.wraptext).
        */
        wxXlTribool GetWrapText();

        /**
        Set to true if Microsoft Excel should wrap the text in the object in all cells in the range.

        [MSDN documentation for Range.WrapText](https://learn.microsoft.com/en-us/office/vba/api/excel.range.wraptext).
        */
        void SetWrapText(bool wrapText);

        /**
        Returns "Range".
        */
        virtual wxString GetAutoExcelObjectName_() const { return wxS("Range"); }

        /**
        See wxAutomationObject::GetConvertVariantFlags()
        */
        long GetConvertVariantFlags_();

        /**
        This method, wrapping @c wxAutomationObject::SetConvertVariantFlags() can be very useful
        for efficient data transfer to/from Excel, using @c wxOleConvertVariant_ReturnSafeArrays
        as a value of @a flags, see bulkdata sample for an example.

        @see GetValue(), SetValue(), GetValue2(), SetValue2()

        */
        bool SetConvertVariantFlags_(long flags);


    private:
        // address can be either "Address" or "AddressLocal"
        wxString DoGetAddress(const wxString& address, wxXlTribool rowAbsolute, wxXlTribool columnAbsolute,
                              XlReferenceStyle* referenceStyle, wxXlTribool external, wxExcelRange* relativeTo);
};

} // namespace wxAutoExcel

#endif // _WXAUTOEXCEL_RANGE_H
