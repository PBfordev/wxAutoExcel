/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Modified by:
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
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

        [MSDN documentation for Range.Activate](http://msdn.microsoft.com/en-us/library/bb178814.aspx).
        */
        bool Activate();

        /**
        Adds a comment to the range.

        [MSDN documentation for Range.AddComment](http://msdn.microsoft.com/en-us/library/bb209547.aspx).
        */
        wxExcelComment AddComment(const wxString& text = wxEmptyString);

        /**
        Filters or copies data from a list based on a criteria range. If the initial selection is a single cell, that cell's current region is used.

        [MSDN documentation for Range.AdvancedFilter](http://msdn.microsoft.com/en-us/library/bb209640.aspx).
        */
        bool AdvancedFilter(enum XlFilterAction action,
                            wxExcelRange* criteriaRange = NULL, wxExcelRange* copyToRange = NULL,
                            wxXlTribool unique = wxDefaultXlTribool);

        /**
        Performs a writeback operation for all edited cells in a range based on an OLAP data source. Since Excel 2010.

        [MSDN documentation for Range.AllocateChanges](http://msdn.microsoft.com/en-us/library/office/ff838963%28v=office.14%29.aspx).
        */
        void AllocateChanges();

        /**
        Applies names to the cells in the specified range.

        [MSDN documentation for Range.ApplyNames](http://msdn.microsoft.com/en-us/library/bb209650.aspx).
        */
        void ApplyNames(wxArrayString* names = NULL,
                        wxXlTribool ignoreRelativeAbsolute =  NULL, wxXlTribool useRowColumnNames = wxDefaultXlTribool,
                        wxXlTribool omitColumn = wxDefaultXlTribool, wxXlTribool omitRow = wxDefaultXlTribool,
                        XlApplyNamesOrder* order = NULL, wxXlTribool appendLast = wxDefaultXlTribool);

        /**
        Applies outlining styles to the specified range.

        [MSDN documentation for Range.ApplyOutlineStyles](http://msdn.microsoft.com/en-us/library/bb209656.aspx).
        */
        void ApplyOutlineStyles();

        /**
        Returns an AutoComplete match from the list. If there's no AutoComplete match or if more than one entry in the list matches the string to complete, this method returns an empty string.

        [MSDN documentation for Range.AutoComplete](http://msdn.microsoft.com/en-us/library/bb209667.aspx).
        */
        wxString AutoComplete(const wxString& str);

        /**
        Performs an autofill on the cells in the specified range.

        [MSDN documentation for Range.AutoFill](http://msdn.microsoft.com/en-us/library/bb209671.aspx).
        */
        bool AutoFill(wxExcelRange destination, XlAutoFillType* type = NULL);

        /**
        Filters a list using the AutoFilter.

        [MSDN documentation for Range.AutoFilter](http://msdn.microsoft.com/en-us/library/bb242013.aspx).
        */
        bool AutoFilter(long* field = NULL, const wxString& criteria1 = wxEmptyString,
                        XlAutoFilterOperator* oper = NULL, const wxString& criteria2 = wxEmptyString,
                        wxXlTribool visibleDropDown = wxDefaultXlTribool);

        /**
        Changes the width of the columns in the range or the height of the rows in the range to achieve the best fit.

        [MSDN documentation for Range.AutoFit](http://msdn.microsoft.com/en-us/library/bb209676.aspx).
        */
        bool AutoFit();

        /**
        Automatically creates an outline for the specified range. If the range is a single cell, Microsoft Excel creates an outline for the entire sheet. The new outline replaces any existing outline.

        [MSDN documentation for Range.AutoOutline](http://msdn.microsoft.com/en-us/library/bb209687.aspx).
        */
        void AutoOutline();

        /**
        Adds a border to a range and sets the Color, LineStyle, and Weight properties for the new border.

        [MSDN documentation for Range.BorderAround](http://msdn.microsoft.com/en-us/library/bb209714.aspx).
        */
        bool BorderAround(XlLineStyle* lineStyle = NULL, XlBorderWeight* weight = NULL,
                          long* colorIndex = NULL, const wxColour* color = NULL);

        /**
        Calculates the range.

        [MSDN documentation for Range.Calculate](http://msdn.microsoft.com/en-us/library/bb178817.aspx).
        */
        void Calculate();

        /**
        Calculates the range the same way Excel 2000 and earlier did.

        [MSDN documentation for Range.CalculateRowMajorOrder](http://msdn.microsoft.com/en-us/library/bb226058.aspx).
        */
        void CalculateRowMajorOrder();

        /**
        Checks the spelling of an object.

        [MSDN documentation for Range.CheckSpelling](http://msdn.microsoft.com/en-us/library/bb178820.aspx).
        */
        void CheckSpelling(const wxString& customDictionary = wxEmptyString, wxXlTribool ignoreUpperCase = wxDefaultXlTribool,
                           wxXlTribool alwaysSuggest = wxDefaultXlTribool, MsoLanguageID* spellLang = NULL);

        /**
        Clears the entire object.

        [MSDN documentation for Range.Clear](http://msdn.microsoft.com/en-us/library/bb178823.aspx).
        */
        bool Clear();

        /**
        Clears all cell comments from the specified range.

        [MSDN documentation for Range.ClearComments](http://msdn.microsoft.com/en-us/library/bb223261.aspx).
        */
        bool ClearComments();

        /**
        Clears the formulas from the range.

        [MSDN documentation for Range.ClearContents](http://msdn.microsoft.com/en-us/library/bb178828.aspx).
        */
        bool ClearContents();

        /**
        Clears the formatting of the object.

        [MSDN documentation for Range.ClearFormats](http://msdn.microsoft.com/en-us/library/bb178831.aspx).
        */
        bool ClearFormats();

         /**
        Removes all hyperlinks from the specified range. Since Excel 2010.

        [MSDN documentation for Range.ClearHyperlinks](http://msdn.microsoft.com/en-us/library/office/ff194741%28v=office.14%29.aspx).
        */
        bool ClearHyperlinks();

        /**
        Clears notes and sound notes from all the cells in the specified range.

        [MSDN documentation for Range.ClearNotes](http://msdn.microsoft.com/en-us/library/bb223264.aspx).
        */
        bool ClearNotes();

        /**
        Clears the outline for the specified range.

        [MSDN documentation for Range.ClearOutline](http://msdn.microsoft.com/en-us/library/bb223266.aspx).
        */
        bool ClearOutline();

        /**
        Returns a Range Represents all the cells whose contents are different from the comparison cell in each column.

        [MSDN documentation for Range.ColumnDifferences](http://msdn.microsoft.com/en-us/library/bb223272.aspx).
        */
        wxExcelRange ColumnDifferences(wxExcelRange comparison);

        /**
        Copies the range to the specified range or to the Clipboard.

        [MSDN documentation for Range.Copy](http://msdn.microsoft.com/en-us/library/bb178833.aspx).
        */
        bool Copy(const wxExcelRange* destination = NULL);

        /**
        Copies the selected object to the Clipboard as a picture.

        [MSDN documentation for Range.CopyPicture](http://msdn.microsoft.com/en-us/library/bb178836.aspx).
        */
        bool CopyPicture(XlPictureAppearance* appearance = NULL, XlCopyPictureFormat* format = NULL);

        /**
        Creates names in the specified range, based on text labels in the sheet.

        [MSDN documentation for Range.CreateNames](http://msdn.microsoft.com/en-us/library/bb223293.aspx).
        */
        bool CreateNames(wxXlTribool top = wxDefaultXlTribool, wxXlTribool left = wxDefaultXlTribool, wxXlTribool bottom = wxDefaultXlTribool, wxXlTribool right = wxDefaultXlTribool);

        /**
        Cuts the object to the Clipboard or pastes it into a specified destination.

        [MSDN documentation for Range.Cut](http://msdn.microsoft.com/en-us/library/bb178840.aspx).
        */
        bool Cut(const wxExcelRange* destination = NULL);

        /**
        Creates a data series in the specified range.

        [MSDN documentation for Range.DataSeries](http://msdn.microsoft.com/en-us/library/bb223314.aspx).
        */
        bool DataSeries(XlRowCol* rowCol = NULL, XlDataSeriesType* type = NULL, XlDataSeriesDate* date = NULL,
                                long* step = NULL, const wxVariant& stop = wxNullVariant, wxXlTribool trend = wxDefaultXlTribool);

        /**
        Deletes the object.

        [MSDN documentation for Range.Delete](http://msdn.microsoft.com/en-us/library/bb178843.aspx).
        */
        bool Delete(XlDeleteShiftDirection* shift = NULL);


        /**
        Designates a range to be recalculated when the next recalculation occurs.

        [MSDN documentation for Range.Dirty](http://msdn.microsoft.com/en-us/library/bb223347.aspx).
        */
        void Dirty();

        /**
        Discards all changes in the edited cells of the range. Since Excel 2010.

        [MSDN documentation for Range.DiscardChanges](http://msdn.microsoft.com/en-us/library/office/ff837815%28v=office.14%29.aspx).
        */
        void DiscardChanges();


        //@{
        /**
        Exports to an XPS or PDF file.

        [MSDN documentation for Range.ExportAsFixedFormat](http://msdn.microsoft.com/en-us/library/bb238866.aspx).
        */
        void ExportAsFixedFormat(XlFixedFormatType type, const wxString& fileName = wxEmptyString,
                                 XlFixedFormatQuality* quality = NULL, wxXlTribool includeDocProperties = wxDefaultXlTribool,
                                 wxXlTribool ignorePrintAreas = wxDefaultXlTribool,
                                 long* from = NULL, long* to = NULL, wxXlTribool openAfterPublish = wxDefaultXlTribool);

        void ExportAsFixedFormat(XlFixedFormatType type, const wxVariantVector& optionalArgs);
        //@}

        /**
        Fills down from the top cell or cells in the specified range to the bottom of the range. The contents and formatting of the cell or cells in the top row of a range are copied into the rest of the rows in the range.

        [MSDN documentation for Range.FillDown](http://msdn.microsoft.com/en-us/library/bb209838.aspx).
        */
        bool FillDown();

        /**
        Fills left from the rightmost cell or cells in the specified range. The contents and formatting of the cell or cells in the rightmost column of a range are copied into the rest of the columns in the range.

        [MSDN documentation for Range.FillLeft](http://msdn.microsoft.com/en-us/library/bb209843.aspx).
        */
        bool FillLeft();

        /**
        Fills right from the leftmost cell or cells in the specified range. The contents and formatting of the cell or cells in the leftmost column of a range are copied into the rest of the columns in the range.

        [MSDN documentation for Range.FillRight](http://msdn.microsoft.com/en-us/library/bb209846.aspx).
        */
        bool FillRight();

        /**
        Fills up from the bottom cell or cells in the specified range to the top of the range. The contents and formatting of the cell or cells in the bottom row of a range are copied into the rest of the rows in the range.

        [MSDN documentation for Range.FillUp](http://msdn.microsoft.com/en-us/library/bb209849.aspx).
        */
        bool FillUp();


        /**
        Starts the Function Wizard for the upper-left cell of the range.

        [MSDN documentation for Range.FunctionWizard](http://msdn.microsoft.com/en-us/library/bb209876.aspx).
        */
        void FunctionWizard();

        /**
        Inserts a cell or a range of cells into the worksheet or macro sheet and shifts other cells away to make space.

        [MSDN documentation for Range.Insert](http://msdn.microsoft.com/en-us/library/bb178850.aspx).
        */
        bool Insert(XlInsertShiftDirection* shift);

        /**
        Adds an indent to the specified range.

        [MSDN documentation for Range.InsertIndent](http://msdn.microsoft.com/en-us/library/bb209953.aspx).
        */
        void InsertIndent(long insertAmount);

        /**
        Rearranges the text in a range so that it fills the range evenly.

        [MSDN documentation for Range.Justify](http://msdn.microsoft.com/en-us/library/bb209965.aspx).
        */
        bool Justify();

        /**
        Pastes a list of all nonhidden names onto the worksheet, beginning with the first cell in the range.

        [MSDN documentation for Range.ListNames](http://msdn.microsoft.com/en-us/library/bb209982.aspx).
        */
        bool ListNames();

        /**
        Creates a merged cell from the specified Range object.

        [MSDN documentation for Range.Merge](http://msdn.microsoft.com/en-us/library/bb178851.aspx).
        */
        void Merge(wxXlTribool across = wxDefaultXlTribool);

        /**
        Navigates a tracer arrow for the specified range to the precedent, dependent, or error-causing cell or cells. Selects the precedent, dependent, or error cells and returns a Range Represents the new selection. This method causes an error if it's applied to a cell without visible tracer arrows.

        [MSDN documentation for Range.NavigateArrow](http://msdn.microsoft.com/en-us/library/bb223485.aspx).
        */
        bool NavigateArrow(wxXlTribool towardPrecedent = wxDefaultXlTribool, long* arrowNumber = NULL, long* linkNumber = NULL);

        /**
        Pastes a Range from the Clipboard into the specified range.

        [MSDN documentation for Range.PasteSpecial](http://msdn.microsoft.com/en-us/library/bb178854.aspx).
        */
        bool PasteSpecial(XlPasteType* paste = NULL, XlPasteSpecialOperation* operation = NULL,
                          wxXlTribool skipBlanks = wxDefaultXlTribool, wxXlTribool transpose = wxDefaultXlTribool);

        //@{
        /**
        Prints the object.

        [MSDN documentation for Range.PrintOut](http://msdn.microsoft.com/en-us/library/bb178858.aspx).
        */
        bool PrintOut(long* from = NULL, long* to = NULL, long* copies = NULL, wxXlTribool preview = wxDefaultXlTribool,
                      const wxString& activePrinter = wxEmptyString, wxXlTribool printToFile = wxDefaultXlTribool,
                      wxXlTribool collate = wxDefaultXlTribool, const wxString& prToFileName= wxEmptyString);
        bool PrintOut(const wxVariantVector& args);
        //@}

        /**
        Shows a preview of the object as it would look when printed.

        [MSDN documentation for Range.PrintPreview](http://msdn.microsoft.com/en-us/library/bb178861.aspx).
        */
        bool PrintPreview(wxXlTribool enableChanges = wxDefaultXlTribool);

        /**
        Removes duplicate values from a range of values.

        [MSDN documentation for Range.RemoveDuplicates](http://msdn.microsoft.com/en-us/library/bb238869.aspx).
        */
        void RemoveDuplicates(const wxArrayLong& columns, XlYesNoGuess* header = NULL);

        /**
        Removes subtotals from a list.

        [MSDN documentation for Range.RemoveSubtotal](http://msdn.microsoft.com/en-us/library/bb223597.aspx).
        */
        void RemoveSubtotal();

        /**
        Returns a Range Represents all the cells whose contents are different from those of the comparison cell in each row.

        [MSDN documentation for Range.RowDifferences](http://msdn.microsoft.com/en-us/library/bb177981.aspx).
        */
        wxExcelRange RowDifferences(wxExcelRange comparison);

        /**
        Selects the object.

        [MSDN documentation for Range.Select](http://msdn.microsoft.com/en-us/library/bb238205.aspx).
        */
        bool Select();


        /**
        Scrolls through the contents of the active window to move the range into view. The range must consist of a single cell in the active document.

        [MSDN documentation for Range.Show](http://msdn.microsoft.com/en-us/library/bb238213.aspx).
        */
        bool Show();

        /**
        Draws tracer arrows to the direct dependents of the range.

        [MSDN documentation for Range.ShowDependents](http://msdn.microsoft.com/en-us/library/bb178120.aspx).
        */
        bool ShowDependents(wxXlTribool remove = wxDefaultXlTribool);

        /**
        Draws tracer arrows through the precedents tree to the cell that's the source of the error, and returns the range that contains that cell.

        [MSDN documentation for Range.ShowErrors](http://msdn.microsoft.com/en-us/library/bb178123.aspx).
        */
        wxExcelRange ShowErrors();

        /**
        Draws tracer arrows to the direct precedents of the range.

        [MSDN documentation for Range.ShowPrecedents](http://msdn.microsoft.com/en-us/library/bb178137.aspx).
        */
        bool ShowPrecedents(wxXlTribool remove = wxDefaultXlTribool);

        /**
        Returns a Range Represents all the cells that match the specified type and value.

        [MSDN documentation for Range.SpecialCells](http://msdn.microsoft.com/en-us/library/bb178148.aspx).
        */
        wxExcelRange SpecialCells(XlCellType type, const wxVariant& value = wxNullVariant);


        /**
        Creates subtotals for the range (or the current region, if the range is a single cell).

        [MSDN documentation for Range.Subtotal](http://msdn.microsoft.com/en-us/library/bb238221.aspx).
        */
        bool Subtotal(XlConsolidationFunction groupBy, XlConsolidationFunction function,
                      wxArrayLong& totalList, wxXlTribool replace = wxDefaultXlTribool, wxXlTribool pageBreaks = wxDefaultXlTribool,
                      XlSummaryRow* summaryBelowData = NULL);

        /**
        Creates a data table based on input values and formulas that you define on a worksheet.

        [MSDN documentation for Range.Table](http://msdn.microsoft.com/en-us/library/bb178155.aspx).
        */
        void Table(wxExcelRange* rowInput = NULL, wxExcelRange* columnInput = NULL);

        /**
        Promotes a range in an outline (that is, decreases its outline level). The specified range must be a row or column, or a range of rows or columns. If the range is in a PivotTable report, this method ungroups the items contained in the range.

        [MSDN documentation for Range.Ungroup](http://msdn.microsoft.com/en-us/library/bb238225.aspx).
        */
        bool Ungroup();

        /**
        Separates a merged area into individual cells.

        [MSDN documentation for Range.UnMerge](http://msdn.microsoft.com/en-us/library/bb210014.aspx).
        */
        bool UnMerge();

        // ***** PROPERTIES *****

        /**
        True if text is automatically indented when the text alignment in a cell is set to equal distribution (either horizontally or vertically.);
        [MSDN documentation for Range.AddIndent](http://msdn.microsoft.com/en-us/library/bb213508.aspx).
        */
        bool GetAddIndent();
        /**
        True if text is automatically indented when the text alignment in a cell is set to equal distribution (either horizontally or vertically.);
        [MSDN documentation for Range.AddIndent](http://msdn.microsoft.com/en-us/library/bb213508.aspx).
        */
        void SetAddIndent(bool addIndent);

        /**
        Returns a value that represents the range reference in the language of the macro.

        [MSDN documentation for Range.Address](http://msdn.microsoft.com/en-us/library/bb213510.aspx).
        */
        wxString GetAddress(wxXlTribool rowAbsolute = wxDefaultXlTribool, wxXlTribool columnAbsolute = wxDefaultXlTribool,
                            XlReferenceStyle* referenceStyle = NULL, wxXlTribool external = wxDefaultXlTribool,
                            wxExcelRange* relativeTo = NULL);
        /**
        Returns the range reference for the specified range in the language of the user.

        [MSDN documentation for Range.AddressLocal](http://msdn.microsoft.com/en-us/library/bb220823.aspx).
        */
        wxString GetAddressLocal(wxXlTribool rowAbsolute = wxDefaultXlTribool, wxXlTribool columnAbsolute = wxDefaultXlTribool,
                                XlReferenceStyle* referenceStyle = NULL, wxXlTribool external = wxDefaultXlTribool,
                                wxExcelRange* relativeTo = NULL);
        /**
        True if the range can be edited on a protected worksheet.

        [MSDN documentation for Range.AllowEdit](http://msdn.microsoft.com/en-us/library/bb213511.aspx).
        */
        bool GetAllowEdit();

        /**
        Returns an Areas collection that represents all the ranges in a multiple-area selection.

        [MSDN documentation for Range.Areas](http://msdn.microsoft.com/en-us/library/bb220845.aspx).
        */
        wxExcelAreas GetAreas();

        /**
        Returns a Borders collection that represents the borders of a style or a range of cells (including a range defined as part of a conditional format).

        [MSDN documentation for Range.Borders](http://msdn.microsoft.com/en-us/library/bb213512.aspx).
        */
        wxExcelBorders GetBorders();

        /**
        Returns a Characters Represents a range of characters within the object text. You can use the Characters object to format characters within a text string.

        [MSDN documentation for Range.Characters](http://msdn.microsoft.com/en-us/library/bb213514.aspx).
        */
        wxExcelCharacters GetCharacters(long start = 1, long* length = NULL);

        /**
        Returns the number of the first column in the first area in the specified range.

        [MSDN documentation for Range.Column](http://msdn.microsoft.com/en-us/library/bb177363.aspx).
        */
        long GetColumn();

        /**
        If all columns in the range have same width returns the width in points else returns -1. Returns -2 on error.

        [MSDN documentation for Range.ColumnWidth](http://msdn.microsoft.com/en-us/library/bb177374.aspx).
        */
        double GetColumnWidth();

         /**
        Sets the column width for all the columns in the range.

        [MSDN documentation for Range.ColumnWidth](http://msdn.microsoft.com/en-us/library/bb177374.aspx).
        */
        void SetColumnWidth(double colWidth);

        /**
        Returns a Comment Represents the comment associated with the cell in the upper-left corner of the range.

        [MSDN documentation for Range.Comment](http://msdn.microsoft.com/en-us/library/bb213516.aspx).
        */
        wxExcelComment GetComment();

        /**
        Returns a value that represents the number of cells in the range.

        [MSDN documentation for Range.Count](http://msdn.microsoft.com/en-us/library/bb213517.aspx).
        */
        long GetCount();

        /**
        Returns a value that represents the number of cells in the range. Since Excel 2007.

        [MSDN documentation for Range.CountLarge](http://msdn.microsoft.com/en-us/library/bb242638.aspx).
        */
        wxLongLong GetCountLarge();

        /**
        If the specified cell is part of an array, returns a Range Represents the entire array.

        [MSDN documentation for Range.CurrentArray](http://msdn.microsoft.com/en-us/library/bb177413.aspx).
        */
        wxExcelRange GetCurrentArray();

        /**
        Returns a Range Represents the current region. The current region is a range bounded by any combination of blank rows and blank columns.

        [MSDN documentation for Range.CurrentRegion](http://msdn.microsoft.com/en-us/library/bb177419.aspx).
        */
        wxExcelRange GetCurrentRegion();

        /**
        Returns a Range Represents the range containing all the dependents of a cell. This can be a multiple selection (a union of Range objects) if there's more than one dependent.

        [MSDN documentation for Range.Dependents](http://msdn.microsoft.com/en-us/library/bb177464.aspx).
        */
        wxExcelRange GetDependents();

        /**
        Returns a Range Represents the range containing all the direct dependents of a cell. This can be a multiple selection (a union of Range objects) if there's more than one dependent.

        [MSDN documentation for Range.DirectDependents](http://msdn.microsoft.com/en-us/library/bb177473.aspx).
        */
        wxExcelRange GetDirectDependents();

        /**
        Returns a Range Represents the range containing all the direct precedents of a cell. This can be a multiple selection (a union of Range objects) if there's more than one precedent.

        [MSDN documentation for Range.DirectPrecedents](http://msdn.microsoft.com/en-us/library/bb177476.aspx).
        */
        wxExcelRange GetDirectPrecedents();

        /**
        Returns a Range Represents the cell at the end of the region that contains the source range. Equivalent to pressing END+UP ARROW, END+DOWN ARROW, END+LEFT ARROW, or END+RIGHT ARROW.

        [MSDN documentation for Range.End](http://msdn.microsoft.com/en-us/library/bb221181.aspx).
        */
        wxExcelRange GetEnd(XlDirection direction);

        /**
        Returns a Range Represents the entire column (or columns) that contains the specified range.

        [MSDN documentation for Range.EntireColumn](http://msdn.microsoft.com/en-us/library/bb208462.aspx).
        */
        wxExcelRange GetEntireColumn();

        /**
        Returns a Range Represents the entire row (or rows) that contains the specified range.

        [MSDN documentation for Range.EntireRow](http://msdn.microsoft.com/en-us/library/bb208465.aspx).
        */
        wxExcelRange GetEntireRow();


        /**
        Allows the user to to access error checking options.

        [MSDN documentation for Range.Errors](http://msdn.microsoft.com/en-us/library/bb208478.aspx).
        */
        wxExcelErrors GetErrors();

        /**
        Returns a Font Represents the font of the specified object.

        [MSDN documentation for Range.Font](http://msdn.microsoft.com/en-us/library/bb213520.aspx).
        */
        wxExcelFont GetFont();

#if WXAUTOEXCEL_USE_CONDFORMAT
        /**
        Returns a FormatConditions collection that represents all the conditional formats for the specified range.

        [MSDN documentation for Range.FormatConditions](http://msdn.microsoft.com/en-us/library/bb208527.aspx).
        */

        wxExcelFormatConditions GetFormatConditions();

#endif  // WXAUTOEXCEL_USE_CONDFORMAT



        /**
        Returns a value that represents the object's formula in A1-style notation and in the language of the macro.

        [MSDN documentation for Range.Formula](http://msdn.microsoft.com/en-us/library/bb213521.aspx).
        */
        wxString GetFormula();
        /**
        Sets a value that represents the object's formula in A1-style notation and in the language of the macro.

        [MSDN documentation for Range.Formula](http://msdn.microsoft.com/en-us/library/bb213521.aspx).
        */
        void SetFormula(const wxString& formula);

        /**
        Returns the array formula of a range. Returns (or can be set to) a single formula or a Visual Basic array. If the specified range doesn't contain an array formula, this property returns null.

        [MSDN documentation for Range.FormulaArray](http://msdn.microsoft.com/en-us/library/bb208529.aspx).
        */
        wxString GetFormulaArray();
        /**
        Sets the array formula of a range.

        [MSDN documentation for Range.FormulaArray](http://msdn.microsoft.com/en-us/library/bb208529.aspx).
        */
        void SetFormulaArray(const wxString& formula);

        /**
        Whether the formula will be hidden when the worksheet is protected.
        Returns @c tb_true if all cells in the range have this property set to true,
        @c tb_false if all cells in the range have this property set to false,
        and @c tb_default if the value of this property is not same for all cells in the range.

        [MSDN documentation for Range.FormulaHidden](http://msdn.microsoft.com/en-us/library/bb213523.aspx).
        */
        wxXlTribool GetFormulaHidden();

        /**
        True if the formula will be hidden when the worksheet is protected.

        [MSDN documentation for Range.FormulaHidden](http://msdn.microsoft.com/en-us/library/bb213523.aspx).
        */
        void SetFormulaHidden(const bool hidden);

        /**
        Returns the formula for the object, using A1-style references in the language of the user.

        [MSDN documentation for Range.FormulaLocal](http://msdn.microsoft.com/en-us/library/bb213525.aspx).
        */
        wxString GetFormulaLocal();
        /**
        Sets the formula for the object, using A1-style references in the language of the user.

        [MSDN documentation for Range.FormulaLocal](http://msdn.microsoft.com/en-us/library/bb213525.aspx).
        */
        void SetFormulaLocal(const wxString& formulaLocal);

        /**
        Returns the formula for the object, using R1C1-style notation in the language of the macro.

        [MSDN documentation for Range.FormulaR1C1](http://msdn.microsoft.com/en-us/library/bb213527.aspx).
        */
        wxString GetFormulaR1C1();
        /**
        Sets the formula for the object, using R1C1-style notation in the language of the macro.

        [MSDN documentation for Range.FormulaR1C1](http://msdn.microsoft.com/en-us/library/bb213527.aspx).
        */
        void SetFormulaR1C1(const wxString& formulaR1C1);

        /**
        Returns the formula for the object, using R1C1-style notation in the language of the user.

        [MSDN documentation for Range.FormulaR1C1Local](http://msdn.microsoft.com/en-us/library/bb213529.aspx).
        */
        wxString GetFormulaR1C1Local();
        /**
        Sets the formula for the object, using R1C1-style notation in the language of the user.

        [MSDN documentation for Range.FormulaR1C1Local](http://msdn.microsoft.com/en-us/library/bb213529.aspx).
        */
        void SetFormulaR1C1Local(const wxString& formulaR1C1Local);

        /**
        True if the specified cell is part of an array formula.

        [MSDN documentation for Range.HasArray](http://msdn.microsoft.com/en-us/library/bb208593.aspx).
        */
        wxXlTribool GetHasArray();
        /**
        Returns @c tb_true if all cells in the range contain formulas, @c tb_false if none of the cells in the range contains a formula,
        and @c tb_default otherwise.

        [MSDN documentation for Range.HasFormula](http://msdn.microsoft.com/en-us/library/bb208626.aspx).
        */
        wxXlTribool GetHasFormula();
        /**
        The range height in points.

        [MSDN documentation for Range.Height](http://msdn.microsoft.com/en-us/library/bb213531.aspx).
        */
        double GetHeight();
        /**
        The range height in points.

        [MSDN documentation for Range.Height](http://msdn.microsoft.com/en-us/library/bb213531.aspx).
        */
        void SetHeight(double height);

        /**
        If true the rows or columns are hidden. The range must span entire row or column.

        [MSDN documentation for Range.Hidden](http://msdn.microsoft.com/en-us/library/bb213533.aspx).
        */
        bool GetHidden();
        /**
        If true the rows or columns are hidden. The range must span entire row or column.

        [MSDN documentation for Range.Hidden](http://msdn.microsoft.com/en-us/library/bb213533.aspx).
        */
        void SetHidden(bool hidden);

        /**
        Horizontal alignment for the range.

        [MSDN documentation for Range.HorizontalAlignment](http://msdn.microsoft.com/en-us/library/bb213535.aspx).
        */
        long GetHorizontalAlignment();
        /**
        Horizontal alignment for the range.

        [MSDN documentation for Range.HorizontalAlignment](http://msdn.microsoft.com/en-us/library/bb213535.aspx).
        */
        void SetHorizontalAlignment(const long alignment);

        /**
        Returns a Hyperlinks collection that represents the hyperlinks for the range.

        [MSDN documentation for Range.Hyperlinks](http://msdn.microsoft.com/en-us/library/bb213536.aspx).
        */
        wxExcelHyperlinks GetHyperlinks();

        /**
        The identifying label for the specified cell when the page is saved as a Web page.

        [MSDN documentation for Range.ID](http://msdn.microsoft.com/en-us/library/bb213537.aspx).
        */
        wxString GetID();
        /**
        The identifying label for the specified cell when the page is saved as a Web page.

        [MSDN documentation for Range.ID](http://msdn.microsoft.com/en-us/library/bb213537.aspx).
        */
        void SetID(const wxString& ID);

        /**
        The indent level for the range. Can be an integer from 0 to 15.

        [MSDN documentation for Range.IndentLevel](http://msdn.microsoft.com/en-us/library/bb213540.aspx).
        */
        long GetIndentLevel();
        /**
        The indent level for the range. Can be an integer from 0 to 15.

        [MSDN documentation for Range.IndentLevel](http://msdn.microsoft.com/en-us/library/bb213540.aspx).
        */
        void SetIndentLevel(long indentLevel);

        /**
        Returns an Interior Represents the interior of the range.

        [MSDN documentation for Range.Interior](http://msdn.microsoft.com/en-us/library/bb213542.aspx).
        */
        wxExcelInterior GetInterior();
        
        //@{
        /**
        Returns a Range Represents a range at an offset to the specified range.

        [MSDN documentation for Range.Item](http://msdn.microsoft.com/en-us/library/bb213544.aspx).
        */
        wxExcelRange GetItem(long rowIndex, long* columnIndex = NULL);
        wxExcelRange GetItem(long rowIndex, const wxString& columnIndex);
        wxExcelRange GetItem(const wxString& cell);
        //@}

        /**
        Returns the distance, in points, from the left edge of column A to the left edge of the range.

        [MSDN documentation for Range.Left](http://msdn.microsoft.com/en-us/library/bb213546.aspx).
        */
        double GetLeft();
        /**
        Returns the number of header rows for the specified range.

        [MSDN documentation for Range.ListHeaderRows](http://msdn.microsoft.com/en-us/library/bb177918.aspx).
        */
        long GetListHeaderRows();

        /**
        Returns a constant that describes the part of the PivotTable report that contains the upper-left corner of the specified range. Can be one of the following XlLocationInTable. constants.

        [MSDN documentation for Range.LocationInTable](http://msdn.microsoft.com/en-us/library/bb208698.aspx).
        */
        XlLocationInTable GetLocationInTable();
        
        /**
        Returns @c tb_true if all cells in the range are locked, @c tb_true if all cells in the range are unlocked,
        and @c tb_default is some cells are locked and some are not.

        [MSDN documentation for Range.Locked](http://msdn.microsoft.com/en-us/library/bb213550.aspx).
        */
        wxXlTribool GetLocked();
        /**
        Sets a value that indicates if the object is locked.

        [MSDN documentation for Range.Locked](http://msdn.microsoft.com/en-us/library/bb213550.aspx).
        */
        void SetLocked(bool locked);

        /**
        Returns the MDX name for the specified Range object. Since Excel 2007.

        [MSDN documentation for Range.MDX](http://msdn.microsoft.com/en-us/library/bb213552.aspx).
        */
        wxString GetMDX();
        /**
        Returns a Range Represents the merged range containing the specified cell. If the specified cell isn't in a merged range, this property returns the specified cell.

        [MSDN documentation for Range.MergeArea](http://msdn.microsoft.com/en-us/library/bb208756.aspx).
        */
        wxExcelRange GetMergeArea();
        /**
        True if the range contains merged cells.

        [MSDN documentation for Range.MergeCells](http://msdn.microsoft.com/en-us/library/bb213554.aspx).
        */
        bool GetMergeCells();

        /**
        Returns the name of the object.

        [MSDN documentation for Range.Name](http://msdn.microsoft.com/en-us/library/bb213557.aspx).
        */
        wxExcelName GetName();

        /**
        Sets the name of the object.

        [MSDN documentation for Range.Name](http://msdn.microsoft.com/en-us/library/bb213557.aspx).
        */
        void SetName(const wxString& name);

        /**
        Returns a range representing the next cell.

        [MSDN documentation for Range.Next](http://msdn.microsoft.com/en-us/library/bb213669.aspx).
        */
        wxExcelRange GetNext();
        
        /**
        Returns the format code for the range or an empty string if all the cells in range don't share the same format

        [MSDN documentation for Range.NumberFormat](http://msdn.microsoft.com/en-us/library/bb213677.aspx).
        */
        wxString GetNumberFormat();
       
        /**
        Sets the format code for the range.

        [MSDN documentation for Range.NumberFormat](http://msdn.microsoft.com/en-us/library/bb213677.aspx).
        */
        void SetNumberFormat(const wxString& numberFormat);

        /**
        Returns the format code for the range as a string in the language of the user or an empty string if all the cells in range don't share the same format.

        [MSDN documentation for Range.NumberFormatLocal](http://msdn.microsoft.com/en-us/library/bb213679.aspx).
        */
        wxString GetNumberFormatLocal();
        
        /**
        Sets the format code for the range as a string in the language of the user.

        [MSDN documentation for Range.NumberFormatLocal](http://msdn.microsoft.com/en-us/library/bb213679.aspx).
        */
        void SetNumberFormatLocal(const wxString& numberFormatLocal);

        /**
        Returns a Range Represents a range that's offset from the specified range.

        [MSDN documentation for Range.Offset](http://msdn.microsoft.com/en-us/library/bb213689.aspx).
        */
        wxExcelRange GetOffset(long rowOffset = 0, long columnOffset = 0);
        
        /**
        Returns the text orientation. See XlOrientation

        [MSDN documentation for Range.Orientation](http://msdn.microsoft.com/en-us/library/bb213693.aspx).
        */
        long GetOrientation();
        
        /**
        Sets the text orientation. See XlOrientation

        [MSDN documentation for Range.Orientation](http://msdn.microsoft.com/en-us/library/bb213693.aspx).
        */
        void SetOrientation(long orientation);

        /**
        Returns the current outline level of the specified row or column.

        [MSDN documentation for Range.OutlineLevel](http://msdn.microsoft.com/en-us/library/bb208909.aspx).
        */
        long GetOutlineLevel();
        
        /**
        Sets the current outline level of the specified row or column.

        [MSDN documentation for Range.OutlineLevel](http://msdn.microsoft.com/en-us/library/bb208909.aspx).
        */
        void SetOutlineLevel(long outlineLevel);

        /**
        Returns the location of a page break. Can be one of the following XlPageBreak constants: xlPageBreakAutomatic, xlPageBreakManual, or xlPageBreakNone.

        [MSDN documentation for Range.PageBreak](http://msdn.microsoft.com/en-us/library/bb208914.aspx).
        */
        XlPageBreak GetPageBreak();
        
        /**
        Sets the location of a page break. Can be one of the following XlPageBreak constants: xlPageBreakAutomatic, xlPageBreakManual, or xlPageBreakNone.

        [MSDN documentation for Range.PageBreak](http://msdn.microsoft.com/en-us/library/bb208914.aspx).
        */
        void SetPageBreak(XlPageBreak pageBreak);

        /**

        [MSDN documentation for Range.Precedents](http://msdn.microsoft.com/en-us/library/bb208963.aspx).
        */
        wxExcelRange GetPrecedents();
        
        /**
        Returns the prefix character for the cell.

        [MSDN documentation for Range.PrefixCharacter](http://msdn.microsoft.com/en-us/library/bb208970.aspx).
        */
        wxString GetPrefixCharacter();
        
        /**
        Returns a range represnting the next cell.

        [MSDN documentation for Range.Previous](http://msdn.microsoft.com/en-us/library/bb213711.aspx).
        */
        wxExcelRange GetPrevious();

        /**
        Returns the reading order for the specified object. Can be one of the following constants: xlRTL (right-to-left), xlLTR (left-to-right), or xlContext.

        [MSDN documentation for Range.ReadingOrder](http://msdn.microsoft.com/en-us/library/bb237496.aspx).
        */
        long GetReadingOrder();
        
        /**
        Resizes the specified range. Returns a Range Represents the resized range.

        [MSDN documentation for Range.Resize](http://msdn.microsoft.com/en-us/library/bb242066.aspx).
        */
        wxExcelRange GetResize(long* rowSize= NULL, long* columnSize = NULL);
        
        /**
        Returns the number of the first row of the first area in the range.

        [MSDN documentation for Range.Row](http://msdn.microsoft.com/en-us/library/bb221550.aspx).
        */
        long GetRow();
        
        /**
        If all rows in the range have same height returns the height in points else returns -1. Returns -2 on error.

        [MSDN documentation for Range.RowHeight](http://msdn.microsoft.com/en-us/library/bb221565.aspx).
        */
        double GetRowHeight();
        
        /**
        Sets the height of all the rows in the range specified, measured in points.

        [MSDN documentation for Range.RowHeight](http://msdn.microsoft.com/en-us/library/bb221565.aspx).
        */
        void SetRowHeight(double rowHeight);

        /**
        True if the outline is expanded for the specified range (so that the detail of the column or row is visible). The specified range must be a single summary column or row in an outline. For the PivotItem object (or the Range object if the range is in a PivotTable report), this property is set to True if the item is showing detail.

        [MSDN documentation for Range.ShowDetail](http://msdn.microsoft.com/en-us/library/bb238593.aspx).
        */
        bool GetShowDetail();
        
        /**
        Returns @c tb_true if the text shrinks to fit into the cell in all cells in the range.
        Returns @c tb_default if this property is not the same in all cells in the range.

        [MSDN documentation for Range.ShrinkToFit](http://msdn.microsoft.com/en-us/library/bb238595.aspx).
        */
        wxXlTribool GetShrinkToFit();
        
        /**
        Set true if the text is to fit into the cell.

        [MSDN documentation for Range.ShrinkToFit](http://msdn.microsoft.com/en-us/library/bb238595.aspx).
        */
        void SetShrinkToFit(bool shrinkToFit);

#if WXAUTOEXCEL_USE_CHARTS
        /**
        Returns a SparklineGroups object that represents an existing group of sparklines from the specified range. Since Excel 2010.

        [MSDN documentation for Range.FormatConditions](http://msdn.microsoft.com/en-us/library/office/ff822140%28v=office.14%29.aspx).
        */

        wxExcelSparklineGroups GetSparklineGroups();

#endif  // WXAUTOEXCEL_USE_CHARTS

        /**
        Returns the style of the range.

        [MSDN documentation for Range.Style](http://msdn.microsoft.com/en-us/library/bb238599.aspx).
        */
        wxExcelStyle GetStyle();
        
        /**
        Sets the style of the range.

        [MSDN documentation for Range.Style](http://msdn.microsoft.com/en-us/library/bb238599.aspx).
        */
        void SetStyle(wxExcelStyle style);

        /**
        True if the range is an outlining summary row or column. The range should be a row or a column.

        [MSDN documentation for Range.Summary](http://msdn.microsoft.com/en-us/library/bb209318.aspx).
        */
        bool GetSummary();
        
        /**
        Returns the text value as seen by user in MS Excel. The range must comprise of a single cell, else an empty string is returned.

        [MSDN documentation for Range.Text](http://msdn.microsoft.com/en-us/library/bb238601.aspx).
        */       
        wxString GetText();
        
        /**
        Returns the distance, in points, from the top edge of row 1 to the top edge of the range.

        [MSDN documentation for Range.Top](http://msdn.microsoft.com/en-us/library/bb238604.aspx).
        */
        double GetTop();
        
        /**
        Returns @c tb_true if the row height of the Range object equals the standard height of the sheet. 
        Returns @c tb_default if the range contains more than one row and the rows are not all the same height.

        [MSDN documentation for Range.UseStandardHeight](http://msdn.microsoft.com/en-us/library/bb221989.aspx).
        */
        wxXlTribool GetUseStandardHeight();
        
        /**
        Whether all cells in the range should use standard row height of the sheet.

        [MSDN documentation for Range.UseStandardHeight](http://msdn.microsoft.com/en-us/library/bb221989.aspx).
        */
        void SetUseStandardHeight(bool useStandardHeight);

        /**
        Returns @c tb_true if the column width of the Range object equals the standard width of the sheet. 
        Returns @c tb_default if the range contains more than one column and the columns are not all the same width.

        [MSDN documentation for Range.UseStandardWidth](http://msdn.microsoft.com/en-us/library/bb221992.aspx).
        */
        wxXlTribool GetUseStandardWidth();
        
        /**
        Whether all cells in the range should use standard column width of the sheet.

        [MSDN documentation for Range.UseStandardWidth](http://msdn.microsoft.com/en-us/library/bb221992.aspx).
        */
        void SetUseStandardWidth(bool useStandardWidth);

        /**

        Returns the Validation object that represents data validation for the specified range.

        [MSDN documentation for Range.Validation](http://msdn.microsoft.com/en-us/library/bb223003.aspx).
        */
        wxExcelValidation GetValidation();

        /**
        Returns the value(s) of the specified range.

        [MSDN documentation for Range.Value](http://msdn.microsoft.com/en-us/library/bb238606.aspx).

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
        Sets the value for a Range.

        [MSDN documentation for Range.Value](http://msdn.microsoft.com/en-us/library/bb238606.aspx).
        
        Be aware that the automation LCID (see wxAutoExcelObject::SetAutomationLCID_()) also affects how MS Excel may
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

        [MSDN documentation for Range.Value2](http://msdn.microsoft.com/en-us/library/bb223007.aspx).
        */
        wxVariant GetValue2();

        /**
        Sets the cell value.

        Also see SetValue() for some oddities of this function.

        [MSDN documentation for Range.Value2](http://msdn.microsoft.com/en-us/library/bb223007.aspx).
        */
        void SetValue2(const wxVariant& value);

        /**
        Returns the vertical alignment of the specified object.

        [MSDN documentation for Range.VerticalAlignment](http://msdn.microsoft.com/en-us/library/bb238610.aspx).
        */
        long GetVerticalAlignment();
        
        /**
        Sets the vertical alignment of the specified object.

        [MSDN documentation for Range.VerticalAlignment](http://msdn.microsoft.com/en-us/library/bb238610.aspx).
        */
        void SetVerticalAlignment(long verticalAlignment);

        /**
        Returns the width, in units, of the range.

        [MSDN documentation for Range.Width](http://msdn.microsoft.com/en-us/library/bb238613.aspx).
        */
        double GetWidth();
        
        /**
        Returns the worksheet containing the specified range.

        [MSDN documentation for Range.Worksheet](http://msdn.microsoft.com/en-us/library/bb223066.aspx).
        */
        wxExcelWorksheet GetWorksheet();
        
        /**              
        Returns @c tb_true if all cells in the range wrap the text, @c tb_false if all cells in the range
        do not wrap the text, and @c tb_default if some cells wrap the text and some do not.

        [MSDN documentation for Range.WrapText](http://msdn.microsoft.com/en-us/library/bb238616.aspx).
        */
        wxXlTribool GetWrapText();
        
        /**
        Set to true if Microsoft Excel should wrap the text in the object in all cells in the range.

        [MSDN documentation for Range.WrapText](http://msdn.microsoft.com/en-us/library/bb238616.aspx).
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
        wxExcelRange DoGetItem(long rowIndex, const wxVariant& columnIndex);
        wxExcelRange DoGetRange(const wxVariant& cell1, const wxVariant& cell2);

        // address can be either "Address" or "AddressLocal"
        wxString DoGetAddress(const wxString& address, wxXlTribool rowAbsolute, wxXlTribool columnAbsolute,
                              XlReferenceStyle* referenceStyle, wxXlTribool external, wxExcelRange* relativeTo);
};

} // namespace wxAutoExcel

#endif // _WXAUTOEXCEL_RANGE_H
