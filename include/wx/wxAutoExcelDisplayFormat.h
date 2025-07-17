/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Copyright:   (c) 2020 PB <pbfordev@gmail.com>
// License:     MIT license
/////////////////////////////////////////////////////////////////////////////

#ifndef _WXAUTOEXCEL_DISPLAYFORMAT_H
#define _WXAUTOEXCEL_DISPLAYFORMAT_H

#include "wx/wxAutoExcel_defs.h"

#include "wx/wxAutoExcel_object.h"

namespace wxAutoExcel
{

/**
@brief Represents the display settings for an associated Range object.
*/
class WXDLLIMPEXP_WXAUTOEXCEL wxExcelDisplayFormat : public wxExcelObject
{
public:
    // ***** PROPERTIES *****

    /**
    Returns a value that indicates if Microsoft Excel automatically indents text of the associated Range object when the text alignment in a cell is set to equal distribution (either horizontally or vertically), as it is displayed in the current user interface.

    [Excel VBA documentation for DisplayFormat.AddIndent](https://docs.microsoft.com/en-us/office/vba/api/excel.displayformat.addindent)
    */
    bool GetAddIndent();

    /**
    Returns a Borders object that represents the borders of the associated Range object as it is displayed in the current user interface.

    [Excel VBA documentation for DisplayFormat.Borders](https://docs.microsoft.com/en-us/office/vba/api/excel.displayformat.borders)
    */
    wxExcelBorders GetBorders();

    /**
    Returns a Characters object that represents a range of characters within the text of the associated Range object as it is displayed in the current user interface.

    [Excel VBA documentation for DisplayFormat.Characters](https://docs.microsoft.com/en-us/office/vba/api/excel.displayformat.characters)
    */
    wxExcelCharacters GetCharacters(long start = 1, long* length = NULL);

    /**
    Returns a Font object that represents the font of the associated Range as it is displayed in the current user interface.

    [Excel VBA documentation for DisplayFormat.Font](https://docs.microsoft.com/en-us/office/vba/api/excel.displayformat.font)
    */
    wxExcelFont GetFont();

    /**
    Whether the formula will be hidden when the worksheet is protected.
    Returns @c tb_true if all cells in the range have this property set to true,
    @c tb_false if all cells in the range have this property set to false,
    and @c tb_default if the value of this property is not same for all cells in the range.

    [Excel VBA documentation for DisplayFormat.FormulaHidden](https://docs.microsoft.com/en-us/office/vba/api/excel.displayformat.formulahidden)
    */
    wxXlTribool GetFormulaHidden();

    /**
    Returns a value that represents the horizontal alignment of the associated Range object as it is displayed in the current user interface.

    [Excel VBA documentation for DisplayFormat.HorizontalAlignment](https://docs.microsoft.com/en-us/office/vba/api/excel.displayformat.horizontalalignment)
    */
    long GetHorizontalAlignment();

    /**
    Returns a value that represents the indent level of the associated Range object as it is displayed in the current user interface.

    [Excel VBA documentation for DisplayFormat.IndentLevel](https://docs.microsoft.com/en-us/office/vba/api/excel.displayformat.indentlevel)
    */
    long GetIndentLevel();

    /**
    Returns an Interior object that represents the interior of the associated Range object as it is displayed in the current user interface.

    [Excel VBA documentation for DisplayFormat.Interior](https://docs.microsoft.com/en-us/office/vba/api/excel.displayformat.interior)
    */
    wxExcelInterior GetInterior();

    /**
    Returns @c tb_true if all cells in the range are locked, @c tb_true if all cells in the range are unlocked,
    and @c tb_default is some cells are locked and some are not.

    [Excel VBA documentation for DisplayFormat.Locked](https://docs.microsoft.com/en-us/office/vba/api/excel.displayformat.locked)
    */
    wxXlTribool GetLocked();

    /**
    Returns a value that indicates if the associated Range object contains merged cells as it is displayed in the current user interface.

    [Excel VBA documentation for DisplayFormat.MergeCells](https://docs.microsoft.com/en-us/office/vba/api/excel.displayformat.mergecells)
    */
    bool GetMergeCells();

    /**
    Returns the format code for the range or an empty string if all the cells in range don't share the same format.

    [Excel VBA documentation for DisplayFormat.NumberFormat](https://docs.microsoft.com/en-us/office/vba/api/excel.displayformat.numberformat)
    */
    wxString GetNumberFormat();

    /**
    Returns a value that represents the format code of the associated Range as a string in the language of the user as it is displayed in the current user interface
    or an empty string if all the cells in range don't share the same format.

    [Excel VBA documentation for DisplayFormat.NumberFormatLocal](https://docs.microsoft.com/en-us/office/vba/api/excel.displayformat.numberformatlocal)
    */
    wxString GetNumberFormatLocal();

    /**
    Returns a value that represents the text orientation of the associated Range object as it is displayed in the current user interface.

    [Excel VBA documentation for DisplayFormat.Orientation](https://docs.microsoft.com/en-us/office/vba/api/excel.displayformat.orientation)
    */
    long GetOrientation();

    /**
    Returns the reading order of the associated Range object as it is displayed in the current user interface.

    [Excel VBA documentation for DisplayFormat.ReadingOrder](https://docs.microsoft.com/en-us/office/vba/api/excel.displayformat.readingorder)
    */
    long GetReadingOrder();

    /**
    Returns @c tb_true if the text shrinks to fit into the cell in all cells in the range.
    Returns @c tb_default if this property is not the same in all cells in the range.

    [Excel VBA documentation for DisplayFormat.ShrinkToFit](https://docs.microsoft.com/en-us/office/vba/api/excel.displayformat.shrinktofit)
    */
    wxXlTribool GetShrinkToFit();

    /**
    Returns a value, containing a Style object, that represents the style of the associated Range object as it is displayed in the current user interface.

    [Excel VBA documentation for DisplayFormat.Style](https://docs.microsoft.com/en-us/office/vba/api/excel.displayformat.style)
    */
    wxExcelStyle GetStyle();

    /**
    Returns a value that represents the vertical alignment of the associated Range object as it is displayed in the current user interface.

    [Excel VBA documentation for DisplayFormat.VerticalAlignment](https://docs.microsoft.com/en-us/office/vba/api/excel.displayformat.verticalalignment)
    */
    long GetVerticalAlignment();

    /**
    Returns @c tb_true if all cells in the range wrap the text, @c tb_false if all cells in the range
    do not wrap the text, and @c tb_default if some cells wrap the text and some do not.

    [Excel VBA documentation for DisplayFormat.WrapText](https://docs.microsoft.com/en-us/office/vba/api/excel.displayformat.wraptext)
    */
    wxXlTribool GetWrapText();

    /**
    Returns "DisplayFormat".
    */
    virtual wxString GetAutoExcelObjectName_() const { return wxS("DisplayFormat"); }

}; // class wxExcelDisplayFormat

} // namespace wxAutoExcel

#endif // #ifndef _WXAUTOEXCEL_DISPLAYFORMAT_H
