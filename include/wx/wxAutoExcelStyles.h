/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// License:     MIT license
/////////////////////////////////////////////////////////////////////////////


#ifndef _WXAUTOEXCEL_STYLES_H
#define _WXAUTOEXCEL_STYLES_H

#include "wx/wxAutoExcel_defs.h"
#include "wx/wxAutoExcel_object.h"

namespace wxAutoExcel {

    /**
    @brief Represents Microsoft Excel Style object.
    */
   class WXDLLIMPEXP_WXAUTOEXCEL wxExcelStyle: public wxExcelObject
    {
    public:
        // ***** METHODS *****

        /**
        Deletes the object.

        [MSDN documentation for Style.Delete](https://learn.microsoft.com/en-us/office/vba/api/excel.style.delete).
        */
        bool Delete();

        // ***** PROPERTIES *****

        /**
        If true then text is automatically indented when the text alignment in a cell is set to equal distribution (either horizontally or vertically.)

        [MSDN documentation for Style.AddIndent]().
        */
        bool GetAddIndent();

        /**
        If true then text is automatically indented when the text alignment in a cell is set to equal distribution (either horizontally or vertically.)

        [MSDN documentation for Style.AddIndent]().
        */
        void SetAddIndent(bool addIndent);

        /**
        Returns a Borders collection that represents the borders of a style or a range of cells (including a range defined as part of a conditional format).

        [MSDN documentation for Style.Borders](https://learn.microsoft.com/en-us/office/vba/api/excel.style.borders).
        */
        wxExcelBorders GetBorders();

        /**
        Returns true if the style is one of the built-in ones.

        [MSDN documentation for Style.BuiltIn](https://learn.microsoft.com/en-us/office/vba/api/excel.style.builtin).
        */
        bool GetBuiltIn();

        /**
        Returns a Font Represents the style font.

        [MSDN documentation for Style.Font](https://learn.microsoft.com/en-us/office/vba/api/excel.style.font).
        */
        wxExcelFont GetFont();

        /**
        If true the formula will be hidden when the worksheet is protected.

        [MSDN documentation for Style.FormulaHidden](https://learn.microsoft.com/en-us/office/vba/api/excel.style.formulahidden).
        */
        bool GetFormulaHidden();

        /**
        If true the formula will be hidden when the worksheet is protected.

        [MSDN documentation for Style.FormulaHidden](https://learn.microsoft.com/en-us/office/vba/api/excel.style.formulahidden).
        */
        void SetFormulaHidden(bool formulaHidden);

        /**
        Returns a XlHAlign value that represents the horizontal alignment.

        [MSDN documentation for Style.HorizontalAlignment](https://learn.microsoft.com/en-us/office/vba/api/excel.style.horizontalalignment).
        */
        XlHAlign GetHorizontalAlignment();

        /**
        Sets a XlHAlign value that represents the horizontal alignment.

        [MSDN documentation for Style.HorizontalAlignment](https://learn.microsoft.com/en-us/office/vba/api/excel.style.horizontalalignment).
        */
        void SetHorizontalAlignment(XlHAlign horizontalAlignment);

        /**
        True if the style includes the AddIndent, HorizontalAlignment, VerticalAlignment, WrapText, IndentLevel, and Orientation properties.

        [MSDN documentation for Style.IncludeAlignment](https://learn.microsoft.com/en-us/office/vba/api/excel.style.includealignment).
        */
        bool GetIncludeAlignment();

        /**
        True if the style includes the AddIndent, HorizontalAlignment, VerticalAlignment, WrapText, IndentLevel, and Orientation properties.

        [MSDN documentation for Style.IncludeAlignment](https://learn.microsoft.com/en-us/office/vba/api/excel.style.includealignment).
        */
        void SetIncludeAlignment(bool includeAlignment);

        /**
        True if the style includes the Color, ColorIndex, LineStyle, and Weight border properties.

        [MSDN documentation for Style.IncludeBorder](https://learn.microsoft.com/en-us/office/vba/api/excel.style.includeborder).
        */
        bool GetIncludeBorder();

        /**
        True if the style includes the Color, ColorIndex, LineStyle, and Weight border properties.

        [MSDN documentation for Style.IncludeBorder](https://learn.microsoft.com/en-us/office/vba/api/excel.style.includeborder).
        */
        void SetIncludeBorder(bool includeBorder);

        /**
        True if the style includes the Background, Bold, Color, ColorIndex, FontStyle, Italic, Name, Size, Strikethrough, Subscript, Superscript, and Underline font properties.

        [MSDN documentation for Style.IncludeFont](https://learn.microsoft.com/en-us/office/vba/api/excel.style.includefont).
        */
        bool GetIncludeFont();

        /**
        True if the style includes the Background, Bold, Color, ColorIndex, FontStyle, Italic, Name, Size, Strikethrough, Subscript, Superscript, and Underline font properties.

        [MSDN documentation for Style.IncludeFont](https://learn.microsoft.com/en-us/office/vba/api/excel.style.includefont).
        */
        void SetIncludeFont(bool includeFont);

        /**
        True if the style includes the NumberFormat property.

        [MSDN documentation for Style.IncludeNumber](https://learn.microsoft.com/en-us/office/vba/api/excel.style.includenumber).
        */
        bool GetIncludeNumber();

        /**
        True if the style includes the NumberFormat property. Read/write Boolean

        [MSDN documentation for Style.IncludeNumber](https://learn.microsoft.com/en-us/office/vba/api/excel.style.includenumber).
        */
        void SetIncludeNumber(bool includeNumber);

        /**
        True if the style includes the Color, ColorIndex, InvertIfNegative, Pattern, PatternColor, and PatternColorIndex interior properties.

        [MSDN documentation for Style.IncludePatterns](https://learn.microsoft.com/en-us/office/vba/api/excel.style.includepatterns).
        */
        bool GetIncludePatterns();

        /**
        True if the style includes the Color, ColorIndex, InvertIfNegative, Pattern, PatternColor, and PatternColorIndex interior properties.

        [MSDN documentation for Style.IncludePatterns](https://learn.microsoft.com/en-us/office/vba/api/excel.style.includepatterns).
        */
        void SetIncludePatterns(bool includePatterns);

        /**
        True if the style includes the FormulaHidden and Locked protection properties.

        [MSDN documentation for Style.IncludeProtection](https://learn.microsoft.com/en-us/office/vba/api/excel.style.includeprotection).
        */
        bool GetIncludeProtection();

        /**
        True if the style includes the FormulaHidden and Locked protection properties.

        [MSDN documentation for Style.IncludeProtection](https://learn.microsoft.com/en-us/office/vba/api/excel.style.includeprotection).
        */
        void SetIncludeProtection(bool includeProtection);

        /**
        Returns a the indent level for the style.

        [MSDN documentation for Style.IndentLevel](https://learn.microsoft.com/en-us/office/vba/api/excel.style.indentlevel).
        */
        long GetIndentLevel();

        /**
        Sets the indent level for the style.

        [MSDN documentation for Style.IndentLevel](https://learn.microsoft.com/en-us/office/vba/api/excel.style.indentlevel).
        */
        void SetIndentLevel(long indentLevel);

        /**
        Returns an Interior Represents the style interior.

        [MSDN documentation for Style.Interior](https://learn.microsoft.com/en-us/office/vba/api/excel.style.interior).
        */
        wxExcelInterior GetInterior();

        /**
        True if the style is locked.

        [MSDN documentation for Style.Locked](https://learn.microsoft.com/en-us/office/vba/api/excel.style.locked).
        */
        bool GetLocked();

        /**
        True if the style is locked.

        [MSDN documentation for Style.Locked](https://learn.microsoft.com/en-us/office/vba/api/excel.style.locked).
        */
        void SetLocked(bool locked);

        /**
        True if the style contains merged cells.

        [MSDN documentation for Style.MergeCells](https://learn.microsoft.com/en-us/office/vba/api/excel.style.mergecells).
        */
        bool GetMergeCells();

        /**
        Returns the name of the object.

        [MSDN documentation for Style.Name](https://learn.microsoft.com/en-us/office/vba/api/excel.style.name).
        */
        wxString GetName();

        /**
        Returns the name of the object, in the language of the user.

        [MSDN documentation for Style.NameLocal](https://learn.microsoft.com/en-us/office/vba/api/excel.style.namelocal).
        */
        wxString GetNameLocal();

        /**
        Returns a value that represents the format code for the object.

        The name is spelled Numberformat rather than NumberFormat as in Excel, because the Windows headers define GetNumberFormat as a macro, which would otherwise rename the method.

        [MSDN documentation for Style.NumberFormat](https://learn.microsoft.com/en-us/office/vba/api/excel.style.numberformat).
        */
        wxString GetNumberformat();

        /**
        Sets a value that represents the format code for the object.

        The name is spelled Numberformat rather than NumberFormat as in Excel, because the Windows headers define GetNumberFormat as a macro, which would otherwise rename the method.

        [MSDN documentation for Style.NumberFormat](https://learn.microsoft.com/en-us/office/vba/api/excel.style.numberformat).
        */
        void SetNumberformat(const wxString& numberFormat);

        /**
        Returns a value that represents the format code for the object as a string in the language of the user.

        The name is spelled Numberformat rather than NumberFormat as in Excel, because the Windows headers define GetNumberFormat as a macro, which would otherwise rename the method.

        [MSDN documentation for Style.NumberFormatLocal](https://learn.microsoft.com/en-us/office/vba/api/excel.style.numberformatlocal).
        */
        wxString GetNumberformatLocal();

        /**
        Sets a value that represents the format code for the object as a string in the language of the user.

        The name is spelled Numberformat rather than NumberFormat as in Excel, because the Windows headers define GetNumberFormat as a macro, which would otherwise rename the method.

        [MSDN documentation for Style.NumberFormatLocal](https://learn.microsoft.com/en-us/office/vba/api/excel.style.numberformatlocal).
        */
        void SetNumberformatLocal(const wxString& numberFormatLocal);

        /**
        Returns the text orientation.

        [MSDN documentation for Style.Orientation](https://learn.microsoft.com/en-us/office/vba/api/excel.style.orientation).
        */
        XlOrientation GetOrientation();

        /**
        Sets athe text orientation.

        [MSDN documentation for Style.Orientation](https://learn.microsoft.com/en-us/office/vba/api/excel.style.orientation).
        */
        void SetOrientation(XlOrientation orientation);

        /**
        Returns the style reading order. Can be one of the following constants: xlRTL (right-to-left), xlLTR (left-to-right), or xlContext.

        [MSDN documentation for Style.ReadingOrder](https://learn.microsoft.com/en-us/office/vba/api/excel.style.readingorder).
        */
        long GetReadingOrder();

        /**
        Sets the style reading order. Can be one of the following constants: xlRTL (right-to-left), xlLTR (left-to-right), or xlContext.

        [MSDN documentation for Style.ReadingOrder](https://learn.microsoft.com/en-us/office/vba/api/excel.style.readingorder).
        */
        void SetReadingOrder(long readingOrder);

        /**
        True if text automatically shrinks to fit in the available column width.

        [MSDN documentation for Style.ShrinkToFit](https://learn.microsoft.com/en-us/office/vba/api/excel.style.shrinktofit).
        */
        bool GetShrinkToFit();

        /**
        True if text automatically shrinks to fit in the available column width.

        [MSDN documentation for Style.ShrinkToFit](https://learn.microsoft.com/en-us/office/vba/api/excel.style.shrinktofit).
        */
        void SetShrinkToFit(bool shrinkToFit);

        /**
        Returns the style vertical alignment.

        [MSDN documentation for Style.VerticalAlignment](https://learn.microsoft.com/en-us/office/vba/api/excel.style.verticalalignment).
        */
        XlVAlign GetVerticalAlignment();

        /**
        Sets the style vertical alignment.

        [MSDN documentation for Style.VerticalAlignment](https://learn.microsoft.com/en-us/office/vba/api/excel.style.verticalalignment).
        */
        void SetVerticalAlignment(XlVAlign verticalAlignment);

        /**
        True if Microsoft Excel wraps the text.

        [MSDN documentation for Style.WrapText](https://learn.microsoft.com/en-us/office/vba/api/excel.style.wraptext).
        */
        bool GetWrapText();

        /**
        True if Microsoft Excel wraps the text.

        [MSDN documentation for Style.WrapText](https://learn.microsoft.com/en-us/office/vba/api/excel.style.wraptext).
        */
        void SetWrapText(bool wrapText);


        /**
        Returns "Style".
        */
        virtual wxString GetAutoExcelObjectName_() const { return wxS("Style"); }
    };

    /**
    @brief Represents Microsoft Excel Styles collection.
    */
   class WXDLLIMPEXP_WXAUTOEXCEL wxExcelStyles: public wxExcelObject
    {
    public:
        // ***** METHODS *****

        /**
        Creates a new style and adds it to the list of styles that are available for the current workbook.

        [MSDN documentation for Styles.Add](https://learn.microsoft.com/en-us/office/vba/api/excel.styles.add).
        */
        wxExcelStyle Add(const wxString& name, wxExcelStyle* basedOn);

        /**
        Merges the styles from another workbook.

        [MSDN documentation for Styles.Merge](https://learn.microsoft.com/en-us/office/vba/api/excel.styles.merge).
        */
        void Merge(wxExcelWorkbook workbook);

        // ***** PROPERTIES *****

        /**
        Number of items in the collection

        [MSDN documentation for Styles.Count](https://learn.microsoft.com/en-us/office/vba/api/excel.styles.count).
        */
        long GetCount();

        //@{
        /**
        Returns the style.

        [MSDN documentation for Styles.Item](https://learn.microsoft.com/en-us/office/vba/api/excel.styles.item).
        */
        wxExcelStyle GetItem(long index);
        wxExcelStyle operator[](long index);
        wxExcelStyle GetItem(const wxString& name);
        wxExcelStyle operator[](const wxString& name);
        //@}

        /**
        Returns "Styles".
        */
        virtual wxString GetAutoExcelObjectName_() const { return wxS("Styles"); }
    };


} // namespace wxAutoExcel

#endif //_WXAUTOEXCEL_STYLES_H
