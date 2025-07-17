/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
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

        [MSDN documentation for Style.Delete](http://msdn.microsoft.com/en-us/library/bb179049).
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

        [MSDN documentation for Style.Borders](http://msdn.microsoft.com/en-us/library/bb238019).
        */
        wxExcelBorders GetBorders();

        /**
        Returns true if the style is one of the built-in ones.

        [MSDN documentation for Style.BuiltIn](http://msdn.microsoft.com/en-us/library/bb215808).
        */
        bool GetBuiltIn();

        /**
        Returns a Font Represents the style font.

        [MSDN documentation for Style.Font](http://msdn.microsoft.com/en-us/library/bb238026).
        */
        wxExcelFont GetFont();

        /**
        If true the formula will be hidden when the worksheet is protected.

        [MSDN documentation for Style.FormulaHidden](http://msdn.microsoft.com/en-us/library/bb238033).
        */
        bool GetFormulaHidden();

        /**
        If true the formula will be hidden when the worksheet is protected.

        [MSDN documentation for Style.FormulaHidden](http://msdn.microsoft.com/en-us/library/bb238033).
        */
        void SetFormulaHidden(bool formulaHidden);

        /**
        Returns a XlHAlign value that represents the horizontal alignment.

        [MSDN documentation for Style.HorizontalAlignment](http://msdn.microsoft.com/en-us/library/bb238035).
        */
        XlHAlign GetHorizontalAlignment();

        /**
        Sets a XlHAlign value that represents the horizontal alignment.

        [MSDN documentation for Style.HorizontalAlignment](http://msdn.microsoft.com/en-us/library/bb238035).
        */
        void SetHorizontalAlignment(XlHAlign horizontalAlignment);

        /**
        True if the style includes the AddIndent, HorizontalAlignment, VerticalAlignment, WrapText, IndentLevel, and Orientation properties.

        [MSDN documentation for Style.IncludeAlignment](http://msdn.microsoft.com/en-us/library/bb177611).
        */
        bool GetIncludeAlignment();

        /**
        True if the style includes the AddIndent, HorizontalAlignment, VerticalAlignment, WrapText, IndentLevel, and Orientation properties.

        [MSDN documentation for Style.IncludeAlignment](http://msdn.microsoft.com/en-us/library/bb177611).
        */
        void SetIncludeAlignment(bool includeAlignment);

        /**
        True if the style includes the Color, ColorIndex, LineStyle, and Weight border properties.

        [MSDN documentation for Style.IncludeBorder](http://msdn.microsoft.com/en-us/library/bb177613).
        */
        bool GetIncludeBorder();

        /**
        True if the style includes the Color, ColorIndex, LineStyle, and Weight border properties.

        [MSDN documentation for Style.IncludeBorder](http://msdn.microsoft.com/en-us/library/bb177613).
        */
        void SetIncludeBorder(bool includeBorder);

        /**
        True if the style includes the Background, Bold, Color, ColorIndex, FontStyle, Italic, Name, Size, Strikethrough, Subscript, Superscript, and Underline font properties.

        [MSDN documentation for Style.IncludeFont](http://msdn.microsoft.com/en-us/library/bb177617).
        */
        bool GetIncludeFont();

        /**
        True if the style includes the Background, Bold, Color, ColorIndex, FontStyle, Italic, Name, Size, Strikethrough, Subscript, Superscript, and Underline font properties.

        [MSDN documentation for Style.IncludeFont](http://msdn.microsoft.com/en-us/library/bb177617).
        */
        void SetIncludeFont(bool includeFont);

        /**
        True if the style includes the NumberFormat property.

        [MSDN documentation for Style.IncludeNumber](http://msdn.microsoft.com/en-us/library/bb177620).
        */
        bool GetIncludeNumber();

        /**
        True if the style includes the NumberFormat property. Read/write Boolean

        [MSDN documentation for Style.IncludeNumber](http://msdn.microsoft.com/en-us/library/bb177620).
        */
        void SetIncludeNumber(bool includeNumber);

        /**
        True if the style includes the Color, ColorIndex, InvertIfNegative, Pattern, PatternColor, and PatternColorIndex interior properties.

        [MSDN documentation for Style.IncludePatterns](http://msdn.microsoft.com/en-us/library/bb177623).
        */
        bool GetIncludePatterns();

        /**
        True if the style includes the Color, ColorIndex, InvertIfNegative, Pattern, PatternColor, and PatternColorIndex interior properties.

        [MSDN documentation for Style.IncludePatterns](http://msdn.microsoft.com/en-us/library/bb177623).
        */
        void SetIncludePatterns(bool includePatterns);

        /**
        True if the style includes the FormulaHidden and Locked protection properties.

        [MSDN documentation for Style.IncludeProtection](http://msdn.microsoft.com/en-us/library/bb177627).
        */
        bool GetIncludeProtection();

        /**
        True if the style includes the FormulaHidden and Locked protection properties.

        [MSDN documentation for Style.IncludeProtection](http://msdn.microsoft.com/en-us/library/bb177627).
        */
        void SetIncludeProtection(bool includeProtection);

        /**
        Returns a the indent level for the style.

        [MSDN documentation for Style.IndentLevel](http://msdn.microsoft.com/en-us/library/bb238044).
        */
        long GetIndentLevel();

        /**
        Sets the indent level for the style.

        [MSDN documentation for Style.IndentLevel](http://msdn.microsoft.com/en-us/library/bb238044).
        */
        void SetIndentLevel(long indentLevel);

        /**
        Returns an Interior Represents the style interior.

        [MSDN documentation for Style.Interior](http://msdn.microsoft.com/en-us/library/bb238047).
        */
        wxExcelInterior GetInterior();

        /**
        True if the style is locked.

        [MSDN documentation for Style.Locked](http://msdn.microsoft.com/en-us/library/bb238055).
        */
        bool GetLocked();

        /**
        True if the style is locked.

        [MSDN documentation for Style.Locked](http://msdn.microsoft.com/en-us/library/bb238055).
        */
        void SetLocked(bool locked);

        /**
        True if the style contains merged cells.

        [MSDN documentation for Style.MergeCells](http://msdn.microsoft.com/en-us/library/bb238161).
        */
        bool GetMergeCells();

        /**
        Returns the name of the object.

        [MSDN documentation for Style.Name](http://msdn.microsoft.com/en-us/library/bb238165).
        */
        wxString GetName();

        /**
        Returns the name of the object, in the language of the user.

        [MSDN documentation for Style.NameLocal](http://msdn.microsoft.com/en-us/library/bb238170).
        */
        wxString GetNameLocal();

        /**
        Returns a value that represents the format code for the object.

        [MSDN documentation for Style.NumberFormat](http://msdn.microsoft.com/en-us/library/bb238176).
        */
        wxString GetNumberFormat();

        /**
        Sets a value that represents the format code for the object.

        [MSDN documentation for Style.NumberFormat](http://msdn.microsoft.com/en-us/library/bb238176).
        */
        void SetNumberFormat(const wxString& numberFormat);

        /**
        Returns a value that represents the format code for the object as a string in the language of the user.

        [MSDN documentation for Style.NumberFormatLocal](http://msdn.microsoft.com/en-us/library/bb238182).
        */
        wxString GetNumberFormatLocal();

        /**
        Sets a value that represents the format code for the object as a string in the language of the user.

        [MSDN documentation for Style.NumberFormatLocal](http://msdn.microsoft.com/en-us/library/bb238182).
        */
        void SetNumberFormatLocal(const wxString& numberFormatLocal);

        /**
        Returns the text orientation.

        [MSDN documentation for Style.Orientation](http://msdn.microsoft.com/en-us/library/bb238187).
        */
        XlOrientation GetOrientation();

        /**
        Sets athe text orientation.

        [MSDN documentation for Style.Orientation](http://msdn.microsoft.com/en-us/library/bb238187).
        */
        void SetOrientation(XlOrientation orientation);

        /**
        Returns the style reading order. Can be one of the following constants: xlRTL (right-to-left), xlLTR (left-to-right), or xlContext.

        [MSDN documentation for Style.ReadingOrder](http://msdn.microsoft.com/en-us/library/bb238192).
        */
        long GetReadingOrder();

        /**
        Sets the style reading order. Can be one of the following constants: xlRTL (right-to-left), xlLTR (left-to-right), or xlContext.

        [MSDN documentation for Style.ReadingOrder](http://msdn.microsoft.com/en-us/library/bb238192).
        */
        void SetReadingOrder(long readingOrder);

        /**
        True if text automatically shrinks to fit in the available column width.

        [MSDN documentation for Style.ShrinkToFit](http://msdn.microsoft.com/en-us/library/bb215134).
        */
        bool GetShrinkToFit();

        /**
        True if text automatically shrinks to fit in the available column width.

        [MSDN documentation for Style.ShrinkToFit](http://msdn.microsoft.com/en-us/library/bb215134).
        */
        void SetShrinkToFit(bool shrinkToFit);

        /**
        Returns the style vertical alignment.

        [MSDN documentation for Style.VerticalAlignment](http://msdn.microsoft.com/en-us/library/bb215140).
        */
        XlVAlign GetVerticalAlignment();

        /**
        Sets the style vertical alignment.

        [MSDN documentation for Style.VerticalAlignment](http://msdn.microsoft.com/en-us/library/bb215140).
        */
        void SetVerticalAlignment(XlVAlign verticalAlignment);

        /**
        True if Microsoft Excel wraps the text.

        [MSDN documentation for Style.WrapText](http://msdn.microsoft.com/en-us/library/bb215143).
        */
        bool GetWrapText();

        /**
        True if Microsoft Excel wraps the text.

        [MSDN documentation for Style.WrapText](http://msdn.microsoft.com/en-us/library/bb215143).
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

        [MSDN documentation for Styles.Add](http://msdn.microsoft.com/en-us/library/bb179052).
        */
        wxExcelStyle Add(const wxString& name, wxExcelStyle* basedOn);

        /**
        Merges the styles from another workbook.

        [MSDN documentation for Styles.Merge](http://msdn.microsoft.com/en-us/library/bb179054).
        */
        void Merge(wxExcelWorkbook workbook);

        // ***** PROPERTIES *****

        /**
        Number of items in the collection

        [MSDN documentation for Styles.Count](http://msdn.microsoft.com/en-us/library/bb238198).
        */
        long GetCount();

        //@{
        /**
        Returns the style.

        [MSDN documentation for Styles.Item](http://msdn.microsoft.com/en-us/library/bb238201).
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
