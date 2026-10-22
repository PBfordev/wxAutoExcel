/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// License:     MIT license
/////////////////////////////////////////////////////////////////////////////


#ifndef _WXAUTOEXCEL_ERRORCHECKINGOPTIONS_H
#define _WXAUTOEXCEL_ERRORCHECKINGOPTIONS_H

#include "wx/wxAutoExcel_defs.h"

#include "wx/wxAutoExcel_object.h"


namespace wxAutoExcel {

    /**
    @brief Represents Microsoft Excel ErrorCheckingOptions object.
    */
   class WXDLLIMPEXP_WXAUTOEXCEL wxExcelErrorCheckingOptions: public wxExcelObject
    {
    public:
        // ***** PROPERTIES *****

        /**
        Alerts the user for all cells that violate enabled error-checking rules. When this property is set to True (default), the AutoCorrect Options button appears next to all cells that violate enabled errors. False disables background checking for errors.

        [MSDN documentation for ErrorCheckingOptions.BackgroundChecking](https://learn.microsoft.com/en-us/office/vba/api/excel.errorcheckingoptions.backgroundchecking).
        */
        bool GetBackgroundChecking();

        /**
        Alerts the user for all cells that violate enabled error-checking rules. When this property is set to True (default), the AutoCorrect Options button appears next to all cells that violate enabled errors. False disables background checking for errors.

        [MSDN documentation for ErrorCheckingOptions.BackgroundChecking](https://learn.microsoft.com/en-us/office/vba/api/excel.errorcheckingoptions.backgroundchecking).
        */
        void SetBackgroundChecking(bool backgroundChecking);

        /**
        When set to True (default), Microsoft Excel identifies, with an AutoCorrect Options button, selected cells containing formulas that refer to empty cells. False disables empty cell reference checking.

        [MSDN documentation for ErrorCheckingOptions.EmptyCellReferences](https://learn.microsoft.com/en-us/office/vba/api/excel.errorcheckingoptions.emptycellreferences).
        */
        bool GetEmptyCellReferences();

        /**
        When set to True (default), Microsoft Excel identifies, with an AutoCorrect Options button, selected cells containing formulas that refer to empty cells. False disables empty cell reference checking.

        [MSDN documentation for ErrorCheckingOptions.EmptyCellReferences](https://learn.microsoft.com/en-us/office/vba/api/excel.errorcheckingoptions.emptycellreferences).
        */
        void SetEmptyCellReferences(bool emptyCellReferences);

        /**
        When set to True (default), Microsoft Excel identifies, with an AutoCorrect Options button, selected cells that contain formulas evaluating to an error. False disables error checking for cells that evaluate to an error value.

        [MSDN documentation for ErrorCheckingOptions.EvaluateToError](https://learn.microsoft.com/en-us/office/vba/api/excel.errorcheckingoptions.evaluatetoerror).
        */
        bool GetEvaluateToError();

        /**
        When set to True (default), Microsoft Excel identifies, with an AutoCorrect Options button, selected cells that contain formulas evaluating to an error. False disables error checking for cells that evaluate to an error value.

        [MSDN documentation for ErrorCheckingOptions.EvaluateToError](https://learn.microsoft.com/en-us/office/vba/api/excel.errorcheckingoptions.evaluatetoerror).
        */
        void SetEvaluateToError(bool evaluateToError);

        /**
        When set to True (default), Microsoft Excel identifies cells containing an inconsistent formula in a region. False disables the inconsistent formula check.

        [MSDN documentation for ErrorCheckingOptions.InconsistentFormula](https://learn.microsoft.com/en-us/office/vba/api/excel.errorcheckingoptions.inconsistentformula).
        */
        bool GetInconsistentFormula();

        /**
        When set to True (default), Microsoft Excel identifies cells containing an inconsistent formula in a region. False disables the inconsistent formula check.

        [MSDN documentation for ErrorCheckingOptions.InconsistentFormula](https://learn.microsoft.com/en-us/office/vba/api/excel.errorcheckingoptions.inconsistentformula).
        */
        void SetInconsistentFormula(bool inconsistentFormula);

        /**
        When set to True (default), Microsoft Excel identifies cells containing an inconsistent table formula . False disables the inconsistent formula check.

        [MSDN documentation for ErrorCheckingOptions.InconsistentTableFormula](https://learn.microsoft.com/en-us/office/vba/api/excel.errorcheckingoptions.inconsistenttableformula).
        */
        bool GetInconsistentTableFormula();

        /**
        When set to True (default), Microsoft Excel identifies cells containing an inconsistent table formula . False disables the inconsistent formula check.

        [MSDN documentation for ErrorCheckingOptions.InconsistentTableFormula](https://learn.microsoft.com/en-us/office/vba/api/excel.errorcheckingoptions.inconsistenttableformula).
        */
        void SetInconsistentTableFormula(bool inconsistentTableFormula);

        /**
        Returns the color of the indicator for error checking options. Read/write XlColorIndex.

        [MSDN documentation for ErrorCheckingOptions.IndicatorColorIndex](https://learn.microsoft.com/en-us/office/vba/api/excel.errorcheckingoptions.indicatorcolorindex).
        */
        XlColorIndex GetIndicatorColorIndex();

        /**
        Sets the color of the indicator for error checking options. Read/write XlColorIndex.

        [MSDN documentation for ErrorCheckingOptions.IndicatorColorIndex](https://learn.microsoft.com/en-us/office/vba/api/excel.errorcheckingoptions.indicatorcolorindex).
        */
        void SetIndicatorColorIndex(XlColorIndex indicatorColorIndex);

        /**
        A Boolean value that is True if data validation is enabled in a list.

        [MSDN documentation for ErrorCheckingOptions.ListDataValidation](https://learn.microsoft.com/en-us/office/vba/api/excel.errorcheckingoptions.listdatavalidation).
        */
        bool GetListDataValidation();

        /**
        A Boolean value that is True if data validation is enabled in a list.

        [MSDN documentation for ErrorCheckingOptions.ListDataValidation](https://learn.microsoft.com/en-us/office/vba/api/excel.errorcheckingoptions.listdatavalidation).
        */
        void SetListDataValidation(bool listDataValidation);

        /**
        When set to True (default), Microsoft Excel identifies, with an AutoCorrect Options button, selected cells that contain numbers written as text. False disables error checking for numbers written as text.

        [MSDN documentation for ErrorCheckingOptions.NumberAsText](https://learn.microsoft.com/en-us/office/vba/api/excel.errorcheckingoptions.numberastext).
        */
        bool GetNumberAsText();

        /**
        When set to True (default), Microsoft Excel identifies, with an AutoCorrect Options button, selected cells that contain numbers written as text. False disables error checking for numbers written as text.

        [MSDN documentation for ErrorCheckingOptions.NumberAsText](https://learn.microsoft.com/en-us/office/vba/api/excel.errorcheckingoptions.numberastext).
        */
        void SetNumberAsText(bool numberAsText);

        /**
        When set to True (default), Microsoft Excel identifies, with an AutoCorrect Options button, the selected cells that contain formulas referring to a range that omits adjacent cells that could be included. False disables error checking for omitted cells.

        [MSDN documentation for ErrorCheckingOptions.OmittedCells](https://learn.microsoft.com/en-us/office/vba/api/excel.errorcheckingoptions.omittedcells).
        */
        bool GetOmittedCells();

        /**
        When set to True (default), Microsoft Excel identifies, with an AutoCorrect Options button, the selected cells that contain formulas referring to a range that omits adjacent cells that could be included. False disables error checking for omitted cells.

        [MSDN documentation for ErrorCheckingOptions.OmittedCells](https://learn.microsoft.com/en-us/office/vba/api/excel.errorcheckingoptions.omittedcells).
        */
        void SetOmittedCells(bool omittedCells);


        /**
        When set to True (default), Microsoft Excel identifies, with an AutoCorrect Options button, cells that contain a text date with a two-digit year. False disables error checking for cells containing a text date with a two-digit year.

        [MSDN documentation for ErrorCheckingOptions.TextDate](https://learn.microsoft.com/en-us/office/vba/api/excel.errorcheckingoptions.textdate).
        */
        bool GetTextDate();

        /**
        When set to True (default), Microsoft Excel identifies, with an AutoCorrect Options button, cells that contain a text date with a two-digit year. False disables error checking for cells containing a text date with a two-digit year.

        [MSDN documentation for ErrorCheckingOptions.TextDate](https://learn.microsoft.com/en-us/office/vba/api/excel.errorcheckingoptions.textdate).
        */
        void SetTextDate(bool textDate);

        /**
        When set to True (default), Microsoft Excel identifies selected cells that are unlocked and contain a formula. False disables error checking for unlocked cells that contain formulas.

        [MSDN documentation for ErrorCheckingOptions.UnlockedFormulaCells](https://learn.microsoft.com/en-us/office/vba/api/excel.errorcheckingoptions.unlockedformulacells).
        */
        bool GetUnlockedFormulaCells();

        /**
        When set to True (default), Microsoft Excel identifies selected cells that are unlocked and contain a formula. False disables error checking for unlocked cells that contain formulas.

        [MSDN documentation for ErrorCheckingOptions.UnlockedFormulaCells](https://learn.microsoft.com/en-us/office/vba/api/excel.errorcheckingoptions.unlockedformulacells).
        */
        void SetUnlockedFormulaCells(bool unlockedFormulaCells);



        /**
        Returns "ErrorCheckingOptions".
        */
        virtual wxString GetAutoExcelObjectName_() const { return wxS("ErrorCheckingOptions"); }

};



} // namespace wxAutoExcel

#endif // _WXAUTOEXCEL_ERRORCHECKINGOPTIONS_H
