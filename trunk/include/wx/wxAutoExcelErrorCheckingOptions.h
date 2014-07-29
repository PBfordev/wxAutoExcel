/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Modified by:
// RCS-ID:      $Id: $
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
/////////////////////////////////////////////////////////////////////////////


#ifndef _WXAUTOEXCEL_ERRORCHECKINGOPTIONS_H
#define _WXAUTOEXCEL_ERRORCHECKINGOPTIONS_H

#include "wx/wxAutoExcel_defs.h"

#include "wx/wxAutoExcelObject.h"


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

        [MSDN documentation for ErrorCheckingOptions.BackgroundChecking](http://msdn.microsoft.com/en-us/library/bb220876.aspx).
        */
        bool GetBackgroundChecking();

        /**
        Alerts the user for all cells that violate enabled error-checking rules. When this property is set to True (default), the AutoCorrect Options button appears next to all cells that violate enabled errors. False disables background checking for errors.

        [MSDN documentation for ErrorCheckingOptions.BackgroundChecking](http://msdn.microsoft.com/en-us/library/bb220876.aspx).
        */
        void SetBackgroundChecking(bool backgroundChecking);

        /**
        When set to True (default), Microsoft Excel identifies, with an AutoCorrect Options button, selected cells containing formulas that refer to empty cells. False disables empty cell reference checking.

        [MSDN documentation for ErrorCheckingOptions.EmptyCellReferences](http://msdn.microsoft.com/en-us/library/bb221100.aspx).
        */
        bool GetEmptyCellReferences();

        /**
        When set to True (default), Microsoft Excel identifies, with an AutoCorrect Options button, selected cells containing formulas that refer to empty cells. False disables empty cell reference checking.

        [MSDN documentation for ErrorCheckingOptions.EmptyCellReferences](http://msdn.microsoft.com/en-us/library/bb221100.aspx).
        */
        void SetEmptyCellReferences(bool emptyCellReferences);

        /**
        When set to True (default), Microsoft Excel identifies, with an AutoCorrect Options button, selected cells that contain formulas evaluating to an error. False disables error checking for cells that evaluate to an error value.

        [MSDN documentation for ErrorCheckingOptions.EvaluateToError](http://msdn.microsoft.com/en-us/library/bb208484.aspx).
        */
        bool GetEvaluateToError();

        /**
        When set to True (default), Microsoft Excel identifies, with an AutoCorrect Options button, selected cells that contain formulas evaluating to an error. False disables error checking for cells that evaluate to an error value.

        [MSDN documentation for ErrorCheckingOptions.EvaluateToError](http://msdn.microsoft.com/en-us/library/bb208484.aspx).
        */
        void SetEvaluateToError(bool evaluateToError);

        /**
        When set to True (default), Microsoft Excel identifies cells containing an inconsistent formula in a region. False disables the inconsistent formula check.

        [MSDN documentation for ErrorCheckingOptions.InconsistentFormula](http://msdn.microsoft.com/en-us/library/bb177630.aspx).
        */
        bool GetInconsistentFormula();

        /**
        When set to True (default), Microsoft Excel identifies cells containing an inconsistent formula in a region. False disables the inconsistent formula check.

        [MSDN documentation for ErrorCheckingOptions.InconsistentFormula](http://msdn.microsoft.com/en-us/library/bb177630.aspx).
        */
        void SetInconsistentFormula(bool inconsistentFormula);

        /**
        When set to True (default), Microsoft Excel identifies cells containing an inconsistent table formula . False disables the inconsistent formula check.

        [MSDN documentation for ErrorCheckingOptions.InconsistentTableFormula](http://msdn.microsoft.com).
        */
        bool GetInconsistentTableFormula();

        /**
        When set to True (default), Microsoft Excel identifies cells containing an inconsistent table formula . False disables the inconsistent formula check.

        [MSDN documentation for ErrorCheckingOptions.InconsistentTableFormula](http://msdn.microsoft.com).
        */
        void SetInconsistentTableFormula(bool inconsistentTableFormula);

        /**
        Returns the color of the indicator for error checking options. Read/write XlColorIndex.

        [MSDN documentation for ErrorCheckingOptions.IndicatorColorIndex](http://msdn.microsoft.com/en-us/library/bb177633.aspx).
        */
        XlColorIndex GetIndicatorColorIndex();

        /**
        Sets the color of the indicator for error checking options. Read/write XlColorIndex.

        [MSDN documentation for ErrorCheckingOptions.IndicatorColorIndex](http://msdn.microsoft.com/en-us/library/bb177633.aspx).
        */
        void SetIndicatorColorIndex(XlColorIndex indicatorColorIndex);

        /**
        A Boolean value that is True if data validation is enabled in a list.

        [MSDN documentation for ErrorCheckingOptions.ListDataValidation](http://msdn.microsoft.com/en-us/library/bb177915.aspx).
        */
        bool GetListDataValidation();

        /**
        A Boolean value that is True if data validation is enabled in a list.

        [MSDN documentation for ErrorCheckingOptions.ListDataValidation](http://msdn.microsoft.com/en-us/library/bb177915.aspx).
        */
        void SetListDataValidation(bool listDataValidation);

        /**
        When set to True (default), Microsoft Excel identifies, with an AutoCorrect Options button, selected cells that contain numbers written as text. False disables error checking for numbers written as text.

        [MSDN documentation for ErrorCheckingOptions.NumberAsText](http://msdn.microsoft.com/en-us/library/bb208838.aspx).
        */
        bool GetNumberAsText();

        /**
        When set to True (default), Microsoft Excel identifies, with an AutoCorrect Options button, selected cells that contain numbers written as text. False disables error checking for numbers written as text.

        [MSDN documentation for ErrorCheckingOptions.NumberAsText](http://msdn.microsoft.com/en-us/library/bb208838.aspx).
        */
        void SetNumberAsText(bool numberAsText);

        /**
        When set to True (default), Microsoft Excel identifies, with an AutoCorrect Options button, the selected cells that contain formulas referring to a range that omits adjacent cells that could be included. False disables error checking for omitted cells.

        [MSDN documentation for ErrorCheckingOptions.OmittedCells](http://msdn.microsoft.com/en-us/library/bb208874.aspx).
        */
        bool GetOmittedCells();

        /**
        When set to True (default), Microsoft Excel identifies, with an AutoCorrect Options button, the selected cells that contain formulas referring to a range that omits adjacent cells that could be included. False disables error checking for omitted cells.

        [MSDN documentation for ErrorCheckingOptions.OmittedCells](http://msdn.microsoft.com/en-us/library/bb208874.aspx).
        */
        void SetOmittedCells(bool omittedCells);

        
        /**
        When set to True (default), Microsoft Excel identifies, with an AutoCorrect Options button, cells that contain a text date with a two-digit year. False disables error checking for cells containing a text date with a two-digit year.

        [MSDN documentation for ErrorCheckingOptions.TextDate](http://msdn.microsoft.com/en-us/library/bb221725.aspx).
        */
        bool GetTextDate();

        /**
        When set to True (default), Microsoft Excel identifies, with an AutoCorrect Options button, cells that contain a text date with a two-digit year. False disables error checking for cells containing a text date with a two-digit year.

        [MSDN documentation for ErrorCheckingOptions.TextDate](http://msdn.microsoft.com/en-us/library/bb221725.aspx).
        */
        void SetTextDate(bool textDate);

        /**
        When set to True (default), Microsoft Excel identifies selected cells that are unlocked and contain a formula. False disables error checking for unlocked cells that contain formulas.

        [MSDN documentation for ErrorCheckingOptions.UnlockedFormulaCells](http://msdn.microsoft.com/en-us/library/bb221943.aspx).
        */
        bool GetUnlockedFormulaCells();

        /**
        When set to True (default), Microsoft Excel identifies selected cells that are unlocked and contain a formula. False disables error checking for unlocked cells that contain formulas.

        [MSDN documentation for ErrorCheckingOptions.UnlockedFormulaCells](http://msdn.microsoft.com/en-us/library/bb221943.aspx).
        */
        void SetUnlockedFormulaCells(bool unlockedFormulaCells);


        
        /**
        Returns "ErrorCheckingOptions".
        */
        virtual wxString GetAutoExcelObjectName_() const { return wxS("ErrorCheckingOptions"); }

};



} // namespace wxAutoExcel

#endif // _WXAUTOEXCEL_ERRORCHECKINGOPTIONS_H
