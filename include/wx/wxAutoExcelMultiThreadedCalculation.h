/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Modified by:
// Copyright:   (c) 2020 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
/////////////////////////////////////////////////////////////////////////////

#ifndef _WXAUTOEXCEL_MULTITHREADEDCALCULATION_H
#define _WXAUTOEXCEL_MULTITHREADEDCALCULATION_H

#include "wx/wxAutoExcel_defs.h"

#include "wx/wxAutoExcel_object.h"

namespace wxAutoExcel
{

/**
@brief Represents an object that returns or sets the concurrent calculation mode.
*/
class WXDLLIMPEXP_WXAUTOEXCEL wxExcelMultiThreadedCalculation : public wxExcelObject
{
public:
    // ***** PROPERTIES *****

    /**
    The Enabled property allows MultiThreadedCalculation objects to be enabled or disabled at run time.

    [Excel VBA documentation for MultiThreadedCalculation.Enabled](https://docs.microsoft.com/en-us/office/vba/api/excel.multithreadedcalculation.enabled)
    */
    bool GetEnabled();

    /**
    The Enabled property allows MultiThreadedCalculation objects to be enabled or disabled at run time.

    [Excel VBA documentation for MultiThreadedCalculation.Enabled](https://docs.microsoft.com/en-us/office/vba/api/excel.multithreadedcalculation.enabled)
    */
    void SetEnabled(bool enabled);

    /**
    Gets the total count of the process threads that are a part of the specified MultiThreadedCalculation object.

    [Excel VBA documentation for MultiThreadedCalculation.ThreadCount](https://docs.microsoft.com/en-us/office/vba/api/excel.multithreadedcalculation.threadcount)
    */
    long GetThreadCount();

    /**
    Returns the thread mode for the specified MultiThreadedCalculation object. Read/write XlThreadMode.

    [Excel VBA documentation for MultiThreadedCalculation.ThreadMode](https://docs.microsoft.com/en-us/office/vba/api/excel.multithreadedcalculation.threadmode)
    */
    XlThreadMode GetThreadMode();

    /**
    Sets the thread mode for the specified MultiThreadedCalculation object. Read/write XlThreadMode.

    [Excel VBA documentation for MultiThreadedCalculation.ThreadMode](https://docs.microsoft.com/en-us/office/vba/api/excel.multithreadedcalculation.threadmode)
    */
    void SetThreadMode(XlThreadMode threadMode);

    /**
    Returns "MultiThreadedCalculation".
    */
    virtual wxString GetAutoExcelObjectName_() const { return wxS("MultiThreadedCalculation"); }

}; // class wxExcelMultiThreadedCalculation

} // namespace wxAutoExcel 

#endif // #ifndef _WXAUTOEXCEL_MULTITHREADEDCALCULATION_H
