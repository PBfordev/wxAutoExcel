/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Modified by:
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
/////////////////////////////////////////////////////////////////////////////


#ifndef _WXAUTOEXCEL_ERRORS_H
#define _WXAUTOEXCEL_ERRORS_H

#include "wx/wxAutoExcel_defs.h"

#include "wx/wxAutoExcel_object.h"


namespace wxAutoExcel {

    /**
    @brief Represents Microsoft Excel Error object.
    */
   class WXDLLIMPEXP_WXAUTOEXCEL wxExcelError: public wxExcelObject
    {
    public:

        // ***** PROPERTIES *****

        /**
        Allows the user to set or return the state of an error checking option for a range. False enables an error checking option for a range. True disables an error checking option for a range.

        [MSDN documentation for Error.Ignore](http://msdn.microsoft.com/en-us/library/bb177586.aspx).
        */
        bool GetIgnore();

        /**
        Allows the user to set or return the state of an error checking option for a range. False enables an error checking option for a range. True disables an error checking option for a range.

        [MSDN documentation for Error.Ignore](http://msdn.microsoft.com/en-us/library/bb177586.aspx).
        */
        void SetIgnore(bool ignore);

        /**
        Returns a Boolean value that indicates if all the validation criteria are met (that is, if the range contains valid data).

        [MSDN documentation for Error.Value](http://msdn.microsoft.com/en-us/library/bb214602.aspx).
        */
        bool GetValue();


        /**
        Returns "Error".
        */
        virtual wxString GetAutoExcelObjectName_() const { return wxS("Error"); }

};



    /**
    @brief Represents Microsoft Excel Errors collection.
    */
   class WXDLLIMPEXP_WXAUTOEXCEL wxExcelErrors : public wxExcelObject
    {
    public:
        // ***** PROPERTIES *****

        //@{
        /**
        Returns a single member of the Error object. Item can be either an index or
        one of XlErrorChecks constants

        [MSDN documentation for Errors.Item](http://msdn.microsoft.com/en-us/library/bb236967.aspx).
        */
        wxExcelError GetItem(long index);
        wxExcelError operator[](long index);
        //@}

        /**
        Returns "Errors".
        */
        virtual wxString GetAutoExcelObjectName_() const { return wxS("Errors"); }

    };


} // namespace wxAutoExcel

#endif // _WXAUTOEXCEL_ERRORS_H
