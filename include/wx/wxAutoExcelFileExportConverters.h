/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Copyright:   (c) 2020 PB <pbfordev@gmail.com>
// License:     MIT license
/////////////////////////////////////////////////////////////////////////////

#ifndef _WXAUTOEXCEL_FILEEXPORTCONVERTERS_H
#define _WXAUTOEXCEL_FILEEXPORTCONVERTERS_H

#include "wx/wxAutoExcel_defs.h"

#include "wx/wxAutoExcel_object.h"

namespace wxAutoExcel
{

/**
@brief Represents a file converter that is used to save files.
*/
class WXDLLIMPEXP_WXAUTOEXCEL wxExcelFileExportConverter : public wxExcelObject
{
public:
    // ***** PROPERTIES *****

    /**
    Returns the description for the file converter.

    [Excel VBA documentation for FileExportConverter.Description](https://docs.microsoft.com/en-us/office/vba/api/excel.fileexportconverter.description)
    */
    wxString GetDescription();

    /**
    Returns the file name extensions associated with the specified FileExportConverter object.

    [Excel VBA documentation for FileExportConverter.Extensions](https://docs.microsoft.com/en-us/office/vba/api/excel.fileexportconverter.extensions)
    */
    wxString GetExtensions();

    /**
    Returns an integer that identifies the file format associated with the specified FileExportConverter object.

    [Excel VBA documentation for FileExportConverter.FileFormat](https://docs.microsoft.com/en-us/office/vba/api/excel.fileexportconverter.fileformat)
    */
    long GetFileFormat();

    /**
    Returns "FileExportConverter".
    */
    virtual wxString GetAutoExcelObjectName_() const { return wxS("FileExportConverter"); }

}; // class wxExcelFileExportConverter

/**
    @brief Represents a collection of FileExportConverter objects that represent all the file converters available for saving files.
*/
class WXDLLIMPEXP_WXAUTOEXCEL wxExcelFileExportConverters : public wxExcelObject
{
public:
    // ***** PROPERTIES *****

    /**
        Returns the number of objects in the collection.
    */
    long GetCount();

     //@{
    /**
        Returns the FileExportConverter with the given index.

        [MSDN documentation for FileExportConverters.Item](https://docs.microsoft.com/en-us/office/vba/api/excel.fileexportconverters.item)
    */
    wxExcelFileExportConverter GetItem(long index);
    wxExcelFileExportConverter operator[](long index);
    //@}

    /**
    Returns "FileExportConverters".
    */
    virtual wxString GetAutoExcelObjectName_() const { return wxS("FileExportConverters"); }

}; // class wxExcelFileExportConverters

} // namespace wxAutoExcel

#endif // #ifndef _WXAUTOEXCEL_FILEEXPORTCONVERTERs_H
