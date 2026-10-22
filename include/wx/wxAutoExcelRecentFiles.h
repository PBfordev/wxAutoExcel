/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// License:     MIT license
/////////////////////////////////////////////////////////////////////////////


#ifndef _WXAUTOEXCEL_RECENTFILES_H
#define _WXAUTOEXCEL_RECENTFILES_H

#include "wx/wxAutoExcel_defs.h"

#include "wx/wxAutoExcel_object.h"
#include "wx/wxAutoExcel_enums.h"

namespace wxAutoExcel {

    /**
    @brief Represents a Microsoft Excel RecentFile object.
    */
   class WXDLLIMPEXP_WXAUTOEXCEL wxExcelRecentFile : public wxExcelObject
    {
    public:
        // ***** METHODS *****

        /**
        Deletes the object.

        [MSDN documentation for RecentFile.Delete](https://learn.microsoft.com/en-us/office/vba/api/excel.recentfile.delete).
        */
        void Delete();

        /**
        Opens a recent workbook.

        [MSDN documentation for RecentFile.Open](https://learn.microsoft.com/en-us/office/vba/api/excel.recentfile.open).
        */
        wxExcelWorkbook Open();

        // ***** PROPERTIES *****

        /**
        Returns a Long value that represents the index number of the object within the collection of similar objects.

        [MSDN documentation for RecentFile.Index](https://learn.microsoft.com/en-us/office/vba/api/excel.recentfile.index).
        */
        long GetIndex();

        /**
        Returns a String value that represents the name of the object.

        [MSDN documentation for RecentFile.Name](https://learn.microsoft.com/en-us/office/vba/api/excel.recentfile.name).
        */
        wxString GetName();

        /**
        Returns a String value that represents the complete path to the workbook/file that this RecentFile object respresents.

        [MSDN documentation for RecentFile.Path](https://learn.microsoft.com/en-us/office/vba/api/excel.recentfile.path).
        */
        wxString GetPath();

        /**
        Returns "RecentFile".
        */
        virtual wxString GetAutoExcelObjectName_() const { return wxS("RecentFile"); }
    };

    /**
    @brief Represents a Microsoft Excel (Office) RecentFiles collection.
    */
   class WXDLLIMPEXP_WXAUTOEXCEL wxExcelRecentFiles : public wxExcelObject
    {
    public:
        // ***** METHODS *****

        /**
        Adds a file to the list of recently used files.

        [MSDN documentation for RecentFiles.Add](https://learn.microsoft.com/en-us/office/vba/api/excel.recentfiles.add).
        */
        wxExcelRecentFile Add(const wxString& name);

        // ***** PROPERTIES *****

        /**
        Returns a Long value that represents the number of objects in the collection.

        [MSDN documentation for RecentFiles.Count](https://learn.microsoft.com/en-us/office/vba/api/excel.recentfiles.count).
        */
        long GetCount();

        //@{
        /**
        Returns a single object from a collection.

        [MSDN documentation for RecentFiles.Item](https://learn.microsoft.com/en-us/office/vba/api/excel.recentfiles.item).
        */
        wxExcelRecentFile GetItem(long index);
        wxExcelRecentFile operator[](long index);
        //@}

        /**
        Returns the maximum number of files in the list of recently used files. Can be a value from 0 (zero) through 50.

        [MSDN documentation for RecentFiles.Maximum](https://learn.microsoft.com/en-us/office/vba/api/excel.recentfiles.maximum).
        */
        long GetMaximum();

        /**
        Sets the maximum number of files in the list of recently used files. Can be a value from 0 (zero) through 50.

        [MSDN documentation for RecentFiles.Maximum](https://learn.microsoft.com/en-us/office/vba/api/excel.recentfiles.maximum).
        */
        void SetMaximum(long maximum);


        /**
        Returns "RecentFiles".
        */
        virtual wxString GetAutoExcelObjectName_() const { return wxS("RecentFiles"); }
    };


} // namespace wxAutoExcel


#endif //_WXAUTOEXCEL_RECENTFILES_H
