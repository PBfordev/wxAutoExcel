/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Modified by:
// RCS-ID:      $Id: $
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
/////////////////////////////////////////////////////////////////////////////


#ifndef _WXAUTOEXCEL_RECENTFILES_H
#define _WXAUTOEXCEL_RECENTFILES_H

#include "wx/wxAutoExcel_defs.h"

#include "wx/wxAutoExcelObject.h"
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

        [MSDN documentation for RecentFile.Delete](http://msdn.microsoft.com/en-us/library/bb178870.aspx).
        */
        void Delete();

        /**
        Opens a recent workbook.

        [MSDN documentation for RecentFile.Open](http://msdn.microsoft.com/en-us/library/bb178873.aspx).
        */
        wxExcelWorkbook Open();

        // ***** PROPERTIES *****

        /**
        Returns a Long value that represents the index number of the object within the collection of similar objects.

        [MSDN documentation for RecentFile.Index](http://msdn.microsoft.com/en-us/library/bb237498.aspx).
        */
        long GetIndex();

        /**
        Returns a String value that represents the name of the object.

        [MSDN documentation for RecentFile.Name](http://msdn.microsoft.com/en-us/library/bb237502.aspx).
        */
        wxString GetName();

        /**
        Returns a String value that represents the complete path to the workbook/file that this RecentFile object respresents.

        [MSDN documentation for RecentFile.Path](http://msdn.microsoft.com/en-us/library/bb237505.aspx).
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

        [MSDN documentation for RecentFiles.Add](http://msdn.microsoft.com/en-us/library/bb178877.aspx).
        */
        wxExcelRecentFile Add(const wxString& name);

        // ***** PROPERTIES *****

        /**
        Returns a Long value that represents the number of objects in the collection.

        [MSDN documentation for RecentFiles.Count](http://msdn.microsoft.com/en-us/library/bb237506.aspx).
        */
        long GetCount();

        //@{
        /**
        Returns a single object from a collection.

        [MSDN documentation for RecentFiles.Item](http://msdn.microsoft.com/en-us/library/bb237509.aspx).
        */
        wxExcelRecentFile GetItem(long index);        
        wxExcelRecentFile operator[](long index);        
        //@}                                  

        /**
        Returns the maximum number of files in the list of recently used files. Can be a value from 0 (zero) through 50. 

        [MSDN documentation for RecentFiles.Maximum](http://msdn.microsoft.com/en-us/library/bb208742.aspx).
        */
        long GetMaximum();

        /**
        Sets the maximum number of files in the list of recently used files. Can be a value from 0 (zero) through 50. 

        [MSDN documentation for RecentFiles.Maximum](http://msdn.microsoft.com/en-us/library/bb208742.aspx).
        */
        void SetMaximum(long maximum);        
        
                        
        /**
        Returns "RecentFiles".
        */
        virtual wxString GetAutoExcelObjectName_() const { return wxS("RecentFiles"); }    
    };


} // namespace wxAutoExcel


#endif //_WXAUTOEXCEL_RECENTFILES_H
