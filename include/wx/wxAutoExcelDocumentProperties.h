/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Modified by:
// RCS-ID:      $Id: $
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
/////////////////////////////////////////////////////////////////////////////


#ifndef _WXAUTOEXCEL_DOCUMENTPROPERTIES_H
#define _WXAUTOEXCEL_DOCUMENTPROPERTIES_H

#include "wx/wxAutoExcel_defs.h"

#include "wx/wxAutoExcelObject.h"
#include "wx/wxAutoExcel_enums.h"

namespace wxAutoExcel {

    /**
    @brief Represents a Microsoft Excel DocumentProperty object.
    */
   class WXDLLIMPEXP_WXAUTOEXCEL wxExcelDocumentProperty : public wxExcelObject
    {
    public:
        // ***** METHODS *****

        void Delete();

        // ***** PROPERTIES *****
                
        wxString GetLinkSource();
        void SetLinkSource(const wxString& linkSource);

        bool GetLinkToContent();
        void SetLinkToContent(bool linkToContent);

        wxString GetName();
        void SetName(const wxString& name);

        MsoDocProperties GetType();
        void SetType(MsoDocProperties propertyType);

        wxVariant GetValue();
        
        /**
        Returns "DocumentProperty".
        */
        virtual wxString GetAutoExcelObjectName_() const { return wxS("DocumentProperty"); }    
    };

    /**
    @brief Represents a Microsoft Excel (Office) DocumentProperties collection.
    */
   class WXDLLIMPEXP_WXAUTOEXCEL wxExcelDocumentProperties : public wxExcelObject
    {
    public:
        // ***** METHODS *****

        /**
        Creates a new custom document property. You can add a new document property only to the custom DocumentProperties collection.

        [MSDN documentation for DocumentProperties.Add](http://msdn.microsoft.com/en-us/library/office/ff862806%28v=office.15%29.aspx).
        */

        wxExcelDocumentProperty Add(const wxString& name, bool linkToContent, MsoDocProperties propertyType,
                                    wxVariant* value = NULL, const wxString& linkSource = wxEmptyString);
        
        // ***** PROPERTIES *****

        /**
        Returns the number of items in the collection. 

        [MSDN documentation for DocumentProperties.Count](http://msdn.microsoft.com/en-us/library/bb224553.aspx).
        */
        long GetCount();

         //@{
        /**
        Returns a single DocumentProperty object from the collection.

        [MSDN documentation for DocumentProperties.Item](http://msdn.microsoft.com/en-us/library/office/ff861126%28v=office.15%29.aspx).
        */
        wxExcelDocumentProperty GetItem(long index);
        wxExcelDocumentProperty GetItem(const wxString& name);
        wxExcelDocumentProperty operator[](long index);        
        wxExcelDocumentProperty operator[](const wxString& name);
        //@}        
                        
        /**
        Returns "DocumentProperties".
        */
        virtual wxString GetAutoExcelObjectName_() const { return wxS("DocumentProperties"); }    
    };


} // namespace wxAutoExcel


#endif //_WXAUTOEXCEL_DOCUMENTPROPERTIES_H
