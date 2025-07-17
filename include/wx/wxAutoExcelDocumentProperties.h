/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
/////////////////////////////////////////////////////////////////////////////


#ifndef _WXAUTOEXCEL_DOCUMENTPROPERTIES_H
#define _WXAUTOEXCEL_DOCUMENTPROPERTIES_H

#include "wx/wxAutoExcel_defs.h"

#include "wx/wxAutoExcel_object.h"
#include "wx/wxAutoExcel_enums.h"

namespace wxAutoExcel {

    /**
    @brief Represents a Microsoft Excel DocumentProperty object.
    */
   class WXDLLIMPEXP_WXAUTOEXCEL wxExcelDocumentProperty : public wxExcelObject
    {
    public:
        // ***** METHODS *****

         /**
        Removes a custom document property.

        [MSDN documentation for DocumentProperty.Delete](http://msdn.microsoft.com/en-us/library/office/ff860911%28v=office.14%29.aspx).
        */
        void Delete();

        // ***** PROPERTIES *****

        /**
        Gets the source of a linked custom document property.

        [MSDN documentation for DocumentProperty.LinkSource](http://msdn.microsoft.com/en-us/library/office/ff861227%28v=office.14%29.aspx).
        */
        wxString GetLinkSource();
        /**
        Sets the source of a linked custom document property.

        [MSDN documentation for DocumentProperty.LinkSource](http://msdn.microsoft.com/en-us/library/office/ff861227%28v=office.14%29.aspx).
        */
        void SetLinkSource(const wxString& linkSource);

        /**
        Is true if the value of the custom document property is linked to the content of the container document. False if the value is static.

        [MSDN documentation for DocumentProperty.LinkToContent](http://msdn.microsoft.com/en-us/library/office/ff860252%28v=office.14%29.aspx).
        */
        bool GetLinkToContent();
        /**
        Is true if the value of the custom document property is linked to the content of the container document. False if the value is static.

        [MSDN documentation for DocumentProperty.LinkToContent](http://msdn.microsoft.com/en-us/library/office/ff860252%28v=office.14%29.aspx).
        */
        void SetLinkToContent(bool linkToContent);

        /**
        The name of a document property.

        [MSDN documentation for DocumentProperty.Name](http://msdn.microsoft.com/en-us/library/office/ff863966%28v=office.14%29.aspx).
        */
        wxString GetName();
        /**
        The name of a document property.

        [MSDN documentation for DocumentProperty.Name](http://msdn.microsoft.com/en-us/library/office/ff863966%28v=office.14%29.aspx).
        */
        void SetName(const wxString& name);

        /**
        The document property type. Read-only for built-in document properties; read/write for custom document properties.

        [MSDN documentation for DocumentProperty.Type](http://msdn.microsoft.com/en-us/library/office/ff863525%28v=office.14%29.aspx).
        */
        MsoDocProperties GetType();

        /**
        The document property type. Read-only for built-in document properties; read/write for custom document properties.

        [MSDN documentation for DocumentProperty.Type](http://msdn.microsoft.com/en-us/library/office/ff863525%28v=office.14%29.aspx).
        */
        void SetType(MsoDocProperties propertyType);

        /**
        The value of a document property.

        [MSDN documentation for DocumentProperty.Value](http://msdn.microsoft.com/en-us/library/office/ff861055%28v=office.14%29.aspx).
        */
        wxVariant GetValue();

        /**
        The value of a document property.

        [MSDN documentation for DocumentProperty.Value](http://msdn.microsoft.com/en-us/library/office/ff861055%28v=office.14%29.aspx).
        */
        void SetValue(const wxVariant& value);

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
