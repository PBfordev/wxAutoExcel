/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// License:     MIT license
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

        [MSDN documentation for DocumentProperty.Delete](https://learn.microsoft.com/en-us/office/vba/api/office.documentproperty.delete).
        */
        void Delete();

        // ***** PROPERTIES *****

        /**
        Gets the source of a linked custom document property.

        [MSDN documentation for DocumentProperty.LinkSource](https://learn.microsoft.com/en-us/office/vba/api/office.documentproperty.linksource).
        */
        wxString GetLinkSource();
        /**
        Sets the source of a linked custom document property.

        [MSDN documentation for DocumentProperty.LinkSource](https://learn.microsoft.com/en-us/office/vba/api/office.documentproperty.linksource).
        */
        void SetLinkSource(const wxString& linkSource);

        /**
        Is true if the value of the custom document property is linked to the content of the container document. False if the value is static.

        [MSDN documentation for DocumentProperty.LinkToContent](https://learn.microsoft.com/en-us/office/vba/api/office.documentproperty.linktocontent).
        */
        bool GetLinkToContent();
        /**
        Is true if the value of the custom document property is linked to the content of the container document. False if the value is static.

        [MSDN documentation for DocumentProperty.LinkToContent](https://learn.microsoft.com/en-us/office/vba/api/office.documentproperty.linktocontent).
        */
        void SetLinkToContent(bool linkToContent);

        /**
        The name of a document property.

        [MSDN documentation for DocumentProperty.Name](https://learn.microsoft.com/en-us/office/vba/api/office.documentproperty.name).
        */
        wxString GetName();
        /**
        The name of a document property.

        [MSDN documentation for DocumentProperty.Name](https://learn.microsoft.com/en-us/office/vba/api/office.documentproperty.name).
        */
        void SetName(const wxString& name);

        /**
        The document property type. Read-only for built-in document properties; read/write for custom document properties.

        [MSDN documentation for DocumentProperty.Type](https://learn.microsoft.com/en-us/office/vba/api/office.documentproperty.type).
        */
        MsoDocProperties GetType();

        /**
        The document property type. Read-only for built-in document properties; read/write for custom document properties.

        [MSDN documentation for DocumentProperty.Type](https://learn.microsoft.com/en-us/office/vba/api/office.documentproperty.type).
        */
        void SetType(MsoDocProperties propertyType);

        /**
        The value of a document property.

        [MSDN documentation for DocumentProperty.Value](https://learn.microsoft.com/en-us/office/vba/api/office.documentproperty.value).
        */
        wxVariant GetValue();

        /**
        The value of a document property.

        [MSDN documentation for DocumentProperty.Value](https://learn.microsoft.com/en-us/office/vba/api/office.documentproperty.value).
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

        [MSDN documentation for DocumentProperties.Add](https://learn.microsoft.com/en-us/office/vba/api/office.documentproperties.add).
        */

        wxExcelDocumentProperty Add(const wxString& name, bool linkToContent, MsoDocProperties propertyType,
                                    wxVariant* value = NULL, const wxString& linkSource = wxEmptyString);

        // ***** PROPERTIES *****

        /**
        Returns the number of items in the collection.

        [MSDN documentation for DocumentProperties.Count](https://learn.microsoft.com/en-us/office/vba/api/office.documentproperties.count).
        */
        long GetCount();

         //@{
        /**
        Returns a single DocumentProperty object from the collection.

        [MSDN documentation for DocumentProperties.Item](https://learn.microsoft.com/en-us/office/vba/api/office.documentproperties.item).
        */
        wxExcelDocumentProperty GetItem(long index);
        wxExcelDocumentProperty GetItem(const wxString& name);
        wxExcelDocumentProperty operator[](long index);
        wxExcelDocumentProperty operator[](const wxString& name);
        //@}

        /**
        The iterator over the items, see wxExcelCollectionIterator.
        */
        typedef wxExcelCollectionIterator<wxExcelDocumentProperties, wxExcelDocumentProperty> iterator;

        //@{
        /**
        Returns the iterator referring to the first item, or past the last item,
        which allows iterating over the items with a range-based for loop,
        see wxExcelCollectionIterator.
        */
        iterator begin() { return iterator(this, 1); }
        iterator end() { return iterator(this, GetCount() + 1); }
        //@}

        /**
        Returns "DocumentProperties".
        */
        virtual wxString GetAutoExcelObjectName_() const { return wxS("DocumentProperties"); }
    };


} // namespace wxAutoExcel


#endif //_WXAUTOEXCEL_DOCUMENTPROPERTIES_H
