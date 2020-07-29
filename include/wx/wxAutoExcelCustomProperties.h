/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Modified by:
// Copyright:   (c) 2020 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
/////////////////////////////////////////////////////////////////////////////

#ifndef _WXAUTOEXCEL_CUSTOMPROPERTIES_H
#define _WXAUTOEXCEL_CUSTOMPROPERTIES_H

#include "wx/wxAutoExcel_defs.h"

#include "wx/wxAutoExcel_object.h"

namespace wxAutoExcel
{

/**
@brief Represents identifier information. Identifier information can be used as metadata for XML.

*/
class WXDLLIMPEXP_WXAUTOEXCEL wxExcelCustomProperty : public wxExcelObject
{
public:
    // ***** METHODS *****

    /**
    Deletes the object.

    [Excel VBA documentation for CustomProperty.Delete](https://docs.microsoft.com/en-us/office/vba/api/excel.customproperty.delete)
    */
    void Delete();

    // ***** PROPERTIES *****

    /**
    Returns a String value representing the name of the object.

    [Excel VBA documentation for CustomProperty.Name](https://docs.microsoft.com/en-us/office/vba/api/excel.customproperty.name)
    */
    wxString GetName();

    /**
    Sets a String value representing the name of the object.

    [Excel VBA documentation for CustomProperty.Name](https://docs.microsoft.com/en-us/office/vba/api/excel.customproperty.name)
    */
    void SetName(const wxString& name);

    /**
    Returns the value of the property.

    [Excel VBA documentation for CustomProperty.Value](https://docs.microsoft.com/en-us/office/vba/api/excel.customproperty.value)
    */
    wxVariant GetValue();

    /**
    Sets the value of the property.

    [Excel VBA documentation for CustomProperty.Value](https://docs.microsoft.com/en-us/office/vba/api/excel.customproperty.value)
    */
    void SetValue(const wxVariant& value);

    /**
    Returns "CustomProperty".
    */
    virtual wxString GetAutoExcelObjectName_() const { return wxS("CustomProperty"); }

}; // class wxExcelCustomProperty

/**
    @brief Represents a collection of CustomProperty objects that represent additional information. The information can be used as metadata for XML.

*/
class WXDLLIMPEXP_WXAUTOEXCEL wxExcelCustomProperties : public wxExcelObject
{
public:
    // ***** METHODS *****

    /**
       Adds custom property information.

        [MSDN documentation for CustomProperties.Add](https://docs.microsoft.com/en-us/office/vba/api/excel.customproperties.add)
    */
    wxExcelCustomProperty Add(const wxString& name, const wxVariant& value);

    // ***** PROPERTIES *****

    /**
        Returns the number of objects in the collection.
    */
    long GetCount();

    //@{
    /**
        Returns the CustomProperty with the given index or name.

        [MSDN documentation for CustomProperties.Item](https://docs.microsoft.com/en-us/office/vba/api/excel.customproperties.item)
    */
    wxExcelCustomProperty GetItem(long index);
    wxExcelCustomProperty GetItem(const wxString& name);
    wxExcelCustomProperty operator[](long index);
    wxExcelCustomProperty operator[](const wxString& name);
    //@}

    /**
    Returns "CustomProperties".
    */
    virtual wxString GetAutoExcelObjectName_() const { return wxS("CustomProperties"); }

}; // class wxExcelCustomProperties


} // namespace wxAutoExcel 

#endif // #ifndef _WXAUTOEXCEL_CUSTOMPROPERTIES_H
