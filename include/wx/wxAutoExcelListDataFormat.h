/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Copyright:   (c) 2020 PB <pbfordev@gmail.com>
// License:     MIT license
/////////////////////////////////////////////////////////////////////////////

#ifndef _WXAUTOEXCEL_LISTDATAFORMAT_H
#define _WXAUTOEXCEL_LISTDATAFORMAT_H

#include "wx/wxAutoExcel_defs.h"

#include "wx/wxAutoExcel_object.h"

namespace wxAutoExcel
{

/**
@brief Represents an object holding all data type properties for ListColumn.
*/
class WXDLLIMPEXP_WXAUTOEXCEL wxExcelListDataFormat : public wxExcelObject
{
public:
    // ***** PROPERTIES *****

    /**
    Returns a Boolean value indicating whether users can provide their own data for cells in a column (rather than being restricted to a list of values) for those columns that supply a list of values. Returns False for lists that are not linked to a SharePoint site. Also returns False if the column is not a specified as choice or multi-choice.

    [MSDN documentation for ListDataFormat.AllowFillIn](https://docs.microsoft.com/en-us/office/vba/api/excel.listdataformat.allowfillin)
    */
    bool GetAllowFillIn();

    /**
    Returns an Array of String values that contains the choices offered to the user by the ListLookUp, ChoiceMulti, and Choice data types of the DefaultValue property.

    [MSDN documentation for ListDataFormat.Choices](https://docs.microsoft.com/en-us/office/vba/api/excel.listdataformat.choices)
    */
    wxArrayString GetChoices();

    /**
    Returns a Long value that represents the number of decimal places to show for the numbers in the ListColumn object.

    [MSDN documentation for ListDataFormat.DecimalPlaces](https://docs.microsoft.com/en-us/office/vba/api/excel.listdataformat.decimalplaces)
    */
    long GetDecimalPlaces();

    /**
    Returns Variant representing the default data type value for a new row in a column. The Nothing object is returned when the schema does not specify a default value.

    [MSDN documentation for ListDataFormat.DefaultValue](https://docs.microsoft.com/en-us/office/vba/api/excel.listdataformat.defaultvalue)
    */
    wxVariant GetDefaultValue();

    /**
    Returns a Boolean value. Returns True only if the number data for the ListColumn object will be shown in percentage formatting.

    [MSDN documentation for ListDataFormat.IsPercent](https://docs.microsoft.com/en-us/office/vba/api/excel.listdataformat.ispercent)
    */
    bool GetIsPercent();

    /**
    Returns a Long value that represents the LCID for the ListColumn object that is specified in the schema definition.

    [MSDN documentation for ListDataFormat.lcid](https://docs.microsoft.com/en-us/office/vba/api/excel.listdataformat.lcid)
    */
    long Getlcid();

    /**
    Returns a Long containing the maximum number of characters allowed in the ListColumn object if the Type property is set to xlListDataTypeText or xlListDataTypeMultiLineText.

    [MSDN documentation for ListDataFormat.MaxCharacters](https://docs.microsoft.com/en-us/office/vba/api/excel.listdataformat.maxcharacters)
    */
    long GetMaxCharacters();

    /**
    Returns a Variant containing the maximum value allowed in this field in the list column.

    [MSDN documentation for ListDataFormat.MaxNumber](https://docs.microsoft.com/en-us/office/vba/api/excel.listdataformat.maxnumber)
    */
    wxVariant GetMaxNumber();

    /**
    Returns a Variant containing the minimum value allowed in this field in the list column. This can be a negative floating point number.

    [MSDN documentation for ListDataFormat.MinNumber](https://docs.microsoft.com/en-us/office/vba/api/excel.listdataformat.minnumber)
    */
    wxVariant GetMinNumber();

    /**
    Returns True if the object has been opened as read-only.

    [MSDN documentation for ListDataFormat.ReadOnly](https://docs.microsoft.com/en-us/office/vba/api/excel.listdataformat.readonly)
    */
    bool GetReadOnly();

    /**
    Returns a Boolean value indicating whether the schema definition of a column requires data before the row is committed.

    [MSDN documentation for ListDataFormat.Required](https://docs.microsoft.com/en-us/office/vba/api/excel.listdataformat.required)
    */
    bool GetRequired();

    /**
    Returns an XlListDataType value that represents the data type of the list column.

    [MSDN documentation for ListDataFormat.Type](https://docs.microsoft.com/en-us/office/vba/api/excel.listdataformat.type)
    */
    XlListDataType GetType();

    /**
    Returns "ListDataFormat".
    */
    virtual wxString GetAutoExcelObjectName_() const { return wxS("ListDataFormat"); }

}; // class wxExcelListDataFormat

} // namespace wxAutoExcel

#endif // #ifndef _WXAUTOEXCEL_LISTDATAFORMAT_H
