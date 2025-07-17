/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Copyright:   (c) 2020 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
/////////////////////////////////////////////////////////////////////////////

#ifndef _WXAUTOEXCEL_TABLESTYLEELEMENT_H
#define _WXAUTOEXCEL_TABLESTYLEELEMENT_H

#include "wx/wxAutoExcel_defs.h"

#include "wx/wxAutoExcel_object.h"

namespace wxAutoExcel
{

/**
@brief Represents a table style element.
*/
class WXDLLIMPEXP_WXAUTOEXCEL wxExcelTableStyleElement : public wxExcelObject
{
public:
    // ***** METHODS *****

    /**
    Clears the formatting.

    [MSDN documentation for TableStyleElement.Clear](https://docs.microsoft.com/en-us/office/vba/api/excel.tablestyleelement.clear)
    */
    void Clear();

    // ***** PROPERTIES *****

    /**
    Returns a Borders collection that represents the borders of a table style element.

    [MSDN documentation for TableStyleElement.Borders](https://docs.microsoft.com/en-us/office/vba/api/excel.tablestyleelement.borders)
    */
    wxExcelBorders GetBorders();

    /**
    Returns a Font object that represents the font of the specified object.

    [MSDN documentation for TableStyleElement.Font](https://docs.microsoft.com/en-us/office/vba/api/excel.tablestyleelement.font)
    */
    wxExcelFont GetFont();

    /**
    Returns whether a table style element has formatting applied to the specified element.

    [MSDN documentation for TableStyleElement.HasFormat](https://docs.microsoft.com/en-us/office/vba/api/excel.tablestyleelement.hasformat)
    */
    bool GetHasFormat();

    /**
    Returns an Interior object that represents the interior of the specified object.

    [MSDN documentation for TableStyleElement.Interior](https://docs.microsoft.com/en-us/office/vba/api/excel.tablestyleelement.interior)
    */
    wxExcelInterior GetInterior();

    /**
    Returns the size of banding.

    [MSDN documentation for TableStyleElement.StripeSize](https://docs.microsoft.com/en-us/office/vba/api/excel.tablestyleelement.stripesize)
    */
    long GetStripeSize();

    /**
    Sets the size of banding.

    [MSDN documentation for TableStyleElement.StripeSize](https://docs.microsoft.com/en-us/office/vba/api/excel.tablestyleelement.stripesize)
    */
    void SetStripeSize(long stripeSize);

    /**
    Returns "TableStyleElement".
    */
    virtual wxString GetAutoExcelObjectName_() const { return wxS("TableStyleElement"); }

}; // class wxExcelTableStyleElement


/**
    @brief Represents a collection of TableStyleElement objects.
*/
class WXDLLIMPEXP_WXAUTOEXCEL wxExcelTableStyleElements : public wxExcelObject
{
public:
    // ***** METHODS *****

    //@{
    /**
        Returns the TableStyleElement with the given index.

        [MSDN documentation for TableStyleElements.Item](https://docs.microsoft.com/en-us/office/vba/api/excel.tablestyleelements.item)
    */
    wxExcelTableStyleElement GetItem(XlTableStyleElementType index);
    wxExcelTableStyleElement operator[](XlTableStyleElementType index);
    //@}

    // ***** PROPERTIES *****

    /**
        Returns the number of objects in the collection.

        [MSDN documentation for TableStyleElements.Count](https://docs.microsoft.com/en-us/office/vba/api/excel.tablestyleelements.count)
    */
    long GetCount();


    /**
    Returns "TableStyleElements".
    */
    virtual wxString GetAutoExcelObjectName_() const { return wxS("TableStyleElements"); }

}; // class wxExcelTableStyleElements


} // namespace wxAutoExcel

#endif // #ifndef _WXAUTOEXCEL_TABLESTYLEELEMENT_H
