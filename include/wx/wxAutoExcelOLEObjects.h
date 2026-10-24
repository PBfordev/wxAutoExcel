/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// License:     MIT license
/////////////////////////////////////////////////////////////////////////////


#ifndef _WXAUTOEXCEL_OLEOBJECTS_H
#define _WXAUTOEXCEL_OLEOBJECTS_H

#include "wx/wxAutoExcel_defs.h"

#include "wx/wxAutoExcel_object.h"
#include "wx/wxAutoExcel_enums.h"

namespace wxAutoExcel {

    /**
    @brief Represents Microsoft Excel OLEObject object.
    */
   class WXDLLIMPEXP_WXAUTOEXCEL wxExcelOLEObject : public wxExcelObject
    {
    public:
        // ***** METHODS *****

        /**
        Activates the object.

        [MSDN documentation for OLEObject.Activate](https://learn.microsoft.com/en-us/office/vba/api/excel.oleobject.activate).
        */
        bool Activate();

        /**
        Brings the object to the front of the z-order.

        [MSDN documentation for OLEObject.BringToFront](https://learn.microsoft.com/en-us/office/vba/api/excel.oleobject.bringtofront).
        */
        bool BringToFront();

        /**
        Copies the object to the Clipboard.

        [MSDN documentation for OLEObject.Copy](https://learn.microsoft.com/en-us/office/vba/api/excel.oleobject.copy).
        */
        bool Copy();

        /**
        Copies the selected object to the Clipboard as a picture. Variant.

        [MSDN documentation for OLEObject.CopyPicture](https://learn.microsoft.com/en-us/office/vba/api/excel.oleobject.copypicture).
        */
        bool CopyPicture(XlPictureAppearance* appearance, XlCopyPictureFormat* format);

        /**
        Cuts the object to the Clipboard or pastes it into a specified destination.

        [MSDN documentation for OLEObject.Cut](https://learn.microsoft.com/en-us/office/vba/api/excel.oleobject.cut).
        */
        bool Cut();

        /**
        Deletes the object.

        [MSDN documentation for OLEObject.Delete](https://learn.microsoft.com/en-us/office/vba/api/excel.oleobject.delete).
        */
        bool Delete();

        /**
        Duplicates the object and returns a reference to the new copy.

        [MSDN documentation for OLEObject.Duplicate](https://learn.microsoft.com/en-us/office/vba/api/excel.oleobject.duplicate).
        */
        wxExcelObject Duplicate();

        /**
        Selects the object.

        [MSDN documentation for OLEObject.Select](https://learn.microsoft.com/en-us/office/vba/api/excel.oleobject.select).
        */
        bool Select(wxXlTribool replace = wxDefaultXlTribool);

        /**
        Sends the object to the back of the z-order.

        [MSDN documentation for OLEObject.SendToBack](https://learn.microsoft.com/en-us/office/vba/api/excel.oleobject.sendtoback).
        */
        bool SendToBack();

        /**
        Updates the link.

        [MSDN documentation for OLEObject.Update](https://learn.microsoft.com/en-us/office/vba/api/excel.oleobject.update).
        */
        bool Update();

        /**
        Sends a verb to the server of the specified OLE object.

        [MSDN documentation for OLEObject.Verb](https://learn.microsoft.com/en-us/office/vba/api/excel.oleobject.verb).
        */
        bool Verb(XlOLEVerb* verb = NULL);

        // ***** PROPERTIES *****

        /**
        True if the OLE object is automatically loaded when the workbook that contains it is opened.

        [MSDN documentation for OLEObject.AutoLoad](https://learn.microsoft.com/en-us/office/vba/api/excel.oleobject.autoload).
        */
        bool GetAutoLoad();

        /**
        True if the OLE object is automatically loaded when the workbook that contains it is opened.

        [MSDN documentation for OLEObject.AutoLoad](https://learn.microsoft.com/en-us/office/vba/api/excel.oleobject.autoload).
        */
        void SetAutoLoad(bool autoLoad);

        /**
        True if the OLE object is updated automatically when the source changes. Valid only if the object is linked (its OLEType property must be xlOLELink).

        [MSDN documentation for OLEObject.AutoUpdate](https://learn.microsoft.com/en-us/office/vba/api/excel.oleobject.autoupdate).
        */
        bool GetAutoUpdate();

        /**
        Returns a Border Represents the border of the object.

        [MSDN documentation for OLEObject.Border](https://learn.microsoft.com/en-us/office/vba/api/excel.oleobject.border).
        */
        wxExcelBorder GetBorder();

        /**
        Returns a Range Represents the cell that lies under the lower-right corner of the object.

        [MSDN documentation for OLEObject.BottomRightCell](https://learn.microsoft.com/en-us/office/vba/api/excel.oleobject.bottomrightcell).
        */
        wxExcelRange GetBottomRightCell();

        /**
        True if the object is enabled.

        [MSDN documentation for OLEObject.Enabled](https://learn.microsoft.com/en-us/office/vba/api/excel.oleobject.enabled).
        */
        bool GetEnabled();

        /**
        True if the object is enabled.

        [MSDN documentation for OLEObject.Enabled](https://learn.microsoft.com/en-us/office/vba/api/excel.oleobject.enabled).
        */
        void SetEnabled(bool enabled);

        /**
        Returns a Double value that represents the height, in points, of the object.

        [MSDN documentation for OLEObject.Height](https://learn.microsoft.com/en-us/office/vba/api/excel.oleobject.height).
        */
        double GetHeight();

        /**
        Sets a Double value that represents the height, in points, of the object.

        [MSDN documentation for OLEObject.Height](https://learn.microsoft.com/en-us/office/vba/api/excel.oleobject.height).
        */
        void SetHeight(double height);

        /**
        Returns a Long value that represents the index number of the object within the collection of similar objects.

        [MSDN documentation for OLEObject.Index](https://learn.microsoft.com/en-us/office/vba/api/excel.oleobject.index).
        */
        long GetIndex();

        /**
        Returns an Interior Represents the interior of the specified object.

        [MSDN documentation for OLEObject.Interior](https://learn.microsoft.com/en-us/office/vba/api/excel.oleobject.interior).
        */
        wxExcelInterior GetInterior();

        /**
        Returns a Double value that represents the distance, in points, from the left edge of the object to the left edge of column A (on a worksheet) or the left edge of the chart area (on a chart).

        [MSDN documentation for OLEObject.Left](https://learn.microsoft.com/en-us/office/vba/api/excel.oleobject.left).
        */
        double GetLeft();

        /**
        Sets a Double value that represents the distance, in points, from the left edge of the object to the left edge of column A (on a worksheet) or the left edge of the chart area (on a chart).

        [MSDN documentation for OLEObject.Left](https://learn.microsoft.com/en-us/office/vba/api/excel.oleobject.left).
        */
        void SetLeft(double left);

        /**
        Returns the worksheet range linked to the control's value. If you place a value in the cell, the control takes this value. Likewise, if you change the value of the control, that value is also placed in the cell.

        [MSDN documentation for OLEObject.LinkedCell](https://learn.microsoft.com/en-us/office/vba/api/excel.oleobject.linkedcell).
        */
        wxString GetLinkedCell();

        /**
        Sets the worksheet range linked to the control's value. If you place a value in the cell, the control takes this value. Likewise, if you change the value of the control, that value is also placed in the cell.

        [MSDN documentation for OLEObject.LinkedCell](https://learn.microsoft.com/en-us/office/vba/api/excel.oleobject.linkedcell).
        */
        void SetLinkedCell(const wxString& linkedCell);

        /**
        Returns the worksheet range used to fill the specified list box. Setting this property destroys any existing list in the list box.

        [MSDN documentation for OLEObject.ListFillRange](https://learn.microsoft.com/en-us/office/vba/api/excel.oleobject.listfillrange).
        */
        wxString GetListFillRange();

        /**
        Sets the worksheet range used to fill the specified list box. Setting this property destroys any existing list in the list box.

        [MSDN documentation for OLEObject.ListFillRange](https://learn.microsoft.com/en-us/office/vba/api/excel.oleobject.listfillrange).
        */
        void SetListFillRange(const wxString& listFillRange);

        /**
        Returns a Boolean value that indicates if the object is locked.

        [MSDN documentation for OLEObject.Locked](https://learn.microsoft.com/en-us/office/vba/api/excel.oleobject.locked).
        */
        bool GetLocked();

        /**
        Sets a Boolean value that indicates if the object is locked.

        [MSDN documentation for OLEObject.Locked](https://learn.microsoft.com/en-us/office/vba/api/excel.oleobject.locked).
        */
        void SetLocked(bool locked);

        /**
        Returns a String value representing the name of the object.

        [MSDN documentation for OLEObject.Name](https://learn.microsoft.com/en-us/office/vba/api/excel.oleobject.name).
        */
        wxString GetName();

        /**
        Sets a String value representing the name of the object.

        [MSDN documentation for OLEObject.Name](https://learn.microsoft.com/en-us/office/vba/api/excel.oleobject.name).
        */
        void SetName(const wxString& name);

        /**
        Returns the OLE object type. Can be one of the following XlOLEType constants: xlOLELink or xlOLEEmbed. Returns xlOLELink if the object is linked (it exists outside of the file), or returns xlOLEEmbed if the object is embedded (it's entirely contained within the file).

        [MSDN documentation for OLEObject.OLEType](https://learn.microsoft.com/en-us/office/vba/api/excel.oleobject.oletype).
        */
        XlOLEType GetOLEType();

        /**
        Returns a Variant value, containing an XlPlacement constant, that represents the way the object is attached to the cells below it.

        [MSDN documentation for OLEObject.Placement](https://learn.microsoft.com/en-us/office/vba/api/excel.oleobject.placement).
        */
        XlPlacement GetPlacement();

        /**
        Sets a Variant value, containing an XlPlacement constant, that represents the way the object is attached to the cells below it.

        [MSDN documentation for OLEObject.Placement](https://learn.microsoft.com/en-us/office/vba/api/excel.oleobject.placement).
        */
        void SetPlacement(XlPlacement placement);

        /**
        True if the object will be printed when the document is printed.

        [MSDN documentation for OLEObject.PrintObject](https://learn.microsoft.com/en-us/office/vba/api/excel.oleobject.printobject).
        */
        bool GetPrintObject();

        /**
        True if the object will be printed when the document is printed.

        [MSDN documentation for OLEObject.PrintObject](https://learn.microsoft.com/en-us/office/vba/api/excel.oleobject.printobject).
        */
        void SetPrintObject(bool printObject);

        /**
        Returns the programmatic identifiers for the object.

        [MSDN documentation for OLEObject.progID](https://learn.microsoft.com/en-us/office/vba/api/excel.oleobject.progid).
        */
        wxString GetprogID();

        /**
        Returns a Boolean value that determines if the object has a shadow.

        [MSDN documentation for OLEObject.Shadow](https://learn.microsoft.com/en-us/office/vba/api/excel.oleobject.shadow).
        */
        bool GetShadow();

        /**
        Sets a Boolean value that determines if the object has a shadow.

        [MSDN documentation for OLEObject.Shadow](https://learn.microsoft.com/en-us/office/vba/api/excel.oleobject.shadow).
        */
        void SetShadow(bool shadow);

#if WXAUTOEXCEL_USE_SHAPES

        /**
        Returns a ShapeRange Represents the specified object or objects.

        [MSDN documentation for OLEObject.ShapeRange](https://learn.microsoft.com/en-us/office/vba/api/excel.oleobject.shaperange).
        */
        wxExcelShapeRange GetShapeRange();

#endif // #if WXAUTOEXCEL_USE_SHAPES

        /**
        Returns a String value that represents the specified object's link source name.

        [MSDN documentation for OLEObject.SourceName](https://learn.microsoft.com/en-us/office/vba/api/excel.oleobject.sourcename).
        */
        wxString GetSourceName();

        /**
        Sets a String value that represents the specified object's link source name.

        [MSDN documentation for OLEObject.SourceName](https://learn.microsoft.com/en-us/office/vba/api/excel.oleobject.sourcename).
        */
        void SetSourceName(const wxString& sourceName);

        /**
        Returns a Double value that represents the distance, in points, from the top edge of the object to the top of row 1 (on a worksheet) or the top of the chart area (on a chart).

        [MSDN documentation for OLEObject.Top](https://learn.microsoft.com/en-us/office/vba/api/excel.oleobject.top).
        */
        double GetTop();

        /**
        Sets a Double value that represents the distance, in points, from the top edge of the object to the top of row 1 (on a worksheet) or the top of the chart area (on a chart).

        [MSDN documentation for OLEObject.Top](https://learn.microsoft.com/en-us/office/vba/api/excel.oleobject.top).
        */
        void SetTop(double top);

        /**
        Returns a Range Represents the cell that lies under the upper-left corner of the specified object.

        [MSDN documentation for OLEObject.TopLeftCell](https://learn.microsoft.com/en-us/office/vba/api/excel.oleobject.topleftcell).
        */
        wxExcelRange GetTopLeftCell();

        /**
        Returns a Boolean value that determines whether the object is visible.

        [MSDN documentation for OLEObject.Visible](https://learn.microsoft.com/en-us/office/vba/api/excel.oleobject.visible).
        */
        bool GetVisible();

        /**
        Sets a Boolean value that determines whether the object is visible.

        [MSDN documentation for OLEObject.Visible](https://learn.microsoft.com/en-us/office/vba/api/excel.oleobject.visible).
        */
        void SetVisible(bool visible);

        /**
        Returns a Double value that represents the width, in points, of the object.

        [MSDN documentation for OLEObject.Width](https://learn.microsoft.com/en-us/office/vba/api/excel.oleobject.width).
        */
        double GetWidth();

        /**
        Sets a Double value that represents the width, in points, of the object.

        [MSDN documentation for OLEObject.Width](https://learn.microsoft.com/en-us/office/vba/api/excel.oleobject.width).
        */
        void SetWidth(double width);

        /**
        Returns the z-order position of the object.

        [MSDN documentation for OLEObject.ZOrder](https://learn.microsoft.com/en-us/office/vba/api/excel.oleobject.zorder).
        */
        long GetZOrder();

        /**
        Returns "OLEObject".
        */
        virtual wxString GetAutoExcelObjectName_() const { return wxS("OLEObject"); }
    };

    /**
    @brief Represents Microsoft Excel OLEObjects collection.
    */
   class WXDLLIMPEXP_WXAUTOEXCEL wxExcelOLEObjects : public wxExcelObject
    {
    public:
        // ***** METHODS *****

        /**
        Adds a new OLE object to a sheet.

        [MSDN documentation for OLEObjects.Add](https://learn.microsoft.com/en-us/office/vba/api/excel.oleobjects.add).
        */
        wxExcelOLEObject Add(const wxString& classType = wxEmptyString, const wxString& filename = wxEmptyString,
                             double* height = NULL,  wxXlTribool link = wxDefaultXlTribool, wxXlTribool displayAsIcon = wxDefaultXlTribool,
                             const wxString& iconFileName = wxEmptyString, long* iconIndex = NULL,
                             const wxString& iconLabel = wxEmptyString, double* left = NULL, double* width = NULL);

        //@{
        /**
        Returns a single object from a collection.

        [MSDN documentation for OLEObjects.Item](https://learn.microsoft.com/en-us/office/vba/api/excel.oleobjects.item).
        */
        wxExcelOLEObject Item(long index);
        wxExcelOLEObject operator[](long index);
        wxExcelOLEObject Item(const wxString& name);
        wxExcelOLEObject operator[](const wxString& name);
        //@}

        /**
        Returns a Long value that represents the number of objects in the collection.

        [MSDN documentation for OLEObjects.Count](https://learn.microsoft.com/en-us/office/vba/api/excel.oleobjects.count).
        */
        long GetCount();

        /**
        The iterator over the items, see wxExcelCollectionIterator.
        */
        typedef wxExcelCollectionIterator<wxExcelOLEObjects, wxExcelOLEObject> iterator;

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
        Returns "OLEObjects".
        */
        virtual wxString GetAutoExcelObjectName_() const { return wxS("OLEObjects"); }
    };


} // namespace wxAutoExcel

#endif //_WXAUTOEXCEL_OLEOBJECTS_H
