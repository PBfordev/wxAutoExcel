/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Modified by:
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
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

        [MSDN documentation for OLEObject.Activate](http://msdn.microsoft.com/en-us/library/bb211894).
        */
        bool Activate();

        /**
        Brings the object to the front of the z-order.

        [MSDN documentation for OLEObject.BringToFront](http://msdn.microsoft.com/en-us/library/bb211899).
        */
        bool BringToFront();

        /**
        Copies the object to the Clipboard.

        [MSDN documentation for OLEObject.Copy](http://msdn.microsoft.com/en-us/library/bb211903).
        */
        bool Copy();

        /**
        Copies the selected object to the Clipboard as a picture. Variant.

        [MSDN documentation for OLEObject.CopyPicture](http://msdn.microsoft.com/en-us/library/bb211908).
        */
        bool CopyPicture(XlPictureAppearance* appearance, XlCopyPictureFormat* format);

        /**
        Cuts the object to the Clipboard or pastes it into a specified destination.

        [MSDN documentation for OLEObject.Cut](http://msdn.microsoft.com/en-us/library/bb211910).
        */
        bool Cut();

        /**
        Deletes the object.

        [MSDN documentation for OLEObject.Delete](http://msdn.microsoft.com/en-us/library/bb211914).
        */
        bool Delete();

        /**
        Duplicates the object and returns a reference to the new copy.

        [MSDN documentation for OLEObject.Duplicate](http://msdn.microsoft.com/en-us/library/bb211917).
        */
        wxExcelObject Duplicate();

        /**
        Selects the object.

        [MSDN documentation for OLEObject.Select](http://msdn.microsoft.com/en-us/library/bb238011).
        */
        bool Select(wxXlTribool replace = wxDefaultXlTribool);

        /**
        Sends the object to the back of the z-order.

        [MSDN documentation for OLEObject.SendToBack](http://msdn.microsoft.com/en-us/library/bb238016).
        */
        bool SendToBack();

        /**
        Updates the link.

        [MSDN documentation for OLEObject.Update](http://msdn.microsoft.com/en-us/library/bb238039).
        */
        bool Update();

        /**
        Sends a verb to the server of the specified OLE object.

        [MSDN documentation for OLEObject.Verb](http://msdn.microsoft.com/en-us/library/bb238045).
        */
        bool Verb(XlOLEVerb* verb = NULL);

        // ***** PROPERTIES *****

        /**
        True if the OLE object is automatically loaded when the workbook that contains it is opened.

        [MSDN documentation for OLEObject.AutoLoad](http://msdn.microsoft.com/en-us/library/bb237188).
        */
        bool GetAutoLoad();

        /**
        True if the OLE object is automatically loaded when the workbook that contains it is opened.

        [MSDN documentation for OLEObject.AutoLoad](http://msdn.microsoft.com/en-us/library/bb237188).
        */
        void SetAutoLoad(bool autoLoad);

        /**
        True if the OLE object is updated automatically when the source changes. Valid only if the object is linked (its OLEType property must be xlOLELink).

        [MSDN documentation for OLEObject.AutoUpdate](http://msdn.microsoft.com/en-us/library/bb237192).
        */
        bool GetAutoUpdate();

        /**
        Returns a Border Represents the border of the object.

        [MSDN documentation for OLEObject.Border](http://msdn.microsoft.com/en-us/library/bb237195).
        */
        wxExcelBorder GetBorder();

        /**
        Returns a Range Represents the cell that lies under the lower-right corner of the object.

        [MSDN documentation for OLEObject.BottomRightCell](http://msdn.microsoft.com/en-us/library/bb237198).
        */
        wxExcelRange GetBottomRightCell();

        /**
        True if the object is enabled.

        [MSDN documentation for OLEObject.Enabled](http://msdn.microsoft.com/en-us/library/bb237200).
        */
        bool GetEnabled();

        /**
        True if the object is enabled.

        [MSDN documentation for OLEObject.Enabled](http://msdn.microsoft.com/en-us/library/bb237200).
        */
        void SetEnabled(bool enabled);

        /**
        Returns a Double value that represents the height, in points, of the object.

        [MSDN documentation for OLEObject.Height](http://msdn.microsoft.com/en-us/library/bb237202).
        */
        double GetHeight();

        /**
        Sets a Double value that represents the height, in points, of the object.

        [MSDN documentation for OLEObject.Height](http://msdn.microsoft.com/en-us/library/bb237202).
        */
        void SetHeight(double height);

        /**
        Returns a Long value that represents the index number of the object within the collection of similar objects.

        [MSDN documentation for OLEObject.Index](http://msdn.microsoft.com/en-us/library/bb237205).
        */
        long GetIndex();

        /**
        Returns an Interior Represents the interior of the specified object.

        [MSDN documentation for OLEObject.Interior](http://msdn.microsoft.com/en-us/library/bb237208).
        */
        wxExcelInterior GetInterior();

        /**
        Returns a Double value that represents the distance, in points, from the left edge of the object to the left edge of column A (on a worksheet) or the left edge of the chart area (on a chart).

        [MSDN documentation for OLEObject.Left](http://msdn.microsoft.com/en-us/library/bb237212).
        */
        double GetLeft();

        /**
        Sets a Double value that represents the distance, in points, from the left edge of the object to the left edge of column A (on a worksheet) or the left edge of the chart area (on a chart).

        [MSDN documentation for OLEObject.Left](http://msdn.microsoft.com/en-us/library/bb237212).
        */
        void SetLeft(double left);

        /**
        Returns the worksheet range linked to the control's value. If you place a value in the cell, the control takes this value. Likewise, if you change the value of the control, that value is also placed in the cell.

        [MSDN documentation for OLEObject.LinkedCell](http://msdn.microsoft.com/en-us/library/bb237214).
        */
        wxString GetLinkedCell();

        /**
        Sets the worksheet range linked to the control's value. If you place a value in the cell, the control takes this value. Likewise, if you change the value of the control, that value is also placed in the cell.

        [MSDN documentation for OLEObject.LinkedCell](http://msdn.microsoft.com/en-us/library/bb237214).
        */
        void SetLinkedCell(const wxString& linkedCell);

        /**
        Returns the worksheet range used to fill the specified list box. Setting this property destroys any existing list in the list box.

        [MSDN documentation for OLEObject.ListFillRange](http://msdn.microsoft.com/en-us/library/bb237217).
        */
        wxString GetListFillRange();

        /**
        Sets the worksheet range used to fill the specified list box. Setting this property destroys any existing list in the list box.

        [MSDN documentation for OLEObject.ListFillRange](http://msdn.microsoft.com/en-us/library/bb237217).
        */
        void SetListFillRange(const wxString& listFillRange);

        /**
        Returns a Boolean value that indicates if the object is locked.

        [MSDN documentation for OLEObject.Locked](http://msdn.microsoft.com/en-us/library/bb237219).
        */
        bool GetLocked();

        /**
        Sets a Boolean value that indicates if the object is locked.

        [MSDN documentation for OLEObject.Locked](http://msdn.microsoft.com/en-us/library/bb237219).
        */
        void SetLocked(bool locked);

        /**
        Returns a String value representing the name of the object.

        [MSDN documentation for OLEObject.Name](http://msdn.microsoft.com/en-us/library/bb237220).
        */
        wxString GetName();

        /**
        Sets a String value representing the name of the object.

        [MSDN documentation for OLEObject.Name](http://msdn.microsoft.com/en-us/library/bb237220).
        */
        void SetName(const wxString& name);        

        /**
        Returns the OLE object type. Can be one of the following XlOLEType constants: xlOLELink or xlOLEEmbed. Returns xlOLELink if the object is linked (it exists outside of the file), or returns xlOLEEmbed if the object is embedded (it's entirely contained within the file).

        [MSDN documentation for OLEObject.OLEType](http://msdn.microsoft.com/en-us/library/bb208870).
        */
        XlOLEType GetOLEType();

        /**
        Returns a Variant value, containing an XlPlacement constant, that represents the way the object is attached to the cells below it.

        [MSDN documentation for OLEObject.Placement](http://msdn.microsoft.com/en-us/library/bb237225).
        */
        XlPlacement GetPlacement();

        /**
        Sets a Variant value, containing an XlPlacement constant, that represents the way the object is attached to the cells below it.

        [MSDN documentation for OLEObject.Placement](http://msdn.microsoft.com/en-us/library/bb237225).
        */
        void SetPlacement(XlPlacement placement);

        /**
        True if the object will be printed when the document is printed.

        [MSDN documentation for OLEObject.PrintObject](http://msdn.microsoft.com/en-us/library/bb237226).
        */
        bool GetPrintObject();

        /**
        True if the object will be printed when the document is printed.

        [MSDN documentation for OLEObject.PrintObject](http://msdn.microsoft.com/en-us/library/bb237226).
        */
        void SetPrintObject(bool printObject);

        /**
        Returns the programmatic identifiers for the object.

        [MSDN documentation for OLEObject.progID](http://msdn.microsoft.com/en-us/library/bb237228).
        */
        wxString GetprogID();

        /**
        Returns a Boolean value that determines if the object has a shadow.

        [MSDN documentation for OLEObject.Shadow](http://msdn.microsoft.com/en-us/library/bb214886).
        */
        bool GetShadow();

        /**
        Sets a Boolean value that determines if the object has a shadow.

        [MSDN documentation for OLEObject.Shadow](http://msdn.microsoft.com/en-us/library/bb214886).
        */
        void SetShadow(bool shadow);

#if WXAUTOEXCEL_USE_SHAPES

        /**
        Returns a ShapeRange Represents the specified object or objects.

        [MSDN documentation for OLEObject.ShapeRange](http://msdn.microsoft.com/en-us/library/bb214890).
        */
        wxExcelShapeRange GetShapeRange();

#endif // #if WXAUTOEXCEL_USE_SHAPES

        /**
        Returns a String value that represents the specified object's link source name.

        [MSDN documentation for OLEObject.SourceName](http://msdn.microsoft.com/en-us/library/bb214893).
        */
        wxString GetSourceName();

        /**
        Sets a String value that represents the specified object's link source name.

        [MSDN documentation for OLEObject.SourceName](http://msdn.microsoft.com/en-us/library/bb214893).
        */
        void SetSourceName(const wxString& sourceName);

        /**
        Returns a Double value that represents the distance, in points, from the top edge of the object to the top of row 1 (on a worksheet) or the top of the chart area (on a chart).

        [MSDN documentation for OLEObject.Top](http://msdn.microsoft.com/en-us/library/bb214920).
        */
        double GetTop();

        /**
        Sets a Double value that represents the distance, in points, from the top edge of the object to the top of row 1 (on a worksheet) or the top of the chart area (on a chart).

        [MSDN documentation for OLEObject.Top](http://msdn.microsoft.com/en-us/library/bb214920).
        */
        void SetTop(double top);

        /**
        Returns a Range Represents the cell that lies under the upper-left corner of the specified object.

        [MSDN documentation for OLEObject.TopLeftCell](http://msdn.microsoft.com/en-us/library/bb214922).
        */
        wxExcelRange GetTopLeftCell();

        /**
        Returns a Boolean value that determines whether the object is visible.

        [MSDN documentation for OLEObject.Visible](http://msdn.microsoft.com/en-us/library/bb214924).
        */
        bool GetVisible();

        /**
        Sets a Boolean value that determines whether the object is visible.

        [MSDN documentation for OLEObject.Visible](http://msdn.microsoft.com/en-us/library/bb214924).
        */
        void SetVisible(bool visible);

        /**
        Returns a Double value that represents the width, in points, of the object.

        [MSDN documentation for OLEObject.Width](http://msdn.microsoft.com/en-us/library/bb214926).
        */
        double GetWidth();

        /**
        Sets a Double value that represents the width, in points, of the object.

        [MSDN documentation for OLEObject.Width](http://msdn.microsoft.com/en-us/library/bb214926).
        */
        void SetWidth(double width);

        /**
        Returns the z-order position of the object.

        [MSDN documentation for OLEObject.ZOrder](http://msdn.microsoft.com/en-us/library/bb214927).
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

        [MSDN documentation for OLEObjects.Add](http://msdn.microsoft.com/en-us/library/bb211920).
        */
        wxExcelOLEObject Add(const wxString& classType = wxEmptyString, const wxString& filename = wxEmptyString, 
                             double* height = NULL,  wxXlTribool link = wxDefaultXlTribool, wxXlTribool displayAsIcon = wxDefaultXlTribool, 
                             const wxString& iconFileName = wxEmptyString, long* iconIndex = NULL, 
                             const wxString& iconLabel = wxEmptyString, double* left = NULL, double* width = NULL);

        //@{
        /**
        Returns a single object from a collection.

        [MSDN documentation for OLEObjects.Item](http://msdn.microsoft.com/en-us/library/bb211942).
        */
        wxExcelOLEObject Item(long index);
        wxExcelOLEObject operator[](long index);
        wxExcelOLEObject Item(const wxString& name);
        wxExcelOLEObject operator[](const wxString& name);
        //@}
        
        /**
        Returns a Long value that represents the number of objects in the collection.

        [MSDN documentation for OLEObjects.Count](http://msdn.microsoft.com/en-us/library/bb213196).
        */
        long GetCount();

        /**
        Returns "OLEObjects".
        */
        virtual wxString GetAutoExcelObjectName_() const { return wxS("OLEObjects"); }
    };


} // namespace wxAutoExcel

#endif //_WXAUTOEXCEL_OLEOBJECTS_H
