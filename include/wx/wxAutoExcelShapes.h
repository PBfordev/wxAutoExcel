/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// License:     MIT license
/////////////////////////////////////////////////////////////////////////////


#ifndef _WXAUTOEXCEL_SHAPES_H
#define _WXAUTOEXCEL_SHAPES_H

#include "wx/wxAutoExcel_defs.h"


#if WXAUTOEXCEL_USE_SHAPES

#include <wx/geometry.h>

#include "wx/wxAutoExcel_object.h"
#include "wx/wxAutoExcel_enums.h"


namespace wxAutoExcel {


    /**
    @brief Represents Microsoft Excel Shape object.
    */
   class WXDLLIMPEXP_WXAUTOEXCEL wxExcelShapes : public wxExcelObject
    {
    public:
        // ***** METHODS *****

         /**
        Creates a 3D model from an existing file.

        [Excel VBA documentation for Shapes.Add3DModel](https://learn.microsoft.com/en-us/office/vba/api/excel.shapes.add3dmodel)
        */
        wxExcelShape Add3DModel(const wxString& fileName, 
                                wxXlTribool linkToFile = wxDefaultXlTribool,
                                wxXlTribool saveWithDocument = wxDefaultXlTribool, 
                                double* left = NULL, double* top = NULL,
                                double* width = NULL, double* height = NULL);


        /**
        Creates a borderless line callout. Returns a Shape that represents the new callout.

        [MSDN documentation for Shapes.AddCallout](https://learn.microsoft.com/en-us/office/vba/api/excel.shapes.addcallout).
        */
        wxExcelShape AddCallout(MsoCalloutType type, double left, double top,
                                double width, double height);

#if WXAUTOEXCEL_USE_CHARTS
        /**
        Creates a chart at the specified location on the active sheet.
        For Excel 2013 and newer it is recommended to use AddChart2 instead.

        [MSDN documentation for Shapes.AddChart](https://learn.microsoft.com/en-us/previous-versions/office/developer/office-2007/bb238877%28v=office.12%29).
        */
        wxExcelShape AddChart(XlChartType type, double left, double top,
                              double width, double height);

        /**
        Adds a chart to the document. Returns a Shape object that represents a chart and adds it to the specified collection.
        Since Excel 2013.

        [MSDN documentation for Shapes.AddChart2](https://learn.microsoft.com/en-us/office/vba/api/excel.shapes.addchart2).
        */
        wxExcelShape AddChart2(long style = -1, XlChartType* type = NULL,
            double* left = NULL, double* top = NULL, double* width = NULL, double* height = NULL,
            wxXlTribool newLayout = wxDefaultXlTribool);
#endif // #if WXAUTOEXCEL_USE_CHARTS

        /**
        Creates a connector. Returns a Shape that represents the new connector. When a connector is added, it's not connected to anything. Use the BeginConnect and EndConnect methods to attach the beginning and end of a connector to other shapes in the document.

        [MSDN documentation for Shapes.AddConnector](https://learn.microsoft.com/en-us/office/vba/api/excel.shapes.addconnector).
        */
        wxExcelShape AddConnector(MsoConnectorType type, double beginX, double beginY,
                                  double endX, double endY);



        /**
        Returns a Shape that represents a Bézier curve in a worksheet.

        [MSDN documentation for Shapes.AddCurve](https://learn.microsoft.com/en-us/office/vba/api/excel.shapes.addcurve).
        */
        wxExcelShape AddCurve(const wxVector<wxPoint2DDouble>& points);

        /**
        Creates a Microsoft Excel control. Returns a Shape that represents the new control.

        [MSDN documentation for Shapes.AddFormControl](https://learn.microsoft.com/en-us/office/vba/api/excel.shapes.addformcontrol).
        */
        wxExcelShape AddFormControl(XlFormControl type, double left, double top,
                                    double width, double height);

        /**
        Creates a label. Returns a Shape that represents the new label.

        [MSDN documentation for Shapes.AddLabel](https://learn.microsoft.com/en-us/office/vba/api/excel.shapes.addlabel).
        */
        wxExcelShape AddLabel(MsoTextOrientation orientation, double left, double top,
                              double width, double height);

        /**
        As it applies to the Shapes object, Returns a Shape that represents the new line in a worksheet.

        [MSDN documentation for Shapes.AddLine](https://learn.microsoft.com/en-us/office/vba/api/excel.shapes.addline).
        */
        wxExcelShape AddLine(double beginX, double beginY, double endX, double endY);

        /**
        Creates an OLE object. Returns a Shape that represents the new OLE object.

        [MSDN documentation for Shapes.AddOLEObject](https://learn.microsoft.com/en-us/office/vba/api/excel.shapes.addoleobject).
        */
        wxExcelShape AddOLEObject(const wxString& classType = wxEmptyString, const wxString& filename = wxEmptyString,
                                  wxXlTribool link = wxDefaultXlTribool, wxXlTribool displayAsIcon = wxDefaultXlTribool,
                                  const wxString& iconFileName = wxEmptyString, long* iconIndex = NULL,
                                  const wxString& iconLabel = wxEmptyString,
                                  double* left = NULL, double* top = NULL, double* width = NULL, double* height = NULL);

        /**
        Creates a picture from an existing file. Returns a Shape that represents the new picture.

        [MSDN documentation for Shapes.AddPicture](https://learn.microsoft.com/en-us/office/vba/api/excel.shapes.addpicture).
        */
        wxExcelShape AddPicture(const wxString& fileName, MsoTriState linkToFile,
                                MsoTriState saveWithDocument, double left, double top,
                                double width, double height);

        /**
        Creates a picture from an existing file. Returns a Shape that represents the new picture.

        [MSDN documentation for Shapes.AddPicture2](https://learn.microsoft.com/en-us/office/vba/api/excel.shapes.addpicture2).
        */
        wxExcelShape AddPicture2(const wxString& fileName, MsoTriState linkToFile,
                                 MsoTriState saveWithDocument, double left, double top,
                                 double* width = NULL, double* height = NULL,
                                 MsoPictureCompress* compress = NULL);

        /**
        Creates an open polyline or a closed polygon drawing. Returns a Shape that represents the new polyline or polygon.

        [MSDN documentation for Shapes.AddPolyline](https://learn.microsoft.com/en-us/office/vba/api/excel.shapes.addpolyline).
        */
        wxExcelShape AddPolyline(const wxVector<wxPoint2DDouble>& points);

        /**
        Returns a Shape that represents the new AutoShape in a worksheet.

        [MSDN documentation for Shapes.AddShape](https://learn.microsoft.com/en-us/office/vba/api/excel.shapes.addshape).
        */
        wxExcelShape AddShape(MsoAutoShapeType type, double left, double top, double width, double height);

        /**
        Creates a text box. Returns a Shape that represents the new text box.

        [MSDN documentation for Shapes.AddTextbox](https://learn.microsoft.com/en-us/office/vba/api/excel.shapes.addtextbox).
        */
        wxExcelShape AddTextbox(MsoTextOrientation orientation, double left, double top,
                                double width, double height);

        /**
        Creates a WordArt object. Returns a Shape that represents the new WordArt object.

        [MSDN documentation for Shapes.AddTextEffect](https://learn.microsoft.com/en-us/office/vba/api/excel.shapes.addtexteffect).
        */
        wxExcelShape AddTextEffect(MsoPresetTextEffect presetTextEffect, const wxString& text,
                                   const wxString& fontName, double fontSize,
                                   MsoTriState fontBold, MsoTriState fontItalic,
                                   double left, double top);

        /**
        Builds a freeform object. Returns a FreeformBuilder Represents the freeform as it is being built. Use the AddNodes method to add segments to the freeform. After you have added at least one segment to the freeform, you can use the ConvertToShape method to convert the FreeformBuilder object into a Shape object that has the geometric description you’ve defined in the FreeformBuilder object.

        [MSDN documentation for Shapes.BuildFreeform](https://learn.microsoft.com/en-us/office/vba/api/excel.shapes.buildfreeform).
        */
        wxExcelFreeformBuilder BuildFreeform(MsoEditingType editingType, double X1, double Y1);

        //@{
        /**
        Returns a single object from a collection.

        [MSDN documentation for Shapes.Item](https://learn.microsoft.com/en-us/office/vba/api/excel.shapes.item).
        */
        wxExcelShape Item(long index);
        wxExcelShape Item(const wxString& name);

        wxExcelShape operator[](long index);
        wxExcelShape operator[](const wxString& name);
        //@}

        /**
        Selects all the shapes in the specified Shapes collection.

        [MSDN documentation for Shapes.SelectAll](https://learn.microsoft.com/en-us/office/vba/api/excel.shapes.selectall).
        */
        void SelectAll();

        // ***** PROPERTIES *****

        /**
        Returns a Long value that represents the number of objects in the collection.

        [MSDN documentation for Shapes.Count](https://learn.microsoft.com/en-us/office/vba/api/excel.shapes.count).
        */
        long GetCount();

        //@{
        /**
        Returns a ShapeRange Represents a subset of the shapes in a Shapes collection.

        [MSDN documentation for Shapes.Range](https://learn.microsoft.com/en-us/office/vba/api/excel.shapes.range).
        */
        wxExcelShapeRange GetRange(long index);
        wxExcelShapeRange GetRange(const wxString& name);
        wxExcelShapeRange GetRange(const wxVector<long>& indices);
        wxExcelShapeRange GetRange(const wxVector<wxString>& names);
        //@}

        /**
        The iterator over the items, see wxExcelCollectionIterator.
        */
        typedef wxExcelCollectionIterator<wxExcelShapes, wxExcelShape> iterator;

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
        Returns "Shapes".
        */
        virtual wxString GetAutoExcelObjectName_() const { return wxS("Shapes"); }
    };


} // namespace wxAutoExcel

#endif // #if WXAUTOEXCEL_USE_SHAPES

#endif //_WXAUTOEXCEL_SHAPES_H
