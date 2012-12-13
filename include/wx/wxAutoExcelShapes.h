/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Modified by:
// RCS-ID:      $Id: $
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
/////////////////////////////////////////////////////////////////////////////


#ifndef _WXAUTOEXCEL_SHAPES_H
#define _WXAUTOEXCEL_SHAPES_H

#include "wx/wxAutoExcel_defs.h"

#if WXAUTOEXCEL_USE_SHAPES

#include "wx/wxAutoExcelObject.h"
#include "wx/wxAutoExcel_enums.h"

class wxPoint2DDouble;

namespace wxAutoExcel {


    /**
    @brief Represents Microsoft Excel Shape object.
    */
   class WXDLLIMPEXP_WXAUTOEXCEL wxExcelShapes : public wxExcelObject
    {
    public:        
        // ***** METHODS *****

        /**
        Creates a borderless line callout. Returns a Shape Represents the new callout.

        [MSDN documentation for Shapes.AddCallout](http://msdn.microsoft.com/en-us/library/bb209543).
        */
        wxExcelShape AddCallout(MsoCalloutType type, double left, double top,
                                double width, double height);

#if WXAUTOEXCEL_USE_CHARTS
        /**
        Creates a chart at the specified location on the active sheet.

        [MSDN documentation for Shapes.AddChart](http://msdn.microsoft.com/en-us/library/bb238877).
        */
        wxExcelShape AddChart(XlChartType type, double left, double top,
                              double width, double height);
#endif // #if WXAUTOEXCEL_USE_CHARTS

        /**
        Creates a connector. Returns a Shape Represents the new connector. When a connector is added, it's not connected to anything. Use the BeginConnect and EndConnect methods to attach the beginning and end of a connector to other shapes in the document.

        [MSDN documentation for Shapes.AddConnector](http://msdn.microsoft.com/en-us/library/bb209551).
        */
        wxExcelShape AddConnector(MsoConnectorType type, double beginX, double beginY,
                                  double endX, double endY);



        /**
        Returns a Shape Represents a Bézier curve in a worksheet.

        [MSDN documentation for Shapes.AddCurve](http://msdn.microsoft.com/en-us/library/bb209554).
        */
        wxExcelShape AddCurve(const wxVector<wxPoint2DDouble>& points);

        /**
        Creates a Microsoft Excel control. Returns a Shape Represents the new control.

        [MSDN documentation for Shapes.AddFormControl](http://msdn.microsoft.com/en-us/library/bb209570).
        */
        wxExcelShape AddFormControl(XlFormControl type, double left, double top,
                                    double width, double height);

        /**
        Creates a label. Returns a Shape Represents the new label.

        [MSDN documentation for Shapes.AddLabel](http://msdn.microsoft.com/en-us/library/bb209578).
        */
        wxExcelShape AddLabel(MsoTextOrientation orientation, double left, double top,
                              double width, double height);

        /**
        As it applies to the Shapes object, returns a Shape Represents the new line in a worksheet.

        [MSDN documentation for Shapes.AddLine](http://msdn.microsoft.com/en-us/library/bb209581).
        */
        wxExcelShape AddLine(double beginX, double beginY, double endX, double endY);

        /**
        Creates an OLE object. Returns a Shape Represents the new OLE object.

        [MSDN documentation for Shapes.AddOLEObject](http://msdn.microsoft.com/en-us/library/bb209596).
        */
        wxExcelShape AddOLEObject(const wxString& classType = wxEmptyString, const wxString& filename = wxEmptyString, 
                                  wxXlTribool link = wxDefaultXlTribool, wxXlTribool displayAsIcon = wxDefaultXlTribool, 
                                  const wxString& iconFileName = wxEmptyString, long* iconIndex = NULL, 
                                  const wxString& iconLabel = wxEmptyString, 
                                  double* left = NULL, double* top = NULL, double* width = NULL, double* height = NULL);

        /**
        Creates a picture from an existing file. Returns a Shape Represents the new picture.

        [MSDN documentation for Shapes.AddPicture](http://msdn.microsoft.com/en-us/library/bb209605).
        */
        wxExcelShape AddPicture(const wxString& fileName, MsoTriState linkToFile,
                                MsoTriState saveWithDocument, double left, double top,
                                double width, double height);

        /**
        Creates an open polyline or a closed polygon drawing. Returns a Shape Represents the new polyline or polygon.

        [MSDN documentation for Shapes.AddPolyline](http://msdn.microsoft.com/en-us/library/bb209611).
        */
        wxExcelShape AddPolyline(const wxVector<wxPoint2DDouble>& points);

        /**
        Returns a Shape Represents the new AutoShape in a worksheet.

        [MSDN documentation for Shapes.AddShape](http://msdn.microsoft.com/en-us/library/bb209625).
        */
        wxExcelShape AddShape(MsoAutoShapeType type, double left, double top, double width, double height);

        /**
        Creates a text box. Returns a Shape Represents the new text box.

        [MSDN documentation for Shapes.AddTextbox](http://msdn.microsoft.com/en-us/library/bb209628).
        */
        wxExcelShape AddTextbox(MsoTextOrientation orientation, double left, double top,
                                double width, double height);

        /**
        Creates a WordArt object. Returns a Shape Represents the new WordArt object.

        [MSDN documentation for Shapes.AddTextEffect](http://msdn.microsoft.com/en-us/library/bb209633).
        */
        wxExcelShape AddTextEffect(MsoPresetTextEffect presetTextEffect, const wxString& text,
                                   const wxString& fontName, double fontSize, 
                                   MsoTriState fontBold, MsoTriState fontItalic,
                                   double left, double top);

        /**
        Builds a freeform object. Returns a FreeformBuilder Represents the freeform as it is being built. Use the AddNodes method to add segments to the freeform. After you have added at least one segment to the freeform, you can use the ConvertToShape method to convert the FreeformBuilder object into a Shape object that has the geometric description you’ve defined in the FreeformBuilder object.

        [MSDN documentation for Shapes.BuildFreeform](http://msdn.microsoft.com/en-us/library/bb209724).
        */
        wxExcelFreeformBuilder BuildFreeform(MsoEditingType editingType, double X1, double Y1);

        //@{
        /**
        Returns a single object from a collection.

        [MSDN documentation for Shapes.Item](http://msdn.microsoft.com/en-us/library/bb212199).
        */
        wxExcelShape Item(long index);
        wxExcelShape Item(const wxString& name);
        
        wxExcelShape operator[](long index);
        wxExcelShape operator[](const wxString& name);
        //@}

        /**
        Selects all the shapes in the specified Shapes collection.

        [MSDN documentation for Shapes.SelectAll](http://msdn.microsoft.com/en-us/library/bb238890).
        */
        void SelectAll();

        // ***** PROPERTIES *****

        /**
        Returns a Long value that represents the number of objects in the collection.

        [MSDN documentation for Shapes.Count](http://msdn.microsoft.com/en-us/library/bb237885).
        */
        long GetCount();
        
        //@{
        /**
        Returns a ShapeRange Represents a subset of the shapes in a Shapes collection.

        [MSDN documentation for Shapes.Range](http://msdn.microsoft.com/en-us/library/bb237897).
        */
        wxExcelShapeRange GetRange(long index);
        wxExcelShapeRange GetRange(const wxString& name);
        wxExcelShapeRange GetRange(const wxVector<long>& indices);
        wxExcelShapeRange GetRange(const wxVector<wxString>& names);
        //@}

        /**
        Returns "Shapes".
        */
        virtual wxString GetAutoExcelObjectName_() const { return wxS("Shapes"); }
    };


} // namespace wxAutoExcel

#endif // #if WXAUTOEXCEL_USE_SHAPES

#endif //_WXAUTOEXCEL_SHAPES_H
