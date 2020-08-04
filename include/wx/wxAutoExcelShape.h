/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Modified by:
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
/////////////////////////////////////////////////////////////////////////////


#ifndef _WXAUTOEXCEL_SHAPE_H
#define _WXAUTOEXCEL_SHAPE_H

#include <wx/geometry.h>

#include "wx/wxAutoExcel_defs.h"

#if WXAUTOEXCEL_USE_SHAPES

#include "wx/wxAutoExcel_object.h"
#include "wx/wxAutoExcel_enums.h"

namespace wxAutoExcel {

    /**
    @brief Represents Microsoft Excel Shape object.
    */
   class WXDLLIMPEXP_WXAUTOEXCEL wxExcelShape : public wxExcelObject
    {
    public:
        // ***** METHODS *****

        /**
        Applies to the specified shape formatting that’s been copied by using the PickUp method.

        [MSDN documentation for Shape.Apply](http://msdn.microsoft.com/en-us/library/bb178941).
        */
        void Apply();

        /**
        Copies the object to the Clipboard.

        [MSDN documentation for Shape.Copy](http://msdn.microsoft.com/en-us/library/bb178945).
        */
        void Copy();

        /**
        Copies the selected object to the Clipboard as a picture.

        [MSDN documentation for Shape.CopyPicture](http://msdn.microsoft.com/en-us/library/bb178948).
        */
        void CopyPicture(XlPictureAppearance* appearance = NULL, XlCopyPictureFormat* format = NULL);

        /**
        Cuts the object to the Clipboard.

        [MSDN documentation for Shape.Cut](http://msdn.microsoft.com/en-us/library/bb178952).
        */
        void Cut();

        /**
        Deletes the object.

        [MSDN documentation for Shape.Delete](http://msdn.microsoft.com/en-us/library/bb178954).
        */
        void Delete();

        /**
        Duplicates the object and returns a reference to the new copy.

        [MSDN documentation for Shape.Duplicate](http://msdn.microsoft.com/en-us/library/bb178956).
        */
        void Duplicate();

        /**
        Flips the specified shape around its horizontal or vertical axis.

        [MSDN documentation for Shape.Flip](http://msdn.microsoft.com/en-us/library/bb178959).
        */
        void Flip(MsoFlipCmd flipCmd);

        /**
        Moves the specified shape horizontally by the specified number of points.

        [MSDN documentation for Shape.IncrementLeft](http://msdn.microsoft.com/en-us/library/bb212161).
        */
        void IncrementLeft(double increment);

        /**
        Changes the rotation of the specified shape around the z-axis by the specified number of degrees. Use the Rotation property to set the absolute rotation of the shape.

        [MSDN documentation for Shape.IncrementRotation](http://msdn.microsoft.com/en-us/library/bb212163).
        */
        void IncrementRotation(double increment);

        /**
        Moves the specified shape vertically by the specified number of points.

        [MSDN documentation for Shape.IncrementTop](http://msdn.microsoft.com/en-us/library/bb212165).
        */
        void IncrementTop(double increment);

        /**
        Copies the formatting of the specified shape. Use the Apply method to apply the copied formatting to another shape.

        [MSDN documentation for Shape.PickUp](http://msdn.microsoft.com/en-us/library/bb212175).
        */
        void PickUp();

        /**
        This method reroutes all connectors attached to the specified shape; if the specified shape is a connector, it’s rerouted.

        [MSDN documentation for Shape.RerouteConnections](http://msdn.microsoft.com/en-us/library/bb212198).
        */
        void RerouteConnections();

        /**
        Scales the height of the shape by a specified factor. For pictures and OLE objects, you can indicate whether you want to scale the shape relative to the original or the current size. Shapes other than pictures and OLE objects are always scaled relative to their current height.

        [MSDN documentation for Shape.ScaleHeight](http://msdn.microsoft.com/en-us/library/bb214013).
        */
        void ScaleHeight(double factor, MsoTriState relativeToOriginalSize, MsoScaleFrom* scale);

        /**
        Scales the width of the shape by a specified factor. For pictures and OLE objects, you can indicate whether you want to scale the shape relative to the original or the current size. Shapes other than pictures and OLE objects are always scaled relative to their current width.

        [MSDN documentation for Shape.ScaleWidth](http://msdn.microsoft.com/en-us/library/bb214021).
        */
        void ScaleWidth(double factor, MsoTriState relativeToOriginalSize, MsoScaleFrom* scale);

        /**
        Selects the object.

        [MSDN documentation for Shape.Select](http://msdn.microsoft.com/en-us/library/bb214023).
        */
        void Select(wxXlTribool replace = wxDefaultXlTribool);

        /**
        Makes the formatting of the specified shape the default formatting for the shape.

        [MSDN documentation for Shape.SetShapesDefaultProperties](http://msdn.microsoft.com/en-us/library/bb214037).
        */
        void SetShapesDefaultProperties();

        /**
        Ungroups any grouped shapes in the specified shape or range of shapes. Disassembles pictures and OLE objects within the specified shape or range of shapes.

        [MSDN documentation for Shape.Ungroup](http://msdn.microsoft.com/en-us/library/bb214043).
        */
        wxExcelShapeRange  Ungroup();

        /**
        Moves the specified shape in front of or behind other shapes in the collection (that is, changes the shape's position in the z-order).

        [MSDN documentation for Shape.ZOrder](http://msdn.microsoft.com/en-us/library/bb214054).
        */
        void ZOrder(MsoZOrderCmd ZOrderCmd);

        // ***** PROPERTIES *****

        /**
        Returns an Adjustments object that contains adjustment values for all the adjustments in the specified shape. Applies to any Shape Represents an AutoShape, WordArt, or a connector.

        [MSDN documentation for Shape.Adjustments](http://msdn.microsoft.com/en-us/library/bb237641).
        */
        wxExcelAdjustments GetAdjustments();

        /**
        Returns the descriptive (alternative) text string for a Shape object when the object is saved to a Web page.

        [MSDN documentation for Shape.AlternativeText](http://msdn.microsoft.com/en-us/library/bb237651).
        */
        wxString GetAlternativeText();

        /**
        Sets the descriptive (alternative) text string for a Shape object when the object is saved to a Web page.

        [MSDN documentation for Shape.AlternativeText](http://msdn.microsoft.com/en-us/library/bb237651).
        */
        void SetAlternativeText(const wxString& alternativeText);

        /**
        Returns the shape type for the specified Shape or ShapeRange object, which must represent an AutoShape other than a line, freeform drawing, or connector. Read/write MsoAutoShapeType.

        [MSDN documentation for Shape.AutoShapeType](http://msdn.microsoft.com/en-us/library/bb237659).
        */
        MsoAutoShapeType GetAutoShapeType();

        /**
        Sets the shape type for the specified Shape or ShapeRange object, which must represent an AutoShape other than a line, freeform drawing, or connector. Read/write MsoAutoShapeType.

        [MSDN documentation for Shape.AutoShapeType](http://msdn.microsoft.com/en-us/library/bb237659).
        */
        void SetAutoShapeType(MsoAutoShapeType autoShapeType);

        /**
        Background style. Since Excel 2007.

        [MSDN documentation for Shape.BackgroundStyle](http://msdn.microsoft.com/en-us/library/bb240911).
        */
        MsoBackgroundStyleIndex GetBackgroundStyle();

        /**
        Returns a value that indicates how the specified shape appears when the presentation is viewed in black-and-white mode. Read/write MsoBlackWhiteMode.

        [MSDN documentation for Shape.BlackWhiteMode](http://msdn.microsoft.com/en-us/library/bb237665).
        */
        MsoBlackWhiteMode GetBlackWhiteMode();

        /**
        Sets a value that indicates how the specified shape appears when the presentation is viewed in black-and-white mode. Read/write MsoBlackWhiteMode.

        [MSDN documentation for Shape.BlackWhiteMode](http://msdn.microsoft.com/en-us/library/bb237665).
        */
        void SetBlackWhiteMode(MsoBlackWhiteMode blackWhiteMode);

        /**
        Returns a Range Represents the cell that lies under the lower-right corner of the object.

        [MSDN documentation for Shape.BottomRightCell](http://msdn.microsoft.com/en-us/library/bb237673).
        */
        wxExcelRange GetBottomRightCell();

        /**
        Returns a CalloutFormat object that contains callout formatting properties for the specified shape. Applies to a Shape object that represent line callouts.

        [MSDN documentation for Shape.Callout](http://msdn.microsoft.com/en-us/library/bb237689).
        */
        wxExcelCalloutFormat GetCallout();

#if WXAUTOEXCEL_USE_CHARTS
        /**
        Read-only Since Excel 2007.

        [MSDN documentation for Shape.Chart](http://msdn.microsoft.com/en-us/library/bb237696).
        */
        wxExcelChart GetChart();
#endif // #if WXAUTOEXCEL_USE_CHARTS
        /**
        Returns msoTrue if the specified shape is a child shape or if all shapes in a shape range are child shapes of the same parent. Read-only MsoTriState.

        [MSDN documentation for Shape.Child](http://msdn.microsoft.com/en-us/library/bb237709).
        */
        MsoTriState GetChild();

        /**
        Returns the number of connection sites on the specified shape.

        [MSDN documentation for Shape.ConnectionSiteCount](http://msdn.microsoft.com/en-us/library/bb237718).
        */
        long GetConnectionSiteCount();

        /**
        True if the specified shape is a connector. Read-only MsoTriState.

        [MSDN documentation for Shape.Connector](http://msdn.microsoft.com/en-us/library/bb237732).
        */
        MsoTriState GetConnector();

        /**
        Returns a ConnectorFormat object that contains connector formatting properties. Applies to a Shape that represent connectors.

        [MSDN documentation for Shape.ConnectorFormat](http://msdn.microsoft.com/en-us/library/bb237739).
        */
        wxExcelConnectorFormat  GetConnectorFormat();

        /**
        Returns a ControlFormat object that contains Microsoft Excel control properties.

        [MSDN documentation for Shape.ControlFormat](http://msdn.microsoft.com/en-us/library/bb177393).
        */
        wxExcelControlFormat  GetControlFormat();

        /**
        Returns a FillFormat object for a specified shape or a ChartFillFormat object for a specified chart that contains fill formatting properties for the shape or chart.

        [MSDN documentation for Shape.Fill](http://msdn.microsoft.com/en-us/library/bb237745).
        */
        wxExcelFillFormat GetFill();

        /**
        Returns the Microsoft Excel control type. Read-only XlFormControl.

        [MSDN documentation for Shape.FormControlType](http://msdn.microsoft.com/en-us/library/bb208528).
        */
        XlFormControl GetFormControlType();

        /**
        GlowFormat Since Excel 2007.

        [MSDN documentation for Shape.Glow](http://msdn.microsoft.com/en-us/library/bb242574).
        */
        wxExcelGlowFormat GetGlow();

        /**
        Returns an MsoGraphicStyleIndex constant that represents the style of an SVG graphic.

        [Excel VBA documentation for Shape.GraphicStyle](https://docs.microsoft.com/en-us/office/vba/api/excel.shape.graphicstyle)
        */
        MsoGraphicStyleIndex GetGraphicStyle();

        /**
        Sets an MsoGraphicStyleIndex constant that represents the style of an SVG graphic.

        [Excel VBA documentation for Shape.GraphicStyle](https://docs.microsoft.com/en-us/office/vba/api/excel.shape.graphicstyle)
        */
        void SetGraphicStyle(MsoGraphicStyleIndex graphicStyle);

        /**
        Returns a GroupShapes Represents the individual shapes in the specified group. Use the Item method of the GroupShapes object to return a single shape from the group. Applies to Shape objects that represent grouped shapes.

        [MSDN documentation for Shape.GroupItems](http://msdn.microsoft.com/en-us/library/bb213722).
        */
        wxExcelGroupShapes GetGroupItems();

        /**
        Returns if shape contains a chart. Since Excel 2007.

        [MSDN documentation for Shape.HasChart](http://msdn.microsoft.com/en-us/library/bb213737).
        */
        MsoTriState GetHasChart();

        /**
        Returns a Single value that represents the height, in points, of the object.

        [MSDN documentation for Shape.Height](http://msdn.microsoft.com/en-us/library/bb213745).
        */
        double GetHeight();

        /**
        True if the specified shape is flipped around the horizontal axis. Read-only MsoTriState.

        [MSDN documentation for Shape.HorizontalFlip](http://msdn.microsoft.com/en-us/library/bb213749).
        */
        MsoTriState GetHorizontalFlip();

        /**
        Returns a Hyperlink Represents the hyperlink for the shape.

        [MSDN documentation for Shape.Hyperlink](http://msdn.microsoft.com/en-us/library/bb177584).
        */
        wxExcelHyperlink GetHyperlink();

        /**
        Returns a Long value that represents the type for the specified object.

        [MSDN documentation for Shape.ID](http://msdn.microsoft.com/en-us/library/bb213752).
        */
        long GetID();

        /**
        Returns a Single value that represents the distance, in points, from the left edge of the object to the left edge of column A (on a worksheet) or the left edge of the chart area (on a chart).

        [MSDN documentation for Shape.Left](http://msdn.microsoft.com/en-us/library/bb213756).
        */
        double GetLeft();

        /**
        Returns a LineFormat object that contains line formatting properties for the specified shape. (For a line, the LineFormat object represents the line itself; for a shape with a border, the LineFormat object represents the border).

        [MSDN documentation for Shape.Line](http://msdn.microsoft.com/en-us/library/bb213760).
        */
        wxExcelLineFormat GetLine();

        /**
        Returns a LinkFormat object that contains linked OLE object properties.

        [MSDN documentation for Shape.LinkFormat](http://msdn.microsoft.com/en-us/library/bb177911).
        */
        wxExcelLinkFormat GetLinkFormat();

        /**
        True if the specified shape retains its original proportions when you resize it. False if you can change the height and width of the shape independently of one another when you resize it. Read/write MsoTriState.

        [MSDN documentation for Shape.LockAspectRatio](http://msdn.microsoft.com/en-us/library/bb213766).
        */
        MsoTriState GetLockAspectRatio();

        /**
        True if the specified shape retains its original proportions when you resize it. False if you can change the height and width of the shape independently of one another when you resize it. Read/write MsoTriState.

        [MSDN documentation for Shape.LockAspectRatio](http://msdn.microsoft.com/en-us/library/bb213766).
        */
        void SetLockAspectRatio(MsoTriState lockAspectRatio);

        /**
        Returns a Boolean value that indicates if the object is locked.

        [MSDN documentation for Shape.Locked](http://msdn.microsoft.com/en-us/library/bb213772).
        */
        bool GetLocked();

        /**
        Sets a Boolean value that indicates if the object is locked.

        [MSDN documentation for Shape.Locked](http://msdn.microsoft.com/en-us/library/bb213772).
        */
        void SetLocked(bool locked);

        /**
        Returns a Model3DFormat object that contains Model3D properties.

        [Excel VBA documentation for Shape.Model3D](https://docs.microsoft.com/en-us/office/vba/api/excel.shape.model3d)
        */
        wxExcelModel3DFormat GetModel3D();

        /**
        Returns a String value representing the name of the object.

        [MSDN documentation for Shape.Name](http://msdn.microsoft.com/en-us/library/bb213775).
        */
        wxString GetName();

        /**
        Returns a ShapeNodes collection that represents the geometric description of the specified shape.

        [MSDN documentation for Shape.Nodes](http://msdn.microsoft.com/en-us/library/bb213781).
        */
        wxExcelShapeNodes GetNodes();

        /**
        Returns an OLEFormat object that contains OLE object properties.

        [MSDN documentation for Shape.OLEFormat](http://msdn.microsoft.com/en-us/library/bb208867).
        */
        wxExcelOLEFormat GetOLEFormat();

        /**
        Returns the name of a macro that’s run when the specified object is clicked.

        [MSDN documentation for Shape.OnAction](http://msdn.microsoft.com/en-us/library/bb208891).
        */
        wxString GetOnAction();

        /**
        Sets the name of a macro that’s run when the specified object is clicked.

        [MSDN documentation for Shape.OnAction](http://msdn.microsoft.com/en-us/library/bb208891).
        */
        void SetOnAction(const wxString& onAction);

        /**
        Returns a Shape Represents the common parent shape of a child shape or a range of child shapes.

        [MSDN documentation for Shape.ParentGroup](http://msdn.microsoft.com/en-us/library/bb213785).
        */
        wxExcelShape GetParentGroup();

        /**
        Returns a PictureFormat object that contains picture formatting properties for the specified shape. Applies to a Shape object that represent pictures or OLE objects.

        [MSDN documentation for Shape.PictureFormat](http://msdn.microsoft.com/en-us/library/bb213849).
        */
        wxExcelPictureFormat  GetPictureFormat();

        /**
        Returns an XlPlacement value that represents the way the object is attached to the cells below it.

        [MSDN documentation for Shape.Placement](http://msdn.microsoft.com/en-us/library/bb213855).
        */
        XlPlacement  GetPlacement();

        /**
        Sets an XlPlacement value that represents the way the object is attached to the cells below it.

        [MSDN documentation for Shape.Placement](http://msdn.microsoft.com/en-us/library/bb213855).
        */
        void SetPlacement(XlPlacement  placement);

        /**
        ReflectionFormat Since Excel 2007.

        [MSDN documentation for Shape.Reflection](http://msdn.microsoft.com/en-us/library/bb242582).
        */
        wxExcelReflectionFormat GetReflection();

        /**
        Returns the rotation of the shape, in degrees. Read/write Single.

        [MSDN documentation for Shape.Rotation](http://msdn.microsoft.com/en-us/library/bb215087).
        */
        double GetRotation();

        /**
        Sets the rotation of the shape, in degrees. Read/write Single.

        [MSDN documentation for Shape.Rotation](http://msdn.microsoft.com/en-us/library/bb215087).
        */
        void SetRotation(double rotation);

        /**
        Returns a read-only ShadowFormat object that contains shadow formatting properties for the specified shape or shapes.

        [MSDN documentation for Shape.Shadow](http://msdn.microsoft.com/en-us/library/bb215089).
        */
        wxExcelShadowFormat GetShadow();

        /**
        The style of the shape. Since Excel 2007.

        [MSDN documentation for Shape.ShapeStyle](http://msdn.microsoft.com/en-us/library/bb240883).
        */
        MsoShapeStyleIndex GetShapeStyle();

        /**
        The style of the shape. Since Excel 2007.

        [MSDN documentation for Shape.ShapeStyle](http://msdn.microsoft.com/en-us/library/bb240883).
        */
        void SetShapeStyle(MsoShapeStyleIndex shapeStyle);

        /**
        SoftEdgeFormat Since Excel 2007.

        [MSDN documentation for Shape.SoftEdge](http://msdn.microsoft.com/en-us/library/bb242583).
        */
        wxExcelSoftEdgeFormat GetSoftEdge();

        /**
        Returns a TextEffectFormat object that contains text-effect formatting properties for the specified shape.

        [MSDN documentation for Shape.TextEffect](http://msdn.microsoft.com/en-us/library/bb215092).
        */
        wxExcelTextEffectFormat GetTextEffect();

        /**
        Returns a TextFrame object that contains the alignment and anchoring properties for the specified shape.

        [MSDN documentation for Shape.TextFrame](http://msdn.microsoft.com/en-us/library/bb215095).
        */
        wxExcelTextFrame GetTextFrame();

        /**
        TextFrame2  Since Excel 2007.

        [MSDN documentation for Shape.TextFrame2](http://msdn.microsoft.com/en-us/library/bb215098).
        */
        wxExcelTextFrame2  GetTextFrame2();

        /**
        Returns a ThreeDFormat object that contains 3-D – effect formatting properties for the specified shape.

        [MSDN documentation for Shape.ThreeD](http://msdn.microsoft.com/en-us/library/bb215100).
        */
        wxExcelThreeDFormat GetThreeD();

        /**
        Returns a Single value that represents the distance, in points, from the top edge of the topmost shape in the shape range to the top edge of the worksheet.

        [MSDN documentation for Shape.Top](http://msdn.microsoft.com/en-us/library/bb215103).
        */
        double GetTop();

        /**
        Sets a Single value that represents the distance, in points, from the top edge of the topmost shape in the shape range to the top edge of the worksheet.

        [MSDN documentation for Shape.Top](http://msdn.microsoft.com/en-us/library/bb215103).
        */
        void SetTop(double top);

        /**
        Returns a Range Represents the cell that lies under the upper-left corner of the specified object.

        [MSDN documentation for Shape.TopLeftCell](http://msdn.microsoft.com/en-us/library/bb215106).
        */
        wxExcelRange GetTopLeftCell();

        /**
        Returns a MsoShapeType value that represents the shape type.

        [MSDN documentation for Shape.Type](http://msdn.microsoft.com/en-us/library/bb215108).
        */
        MsoShapeType  GetType();

        /**
        True if the specified shape is flipped around the vertical axis. Read-only MsoTriState.

        [MSDN documentation for Shape.VerticalFlip](http://msdn.microsoft.com/en-us/library/bb215112).
        */
        MsoTriState GetVerticalFlip();

        /**
        Returns the coordinates of the specified freeform drawing's vertices (and control points for Bézier curves) as a series of coordinate pairs. You can use the array returned by this property as an argument to the AddCurve method or AddPolyLine method.

        [MSDN documentation for Shape.Vertices](http://msdn.microsoft.com/en-us/library/bb215114).
        */
        wxVector<wxPoint2DDouble> GetVertices();

        /**
        Returns a MsoTriState value that determines whether the object is visible.

        [MSDN documentation for Shape.Visible](http://msdn.microsoft.com/en-us/library/bb215117).
        */
        MsoTriState GetVisible();

        /**
        Sets a MsoTriState value that determines whether the object is visible.

        [MSDN documentation for Shape.Visible](http://msdn.microsoft.com/en-us/library/bb215117).
        */
        void SetVisible(MsoTriState visible);

        /**
        Returns a Single value that represents the width, in points, of the object.

        [MSDN documentation for Shape.Width](http://msdn.microsoft.com/en-us/library/bb215119).
        */
        double GetWidth();

        /**
        Sets a Single value that represents the width, in points, of the object.

        [MSDN documentation for Shape.Width](http://msdn.microsoft.com/en-us/library/bb215119).
        */
        void SetWidth(double width);

        /**
        Returns the position of the specified shape in the z-order.Read-only

        [MSDN documentation for Shape.ZOrderPosition](http://msdn.microsoft.com/en-us/library/bb215122).
        */
        long GetZOrderPosition();

        /**
        Returns "Shape".
        */
        virtual wxString GetAutoExcelObjectName_() const { return wxS("Shape"); }
    };


} // namespace wxAutoExcel


#endif // #if WXAUTOEXCEL_USE_SHAPES

#endif //_WXAUTOEXCEL_SHAPE_H
