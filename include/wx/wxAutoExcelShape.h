/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// License:     MIT license
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

        [MSDN documentation for Shape.Apply](https://learn.microsoft.com/en-us/office/vba/api/excel.shape.apply).
        */
        void Apply();

        /**
        Copies the object to the Clipboard.

        [MSDN documentation for Shape.Copy](https://learn.microsoft.com/en-us/office/vba/api/excel.shape.copy).
        */
        void Copy();

        /**
        Copies the selected object to the Clipboard as a picture.

        [MSDN documentation for Shape.CopyPicture](https://learn.microsoft.com/en-us/office/vba/api/excel.shape.copypicture).
        */
        void CopyPicture(XlPictureAppearance* appearance = NULL, XlCopyPictureFormat* format = NULL);

        /**
        Cuts the object to the Clipboard.

        [MSDN documentation for Shape.Cut](https://learn.microsoft.com/en-us/office/vba/api/excel.shape.cut).
        */
        void Cut();

        /**
        Deletes the object.

        [MSDN documentation for Shape.Delete](https://learn.microsoft.com/en-us/office/vba/api/excel.shape.delete).
        */
        void Delete();

        /**
        Duplicates the object and returns a reference to the new copy.

        [MSDN documentation for Shape.Duplicate](https://learn.microsoft.com/en-us/office/vba/api/excel.shape.duplicate).
        */
        wxExcelShape Duplicate();

        /**
        Flips the specified shape around its horizontal or vertical axis.

        [MSDN documentation for Shape.Flip](https://learn.microsoft.com/en-us/office/vba/api/excel.shape.flip).
        */
        void Flip(MsoFlipCmd flipCmd);

        /**
        Moves the specified shape horizontally by the specified number of points.

        [MSDN documentation for Shape.IncrementLeft](https://learn.microsoft.com/en-us/office/vba/api/excel.shape.incrementleft).
        */
        void IncrementLeft(double increment);

        /**
        Changes the rotation of the specified shape around the z-axis by the specified number of degrees. Use the Rotation property to set the absolute rotation of the shape.

        [MSDN documentation for Shape.IncrementRotation](https://learn.microsoft.com/en-us/office/vba/api/excel.shape.incrementrotation).
        */
        void IncrementRotation(double increment);

        /**
        Moves the specified shape vertically by the specified number of points.

        [MSDN documentation for Shape.IncrementTop](https://learn.microsoft.com/en-us/office/vba/api/excel.shape.incrementtop).
        */
        void IncrementTop(double increment);

        /**
        Copies the formatting of the specified shape. Use the Apply method to apply the copied formatting to another shape.

        [MSDN documentation for Shape.PickUp](https://learn.microsoft.com/en-us/office/vba/api/excel.shape.pickup).
        */
        void PickUp();

        /**
        This method reroutes all connectors attached to the specified shape; if the specified shape is a connector, it’s rerouted.

        [MSDN documentation for Shape.RerouteConnections](https://learn.microsoft.com/en-us/office/vba/api/excel.shape.rerouteconnections).
        */
        void RerouteConnections();

        /**
        Scales the height of the shape by a specified factor. For pictures and OLE objects, you can indicate whether you want to scale the shape relative to the original or the current size. Shapes other than pictures and OLE objects are always scaled relative to their current height.

        [MSDN documentation for Shape.ScaleHeight](https://learn.microsoft.com/en-us/office/vba/api/excel.shape.scaleheight).
        */
        void ScaleHeight(double factor, MsoTriState relativeToOriginalSize, MsoScaleFrom* scale);

        /**
        Scales the width of the shape by a specified factor. For pictures and OLE objects, you can indicate whether you want to scale the shape relative to the original or the current size. Shapes other than pictures and OLE objects are always scaled relative to their current width.

        [MSDN documentation for Shape.ScaleWidth](https://learn.microsoft.com/en-us/office/vba/api/excel.shape.scalewidth).
        */
        void ScaleWidth(double factor, MsoTriState relativeToOriginalSize, MsoScaleFrom* scale);

        /**
        Selects the object.

        [MSDN documentation for Shape.Select](https://learn.microsoft.com/en-us/office/vba/api/excel.shape.select).
        */
        void Select(wxXlTribool replace = wxDefaultXlTribool);

        /**
        Makes the formatting of the specified shape the default formatting for the shape.

        [MSDN documentation for Shape.SetShapesDefaultProperties](https://learn.microsoft.com/en-us/office/vba/api/excel.shape.setshapesdefaultproperties).
        */
        void SetShapesDefaultProperties();

        /**
        Ungroups any grouped shapes in the specified shape or range of shapes. Disassembles pictures and OLE objects within the specified shape or range of shapes.

        [MSDN documentation for Shape.Ungroup](https://learn.microsoft.com/en-us/office/vba/api/excel.shape.ungroup).
        */
        wxExcelShapeRange  Ungroup();

        /**
        Moves the specified shape in front of or behind other shapes in the collection (that is, changes the shape's position in the z-order).

        [MSDN documentation for Shape.ZOrder](https://learn.microsoft.com/en-us/office/vba/api/excel.shape.zorder).
        */
        void ZOrder(MsoZOrderCmd ZOrderCmd);

        // ***** PROPERTIES *****

        /**
        Returns an Adjustments object that contains adjustment values for all the adjustments in the specified shape. Applies to any Shape Represents an AutoShape, WordArt, or a connector.

        [MSDN documentation for Shape.Adjustments](https://learn.microsoft.com/en-us/office/vba/api/excel.shape.adjustments).
        */
        wxExcelAdjustments GetAdjustments();

        /**
        Returns the descriptive (alternative) text string for a Shape object when the object is saved to a Web page.

        [MSDN documentation for Shape.AlternativeText](https://learn.microsoft.com/en-us/office/vba/api/excel.shape.alternativetext).
        */
        wxString GetAlternativeText();

        /**
        Sets the descriptive (alternative) text string for a Shape object when the object is saved to a Web page.

        [MSDN documentation for Shape.AlternativeText](https://learn.microsoft.com/en-us/office/vba/api/excel.shape.alternativetext).
        */
        void SetAlternativeText(const wxString& alternativeText);

        /**
        Returns the shape type for the specified Shape or ShapeRange object, which must represent an AutoShape other than a line, freeform drawing, or connector. Read/write MsoAutoShapeType.

        [MSDN documentation for Shape.AutoShapeType](https://learn.microsoft.com/en-us/office/vba/api/excel.shape.autoshapetype).
        */
        MsoAutoShapeType GetAutoShapeType();

        /**
        Sets the shape type for the specified Shape or ShapeRange object, which must represent an AutoShape other than a line, freeform drawing, or connector. Read/write MsoAutoShapeType.

        [MSDN documentation for Shape.AutoShapeType](https://learn.microsoft.com/en-us/office/vba/api/excel.shape.autoshapetype).
        */
        void SetAutoShapeType(MsoAutoShapeType autoShapeType);

        /**
        Background style. Since Excel 2007.

        [MSDN documentation for Shape.BackgroundStyle](https://learn.microsoft.com/en-us/office/vba/api/excel.shape.backgroundstyle).
        */
        MsoBackgroundStyleIndex GetBackgroundStyle();

        /**
        Returns a value that indicates how the specified shape appears when the presentation is viewed in black-and-white mode. Read/write MsoBlackWhiteMode.

        [MSDN documentation for Shape.BlackWhiteMode](https://learn.microsoft.com/en-us/office/vba/api/excel.shape.blackwhitemode).
        */
        MsoBlackWhiteMode GetBlackWhiteMode();

        /**
        Sets a value that indicates how the specified shape appears when the presentation is viewed in black-and-white mode. Read/write MsoBlackWhiteMode.

        [MSDN documentation for Shape.BlackWhiteMode](https://learn.microsoft.com/en-us/office/vba/api/excel.shape.blackwhitemode).
        */
        void SetBlackWhiteMode(MsoBlackWhiteMode blackWhiteMode);

        /**
        Returns a Range Represents the cell that lies under the lower-right corner of the object.

        [MSDN documentation for Shape.BottomRightCell](https://learn.microsoft.com/en-us/office/vba/api/excel.shape.bottomrightcell).
        */
        wxExcelRange GetBottomRightCell();

        /**
        Returns a CalloutFormat object that contains callout formatting properties for the specified shape. Applies to a Shape object that represent line callouts.

        [MSDN documentation for Shape.Callout](https://learn.microsoft.com/en-us/office/vba/api/excel.shape.callout).
        */
        wxExcelCalloutFormat GetCallout();

#if WXAUTOEXCEL_USE_CHARTS
        /**
        Read-only Since Excel 2007.

        [MSDN documentation for Shape.Chart](https://learn.microsoft.com/en-us/office/vba/api/excel.shape.chart).
        */
        wxExcelChart GetChart();
#endif // #if WXAUTOEXCEL_USE_CHARTS
        /**
        Returns msoTrue if the specified shape is a child shape or if all shapes in a shape range are child shapes of the same parent. Read-only MsoTriState.

        [MSDN documentation for Shape.Child](https://learn.microsoft.com/en-us/office/vba/api/excel.shape.child).
        */
        MsoTriState GetChild();

        /**
        Returns the number of connection sites on the specified shape.

        [MSDN documentation for Shape.ConnectionSiteCount](https://learn.microsoft.com/en-us/office/vba/api/excel.shape.connectionsitecount).
        */
        long GetConnectionSiteCount();

        /**
        True if the specified shape is a connector. Read-only MsoTriState.

        [MSDN documentation for Shape.Connector](https://learn.microsoft.com/en-us/office/vba/api/excel.shape.connector).
        */
        MsoTriState GetConnector();

        /**
        Returns a ConnectorFormat object that contains connector formatting properties. Applies to a Shape that represent connectors.

        [MSDN documentation for Shape.ConnectorFormat](https://learn.microsoft.com/en-us/office/vba/api/excel.shape.connectorformat).
        */
        wxExcelConnectorFormat  GetConnectorFormat();

        /**
        Returns a ControlFormat object that contains Microsoft Excel control properties.

        [MSDN documentation for Shape.ControlFormat](https://learn.microsoft.com/en-us/office/vba/api/excel.shape.controlformat).
        */
        wxExcelControlFormat  GetControlFormat();

        /**
        Returns a FillFormat object for a specified shape or a ChartFillFormat object for a specified chart that contains fill formatting properties for the shape or chart.

        [MSDN documentation for Shape.Fill](https://learn.microsoft.com/en-us/office/vba/api/excel.shape.fill).
        */
        wxExcelFillFormat GetFill();

        /**
        Returns the Microsoft Excel control type. Read-only XlFormControl.

        [MSDN documentation for Shape.FormControlType](https://learn.microsoft.com/en-us/office/vba/api/excel.shape.formcontroltype).
        */
        XlFormControl GetFormControlType();

        /**
        GlowFormat Since Excel 2007.

        [MSDN documentation for Shape.Glow](https://learn.microsoft.com/en-us/office/vba/api/excel.shape.glow).
        */
        wxExcelGlowFormat GetGlow();

        /**
        Returns an MsoGraphicStyleIndex constant that represents the style of an SVG graphic.

        [Excel VBA documentation for Shape.GraphicStyle](https://learn.microsoft.com/en-us/office/vba/api/excel.shape.graphicstyle)
        */
        MsoGraphicStyleIndex GetGraphicStyle();

        /**
        Sets an MsoGraphicStyleIndex constant that represents the style of an SVG graphic.

        [Excel VBA documentation for Shape.GraphicStyle](https://learn.microsoft.com/en-us/office/vba/api/excel.shape.graphicstyle)
        */
        void SetGraphicStyle(MsoGraphicStyleIndex graphicStyle);

        /**
        Returns a GroupShapes Represents the individual shapes in the specified group. Use the Item method of the GroupShapes object to return a single shape from the group. Applies to Shape objects that represent grouped shapes.

        [MSDN documentation for Shape.GroupItems](https://learn.microsoft.com/en-us/office/vba/api/excel.shape.groupitems).
        */
        wxExcelGroupShapes GetGroupItems();

        /**
        Returns if shape contains a chart. Since Excel 2007.

        [MSDN documentation for Shape.HasChart](https://learn.microsoft.com/en-us/office/vba/api/excel.shape.haschart).
        */
        MsoTriState GetHasChart();

        /**
        Returns a Single value that represents the height, in points, of the object.

        [MSDN documentation for Shape.Height](https://learn.microsoft.com/en-us/office/vba/api/excel.shape.height).
        */
        double GetHeight();

        /**
        True if the specified shape is flipped around the horizontal axis. Read-only MsoTriState.

        [MSDN documentation for Shape.HorizontalFlip](https://learn.microsoft.com/en-us/office/vba/api/excel.shape.horizontalflip).
        */
        MsoTriState GetHorizontalFlip();

        /**
        Returns a Hyperlink Represents the hyperlink for the shape.

        [MSDN documentation for Shape.Hyperlink](https://learn.microsoft.com/en-us/office/vba/api/excel.shape.hyperlink).
        */
        wxExcelHyperlink GetHyperlink();

        /**
        Returns a Long value that represents the type for the specified object.

        [MSDN documentation for Shape.ID](https://learn.microsoft.com/en-us/office/vba/api/excel.shape.id).
        */
        long GetID();

        /**
        Returns a Single value that represents the distance, in points, from the left edge of the object to the left edge of column A (on a worksheet) or the left edge of the chart area (on a chart).

        [MSDN documentation for Shape.Left](https://learn.microsoft.com/en-us/office/vba/api/excel.shape.left).
        */
        double GetLeft();

        /**
        Returns a LineFormat object that contains line formatting properties for the specified shape. (For a line, the LineFormat object represents the line itself; for a shape with a border, the LineFormat object represents the border).

        [MSDN documentation for Shape.Line](https://learn.microsoft.com/en-us/office/vba/api/excel.shape.line).
        */
        wxExcelLineFormat GetLine();

        /**
        Returns a LinkFormat object that contains linked OLE object properties.

        [MSDN documentation for Shape.LinkFormat](https://learn.microsoft.com/en-us/office/vba/api/excel.shape.linkformat).
        */
        wxExcelLinkFormat GetLinkFormat();

        /**
        True if the specified shape retains its original proportions when you resize it. False if you can change the height and width of the shape independently of one another when you resize it. Read/write MsoTriState.

        [MSDN documentation for Shape.LockAspectRatio](https://learn.microsoft.com/en-us/office/vba/api/excel.shape.lockaspectratio).
        */
        MsoTriState GetLockAspectRatio();

        /**
        True if the specified shape retains its original proportions when you resize it. False if you can change the height and width of the shape independently of one another when you resize it. Read/write MsoTriState.

        [MSDN documentation for Shape.LockAspectRatio](https://learn.microsoft.com/en-us/office/vba/api/excel.shape.lockaspectratio).
        */
        void SetLockAspectRatio(MsoTriState lockAspectRatio);

        /**
        Returns a Boolean value that indicates if the object is locked.

        [MSDN documentation for Shape.Locked](https://learn.microsoft.com/en-us/office/vba/api/excel.shape.locked).
        */
        bool GetLocked();

        /**
        Sets a Boolean value that indicates if the object is locked.

        [MSDN documentation for Shape.Locked](https://learn.microsoft.com/en-us/office/vba/api/excel.shape.locked).
        */
        void SetLocked(bool locked);

        /**
        Returns a Model3DFormat object that contains Model3D properties.

        [Excel VBA documentation for Shape.Model3D](https://learn.microsoft.com/en-us/office/vba/api/excel.shape.model3d)
        */
        wxExcelModel3DFormat GetModel3D();

        /**
        Returns a String value representing the name of the object.

        [MSDN documentation for Shape.Name](https://learn.microsoft.com/en-us/office/vba/api/excel.shape.name).
        */
        wxString GetName();

        /**
        Returns a ShapeNodes collection that represents the geometric description of the specified shape.

        [MSDN documentation for Shape.Nodes](https://learn.microsoft.com/en-us/office/vba/api/excel.shape.nodes).
        */
        wxExcelShapeNodes GetNodes();

        /**
        Returns an OLEFormat object that contains OLE object properties.

        [MSDN documentation for Shape.OLEFormat](https://learn.microsoft.com/en-us/office/vba/api/excel.shape.oleformat).
        */
        wxExcelOLEFormat GetOLEFormat();

        /**
        Returns the name of a macro that’s run when the specified object is clicked.

        [MSDN documentation for Shape.OnAction](https://learn.microsoft.com/en-us/office/vba/api/excel.shape.onaction).
        */
        wxString GetOnAction();

        /**
        Sets the name of a macro that’s run when the specified object is clicked.

        [MSDN documentation for Shape.OnAction](https://learn.microsoft.com/en-us/office/vba/api/excel.shape.onaction).
        */
        void SetOnAction(const wxString& onAction);

        /**
        Returns a Shape Represents the common parent shape of a child shape or a range of child shapes.

        [MSDN documentation for Shape.ParentGroup](https://learn.microsoft.com/en-us/office/vba/api/excel.shape.parentgroup).
        */
        wxExcelShape GetParentGroup();

        /**
        Returns a PictureFormat object that contains picture formatting properties for the specified shape. Applies to a Shape object that represent pictures or OLE objects.

        [MSDN documentation for Shape.PictureFormat](https://learn.microsoft.com/en-us/office/vba/api/excel.shape.pictureformat).
        */
        wxExcelPictureFormat  GetPictureFormat();

        /**
        Returns an XlPlacement value that represents the way the object is attached to the cells below it.

        [MSDN documentation for Shape.Placement](https://learn.microsoft.com/en-us/office/vba/api/excel.shape.placement).
        */
        XlPlacement  GetPlacement();

        /**
        Sets an XlPlacement value that represents the way the object is attached to the cells below it.

        [MSDN documentation for Shape.Placement](https://learn.microsoft.com/en-us/office/vba/api/excel.shape.placement).
        */
        void SetPlacement(XlPlacement  placement);

        /**
        ReflectionFormat Since Excel 2007.

        [MSDN documentation for Shape.Reflection](https://learn.microsoft.com/en-us/office/vba/api/excel.shape.reflection).
        */
        wxExcelReflectionFormat GetReflection();

        /**
        Returns the rotation of the shape, in degrees. Read/write Single.

        [MSDN documentation for Shape.Rotation](https://learn.microsoft.com/en-us/office/vba/api/excel.shape.rotation).
        */
        double GetRotation();

        /**
        Sets the rotation of the shape, in degrees. Read/write Single.

        [MSDN documentation for Shape.Rotation](https://learn.microsoft.com/en-us/office/vba/api/excel.shape.rotation).
        */
        void SetRotation(double rotation);

        /**
        Returns a read-only ShadowFormat object that contains shadow formatting properties for the specified shape or shapes.

        [MSDN documentation for Shape.Shadow](https://learn.microsoft.com/en-us/office/vba/api/excel.shape.shadow).
        */
        wxExcelShadowFormat GetShadow();

        /**
        The style of the shape. Since Excel 2007.

        [MSDN documentation for Shape.ShapeStyle](https://learn.microsoft.com/en-us/office/vba/api/excel.shape.shapestyle).
        */
        MsoShapeStyleIndex GetShapeStyle();

        /**
        The style of the shape. Since Excel 2007.

        [MSDN documentation for Shape.ShapeStyle](https://learn.microsoft.com/en-us/office/vba/api/excel.shape.shapestyle).
        */
        void SetShapeStyle(MsoShapeStyleIndex shapeStyle);

        /**
        SoftEdgeFormat Since Excel 2007.

        [MSDN documentation for Shape.SoftEdge](https://learn.microsoft.com/en-us/office/vba/api/excel.shape.softedge).
        */
        wxExcelSoftEdgeFormat GetSoftEdge();

        /**
        Returns a TextEffectFormat object that contains text-effect formatting properties for the specified shape.

        [MSDN documentation for Shape.TextEffect](https://learn.microsoft.com/en-us/office/vba/api/excel.shape.texteffect).
        */
        wxExcelTextEffectFormat GetTextEffect();

        /**
        Returns a TextFrame object that contains the alignment and anchoring properties for the specified shape.

        [MSDN documentation for Shape.TextFrame](https://learn.microsoft.com/en-us/office/vba/api/excel.shape.textframe).
        */
        wxExcelTextFrame GetTextFrame();

        /**
        TextFrame2  Since Excel 2007.

        [MSDN documentation for Shape.TextFrame2](https://learn.microsoft.com/en-us/office/vba/api/excel.shape.textframe2).
        */
        wxExcelTextFrame2  GetTextFrame2();

        /**
        Returns a ThreeDFormat object that contains 3-D – effect formatting properties for the specified shape.

        [MSDN documentation for Shape.ThreeD](https://learn.microsoft.com/en-us/office/vba/api/excel.shape.threed).
        */
        wxExcelThreeDFormat GetThreeD();

        /**
        Returns a Single value that represents the distance, in points, from the top edge of the topmost shape in the shape range to the top edge of the worksheet.

        [MSDN documentation for Shape.Top](https://learn.microsoft.com/en-us/office/vba/api/excel.shape.top).
        */
        double GetTop();

        /**
        Sets a Single value that represents the distance, in points, from the top edge of the topmost shape in the shape range to the top edge of the worksheet.

        [MSDN documentation for Shape.Top](https://learn.microsoft.com/en-us/office/vba/api/excel.shape.top).
        */
        void SetTop(double top);

        /**
        Returns a Range Represents the cell that lies under the upper-left corner of the specified object.

        [MSDN documentation for Shape.TopLeftCell](https://learn.microsoft.com/en-us/office/vba/api/excel.shape.topleftcell).
        */
        wxExcelRange GetTopLeftCell();

        /**
        Returns a MsoShapeType value that represents the shape type.

        [MSDN documentation for Shape.Type](https://learn.microsoft.com/en-us/office/vba/api/excel.shape.type).
        */
        MsoShapeType  GetType();

        /**
        True if the specified shape is flipped around the vertical axis. Read-only MsoTriState.

        [MSDN documentation for Shape.VerticalFlip](https://learn.microsoft.com/en-us/office/vba/api/excel.shape.verticalflip).
        */
        MsoTriState GetVerticalFlip();

        /**
        Returns the coordinates of the specified freeform drawing's vertices (and control points for Bézier curves) as a series of coordinate pairs. You can use the array returned by this property as an argument to the AddCurve method or AddPolyLine method.

        [MSDN documentation for Shape.Vertices](https://learn.microsoft.com/en-us/office/vba/api/excel.shape.vertices).
        */
        wxVector<wxPoint2DDouble> GetVertices();

        /**
        Returns a MsoTriState value that determines whether the object is visible.

        [MSDN documentation for Shape.Visible](https://learn.microsoft.com/en-us/office/vba/api/excel.shape.visible).
        */
        MsoTriState GetVisible();

        /**
        Sets a MsoTriState value that determines whether the object is visible.

        [MSDN documentation for Shape.Visible](https://learn.microsoft.com/en-us/office/vba/api/excel.shape.visible).
        */
        void SetVisible(MsoTriState visible);

        /**
        Returns a Single value that represents the width, in points, of the object.

        [MSDN documentation for Shape.Width](https://learn.microsoft.com/en-us/office/vba/api/excel.shape.width).
        */
        double GetWidth();

        /**
        Sets a Single value that represents the width, in points, of the object.

        [MSDN documentation for Shape.Width](https://learn.microsoft.com/en-us/office/vba/api/excel.shape.width).
        */
        void SetWidth(double width);

        /**
        Returns the position of the specified shape in the z-order.Read-only

        [MSDN documentation for Shape.ZOrderPosition](https://learn.microsoft.com/en-us/office/vba/api/excel.shape.zorderposition).
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
