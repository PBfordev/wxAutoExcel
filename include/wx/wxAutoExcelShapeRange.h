/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// License:     MIT license
/////////////////////////////////////////////////////////////////////////////


#ifndef _WXAUTOEXCEL_SHAPERANGE_H
#define _WXAUTOEXCEL_SHAPERANGE_H

#include <wx/geometry.h>

#include "wx/wxAutoExcel_defs.h"

#if WXAUTOEXCEL_USE_SHAPES

#include "wx/wxAutoExcel_object.h"
#include "wx/wxAutoExcel_enums.h"

namespace wxAutoExcel {

    /**
    @brief Represents Microsoft Excel ShapeRange object.
    */
   class WXDLLIMPEXP_WXAUTOEXCEL wxExcelShapeRange : public wxExcelObject
    {
    public:
        // ***** METHODS *****

        /**
        Aligns the shapes in the specified range of shapes.

        [MSDN documentation for ShapeRange.Align](https://learn.microsoft.com/en-us/office/vba/api/excel.shaperange.align).
        */
        void Align(MsoAlignCmd alignCmd);

        /**
        Applies to the specified shape formatting that’s been copied by using the PickUp method.

        [MSDN documentation for ShapeRange.Apply](https://learn.microsoft.com/en-us/office/vba/api/excel.shaperange.apply).
        */
        void Apply();

        /**
        Deletes the object.

        [MSDN documentation for ShapeRange.Delete](https://learn.microsoft.com/en-us/office/vba/api/excel.shaperange.delete).
        */
        void Delete();

        /**
        Horizontally or vertically distributes the shapes in the specified range of shapes.

        [MSDN documentation for ShapeRange.Distribute](https://learn.microsoft.com/en-us/office/vba/api/excel.shaperange.distribute).
        */
        void Distribute(MsoDistributeCmd distributeCmd);

        /**
        Duplicates the object and returns a reference to the new copy.

        [MSDN documentation for ShapeRange.Duplicate](https://learn.microsoft.com/en-us/office/vba/api/excel.shaperange.duplicate).
        */
        wxExcelShapeRange Duplicate();

        /**
        Flips the specified shape around its horizontal or vertical axis.

        [MSDN documentation for ShapeRange.Flip](https://learn.microsoft.com/en-us/office/vba/api/excel.shaperange.flip).
        */
        void Flip(MsoFlipCmd flipCmd);

        /**
        Groups the shapes in the specified range.

        [MSDN documentation for ShapeRange.Group](https://learn.microsoft.com/en-us/office/vba/api/excel.shaperange.group).
        */
        wxExcelShape Group();

        /**
        Moves the specified shape horizontally by the specified number of points.

        [MSDN documentation for ShapeRange.IncrementLeft](https://learn.microsoft.com/en-us/office/vba/api/excel.shaperange.incrementleft).
        */
        void IncrementLeft(double increment);

        /**
        Changes the rotation of the specified shape around the z-axis by the specified number of degrees. Use the Rotation property to set the absolute rotation of the shape.

        [MSDN documentation for ShapeRange.IncrementRotation](https://learn.microsoft.com/en-us/office/vba/api/excel.shaperange.incrementrotation).
        */
        void IncrementRotation(double increment);

        /**
        Moves the specified shape vertically by the specified number of points.

        [MSDN documentation for ShapeRange.IncrementTop](https://learn.microsoft.com/en-us/office/vba/api/excel.shaperange.incrementtop).
        */
        void IncrementTop(double increment);

        //@{
        /**
        Returns a single object from a collection.

        [MSDN documentation for ShapeRange.Item](https://learn.microsoft.com/en-us/office/vba/api/excel.shaperange.item).
        */
        wxExcelShape Item(long index);
        wxExcelShape Item(const wxString& name);
        wxExcelShape operator[](long index);
        wxExcelShape operator[](const wxString& name);
        //@}

        /**
        Copies the formatting of the specified shape. Use the Apply method to apply the copied formatting to another shape.

        [MSDN documentation for ShapeRange.PickUp](https://learn.microsoft.com/en-us/office/vba/api/excel.shaperange.pickup).
        */
        void PickUp();

        /**
        Regroups the group that the specified shape range belonged to previously. Returns the regrouped shapes as a single Shape object.

        [MSDN documentation for ShapeRange.Regroup](https://learn.microsoft.com/en-us/office/vba/api/excel.shaperange.regroup).
        */
        wxExcelShape Regroup();

        /**
        This method reroutes all connectors attached to the specified shape; if the specified shape is a connector, it’s rerouted.

        [MSDN documentation for ShapeRange.RerouteConnections](https://learn.microsoft.com/en-us/office/vba/api/excel.shaperange.rerouteconnections).
        */
        void RerouteConnections();

        /**
        Scales the height of the shape by a specified factor. For pictures and OLE objects, you can indicate whether you want to scale the shape relative to the original or the current size. Shapes other than pictures and OLE objects are always scaled relative to their current height.

        [MSDN documentation for ShapeRange.ScaleHeight](https://learn.microsoft.com/en-us/office/vba/api/excel.shaperange.scaleheight).
        */
        void ScaleHeight(double factor, MsoTriState relativeToOriginalSize,
                         MsoScaleFrom* scale = NULL);

        /**
        Scales the width of the shape by a specified factor. For pictures and OLE objects, you can indicate whether you want to scale the shape relative to the original or the current size. Shapes other than pictures and OLE objects are always scaled relative to their current width.

        [MSDN documentation for ShapeRange.ScaleWidth](https://learn.microsoft.com/en-us/office/vba/api/excel.shaperange.scalewidth).
        */
        void ScaleWidth(double factor, MsoTriState relativeToOriginalSize,
                        MsoScaleFrom* scale = NULL);

        /**
        Selects the object.

        [MSDN documentation for ShapeRange.Select](https://learn.microsoft.com/en-us/office/vba/api/excel.shaperange.select).
        */
        void Select();

        /**
        Makes the formatting of the specified shape the default formatting for the shape.

        [MSDN documentation for ShapeRange.SetShapesDefaultProperties](https://learn.microsoft.com/en-us/office/vba/api/excel.shaperange.setshapesdefaultproperties).
        */
        void SetShapesDefaultProperties();

        /**
        Ungroups any grouped shapes in the specified shape or range of shapes. Disassembles pictures and OLE objects within the specified shape or range of shapes.

        [MSDN documentation for ShapeRange.Ungroup](https://learn.microsoft.com/en-us/office/vba/api/excel.shaperange.ungroup).
        */
        wxExcelShapeRange Ungroup();

        /**
        Moves the specified shape in front of or behind other shapes in the collection (that is, changes the shape's position in the z-order).

        [MSDN documentation for ShapeRange.ZOrder](https://learn.microsoft.com/en-us/office/vba/api/excel.shaperange.zorder).
        */
        void ZOrder(MsoZOrderCmd ZOrderCmd);

        // ***** PROPERTIES *****

        /**
        Returns an Adjustments object that contains adjustment values for all the adjustments in the specified shape. Applies to any ShapeRange Represents an AutoShape, WordArt, or a connector.

        [MSDN documentation for ShapeRange.Adjustments](https://learn.microsoft.com/en-us/office/vba/api/excel.shaperange.adjustments).
        */
        wxExcelAdjustments GetAdjustments();

        /**
        Returns the descriptive (alternative) text string for a ShapeRange object when the object is saved to a Web page.

        [MSDN documentation for ShapeRange.AlternativeText](https://learn.microsoft.com/en-us/office/vba/api/excel.shaperange.alternativetext).
        */
        wxString GetAlternativeText();

        /**
        Returns the shape type for the specified Shape or ShapeRange object, which must represent an AutoShape other than a line, freeform drawing, or connector. Read/write MsoAutoShapeType.

        [MSDN documentation for ShapeRange.AutoShapeType](https://learn.microsoft.com/en-us/office/vba/api/excel.shaperange.autoshapetype).
        */
        MsoAutoShapeType GetAutoShapeType();

        /**
        Sets the shape type for the specified Shape or ShapeRange object, which must represent an AutoShape other than a line, freeform drawing, or connector. Read/write MsoAutoShapeType.

        [MSDN documentation for ShapeRange.AutoShapeType](https://learn.microsoft.com/en-us/office/vba/api/excel.shaperange.autoshapetype).
        */
        void SetAutoShapeType(MsoAutoShapeType autoShapeType);

        /**
        Background style Since Excel 2007.

        [MSDN documentation for ShapeRange.BackgroundStyle](https://learn.microsoft.com/en-us/office/vba/api/excel.shaperange.backgroundstyle).
        */
        MsoBackgroundStyleIndex GetBackgroundStyle();

        /**
        Background style Since Excel 2007.

        [MSDN documentation for ShapeRange.BackgroundStyle](https://learn.microsoft.com/en-us/office/vba/api/excel.shaperange.backgroundstyle).
        */
        void SetBackgroundStyle(MsoBackgroundStyleIndex backgroundStyle);

        /**
        Returns a value that indicates how the specified shape appears when the presentation is viewed in black-and-white mode. Read/write MsoBlackWhiteMode.

        [MSDN documentation for ShapeRange.BlackWhiteMode](https://learn.microsoft.com/en-us/office/vba/api/excel.shaperange.blackwhitemode).
        */
        MsoBlackWhiteMode GetBlackWhiteMode();

        /**
        Returns a CalloutFormat object that contains callout formatting properties for the specified shape. Applies to a ShapeRange object that represent line callouts.

        [MSDN documentation for ShapeRange.Callout](https://learn.microsoft.com/en-us/office/vba/api/excel.shaperange.callout).
        */
        wxExcelCalloutFormat GetCallout();

#if WXAUTOEXCEL_USE_CHARTS
        /**
        Since Excel 2007.

        [MSDN documentation for ShapeRange.Chart](https://learn.microsoft.com/en-us/office/vba/api/excel.shaperange.chart).
        */
        wxExcelChart GetChart();
#endif // #if WXAUTOEXCEL_USE_CHARTS

        /**
        Returns msoTrue if the specified shape is a child shape or if all shapes in a shape range are child shapes of the same parent. Read-only MsoTriState.

        [MSDN documentation for ShapeRange.Child](https://learn.microsoft.com/en-us/office/vba/api/excel.shaperange.child).
        */
        MsoTriState GetChild();

        /**
        Returns the number of connection sites on the specified shape.

        [MSDN documentation for ShapeRange.ConnectionSiteCount](https://learn.microsoft.com/en-us/office/vba/api/excel.shaperange.connectionsitecount).
        */
        long GetConnectionSiteCount();

        /**
        True if the specified shape is a connector. Read-only MsoTriState.

        [MSDN documentation for ShapeRange.Connector](https://learn.microsoft.com/en-us/office/vba/api/excel.shaperange.connector).
        */
        MsoTriState GetConnector();

        /**
        Returns a ConnectorFormat object that contains connector formatting properties. Applies to a ShapeRange objects that represent connectors.

        [MSDN documentation for ShapeRange.ConnectorFormat](https://learn.microsoft.com/en-us/office/vba/api/excel.shaperange.connectorformat).
        */
        wxExcelConnectorFormat  GetConnectorFormat();

        /**
        Returns a Long value that represents the number of objects in the collection.

        [MSDN documentation for ShapeRange.Count](https://learn.microsoft.com/en-us/office/vba/api/excel.shaperange.count).
        */
        long GetCount();

        /**
        Returns a FillFormat object for a specified shape or a ChartFillFormat object for a specified chart that contains fill formatting properties for the shape or chart.

        [MSDN documentation for ShapeRange.Fill](https://learn.microsoft.com/en-us/office/vba/api/excel.shaperange.fill).
        */
        wxExcelFillFormat GetFill();

        /**
        GlowFormat Since Excel 2007.

        [MSDN documentation for ShapeRange.Glow](https://learn.microsoft.com/en-us/office/vba/api/excel.shaperange.glow).
        */
        wxExcelGlowFormat GetGlow();

        /**
        Returns an MsoGraphicStyleIndex constant that represents the style of an SVG graphic.

        [Excel VBA documentation for ShapeRange.GraphicStyle](https://learn.microsoft.com/en-us/office/vba/api/excel.shaperange.graphicstyle)
        */
        MsoGraphicStyleIndex GetGraphicStyle();

        /**
        Sets an MsoGraphicStyleIndex constant that represents the style of an SVG graphic.

        [Excel VBA documentation for ShapeRange.GraphicStyle](https://learn.microsoft.com/en-us/office/vba/api/excel.shaperange.graphicstyle)
        */
        void SetGraphicStyle(MsoGraphicStyleIndex graphicStyle);


        /**
        Returns a GroupShapes Represents the individual shapes in the specified group. Use the Item method of the GroupShapes object to return a single shape from the group. Applies to ShapeRange objects that represent grouped shapes.

        [MSDN documentation for ShapeRange.GroupItems](https://learn.microsoft.com/en-us/office/vba/api/excel.shaperange.groupitems).
        */
        wxExcelGroupShapes GetGroupItems();

        /**
        Returns msoTrue if the shape range has a chart. Since Excel 2007.

        [MSDN documentation for ShapeRange.HasChart](https://learn.microsoft.com/en-us/office/vba/api/excel.shaperange.haschart).
        */
        MsoTriState GetHasChart();

        /**
        Returns a Single value that represents the height, in points, of the object.

        [MSDN documentation for ShapeRange.Height](https://learn.microsoft.com/en-us/office/vba/api/excel.shaperange.height).
        */
        double GetHeight();

        /**
        True if the specified shape is flipped around the horizontal axis. Read-only MsoTriState.

        [MSDN documentation for ShapeRange.HorizontalFlip](https://learn.microsoft.com/en-us/office/vba/api/excel.shaperange.horizontalflip).
        */
        MsoTriState GetHorizontalFlip();

        /**
        Returns a Long value that represents the type for the specified object.

        [MSDN documentation for ShapeRange.ID](https://learn.microsoft.com/en-us/office/vba/api/excel.shaperange.id).
        */
        long GetID();

        /**
        Returns a Single value that represents the distance, in points, from the left edge of the object to the left edge of column A (on a worksheet) or the left edge of the chart area (on a chart).

        [MSDN documentation for ShapeRange.Left](https://learn.microsoft.com/en-us/office/vba/api/excel.shaperange.left).
        */
        double GetLeft();

        /**
        Sets a Single value that represents the distance, in points, from the left edge of the object to the left edge of column A (on a worksheet) or the left edge of the chart area (on a chart).

        [MSDN documentation for ShapeRange.Left](https://learn.microsoft.com/en-us/office/vba/api/excel.shaperange.left).
        */
        void SetLeft(double left);

        /**
        Returns a LineFormat object that contains line formatting properties for the specified shape. (For a line, the LineFormat object represents the line itself; for a shape with a border, the LineFormat object represents the border).

        [MSDN documentation for ShapeRange.Line](https://learn.microsoft.com/en-us/office/vba/api/excel.shaperange.line).
        */
        wxExcelLineFormat GetLine();

        /**
        True if the specified shape retains its original proportions when you resize it. False if you can change the height and width of the shape independently of one another when you resize it. Read/write MsoTriState.

        [MSDN documentation for ShapeRange.LockAspectRatio](https://learn.microsoft.com/en-us/office/vba/api/excel.shaperange.lockaspectratio).
        */
        MsoTriState GetLockAspectRatio();

        /**
        True if the specified shape retains its original proportions when you resize it. False if you can change the height and width of the shape independently of one another when you resize it. Read/write MsoTriState.

        [MSDN documentation for ShapeRange.LockAspectRatio](https://learn.microsoft.com/en-us/office/vba/api/excel.shaperange.lockaspectratio).
        */
        void SetLockAspectRatio(MsoTriState lockAspectRatio);

        /**
        Returns a Model3DFormat object that contains Model3D properties.

        [Excel VBA documentation for ShapeRange.Model3D](https://learn.microsoft.com/en-us/office/vba/api/excel.shaperange.model3d)
        */
        wxExcelModel3DFormat GetModel3D();

        /**
        Returns a String value representing the name of the object.

        [MSDN documentation for ShapeRange.Name](https://learn.microsoft.com/en-us/office/vba/api/excel.shaperange.name).
        */
        wxString GetName();

        /**
        Sets a String value representing the name of the object.

        [MSDN documentation for ShapeRange.Name](https://learn.microsoft.com/en-us/office/vba/api/excel.shaperange.name).
        */
        void SetName(const wxString& name);

        /**
        Returns a ShapeNodes collection that represents the geometric description of the specified shape.

        [MSDN documentation for ShapeRange.Nodes](https://learn.microsoft.com/en-us/office/vba/api/excel.shaperange.nodes).
        */
        wxExcelShapeNodes GetNodes();

        /**
        Returns a Shape Represents the common parent shape of a child shape or a range of child shapes.

        [MSDN documentation for ShapeRange.ParentGroup](https://learn.microsoft.com/en-us/office/vba/api/excel.shaperange.parentgroup).
        */
        wxExcelShape GetParentGroup();

        /**
        Returns a PictureFormat object that contains picture formatting properties for the specified shape. Applies to a ShapeRange object that represent pictures or OLE objects.

        [MSDN documentation for ShapeRange.PictureFormat](https://learn.microsoft.com/en-us/office/vba/api/excel.shaperange.pictureformat).
        */
        wxExcelPictureFormat GetPictureFormat();

        /**
        Returns a ReflectionFormat object for a specified shape range that contains reflection formatting properties. Since Excel 2007.

        [MSDN documentation for ShapeRange.Reflection](https://learn.microsoft.com/en-us/office/vba/api/excel.shaperange.reflection).
        */
        wxExcelReflectionFormat GetReflection();

        /**
        Returns the rotation of the shape, in degrees. Read/write Single.

        [MSDN documentation for ShapeRange.Rotation](https://learn.microsoft.com/en-us/office/vba/api/excel.shaperange.rotation).
        */
        double GetRotation();

        /**
        Returns a read-only ShadowFormat object that contains shadow formatting properties for the specified shape or shapes.

        [MSDN documentation for ShapeRange.Shadow](https://learn.microsoft.com/en-us/office/vba/api/excel.shaperange.shadow).
        */
        wxExcelShadowFormat GetShadow();

        /**
        Shape style Since Excel 2007.

        [MSDN documentation for ShapeRange.ShapeStyle](https://learn.microsoft.com/en-us/office/vba/api/excel.shaperange.shapestyle).
        */
        MsoShapeStyleIndex GetShapeStyle();

        /**
        Shape style Since Excel 2007.

        [MSDN documentation for ShapeRange.ShapeStyle](https://learn.microsoft.com/en-us/office/vba/api/excel.shaperange.shapestyle).
        */
        void SetShapeStyle(MsoShapeStyleIndex shapeStyle);

        /**
        SoftEdgeFormat Since Excel 2007.

        [MSDN documentation for ShapeRange.SoftEdge](https://learn.microsoft.com/en-us/office/vba/api/excel.shaperange.softedge).
        */
        wxExcelSoftEdgeFormat GetSoftEdge();

        /**
        Returns a TextEffectFormat object that contains text-effect formatting properties for the specified shape.

        [MSDN documentation for ShapeRange.TextEffect](https://learn.microsoft.com/en-us/office/vba/api/excel.shaperange.texteffect).
        */
        wxExcelTextEffectFormat GetTextEffect();

        /**
        Returns a TextFrame object that contains the alignment and anchoring properties for the specified shape.

        [MSDN documentation for ShapeRange.TextFrame](https://learn.microsoft.com/en-us/office/vba/api/excel.shaperange.textframe).
        */
        wxExcelTextFrame GetTextFrame();

        /**
        Read-only Since Excel 2007.

        [MSDN documentation for ShapeRange.TextFrame2](https://learn.microsoft.com/en-us/office/vba/api/excel.shaperange.textframe2).
        */
        wxExcelTextFrame2 GetTextFrame2();

        /**
        Returns a ThreeDFormat object that contains 3-D – effect formatting properties for the specified shape.

        [MSDN documentation for ShapeRange.ThreeD](https://learn.microsoft.com/en-us/office/vba/api/excel.shaperange.threed).
        */
        wxExcelThreeDFormat GetThreeD();

        /**
        Returns a Single value that represents the distance, in points, from the top edge of the topmost shape in the shape range to the top edge of the worksheet.

        [MSDN documentation for ShapeRange.Top](https://learn.microsoft.com/en-us/office/vba/api/excel.shaperange.top).
        */
        double GetTop();

        /**
        Returns a MsoShapeType value that represents the shape type.

        [MSDN documentation for ShapeRange.Type](https://learn.microsoft.com/en-us/office/vba/api/excel.shaperange.type).
        */
        MsoShapeType  GetType();

        /**
        True if the specified shape is flipped around the vertical axis. Read-only MsoTriState.

        [MSDN documentation for ShapeRange.VerticalFlip](https://learn.microsoft.com/en-us/office/vba/api/excel.shaperange.verticalflip).
        */
        MsoTriState GetVerticalFlip();

        /**
        Returns the coordinates of the specified freeform drawing's vertices (and control points for Bézier curves) as a series of coordinate pairs. You can use the array returned by this property as an argument to the AddCurve method or AddPolyLine method.

        [MSDN documentation for ShapeRange.Vertices](https://learn.microsoft.com/en-us/office/vba/api/excel.shaperange.vertices).
        */
        wxVector<wxPoint2DDouble> GetVertices();

        /**
        Returns a MsoTriState value that determines whether the object is visible.

        [MSDN documentation for ShapeRange.Visible](https://learn.microsoft.com/en-us/office/vba/api/excel.shaperange.visible).
        */
        MsoTriState GetVisible();

        /**
        Returns a Single value that represents the width, in points, of the object.

        [MSDN documentation for ShapeRange.Width](https://learn.microsoft.com/en-us/office/vba/api/excel.shaperange.width).
        */
        double GetWidth();

        /**
        Returns the position of the specified shape in the z-order.Read-only

        [MSDN documentation for ShapeRange.ZOrderPosition](https://learn.microsoft.com/en-us/office/vba/api/excel.shaperange.zorderposition).
        */
        long GetZOrderPosition();

        /**
        Returns "ShapeRange".
        */
        virtual wxString GetAutoExcelObjectName_() const { return wxS("ShapeRange"); }
    };


} // namespace wxAutoExcel

#endif // #if WXAUTOEXCEL_USE_SHAPES

#endif //_WXAUTOEXCEL_SHAPERANGE_H
