/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Modified by:
// RCS-ID:      $Id: $
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
/////////////////////////////////////////////////////////////////////////////


#ifndef _WXAUTOEXCEL_SHAPERANGE_H
#define _WXAUTOEXCEL_SHAPERANGE_H

#include "wx/wxAutoExcel_defs.h"

#if WXAUTOEXCEL_USE_SHAPES

#include "wx/wxAutoExcelObject.h"
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

        [MSDN documentation for ShapeRange.Align](http://msdn.microsoft.com/en-us/library/bb209642).
        */
        void Align(MsoAlignCmd alignCmd);

        /**
        Applies to the specified shape formatting that’s been copied by using the PickUp method.

        [MSDN documentation for ShapeRange.Apply](http://msdn.microsoft.com/en-us/library/bb212176).
        */
        void Apply();

        /**
        Deletes the object.

        [MSDN documentation for ShapeRange.Delete](http://msdn.microsoft.com/en-us/library/bb212177).
        */
        void Delete();

        /**
        Horizontally or vertically distributes the shapes in the specified range of shapes.

        [MSDN documentation for ShapeRange.Distribute](http://msdn.microsoft.com/en-us/library/bb209792).
        */
        void Distribute(MsoDistributeCmd distributeCmd);

        /**
        Duplicates the object and returns a reference to the new copy.

        [MSDN documentation for ShapeRange.Duplicate](http://msdn.microsoft.com/en-us/library/bb212179).
        */
        wxExcelShapeRange Duplicate();

        /**
        Flips the specified shape around its horizontal or vertical axis.

        [MSDN documentation for ShapeRange.Flip](http://msdn.microsoft.com/en-us/library/bb212181).
        */
        void Flip(MsoFlipCmd flipCmd);

        /**
        Groups the shapes in the specified range.

        [MSDN documentation for ShapeRange.Group](http://msdn.microsoft.com/en-us/library/bb212182).
        */
        wxExcelShape Group();

        /**
        Moves the specified shape horizontally by the specified number of points.

        [MSDN documentation for ShapeRange.IncrementLeft](http://msdn.microsoft.com/en-us/library/bb212184).
        */
        void IncrementLeft(double increment);

        /**
        Changes the rotation of the specified shape around the z-axis by the specified number of degrees. Use the Rotation property to set the absolute rotation of the shape.

        [MSDN documentation for ShapeRange.IncrementRotation](http://msdn.microsoft.com/en-us/library/bb212188).
        */
        void IncrementRotation(double increment);

        /**
        Moves the specified shape vertically by the specified number of points.

        [MSDN documentation for ShapeRange.IncrementTop](http://msdn.microsoft.com/en-us/library/bb212191).
        */
        void IncrementTop(double increment);

        //@{
        /**
        Returns a single object from a collection.

        [MSDN documentation for ShapeRange.Item](http://msdn.microsoft.com/en-us/library/bb212192).
        */
        wxExcelShape Item(long index);
        wxExcelShape Item(const wxString& name);
        wxExcelShape operator[](long index);
        wxExcelShape operator[](const wxString& name);
        //@}

        /**
        Copies the formatting of the specified shape. Use the Apply method to apply the copied formatting to another shape.

        [MSDN documentation for ShapeRange.PickUp](http://msdn.microsoft.com/en-us/library/bb212196).
        */
        void PickUp();

        /**
        Regroups the group that the specified shape range belonged to previously. Returns the regrouped shapes as a single Shape object.

        [MSDN documentation for ShapeRange.Regroup](http://msdn.microsoft.com/en-us/library/bb223583).
        */
        wxExcelShape Regroup();

        /**
        This method reroutes all connectors attached to the specified shape; if the specified shape is a connector, it’s rerouted.

        [MSDN documentation for ShapeRange.RerouteConnections](http://msdn.microsoft.com/en-us/library/bb212197).
        */
        void RerouteConnections();

        /**
        Scales the height of the shape by a specified factor. For pictures and OLE objects, you can indicate whether you want to scale the shape relative to the original or the current size. Shapes other than pictures and OLE objects are always scaled relative to their current height.

        [MSDN documentation for ShapeRange.ScaleHeight](http://msdn.microsoft.com/en-us/library/bb213974).
        */
        void ScaleHeight(double factor, MsoTriState relativeToOriginalSize,
                         MsoScaleFrom* scale = NULL);

        /**
        Scales the width of the shape by a specified factor. For pictures and OLE objects, you can indicate whether you want to scale the shape relative to the original or the current size. Shapes other than pictures and OLE objects are always scaled relative to their current width.

        [MSDN documentation for ShapeRange.ScaleWidth](http://msdn.microsoft.com/en-us/library/bb213982).
        */
        void ScaleWidth(double factor, MsoTriState relativeToOriginalSize,
                        MsoScaleFrom* scale = NULL);

        /**
        Selects the object.

        [MSDN documentation for ShapeRange.Select](http://msdn.microsoft.com/en-us/library/bb213986).
        */
        void Select();

        /**
        Makes the formatting of the specified shape the default formatting for the shape.

        [MSDN documentation for ShapeRange.SetShapesDefaultProperties](http://msdn.microsoft.com/en-us/library/bb213992).
        */
        void SetShapesDefaultProperties();

        /**
        Ungroups any grouped shapes in the specified shape or range of shapes. Disassembles pictures and OLE objects within the specified shape or range of shapes.

        [MSDN documentation for ShapeRange.Ungroup](http://msdn.microsoft.com/en-us/library/bb213997).
        */
        wxExcelShapeRange Ungroup();

        /**
        Moves the specified shape in front of or behind other shapes in the collection (that is, changes the shape's position in the z-order).

        [MSDN documentation for ShapeRange.ZOrder](http://msdn.microsoft.com/en-us/library/bb214008).
        */
        void ZOrder(MsoZOrderCmd ZOrderCmd);

        // ***** PROPERTIES *****

        /**
        Returns an Adjustments object that contains adjustment values for all the adjustments in the specified shape. Applies to any ShapeRange Represents an AutoShape, WordArt, or a connector.

        [MSDN documentation for ShapeRange.Adjustments](http://msdn.microsoft.com/en-us/library/bb213864).
        */
        wxExcelAdjustments GetAdjustments();

        /**
        Returns the descriptive (alternative) text string for a ShapeRange object when the object is saved to a Web page.

        [MSDN documentation for ShapeRange.AlternativeText](http://msdn.microsoft.com/en-us/library/bb213876).
        */
        wxString GetAlternativeText();

        /**
        Returns the shape type for the specified Shape or ShapeRange object, which must represent an AutoShape other than a line, freeform drawing, or connector. Read/write MsoAutoShapeType.

        [MSDN documentation for ShapeRange.AutoShapeType](http://msdn.microsoft.com/en-us/library/bb213887).
        */
        MsoAutoShapeType GetAutoShapeType();

        /**
        Sets the shape type for the specified Shape or ShapeRange object, which must represent an AutoShape other than a line, freeform drawing, or connector. Read/write MsoAutoShapeType.

        [MSDN documentation for ShapeRange.AutoShapeType](http://msdn.microsoft.com/en-us/library/bb213887).
        */
        void SetAutoShapeType(MsoAutoShapeType autoShapeType);

        /**
        Background style Since Excel 2007.

        [MSDN documentation for ShapeRange.BackgroundStyle](http://msdn.microsoft.com/en-us/library/bb240879).
        */
        MsoBackgroundStyleIndex GetBackgroundStyle();

        /**
        Background style Since Excel 2007.

        [MSDN documentation for ShapeRange.BackgroundStyle](http://msdn.microsoft.com/en-us/library/bb240879).
        */
        void SetBackgroundStyle(MsoBackgroundStyleIndex backgroundStyle);

        /**
        Returns a value that indicates how the specified shape appears when the presentation is viewed in black-and-white mode. Read/write MsoBlackWhiteMode.

        [MSDN documentation for ShapeRange.BlackWhiteMode](http://msdn.microsoft.com/en-us/library/bb213897).
        */
        MsoBlackWhiteMode GetBlackWhiteMode();

        /**
        Returns a CalloutFormat object that contains callout formatting properties for the specified shape. Applies to a ShapeRange object that represent line callouts.

        [MSDN documentation for ShapeRange.Callout](http://msdn.microsoft.com/en-us/library/bb213901).
        */
        wxExcelCalloutFormat GetCallout();

#if WXAUTOEXCEL_USE_CHARTS
        /**
        Since Excel 2007.

        [MSDN documentation for ShapeRange.Chart](http://msdn.microsoft.com/en-us/library/bb213908).
        */
        wxExcelChart GetChart();
#endif // #if WXAUTOEXCEL_USE_CHARTS

        /**
        Returns msoTrue if the specified shape is a child shape or if all shapes in a shape range are child shapes of the same parent. Read-only MsoTriState.

        [MSDN documentation for ShapeRange.Child](http://msdn.microsoft.com/en-us/library/bb213921).
        */
        MsoTriState GetChild();

        /**
        Returns the number of connection sites on the specified shape.

        [MSDN documentation for ShapeRange.ConnectionSiteCount](http://msdn.microsoft.com/en-us/library/bb213929).
        */
        long GetConnectionSiteCount();

        /**
        True if the specified shape is a connector. Read-only MsoTriState.

        [MSDN documentation for ShapeRange.Connector](http://msdn.microsoft.com/en-us/library/bb213937).
        */
        MsoTriState GetConnector();

        /**
        Returns a ConnectorFormat object that contains connector formatting properties. Applies to a ShapeRange objects that represent connectors.

        [MSDN documentation for ShapeRange.ConnectorFormat](http://msdn.microsoft.com/en-us/library/bb213941).
        */
        wxExcelConnectorFormat  GetConnectorFormat();

        /**
        Returns a Long value that represents the number of objects in the collection.

        [MSDN documentation for ShapeRange.Count](http://msdn.microsoft.com/en-us/library/bb213947).
        */
        long GetCount();

        /**
        Returns a FillFormat object for a specified shape or a ChartFillFormat object for a specified chart that contains fill formatting properties for the shape or chart.

        [MSDN documentation for ShapeRange.Fill](http://msdn.microsoft.com/en-us/library/bb213953).
        */
        wxExcelFillFormat GetFill();

        /**
        GlowFormat Since Excel 2007.

        [MSDN documentation for ShapeRange.Glow](http://msdn.microsoft.com/en-us/library/bb242576).
        */
        wxExcelGlowFormat GetGlow();

        /**
        Returns a GroupShapes Represents the individual shapes in the specified group. Use the Item method of the GroupShapes object to return a single shape from the group. Applies to ShapeRange objects that represent grouped shapes.

        [MSDN documentation for ShapeRange.GroupItems](http://msdn.microsoft.com/en-us/library/bb213962).
        */
        wxExcelGroupShapes GetGroupItems();

        /**
        Returns true If has a chart. Since Excel 2007.

        [MSDN documentation for ShapeRange.HasChart](http://msdn.microsoft.com/en-us/library/bb237749).
        */
        bool GetHasChart();

        /**
        Returns a Single value that represents the height, in points, of the object.

        [MSDN documentation for ShapeRange.Height](http://msdn.microsoft.com/en-us/library/bb237759).
        */
        double GetHeight();

        /**
        True if the specified shape is flipped around the horizontal axis. Read-only MsoTriState.

        [MSDN documentation for ShapeRange.HorizontalFlip](http://msdn.microsoft.com/en-us/library/bb237784).
        */
        MsoTriState GetHorizontalFlip();

        /**
        Returns a Long value that represents the type for the specified object.

        [MSDN documentation for ShapeRange.ID](http://msdn.microsoft.com/en-us/library/bb237787).
        */
        long GetID();

        /**
        Returns a Single value that represents the distance, in points, from the left edge of the object to the left edge of column A (on a worksheet) or the left edge of the chart area (on a chart).

        [MSDN documentation for ShapeRange.Left](http://msdn.microsoft.com/en-us/library/bb237802).
        */
        double GetLeft();

        /**
        Sets a Single value that represents the distance, in points, from the left edge of the object to the left edge of column A (on a worksheet) or the left edge of the chart area (on a chart).

        [MSDN documentation for ShapeRange.Left](http://msdn.microsoft.com/en-us/library/bb237802).
        */
        void SetLeft(double left);

        /**
        Returns a LineFormat object that contains line formatting properties for the specified shape. (For a line, the LineFormat object represents the line itself; for a shape with a border, the LineFormat object represents the border).

        [MSDN documentation for ShapeRange.Line](http://msdn.microsoft.com/en-us/library/bb237824).
        */
        wxExcelLineFormat GetLine();

        /**
        True if the specified shape retains its original proportions when you resize it. False if you can change the height and width of the shape independently of one another when you resize it. Read/write MsoTriState.

        [MSDN documentation for ShapeRange.LockAspectRatio](http://msdn.microsoft.com/en-us/library/bb237830).
        */
        MsoTriState GetLockAspectRatio();

        /**
        True if the specified shape retains its original proportions when you resize it. False if you can change the height and width of the shape independently of one another when you resize it. Read/write MsoTriState.

        [MSDN documentation for ShapeRange.LockAspectRatio](http://msdn.microsoft.com/en-us/library/bb237830).
        */
        void SetLockAspectRatio(MsoTriState lockAspectRatio);

        /**
        Returns a String value representing the name of the object.

        [MSDN documentation for ShapeRange.Name](http://msdn.microsoft.com/en-us/library/bb237842).
        */
        wxString GetName();

        /**
        Sets a String value representing the name of the object.

        [MSDN documentation for ShapeRange.Name](http://msdn.microsoft.com/en-us/library/bb237842).
        */
        void SetName(const wxString& name);

        /**
        Returns a ShapeNodes collection that represents the geometric description of the specified shape.

        [MSDN documentation for ShapeRange.Nodes](http://msdn.microsoft.com/en-us/library/bb237852).
        */
        wxExcelShapeNodes GetNodes();        

        /**
        Returns a Shape Represents the common parent shape of a child shape or a range of child shapes.

        [MSDN documentation for ShapeRange.ParentGroup](http://msdn.microsoft.com/en-us/library/bb237863).
        */
        wxExcelShape GetParentGroup();

        /**
        Returns a PictureFormat object that contains picture formatting properties for the specified shape. Applies to a ShapeRange object that represent pictures or OLE objects.

        [MSDN documentation for ShapeRange.PictureFormat](http://msdn.microsoft.com/en-us/library/bb237870).
        */
        wxExcelPictureFormat GetPictureFormat();

        /**
        Returns a ReflectionFormat object for a specified shape range that contains reflection formatting properties. Since Excel 2007.

        [MSDN documentation for ShapeRange.Reflection](http://msdn.microsoft.com/en-us/library/bb242579).
        */
        wxExcelReflectionFormat GetReflection();

        /**
        Returns the rotation of the shape, in degrees. Read/write Single.

        [MSDN documentation for ShapeRange.Rotation](http://msdn.microsoft.com/en-us/library/bb215053).
        */
        double GetRotation();

        /**
        Returns a read-only ShadowFormat object that contains shadow formatting properties for the specified shape or shapes.

        [MSDN documentation for ShapeRange.Shadow](http://msdn.microsoft.com/en-us/library/bb215057).
        */
        wxExcelShadowFormat GetShadow();

        /**
        Shape style Since Excel 2007.

        [MSDN documentation for ShapeRange.ShapeStyle](http://msdn.microsoft.com/en-us/library/bb240891).
        */
        MsoShapeStyleIndex GetShapeStyle();

        /**
        Shape style Since Excel 2007.

        [MSDN documentation for ShapeRange.ShapeStyle](http://msdn.microsoft.com/en-us/library/bb240891).
        */
        void SetShapeStyle(MsoShapeStyleIndex shapeStyle);

        /**
        SoftEdgeFormat Since Excel 2007.

        [MSDN documentation for ShapeRange.SoftEdge](http://msdn.microsoft.com/en-us/library/bb242585).
        */
        wxExcelSoftEdgeFormat GetSoftEdge();

        /**
        Returns a TextEffectFormat object that contains text-effect formatting properties for the specified shape.

        [MSDN documentation for ShapeRange.TextEffect](http://msdn.microsoft.com/en-us/library/bb215060).
        */
        wxExcelTextEffectFormat GetTextEffect();

        /**
        Returns a TextFrame object that contains the alignment and anchoring properties for the specified shape.

        [MSDN documentation for ShapeRange.TextFrame](http://msdn.microsoft.com/en-us/library/bb215063).
        */
        wxExcelTextFrame GetTextFrame();

        /**
        Read-only Since Excel 2007.

        [MSDN documentation for ShapeRange.TextFrame2](http://msdn.microsoft.com/en-us/library/bb215064).
        */
        wxExcelTextFrame2 GetTextFrame2();

        /**
        Returns a ThreeDFormat object that contains 3-D – effect formatting properties for the specified shape.

        [MSDN documentation for ShapeRange.ThreeD](http://msdn.microsoft.com/en-us/library/bb215067).
        */
        wxExcelThreeDFormat GetThreeD();

        /**
        Returns a Single value that represents the distance, in points, from the top edge of the topmost shape in the shape range to the top edge of the worksheet.

        [MSDN documentation for ShapeRange.Top](http://msdn.microsoft.com/en-us/library/bb215068).
        */
        double GetTop();

        /**
        Returns a MsoShapeType value that represents the shape type.

        [MSDN documentation for ShapeRange.Type](http://msdn.microsoft.com/en-us/library/bb215070).
        */
        MsoShapeType  GetType();

        /**
        True if the specified shape is flipped around the vertical axis. Read-only MsoTriState.

        [MSDN documentation for ShapeRange.VerticalFlip](http://msdn.microsoft.com/en-us/library/bb215071).
        */
        MsoTriState GetVerticalFlip();

        /**
        Returns the coordinates of the specified freeform drawing's vertices (and control points for Bézier curves) as a series of coordinate pairs. You can use the array returned by this property as an argument to the AddCurve method or AddPolyLine method.

        [MSDN documentation for ShapeRange.Vertices](http://msdn.microsoft.com/en-us/library/bb215074).
        */
        wxVector<wxPoint2DDouble> GetVertices();

        /**
        Returns a MsoTriState value that determines whether the object is visible.

        [MSDN documentation for ShapeRange.Visible](http://msdn.microsoft.com/en-us/library/bb215077).
        */
        MsoTriState GetVisible();

        /**
        Returns a Single value that represents the width, in points, of the object.

        [MSDN documentation for ShapeRange.Width](http://msdn.microsoft.com/en-us/library/bb215079).
        */
        double GetWidth();

        /**
        Returns the position of the specified shape in the z-order.Read-only

        [MSDN documentation for ShapeRange.ZOrderPosition](http://msdn.microsoft.com/en-us/library/bb215084).
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
