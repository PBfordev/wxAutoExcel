/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// License:     MIT license
/////////////////////////////////////////////////////////////////////////////


#ifndef _WXAUTOEXCEL_TEXTFRAME2_H
#define _WXAUTOEXCEL_TEXTFRAME2_H

#include "wx/wxAutoExcel_defs.h"

#if WXAUTOEXCEL_USE_SHAPES || WXAUTOEXCEL_USE_CHARTS

#include "wx/wxAutoExcel_object.h"
#include "wx/wxAutoExcel_enums.h"

namespace wxAutoExcel {


    /**
    @brief Represents Microsoft Excel TextFrame2.
    */
   class WXDLLIMPEXP_WXAUTOEXCEL wxExcelTextFrame2 : public wxExcelObject
    {
    public:
        // ***** METHODS *****

        /**
        Deletes the text from a text frame and all the associated text properties.

        [MSDN documentation for TextFrame2.DeleteText](https://learn.microsoft.com/en-us/office/vba/api/excel.textframe2.deletetext).
        */
        void DeleteText();

        // ***** PROPERTIES *****

        /**
        The size of the specified object that changes automatically to fit text within its boundaries. Since Excel 2007.

        [MSDN documentation for TextFrame2.AutoSize](https://learn.microsoft.com/en-us/office/vba/api/excel.textframe2.autosize).
        */
        MsoAutoSize GetAutoSize();

        /**
        Returns the TextColumn2 Represents the columns within the text frame. Since Excel 2007.

        [MSDN documentation for TextFrame2.Column](https://learn.microsoft.com/en-us/office/vba/api/excel.textframe2.column).
        */
        wxExcelTextColumn2 GetColumn();

        /**
        Returns a 32-bit integer that indicates the application in which this object was created. Since Excel 2007.

        [MSDN documentation for TextFrame2.Creator](https://learn.microsoft.com/en-us/office/vba/api/excel.textframe2.creator).
        */
        long GetCreator();

        /**
        Returns whether the specified text frame has text. Since Excel 2007.

        [MSDN documentation for TextFrame2.HasText](https://learn.microsoft.com/en-us/office/vba/api/excel.textframe2.hastext).
        */
        MsoTriState GetHasText();

        /**
        Returns the horizontal anchor type for the specified text. Since Excel 2007.

        [MSDN documentation for TextFrame2.HorizontalAnchor](https://learn.microsoft.com/en-us/office/vba/api/excel.textframe2.horizontalanchor).
        */
        MsoHorizontalAnchor GetHorizontalAnchor();

        /**
        Sets the horizontal anchor type for the specified text. Since Excel 2007.

        [MSDN documentation for TextFrame2.HorizontalAnchor](https://learn.microsoft.com/en-us/office/vba/api/excel.textframe2.horizontalanchor).
        */
        void SetHorizontalAnchor(MsoHorizontalAnchor horizontalAnchor);

        /**
        Returns the distance (in points) between the bottom of the text frame and the bottom of the inscribed rectangle of the shape that contains the text. Since Excel 2007.

        [MSDN documentation for TextFrame2.MarginBottom](https://learn.microsoft.com/en-us/office/vba/api/excel.textframe2.marginbottom).
        */
        double GetMarginBottom();

        /**
        Sets the distance (in points) between the bottom of the text frame and the bottom of the inscribed rectangle of the shape that contains the text. Since Excel 2007.

        [MSDN documentation for TextFrame2.MarginBottom](https://learn.microsoft.com/en-us/office/vba/api/excel.textframe2.marginbottom).
        */
        void SetMarginBottom(double marginBottom);

        /**
        Returns the distance (in points) between the left edge of the text frame and the left edge of the inscribed rectangle of the shape that contains the text. Since Excel 2007.

        [MSDN documentation for TextFrame2.MarginLeft](https://learn.microsoft.com/en-us/office/vba/api/excel.textframe2.marginleft).
        */
        double GetMarginLeft();

        /**
        Sets the distance (in points) between the left edge of the text frame and the left edge of the inscribed rectangle of the shape that contains the text. Since Excel 2007.

        [MSDN documentation for TextFrame2.MarginLeft](https://learn.microsoft.com/en-us/office/vba/api/excel.textframe2.marginleft).
        */
        void SetMarginLeft(double marginLeft);

        /**
        Returns the distance (in points) between the right edge of the text frame and the right edge of the inscribed rectangle of the shape that contains the text. Since Excel 2007.

        [MSDN documentation for TextFrame2.MarginRight](https://learn.microsoft.com/en-us/office/vba/api/excel.textframe2.marginright).
        */
        double GetMarginRight();

        /**
        Sets the distance (in points) between the right edge of the text frame and the right edge of the inscribed rectangle of the shape that contains the text. Since Excel 2007.

        [MSDN documentation for TextFrame2.MarginRight](https://learn.microsoft.com/en-us/office/vba/api/excel.textframe2.marginright).
        */
        void SetMarginRight(double marginRight);

        /**
        Returns the distance (in points) between the top of the text frame and the top of the inscribed rectangle of the shape that contains the textSince Excel 2007.

        [MSDN documentation for TextFrame2.MarginTop](https://learn.microsoft.com/en-us/office/vba/api/excel.textframe2.margintop).
        */
        double GetMarginTop();

        /**
        Sets the distance (in points) between the top of the text frame and the top of the inscribed rectangle of the shape that contains the text. Since Excel 2007.

        [MSDN documentation for TextFrame2.MarginTop](https://learn.microsoft.com/en-us/office/vba/api/excel.textframe2.margintop).
        */
        void SetMarginTop(double marginTop);

        /**
        Returns a value that specifies if the text on a shape is rotated if the shape itself is being rotated. 

        [MSDN documentation for TextFrame2.NoTextRotation](https://learn.microsoft.com/en-us/office/vba/api/excel.textframe2.notextrotation).
        */
        MsoTriState GetNoTextRotation();

        /**
        Sets a value that specifies if the text on a shape is rotated if the shape itself is being rotated. 

        [MSDN documentation for TextFrame2.NoTextRotation](https://learn.microsoft.com/en-us/office/vba/api/excel.textframe2.notextrotation).
        */
        void SetNoTextRotation(MsoTriState rotation);

        /**
        Returns a value that represents the text frame orientation. Since Excel 2007.

        [MSDN documentation for TextFrame2.Orientation](https://learn.microsoft.com/en-us/office/vba/api/excel.textframe2.orientation).
        */
        MsoTextOrientation GetOrientation();

        /**
        Sets a value that represents the text frame orientation. Since Excel 2007.

        [MSDN documentation for TextFrame2.Orientation](https://learn.microsoft.com/en-us/office/vba/api/excel.textframe2.orientation).
        */
        void SetOrientation(MsoTextOrientation orientation);

        /**
        Returns the path type for the specified text frame. Since Excel 2007.

        [MSDN documentation for TextFrame2.PathFormat](https://learn.microsoft.com/en-us/office/vba/api/excel.textframe2.pathformat).
        */
        MsoPathFormat GetPathFormat();

        /**
        Sets the path type for the specified text frame.

        [MSDN documentation for TextFrame2.PathFormat](https://learn.microsoft.com/en-us/office/vba/api/excel.textframe2.pathformat).
        */
        void SetPathFormat(MsoPathFormat pathFormat);

        /**
        Returns the TextRange2 Represents the text in the object. Since Excel 2007.

        [MSDN documentation for TextFrame2.TextRange](https://learn.microsoft.com/en-us/office/vba/api/excel.textframe2.textrange).
        */
        wxExcelTextRange2 GetTextRange();

        /**
        Returns a ThreeDFormat object that contains 3-D–effect formatting properties for the specified text. Since Excel 2007.

        [MSDN documentation for TextFrame2.ThreeD](https://learn.microsoft.com/en-us/office/vba/api/excel.textframe2.threed).
        */
        wxExcelThreeDFormat GetThreeD();

        /**
        Returns the vertical anchor type for the specified text. Since Excel 2007.

        [MSDN documentation for TextFrame2.VerticalAnchor](https://learn.microsoft.com/en-us/office/vba/api/excel.textframe2.verticalanchor).
        */
        MsoVerticalAnchor GetVerticalAnchor();

        /**
        Sets the vertical anchor type for the specified text. Since Excel 2007.

        [MSDN documentation for TextFrame2.VerticalAnchor](https://learn.microsoft.com/en-us/office/vba/api/excel.textframe2.verticalanchor).
        */
        void SetVerticalAnchor(MsoVerticalAnchor verticalAnchor);

        /**
        Returns the warp type for the specified text frame. Since Excel 2007.

        [MSDN documentation for TextFrame2.WarpFormat](https://learn.microsoft.com/en-us/office/vba/api/excel.textframe2.warpformat).
        */
        MsoWarpFormat GetWarpFormat();

        /**
        Sets the warp type for the specified text frame. Since Excel 2007.

        [MSDN documentation for TextFrame2.WarpFormat](https://learn.microsoft.com/en-us/office/vba/api/excel.textframe2.warpformat).
        */
        void SetWarpFormat(MsoWarpFormat warpFormat);

        /**
        Returns the Word Art type for the specified text frame. Since Excel 2007.

        [MSDN documentation for TextFrame2.WordArtformat](https://learn.microsoft.com/en-us/office/vba/api/excel.textframe2.wordartformat).
        */
        MsoPresetTextEffect GetWordArtformat();

        /**
        Sets the Word Art type for the specified text frame. Since Excel 2007.

        [MSDN documentation for TextFrame2.WordArtformat](https://learn.microsoft.com/en-us/office/vba/api/excel.textframe2.wordartformat).
        */
        void SetWordArtformat(MsoPresetTextEffect wordArtformat);

        /**
        Returns text break lines within or past the boundaries of the shape. Since Excel 2007.

        [MSDN documentation for TextFrame2.WordWrap](https://learn.microsoft.com/en-us/office/vba/api/excel.textframe2.wordwrap).
        */
        MsoTriState GetWordWrap();

        /**
        Sets text break lines within or past the boundaries of the shape. Since Excel 2007.

        [MSDN documentation for TextFrame2.WordWrap](https://learn.microsoft.com/en-us/office/vba/api/excel.textframe2.wordwrap).
        */
        void SetWordWrap(MsoTriState wordWrap);

        /**
        Returns "TextFrame2".
        */
        virtual wxString GetAutoExcelObjectName_() const { return wxS("TextFrame2"); }
    };


} // namespace wxAutoExcel

#endif // #if WXAUTOEXCEL_USE_SHAPES || WXAUTOEXCEL_USE_CHARTS

#endif //_WXAUTOEXCEL_TEXTFRAME2_H
