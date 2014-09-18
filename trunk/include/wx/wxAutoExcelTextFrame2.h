/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Modified by:
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
/////////////////////////////////////////////////////////////////////////////


#ifndef _WXAUTOEXCEL_TEXTFRAME2_H
#define _WXAUTOEXCEL_TEXTFRAME2_H

#include "wx/wxAutoExcel_defs.h"

#if WXAUTOEXCEL_USE_SHAPES || WXAUTOEXCEL_USE_CHARTS

#include "wx/wxAutoExcelObject.h"
#include "wx/wxAutoExcel_enums.h"

namespace wxAutoExcel {


    /**
    @brief Represents Microsoft Excel TextFrame2.
    */    
   class WXDLLIMPEXP_WXAUTOEXCEL wxExcelTextFrame2 : public wxExcelObject
    {
        // ***** METHODS *****

        /**
        Deletes the text from a text frame and all the associated text properties.

        [MSDN documentation for TextFrame2.DeleteText](http://msdn.microsoft.com/en-us/library/bb256955).
        */
        void DeleteText();

        // ***** PROPERTIES *****

        /**
        The size of the specified object that changes automatically to fit text within its boundaries. Since Excel 2007.

        [MSDN documentation for TextFrame2.AutoSize](http://msdn.microsoft.com/en-us/library/bb256962).
        */
        MsoAutoSize GetAutoSize();

        /**
        Returns the TextColumn2 Represents the columns within the text frame. Since Excel 2007.

        [MSDN documentation for TextFrame2.Column](http://msdn.microsoft.com/en-us/library/bb256964).
        */
        wxExcelTextColumn2 GetColumn();

        /**
        Returns a 32-bit integer that indicates the application in which this object was created. Since Excel 2007.

        [MSDN documentation for TextFrame2.Creator](http://msdn.microsoft.com/en-us/library/bb256966).
        */
        long GetCreator();

        /**
        Returns whether the specified text frame has text. Since Excel 2007.

        [MSDN documentation for TextFrame2.HasText](http://msdn.microsoft.com/en-us/library/bb256968).
        */
        MsoTriState GetHasText();

        /**
        Returns the horizontal anchor type for the specified text. Since Excel 2007.

        [MSDN documentation for TextFrame2.HorizontalAnchor](http://msdn.microsoft.com/en-us/library/bb256970).
        */
        MsoHorizontalAnchor GetHorizontalAnchor();

        /**
        Sets the horizontal anchor type for the specified text. Since Excel 2007.

        [MSDN documentation for TextFrame2.HorizontalAnchor](http://msdn.microsoft.com/en-us/library/bb256970).
        */
        void SetHorizontalAnchor(MsoHorizontalAnchor horizontalAnchor);

        /**
        Returns the distance (in points) between the bottom of the text frame and the bottom of the inscribed rectangle of the shape that contains the text. Since Excel 2007.

        [MSDN documentation for TextFrame2.MarginBottom](http://msdn.microsoft.com/en-us/library/bb256972).
        */
        double GetMarginBottom();

        /**
        Sets the distance (in points) between the bottom of the text frame and the bottom of the inscribed rectangle of the shape that contains the text. Since Excel 2007.

        [MSDN documentation for TextFrame2.MarginBottom](http://msdn.microsoft.com/en-us/library/bb256972).
        */
        void SetMarginBottom(double marginBottom);

        /**
        Returns the distance (in points) between the left edge of the text frame and the left edge of the inscribed rectangle of the shape that contains the text. Since Excel 2007.

        [MSDN documentation for TextFrame2.MarginLeft](http://msdn.microsoft.com/en-us/library/bb256974).
        */
        double GetMarginLeft();

        /**
        Sets the distance (in points) between the left edge of the text frame and the left edge of the inscribed rectangle of the shape that contains the text. Since Excel 2007.

        [MSDN documentation for TextFrame2.MarginLeft](http://msdn.microsoft.com/en-us/library/bb256974).
        */
        void SetMarginLeft(double marginLeft);

        /**
        Returns the distance (in points) between the right edge of the text frame and the right edge of the inscribed rectangle of the shape that contains the text. Since Excel 2007.

        [MSDN documentation for TextFrame2.MarginRight](http://msdn.microsoft.com/en-us/library/bb256976).
        */
        double GetMarginRight();

        /**
        Sets the distance (in points) between the right edge of the text frame and the right edge of the inscribed rectangle of the shape that contains the text. Since Excel 2007.

        [MSDN documentation for TextFrame2.MarginRight](http://msdn.microsoft.com/en-us/library/bb256976).
        */
        void SetMarginRight(double marginRight);

        /**
        Returns the distance (in points) between the top of the text frame and the top of the inscribed rectangle of the shape that contains the textSince Excel 2007.

        [MSDN documentation for TextFrame2.MarginTop](http://msdn.microsoft.com/en-us/library/bb256979).
        */
        double GetMarginTop();

        /**
        Sets the distance (in points) between the top of the text frame and the top of the inscribed rectangle of the shape that contains the text. Since Excel 2007.

        [MSDN documentation for TextFrame2.MarginTop](http://msdn.microsoft.com/en-us/library/bb256979).
        */
        void SetMarginTop(double marginTop);

        /**
        Returns a value that represents the text frame orientation. Since Excel 2007.

        [MSDN documentation for TextFrame2.Orientation](http://msdn.microsoft.com/en-us/library/bb256981).
        */
        MsoTextOrientation GetOrientation();

        /**
        Sets a value that represents the text frame orientation. Since Excel 2007.

        [MSDN documentation for TextFrame2.Orientation](http://msdn.microsoft.com/en-us/library/bb256981).
        */
        void SetOrientation(MsoTextOrientation orientation);       

        /**
        Returns the path type for the specified text frame. Since Excel 2007.

        [MSDN documentation for TextFrame2.PathFormat](http://msdn.microsoft.com/en-us/library/bb256985).
        */
        MsoPathFormat GetPathFormat();

        /**
        Sets the path type for the specified text frame. 

        [MSDN documentation for TextFrame2.PathFormat](http://msdn.microsoft.com/en-us/library/bb256985).
        */
        void SetPathFormat(MsoPathFormat pathFormat);

        /**
        Returns the TextRange2 Represents the text in the object. Since Excel 2007.

        [MSDN documentation for TextFrame2.TextRange](http://msdn.microsoft.com/en-us/library/bb256989).
        */
        wxExcelTextRange2 GetTextRange();

        /**
        Returns a ThreeDFormat object that contains 3-D–effect formatting properties for the specified text. Since Excel 2007.

        [MSDN documentation for TextFrame2.ThreeD](http://msdn.microsoft.com/en-us/library/bb256991).
        */
        wxExcelThreeDFormat GetThreeD();

        /**
        Returns the vertical anchor type for the specified text. Since Excel 2007.

        [MSDN documentation for TextFrame2.VerticalAnchor](http://msdn.microsoft.com/en-us/library/bb256993).
        */
        MsoVerticalAnchor GetVerticalAnchor();

        /**
        Sets the vertical anchor type for the specified text. Since Excel 2007.

        [MSDN documentation for TextFrame2.VerticalAnchor](http://msdn.microsoft.com/en-us/library/bb256993).
        */
        void SetVerticalAnchor(MsoVerticalAnchor verticalAnchor);

        /**
        Returns the warp type for the specified text frame. Since Excel 2007.

        [MSDN documentation for TextFrame2.WarpFormat](http://msdn.microsoft.com/en-us/library/bb256995).
        */
        MsoWarpFormat GetWarpFormat();

        /**
        Sets the warp type for the specified text frame. Since Excel 2007.

        [MSDN documentation for TextFrame2.WarpFormat](http://msdn.microsoft.com/en-us/library/bb256995).
        */
        void SetWarpFormat(MsoWarpFormat warpFormat);

        /**
        Returns the Word Art type for the specified text frame. Since Excel 2007.

        [MSDN documentation for TextFrame2.WordArtformat](http://msdn.microsoft.com/en-us/library/bb256997).
        */
        MsoPresetTextEffect GetWordArtformat();

        /**
        Sets the Word Art type for the specified text frame. Since Excel 2007.

        [MSDN documentation for TextFrame2.WordArtformat](http://msdn.microsoft.com/en-us/library/bb256997).
        */
        void SetWordArtformat(MsoPresetTextEffect wordArtformat);

        /**
        Returns text break lines within or past the boundaries of the shape. Since Excel 2007.

        [MSDN documentation for TextFrame2.WordWrap](http://msdn.microsoft.com/en-us/library/bb257001).
        */
        MsoTriState GetWordWrap();

        /**
        Sets text break lines within or past the boundaries of the shape. Since Excel 2007.

        [MSDN documentation for TextFrame2.WordWrap](http://msdn.microsoft.com/en-us/library/bb257001).
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
