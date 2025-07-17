/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
/////////////////////////////////////////////////////////////////////////////


#ifndef _WXAUTOEXCEL_TEXTFRAME_H
#define _WXAUTOEXCEL_TEXTFRAME_H

#include "wx/wxAutoExcel_defs.h"

#if WXAUTOEXCEL_USE_SHAPES || WXAUTOEXCEL_USE_CHARTS

#include "wx/wxAutoExcel_object.h"
#include "wx/wxAutoExcel_enums.h"

namespace wxAutoExcel {

    /**
    @brief Represents Microsoft Excel TextFrame.
    */
   class WXDLLIMPEXP_WXAUTOEXCEL wxExcelTextFrame : public wxExcelObject
    {
    public:
        // ***** METHODS *****

        /**
        Returns a Characters Represents a range of characters within a shape’s text frame. You can use the Characters object to add and format characters within the text frame.

        [MSDN documentation for TextFrame.Characters](http://msdn.microsoft.com/en-us/library/bb242015).
        */
        wxExcelCharacters Characters(long* start = NULL, long* length = NULL);


        // ***** PROPERTIES *****

        /**
        True if the size of the specified object is changed automatically to fit text within its boundaries.

        [MSDN documentation for TextFrame.AutoSize](http://msdn.microsoft.com/en-us/library/bb220866).
        */
        bool GetAutoSize();

        /**
        True if the size of the specified object is changed automatically to fit text within its boundaries.

        [MSDN documentation for TextFrame.AutoSize](http://msdn.microsoft.com/en-us/library/bb220866).
        */
        void SetAutoSize(bool autoSize);

        /**
        Returns a XlHAlign value that represents the horizontal alignment for the specified object.

        [MSDN documentation for TextFrame.HorizontalAlignment](http://msdn.microsoft.com/en-us/library/bb238224).
        */
        XlHAlign GetHorizontalAlignment();

        /**
        Sets a XlHAlign value that represents the horizontal alignment for the specified object.

        [MSDN documentation for TextFrame.HorizontalAlignment](http://msdn.microsoft.com/en-us/library/bb238224).
        */
        void SetHorizontalAlignment(XlHAlign horizontalAlignment);

        /**
        Returns the distance (in points) between the bottom of the text frame and the bottom of the inscribed rectangle of the shape that contains the text. Read/write Single.

        [MSDN documentation for TextFrame.MarginBottom](http://msdn.microsoft.com/en-us/library/bb208724).
        */
        double GetMarginBottom();

        /**
        Sets the distance (in points) between the bottom of the text frame and the bottom of the inscribed rectangle of the shape that contains the text. Read/write Single.

        [MSDN documentation for TextFrame.MarginBottom](http://msdn.microsoft.com/en-us/library/bb208724).
        */
        void SetMarginBottom(double marginBottom);

        /**
        Returns the distance (in points) between the left edge of the text frame and the left edge of the inscribed rectangle of the shape that contains the text. Read/write Single.

        [MSDN documentation for TextFrame.MarginLeft](http://msdn.microsoft.com/en-us/library/bb208727).
        */
        double GetMarginLeft();

        /**
        Sets the distance (in points) between the left edge of the text frame and the left edge of the inscribed rectangle of the shape that contains the text. Read/write Single.

        [MSDN documentation for TextFrame.MarginLeft](http://msdn.microsoft.com/en-us/library/bb208727).
        */
        void SetMarginLeft(double marginLeft);

        /**
        Returns the distance (in points) between the right edge of the text frame and the right edge of the inscribed rectangle of the shape that contains the text. Read/write Single.

        [MSDN documentation for TextFrame.MarginRight](http://msdn.microsoft.com/en-us/library/bb208728).
        */
        double GetMarginRight();

        /**
        Sets the distance (in points) between the right edge of the text frame and the right edge of the inscribed rectangle of the shape that contains the text. Read/write Single.

        [MSDN documentation for TextFrame.MarginRight](http://msdn.microsoft.com/en-us/library/bb208728).
        */
        void SetMarginRight(double marginRight);

        /**
        Returns the distance (in points) between the top of the text frame and the top of the inscribed rectangle of the shape that contains the text. Read/write Single.

        [MSDN documentation for TextFrame.MarginTop](http://msdn.microsoft.com/en-us/library/bb208731).
        */
        double GetMarginTop();

        /**
        Sets the distance (in points) between the top of the text frame and the top of the inscribed rectangle of the shape that contains the text. Read/write Single.

        [MSDN documentation for TextFrame.MarginTop](http://msdn.microsoft.com/en-us/library/bb208731).
        */
        void SetMarginTop(double marginTop);

        /**
        Returns a Long value that represents the text frame orientation.

        [MSDN documentation for TextFrame.Orientation](http://msdn.microsoft.com/en-us/library/bb238231).
        */
        long GetOrientation();

        /**
        Sets a Long value that represents the text frame orientation.

        [MSDN documentation for TextFrame.Orientation](http://msdn.microsoft.com/en-us/library/bb238231).
        */
        void SetOrientation(long orientation);

        /**
        Returns the reading order for the specified object. Can be one of the following constants: xlRTL (right-to-left), xlLTR (left-to-right), or xlContext.

        [MSDN documentation for TextFrame.ReadingOrder](http://msdn.microsoft.com/en-us/library/bb238234).
        */
        long GetReadingOrder();

        /**
        Sets the reading order for the specified object. Can be one of the following constants: xlRTL (right-to-left), xlLTR (left-to-right), or xlContext.

        [MSDN documentation for TextFrame.ReadingOrder](http://msdn.microsoft.com/en-us/library/bb238234).
        */
        void SetReadingOrder(long readingOrder);

        /**
        Returns a XlVAlign value that represents the vertical alignment of the specified object.

        [MSDN documentation for TextFrame.VerticalAlignment](http://msdn.microsoft.com/en-us/library/bb215155).
        */
        XlVAlign GetVerticalAlignment();

        /**
        Sets a XlVAlign value that represents the vertical alignment of the specified object.

        [MSDN documentation for TextFrame.VerticalAlignment](http://msdn.microsoft.com/en-us/library/bb215155).
        */
        void SetVerticalAlignment(XlVAlign  verticalAlignment);

        /**
        Returns "TextFrame".
        */
        virtual wxString GetAutoExcelObjectName_() const { return wxS("TextFrame"); }
    };


} // namespace wxAutoExcel

#endif // #if WXAUTOEXCEL_USE_SHAPES || WXAUTOEXCEL_USE_CHARTS

#endif //_WXAUTOEXCEL_TEXTFRAME_H
