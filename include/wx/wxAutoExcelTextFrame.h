/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// License:     MIT license
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

        [MSDN documentation for TextFrame.Characters](https://learn.microsoft.com/en-us/office/vba/api/excel.textframe.characters).
        */
        wxExcelCharacters Characters(long* start = NULL, long* length = NULL);


        // ***** PROPERTIES *****

        /**
        True if the size of the specified object is changed automatically to fit text within its boundaries.

        [MSDN documentation for TextFrame.AutoSize](https://learn.microsoft.com/en-us/office/vba/api/excel.textframe.autosize).
        */
        bool GetAutoSize();

        /**
        True if the size of the specified object is changed automatically to fit text within its boundaries.

        [MSDN documentation for TextFrame.AutoSize](https://learn.microsoft.com/en-us/office/vba/api/excel.textframe.autosize).
        */
        void SetAutoSize(bool autoSize);

        /**
        Returns a XlHAlign value that represents the horizontal alignment for the specified object.

        [MSDN documentation for TextFrame.HorizontalAlignment](https://learn.microsoft.com/en-us/office/vba/api/excel.textframe.horizontalalignment).
        */
        XlHAlign GetHorizontalAlignment();

        /**
        Sets a XlHAlign value that represents the horizontal alignment for the specified object.

        [MSDN documentation for TextFrame.HorizontalAlignment](https://learn.microsoft.com/en-us/office/vba/api/excel.textframe.horizontalalignment).
        */
        void SetHorizontalAlignment(XlHAlign horizontalAlignment);

        /**
        Returns the distance (in points) between the bottom of the text frame and the bottom of the inscribed rectangle of the shape that contains the text. Read/write Single.

        [MSDN documentation for TextFrame.MarginBottom](https://learn.microsoft.com/en-us/office/vba/api/excel.textframe.marginbottom).
        */
        double GetMarginBottom();

        /**
        Sets the distance (in points) between the bottom of the text frame and the bottom of the inscribed rectangle of the shape that contains the text. Read/write Single.

        [MSDN documentation for TextFrame.MarginBottom](https://learn.microsoft.com/en-us/office/vba/api/excel.textframe.marginbottom).
        */
        void SetMarginBottom(double marginBottom);

        /**
        Returns the distance (in points) between the left edge of the text frame and the left edge of the inscribed rectangle of the shape that contains the text. Read/write Single.

        [MSDN documentation for TextFrame.MarginLeft](https://learn.microsoft.com/en-us/office/vba/api/excel.textframe.marginleft).
        */
        double GetMarginLeft();

        /**
        Sets the distance (in points) between the left edge of the text frame and the left edge of the inscribed rectangle of the shape that contains the text. Read/write Single.

        [MSDN documentation for TextFrame.MarginLeft](https://learn.microsoft.com/en-us/office/vba/api/excel.textframe.marginleft).
        */
        void SetMarginLeft(double marginLeft);

        /**
        Returns the distance (in points) between the right edge of the text frame and the right edge of the inscribed rectangle of the shape that contains the text. Read/write Single.

        [MSDN documentation for TextFrame.MarginRight](https://learn.microsoft.com/en-us/office/vba/api/excel.textframe.marginright).
        */
        double GetMarginRight();

        /**
        Sets the distance (in points) between the right edge of the text frame and the right edge of the inscribed rectangle of the shape that contains the text. Read/write Single.

        [MSDN documentation for TextFrame.MarginRight](https://learn.microsoft.com/en-us/office/vba/api/excel.textframe.marginright).
        */
        void SetMarginRight(double marginRight);

        /**
        Returns the distance (in points) between the top of the text frame and the top of the inscribed rectangle of the shape that contains the text. Read/write Single.

        [MSDN documentation for TextFrame.MarginTop](https://learn.microsoft.com/en-us/office/vba/api/excel.textframe.margintop).
        */
        double GetMarginTop();

        /**
        Sets the distance (in points) between the top of the text frame and the top of the inscribed rectangle of the shape that contains the text. Read/write Single.

        [MSDN documentation for TextFrame.MarginTop](https://learn.microsoft.com/en-us/office/vba/api/excel.textframe.margintop).
        */
        void SetMarginTop(double marginTop);

        /**
        Returns a Long value that represents the text frame orientation.

        [MSDN documentation for TextFrame.Orientation](https://learn.microsoft.com/en-us/office/vba/api/excel.textframe.orientation).
        */
        long GetOrientation();

        /**
        Sets a Long value that represents the text frame orientation.

        [MSDN documentation for TextFrame.Orientation](https://learn.microsoft.com/en-us/office/vba/api/excel.textframe.orientation).
        */
        void SetOrientation(long orientation);

        /**
        Returns the reading order for the specified object. Can be one of the following constants: xlRTL (right-to-left), xlLTR (left-to-right), or xlContext.

        [MSDN documentation for TextFrame.ReadingOrder](https://learn.microsoft.com/en-us/office/vba/api/excel.textframe.readingorder).
        */
        long GetReadingOrder();

        /**
        Sets the reading order for the specified object. Can be one of the following constants: xlRTL (right-to-left), xlLTR (left-to-right), or xlContext.

        [MSDN documentation for TextFrame.ReadingOrder](https://learn.microsoft.com/en-us/office/vba/api/excel.textframe.readingorder).
        */
        void SetReadingOrder(long readingOrder);

        /**
        Returns a XlVAlign value that represents the vertical alignment of the specified object.

        [MSDN documentation for TextFrame.VerticalAlignment](https://learn.microsoft.com/en-us/office/vba/api/excel.textframe.verticalalignment).
        */
        XlVAlign GetVerticalAlignment();

        /**
        Sets a XlVAlign value that represents the vertical alignment of the specified object.

        [MSDN documentation for TextFrame.VerticalAlignment](https://learn.microsoft.com/en-us/office/vba/api/excel.textframe.verticalalignment).
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
