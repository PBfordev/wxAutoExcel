/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// License:     MIT license
/////////////////////////////////////////////////////////////////////////////


#ifndef _WXAUTOEXCEL_SHADOWFORMAT_H
#define _WXAUTOEXCEL_SHADOWFORMAT_H

#include "wx/wxAutoExcel_defs.h"

#if WXAUTOEXCEL_USE_SHAPES || WXAUTOEXCEL_USE_CHARTS

#include "wx/wxAutoExcel_object.h"
#include "wx/wxAutoExcel_enums.h"

namespace wxAutoExcel {

    /**
    @brief Represents Microsoft Excel ShadowFormat object.
    */
   class WXDLLIMPEXP_WXAUTOEXCEL wxExcelShadowFormat: public wxExcelObject
    {
    public:

        // ***** PROPERTIES *****

        /**
        Returns the degree of blurriness of the specified shadow.  Since Excel 2007.

        [MSDN documentation for ShadowFormat.Blur](https://learn.microsoft.com/en-us/office/vba/api/excel.shadowformat.blur).
        */
        double GetBlur();

        /**
        Sets the degree of blurriness of the specified shadow.  Since Excel 2007.

        [MSDN documentation for ShadowFormat.Blur](https://learn.microsoft.com/en-us/office/vba/api/excel.shadowformat.blur).
        */
        void SetBlur(double blur);


        /**
        Returns a ColorFormat Represents the specified foreground fill or solid color.

        [MSDN documentation for ShadowFormat.ForeColor](https://learn.microsoft.com/en-us/office/vba/api/excel.shadowformat.forecolor).
        */
        wxExcelColorFormat GetForeColor();

        /**
        Sets a ColorFormat Represents the specified foreground fill or solid color.

        [MSDN documentation for ShadowFormat.ForeColor](https://learn.microsoft.com/en-us/office/vba/api/excel.shadowformat.forecolor).
        */
        void SetForeColor(const wxExcelColorFormat& foreColor);

        /**
        True if the shadow of the specified shape appears filled in and is obscured by the shape, even if the shape has no fill. False if the shadow has no fill and the outline of the shadow is visible through the shape if the shape has no fill. Read/write MsoTriState.

        [MSDN documentation for ShadowFormat.Obscured](https://learn.microsoft.com/en-us/office/vba/api/excel.shadowformat.obscured).
        */
        MsoTriState GetObscured();

        /**
        True if the shadow of the specified shape appears filled in and is obscured by the shape, even if the shape has no fill. False if the shadow has no fill and the outline of the shadow is visible through the shape if the shape has no fill. Read/write MsoTriState.

        [MSDN documentation for ShadowFormat.Obscured](https://learn.microsoft.com/en-us/office/vba/api/excel.shadowformat.obscured).
        */
        void SetObscured(MsoTriState obscured);

        /**
        Returns the horizontal offset of the shadow from the specified shape, in points. A positive value offsets the shadow to the right of the shape; a negative value offsets it to the left. Read/write Single.

        [MSDN documentation for ShadowFormat.OffsetX](https://learn.microsoft.com/en-us/office/vba/api/excel.shadowformat.offsetx).
        */
        double GetOffsetX();

        /**
        Sets the horizontal offset of the shadow from the specified shape, in points. A positive value offsets the shadow to the right of the shape; a negative value offsets it to the left. Read/write Single.

        [MSDN documentation for ShadowFormat.OffsetX](https://learn.microsoft.com/en-us/office/vba/api/excel.shadowformat.offsetx).
        */
        void SetOffsetX(double offsetX);

        /**
        Returns the vertical offset of the shadow from the specified shape, in points. A positive value offsets the shadow to the right of the shape; a negative value offsets it to the left. Read/write Single.

        [MSDN documentation for ShadowFormat.OffsetY](https://learn.microsoft.com/en-us/office/vba/api/excel.shadowformat.offsety).
        */
        double GetOffsetY();

        /**
        Sets the vertical offset of the shadow from the specified shape, in points. A positive value offsets the shadow to the right of the shape; a negative value offsets it to the left. Read/write Single.

        [MSDN documentation for ShadowFormat.OffsetY](https://learn.microsoft.com/en-us/office/vba/api/excel.shadowformat.offsety).
        */
        void SetOffsetY(double offsetY);

        /**
        Returns an MsoTriState that represents whether to rotate the shadow when rotating the shape.  Since Excel 2007.

        [MSDN documentation for ShadowFormat.RotateWithShape](https://learn.microsoft.com/en-us/office/vba/api/excel.shadowformat.rotatewithshape).
        */
        MsoTriState GetRotateWithShape();

        /**
        Sets an MsoTriState that represents whether to rotate the shadow when rotating the shape.  Since Excel 2007.

        [MSDN documentation for ShadowFormat.RotateWithShape](https://learn.microsoft.com/en-us/office/vba/api/excel.shadowformat.rotatewithshape).
        */
        void SetRotateWithShape(MsoTriState rotateWithShape);

        /**
        Returns the size of the specified shadow. Since Excel 2007.

        [MSDN documentation for ShadowFormat.Size](https://learn.microsoft.com/en-us/office/vba/api/excel.shadowformat.size).
        */
        double GetSize();

        /**
        Sets the size of the specified shadow. Since Excel 2007.

        [MSDN documentation for ShadowFormat.Size](https://learn.microsoft.com/en-us/office/vba/api/excel.shadowformat.size).
        */
        void SetSize(double size);

        /**
        Returns the style of the specified shadow. Since Excel 2007.

        [MSDN documentation for ShadowFormat.Style](https://learn.microsoft.com/en-us/office/vba/api/excel.shadowformat.style).
        */
        MsoShadowStyle GetStyle();

        /**
        Sets the style of the specified shadow. Since Excel 2007.

        [MSDN documentation for ShadowFormat.Style](https://learn.microsoft.com/en-us/office/vba/api/excel.shadowformat.style).
        */
        void SetStyle(MsoShadowStyle style);

        /**
        Returns the degree of transparency of the specified fill as a value from 0.0 (opaque) through 1.0 (clear). Read/write Double.

        [MSDN documentation for ShadowFormat.Transparency](https://learn.microsoft.com/en-us/office/vba/api/excel.shadowformat.transparency).
        */
        double GetTransparency();

        /**
        Sets the degree of transparency of the specified fill as a value from 0.0 (opaque) through 1.0 (clear). Read/write Double.

        [MSDN documentation for ShadowFormat.Transparency](https://learn.microsoft.com/en-us/office/vba/api/excel.shadowformat.transparency).
        */
        void SetTransparency(double transparency);

        /**
        Returns a MsoShadowType value that represents the shadow format type.

        [MSDN documentation for ShadowFormat.Type](https://learn.microsoft.com/en-us/office/vba/api/excel.shadowformat.type).
        */
        MsoShadowType  GetType();

        /**
        Sets a MsoShadowType value that represents the shadow format type.

        [MSDN documentation for ShadowFormat.Type](https://learn.microsoft.com/en-us/office/vba/api/excel.shadowformat.type).
        */
        void SetType(MsoShadowType  type);

        /**
        Returns a MsoTriState value that determines whether the object is visible.

        [MSDN documentation for ShadowFormat.Visible](https://learn.microsoft.com/en-us/office/vba/api/excel.shadowformat.visible).
        */
        MsoTriState GetVisible();

        /**
        Sets a MsoTriState value that determines whether the object is visible.

        [MSDN documentation for ShadowFormat.Visible](https://learn.microsoft.com/en-us/office/vba/api/excel.shadowformat.visible).
        */
        void SetVisible(MsoTriState visible);

        /**
        Returns "ShadowFormat".
        */
        virtual wxString GetAutoExcelObjectName_() const { return wxS("ShadowFormat"); }
    };


} // namespace wxAutoExcel

#endif // #if WXAUTOEXCEL_USE_SHAPES || WXAUTOEXCEL_USE_CHARTS

#endif //_WXAUTOEXCEL_SHADOWFORMAT_H
