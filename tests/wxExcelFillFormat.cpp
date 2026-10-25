/////////////////////////////////////////////////////////////////////////////
// Purpose:     Tests for the methods of wxExcelFillFormat setting the fill type
// Author:      Claude Opus 5.5 (Anthropic), under the direction of PB
// Copyright:   (c) 2026 PB <pbfordev@gmail.com>
// License:     MIT license
/////////////////////////////////////////////////////////////////////////////

#include <catch2/catch_test_macros.hpp>
#include <catch2/catch_approx.hpp>

#include <wx/ffile.h>
#include <wx/filename.h>

#include "excel_test_fixture.h"

#if WXAUTOEXCEL_USE_SHAPES

using namespace wxAutoExcel;

namespace
{

wxExcelFillFormat AddRectangleFill(ExcelTestWorkbook& book)
{
    wxExcelShapes shapes = book.GetWorksheet().GetShapes();
    REQUIRE(shapes);
    wxExcelShape rectangle = shapes.AddShape(msoShapeRectangle, 10, 10, 120, 60);
    REQUIRE(rectangle);
    wxExcelFillFormat fill = rectangle.GetFill();
    REQUIRE(fill);
    return fill;
}

MsoFillType GetType(wxExcelFillFormat& fill)
{
    const MsoFillType type = fill.GetType();
    REQUIRE(fill);
    return type;
}

// A temporary 2x2 pixel BMP picture, deleted when the object is destroyed.
class TempPicture
{
public:
    TempPicture()
    {
        path = wxFileName::CreateTempFileName("wxae");
        wxRemoveFile(path);
        path += ".bmp";

        // the file and info headers, followed by two rows of 2 pixels,
        // each row padded to 4 bytes
        static const unsigned char bmp[] =
        {
            'B', 'M', 70, 0, 0, 0, 0, 0, 0, 0, 54, 0, 0, 0,
            40, 0, 0, 0, 2, 0, 0, 0, 2, 0, 0, 0, 1, 0, 24, 0,
            0, 0, 0, 0, 16, 0, 0, 0, 0x13, 0x0B, 0, 0, 0x13, 0x0B, 0, 0,
            0, 0, 0, 0, 0, 0, 0, 0,
            0, 0, 255, 0, 255, 0, 0, 0,
            255, 0, 0, 255, 255, 255, 0, 0
        };

        wxFFile file(path, "wb");
        REQUIRE(file.IsOpened());
        REQUIRE(file.Write(bmp, sizeof(bmp)) == sizeof(bmp));
    }

    ~TempPicture()
    {
        wxRemoveFile(path);
    }

    const wxString& GetPath() const { return path; }

private:
    wxString path;
};

} // unnamed namespace

TEST_CASE("wxExcelFillFormat sets a solid or patterned fill",
          "[excel][shapes][fill]")
{
    ExcelTestWorkbook book;
    REQUIRE(book.IsReady());

    wxExcelFillFormat fill = AddRectangleFill(book);

    fill.Patterned(msoPattern30Percent);
    REQUIRE(fill);
    CHECK(GetType(fill) == msoFillPatterned);
    CHECK(fill.GetPattern() == msoPattern30Percent);
    REQUIRE(fill);

    fill.Solid();
    REQUIRE(fill);
    CHECK(GetType(fill) == msoFillSolid);

    wxExcelColorFormat foreColor = fill.GetForeColor();
    REQUIRE(foreColor);
    foreColor.SetRGB(*wxRED);
    REQUIRE(foreColor);
    CHECK(fill.GetForeColor().GetRGB() == *wxRED);
}

TEST_CASE("wxExcelFillFormat sets gradient fills",
          "[excel][shapes][fill]")
{
    ExcelTestWorkbook book;
    REQUIRE(book.IsReady());

    wxExcelFillFormat fill = AddRectangleFill(book);

    SECTION("one-color gradient")
    {
        fill.OneColorGradient(msoGradientHorizontal, 1, 0.5);
        REQUIRE(fill);
        CHECK(GetType(fill) == msoFillGradient);
        CHECK(fill.GetGradientStyle() == msoGradientHorizontal);
        CHECK(fill.GetGradientVariant() == 1);
        CHECK(fill.GetGradientColorType() == msoGradientOneColor);
        CHECK(fill.GetGradientDegree() == Catch::Approx(0.5));
        REQUIRE(fill);
    }

    SECTION("two-color gradient")
    {
        fill.TwoColorGradient(msoGradientVertical, 2);
        REQUIRE(fill);
        CHECK(GetType(fill) == msoFillGradient);
        CHECK(fill.GetGradientStyle() == msoGradientVertical);
        CHECK(fill.GetGradientVariant() == 2);
        REQUIRE(fill);
    }

    SECTION("preset gradient")
    {
        fill.PresetGradient(msoGradientDiagonalUp, 1, msoGradientEarlySunset);
        REQUIRE(fill);
        CHECK(GetType(fill) == msoFillGradient);
        CHECK(fill.GetGradientStyle() == msoGradientDiagonalUp);
        CHECK(fill.GetGradientColorType() == msoGradientPresetColors);
        CHECK(fill.GetPresetGradientType() == msoGradientEarlySunset);
        REQUIRE(fill);
    }
}

TEST_CASE("wxExcelFillFormat sets texture and picture fills",
          "[excel][shapes][fill]")
{
    ExcelTestWorkbook book;
    REQUIRE(book.IsReady());

    wxExcelFillFormat fill = AddRectangleFill(book);

    SECTION("preset texture")
    {
        fill.PresetTextured(msoTexturePapyrus);
        REQUIRE(fill);
        CHECK(GetType(fill) == msoFillTextured);
    }

    SECTION("user texture")
    {
        TempPicture picture;
        fill.UserTextured(picture.GetPath());
        REQUIRE(fill);
        // Excel returns msoTextureTypeMixed as its TextureType
        CHECK(GetType(fill) == msoFillTextured);
    }

    SECTION("user picture")
    {
        TempPicture picture;
        fill.UserPicture(picture.GetPath());
        REQUIRE(fill);
        CHECK(GetType(fill) == msoFillPicture);
    }
}

#endif // WXAUTOEXCEL_USE_SHAPES
