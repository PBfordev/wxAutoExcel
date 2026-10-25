/////////////////////////////////////////////////////////////////////////////
// Purpose:     Tests for the methods of wxExcelFillFormat setting the fill type
// Author:      Claude Opus 5.5 (Anthropic), under the direction of PB
// Copyright:   (c) 2026 PB <pbfordev@gmail.com>
// License:     MIT license
/////////////////////////////////////////////////////////////////////////////

#include <catch2/catch_test_macros.hpp>
#include <catch2/catch_approx.hpp>

#include "excel_test_fixture.h"
#include "temp_files.h"

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

TEST_CASE("wxExcelGradientStops inserts and deletes the stops of a gradient",
          "[excel][shapes][fill]")
{
    ExcelTestWorkbook book;
    REQUIRE(book.IsReady());

    wxExcelFillFormat fill = AddRectangleFill(book);
    fill.TwoColorGradient(msoGradientHorizontal, 1);
    REQUIRE(fill);

    wxExcelGradientStops stops = fill.GetGradientStops();
    REQUIRE(stops);
    CHECK(stops.GetCount() == 2);
    REQUIRE(stops);

    wxExcelGradientStop first = stops.GetItem(1);
    REQUIRE(first);
    CHECK(first.GetPosition() == 0.0);
    REQUIRE(first);

    stops.Insert(*wxGREEN, 0.5, 0.25);
    REQUIRE(stops);
    CHECK(stops.GetCount() == 3);
    REQUIRE(stops);

    // the stop is added after the last one, whatever its position
    wxExcelGradientStop middle = stops.GetItem(3);
    REQUIRE(middle);
    CHECK(middle.GetPosition() == Catch::Approx(0.5));
    CHECK(middle.GetTransparency() == Catch::Approx(0.25));
    CHECK(middle.GetColor().GetRGB() == *wxGREEN);
    REQUIRE(middle);

    middle.SetPosition(0.4);
    middle.SetTransparency(0);
    REQUIRE(middle);
    CHECK(middle.GetPosition() == Catch::Approx(0.4));
    CHECK(middle.GetTransparency() == Catch::Approx(0.0));
    REQUIRE(middle);

    stops.Insert2(*wxBLUE, 0.8, 0, -1, 0.5);
    REQUIRE(stops);
    CHECK(stops.GetCount() == 4);
    REQUIRE(stops);

    long count = 0;
    for ( wxExcelGradientStop stop : stops )
    {
        CHECK(stop);
        ++count;
    }
    CHECK(count == 4);

    stops.Delete(2);
    REQUIRE(stops);
    CHECK(stops.GetCount() == 3);
    stops.Delete();
    REQUIRE(stops);
    CHECK(stops.GetCount() == 2);
}

#endif // WXAUTOEXCEL_USE_SHAPES
