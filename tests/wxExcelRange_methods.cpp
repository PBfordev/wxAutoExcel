/////////////////////////////////////////////////////////////////////////////
// Purpose:     Tests for wxExcelRange methods drawing borders, filling
//              series, creating and applying names, finding differences,
//              outlining, and copying ranges as pictures
// Author:      Claude Opus 5.5 (Anthropic), under the direction of PB
// Copyright:   (c) 2026 PB <pbfordev@gmail.com>
// License:     MIT license
/////////////////////////////////////////////////////////////////////////////

#include <catch2/catch_test_macros.hpp>

#include "excel_test_fixture.h"

using namespace wxAutoExcel;

namespace
{

wxExcelBorder GetBorder(wxExcelRange range, XlBordersIndex index)
{
    wxExcelBorder border = range.GetBorders().GetItem(index);
    REQUIRE(border);
    return border;
}

void SetValues(ExcelTestWorkbook& book, const wxString& address,
               double first, double second, double third)
{
    wxExcelRange range = book.GetRange(address);
    const double values[] = { first, second, third };
    for ( long i = 0; i < 3; ++i )
    {
        wxExcelRange cell = range.GetItem(i + 1);
        REQUIRE(cell);
        cell.SetValue(wxVariant(values[i]));
        REQUIRE(cell);
    }
}

wxString GetFormula(ExcelTestWorkbook& book, const wxString& address)
{
    wxExcelRange cell = book.GetRange(address);
    const wxString formula = cell.GetFormula();
    REQUIRE(cell);
    return formula;
}

} // unnamed namespace

TEST_CASE("wxExcelRange draws a border around the range",
          "[excel][range][borders]")
{
    ExcelTestWorkbook book;
    REQUIRE(book.IsReady());

    wxExcelRange range = book.GetRange("B2:C3");
    XlLineStyle lineStyle = xlContinuous;
    XlBorderWeight weight = xlThick;
    REQUIRE(range.BorderAround(&lineStyle, &weight, NULL, wxRED));

    const XlBordersIndex edges[] = { xlEdgeTop, xlEdgeBottom, xlEdgeLeft, xlEdgeRight };
    for ( size_t i = 0; i < WXSIZEOF(edges); ++i )
    {
        wxExcelBorder border = GetBorder(range, edges[i]);
        CHECK(border.GetLineStyle() == xlContinuous);
        CHECK(border.GetWeight() == xlThick);
        CHECK(border.GetColor() == *wxRED);
        REQUIRE(border);
    }

    // only the edges have the border, not the cells inside
    CHECK(GetBorder(range, xlInsideHorizontal).GetLineStyle() == xlLineStyleNone);
    CHECK(GetBorder(book.GetRange("B2"), xlEdgeBottom).GetLineStyle() == xlLineStyleNone);
}

TEST_CASE("wxExcelRange fills a data series",
          "[excel][range][series]")
{
    ExcelTestWorkbook book;
    REQUIRE(book.IsReady());

    SECTION("linear series in a column")
    {
        book.GetRange("A1").SetValue(wxVariant(1L));
        wxExcelRange range = book.GetRange("A1:A5");
        XlRowCol rowCol = xlColumns;
        XlDataSeriesType type = xlDataSeriesLinear;
        long step = 2;
        REQUIRE(range.DataSeries(&rowCol, &type, NULL, &step));
        CHECK(book.GetRange("A2").GetValue().GetDouble() == 3.0);
        CHECK(book.GetRange("A5").GetValue().GetDouble() == 9.0);
    }

    SECTION("growth series in a row up to a stop value")
    {
        book.GetRange("A1").SetValue(wxVariant(1L));
        wxExcelRange range = book.GetRange("A1:F1");
        XlRowCol rowCol = xlRows;
        XlDataSeriesType type = xlGrowth;
        long step = 3;
        REQUIRE(range.DataSeries(&rowCol, &type, NULL, &step, wxVariant(30L)));
        CHECK(book.GetRange("D1").GetValue().GetDouble() == 27.0);
        CHECK(book.GetRange("E1").GetValue().IsNull());
    }
}

TEST_CASE("wxExcelRange creates and applies names",
          "[excel][range][names]")
{
    ExcelTestWorkbook book;
    REQUIRE(book.IsReady());

    wxExcelNames names = book.GetWorkbook().GetNames();
    REQUIRE(names);
    CHECK(names.GetCount() == 0);

    // the names of the columns are created from the labels in the top row
    book.GetRange("A1").SetValue("Price");
    book.GetRange("B1").SetValue("Amount");
    SetValues(book, "A2:A4", 10, 20, 30);
    SetValues(book, "B2:B4", 1, 2, 3);
    REQUIRE(book.GetRange("A1:B4").CreateNames(true, false, false, false));
    CHECK(names.GetCount() == 2);
    wxExcelName price = names.Item("Price");
    REQUIRE(price);
    CHECK(price.GetRefersToRange().GetAddress(false, false) == "A2:A4");
    REQUIRE(price);

    // all names are applied without any given
    wxExcelRange total = book.GetRange("D1");
    total.SetFormula("=SUM($A$2:$A$4)+SUM($B$2:$B$4)");
    REQUIRE(total);
    total.ApplyNames();
    REQUIRE(total);
    CHECK(GetFormula(book, "D1") == "=SUM(Price)+SUM(Amount)");
    CHECK(book.GetRange("D1").GetValue().GetDouble() == 66.0);

    // only the given names are applied (Excel then applies only them also
    // when no names are given later, so this is done after the above)
    wxExcelRange other = book.GetRange("D2");
    other.SetFormula("=SUM($A$2:$A$4)+SUM($B$2:$B$4)");
    REQUIRE(other);
    wxArrayString applied;
    applied.push_back("Price");
    other.ApplyNames(&applied);
    REQUIRE(other);
    CHECK(GetFormula(book, "D2") == "=SUM(Price)+SUM($B$2:$B$4)");
}

TEST_CASE("wxExcelRange finds cells differing from the comparison cell",
          "[excel][range][find]")
{
    ExcelTestWorkbook book;
    REQUIRE(book.IsReady());

    SetValues(book, "A1:A3", 1, 1, 2);
    wxExcelRange column = book.GetRange("A1:A3");
    wxExcelRange columnDifferences = column.ColumnDifferences(book.GetRange("A1"));
    REQUIRE(columnDifferences);
    CHECK(columnDifferences.GetAddress(false, false) == "A3");

    SetValues(book, "C1:E1", 5, 6, 5);
    wxExcelRange row = book.GetRange("C1:E1");
    wxExcelRange rowDifferences = row.RowDifferences(book.GetRange("C1"));
    REQUIRE(rowDifferences);
    CHECK(rowDifferences.GetAddress(false, false) == "D1");
}

TEST_CASE("wxExcelRange creates and clears an outline",
          "[excel][range][outline]")
{
    ExcelTestWorkbook book;
    REQUIRE(book.IsReady());

    SetValues(book, "A1:A3", 1, 2, 3);
    book.GetRange("A4").SetFormula("=SUM(A1:A3)");

    wxExcelRange range = book.GetRange("A1:A4");
    range.AutoOutline();
    REQUIRE(range);

    // the summed rows become a group, the summary row stays at the top level
    CHECK(book.GetRange("A1").GetEntireRow().GetOutlineLevel() == 2);
    CHECK(book.GetRange("A3").GetEntireRow().GetOutlineLevel() == 2);
    CHECK(book.GetRange("A4").GetEntireRow().GetOutlineLevel() == 1);

    REQUIRE(range.ClearOutline());
    CHECK(book.GetRange("A1").GetEntireRow().GetOutlineLevel() == 1);
}

#if WXAUTOEXCEL_USE_SHAPES

TEST_CASE("wxExcelRange copies the range as a picture",
          "[excel][range][picture]")
{
    ExcelTestWorkbook book;
    REQUIRE(book.IsReady());

    book.GetRange("A1").SetValue("Picture");
    XlPictureAppearance appearance = xlScreen;
    XlCopyPictureFormat format = xlBitmap;
    REQUIRE(book.GetRange("A1:B2").CopyPicture(&appearance, &format));

    wxExcelRange destination = book.GetRange("D4");
    REQUIRE(book.GetWorksheet().Paste(&destination));

    wxExcelShapes shapes = book.GetWorksheet().GetShapes();
    REQUIRE(shapes);
    CHECK(shapes.GetCount() == 1);
    wxExcelShape picture = shapes.Item(1);
    REQUIRE(picture);
    CHECK(picture.GetType() == msoPicture);
    CHECK(picture.GetTopLeftCell().GetAddress(false, false) == "D4");
}

#endif // WXAUTOEXCEL_USE_SHAPES
