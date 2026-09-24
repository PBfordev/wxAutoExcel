/////////////////////////////////////////////////////////////////////////////
// Purpose:     Tests for wxExcelRange formatting properties
// Author:      OpenAI Codex, under the direction of PB
// Generated:   AI-generated using GPT-5.6 Sol with High reasoning effort
// Copyright:   (c) 2026 PB <pbfordev@gmail.com>
// License:     MIT license
/////////////////////////////////////////////////////////////////////////////

#include <catch2/catch_approx.hpp>
#include <catch2/catch_message.hpp>
#include <catch2/catch_test_macros.hpp>

#include "excel_test_fixture.h"

using namespace wxAutoExcel;

TEST_CASE("wxExcelRange font properties round-trip",
          "[excel][range][format][font]")
{
    ExcelTestWorkbook excel;
    REQUIRE(excel.IsReady());

    wxExcelRange range = excel.GetRange("B2:D4");
    wxExcelFont font = range.GetFont();
    REQUIRE(font);

    const wxColour expectedColor(12, 34, 56);
    font.SetName("Arial");
    REQUIRE(font);
    font.SetSize(14.0);
    REQUIRE(font);
    font.SetBold(true);
    REQUIRE(font);
    font.SetItalic(true);
    REQUIRE(font);
    font.SetColor(expectedColor);
    REQUIRE(font);

    CHECK(font.GetName() == "Arial");
    REQUIRE(font);
    CHECK(font.GetSize() == Catch::Approx(14.0));
    REQUIRE(font);
    CHECK(font.GetBold());
    REQUIRE(font);
    CHECK(font.GetItalic());
    REQUIRE(font);
    CHECK(font.GetColor() == expectedColor);
    REQUIRE(font);
}

TEST_CASE("wxExcelRange font effects and theme properties round-trip",
          "[excel][range][format][font]")
{
    ExcelTestWorkbook excel;
    REQUIRE(excel.IsReady());

    wxExcelFont font = excel.GetRange("B2:D4").GetFont();
    REQUIRE(font);

    font.SetFontStyle("Bold Italic");
    REQUIRE(font);
    font.SetStrikethrough(true);
    REQUIRE(font);
    font.SetThemeColor(xlThemeColorAccent2);
    REQUIRE(font);
    font.SetThemeFont(xlThemeFontMajor);
    REQUIRE(font);
    font.SetTintAndShade(0.25);
    REQUIRE(font);

    CHECK(font.GetFontStyle() == "Bold Italic");
    REQUIRE(font);
    CHECK(font.GetStrikethrough());
    REQUIRE(font);
    CHECK(font.GetThemeColor() == xlThemeColorAccent2);
    REQUIRE(font);
    CHECK(font.GetThemeFont() == xlThemeFontMajor);
    REQUIRE(font);
    CHECK(font.GetTintAndShade() == Catch::Approx(0.25).margin(0.0001));
    REQUIRE(font);

    font.SetSubscript(true);
    REQUIRE(font);
    CHECK(font.GetSubscript());
    REQUIRE(font);
    CHECK_FALSE(font.GetSuperscript());
    REQUIRE(font);

    font.SetSubscript(false);
    REQUIRE(font);
    font.SetSuperscript(true);
    REQUIRE(font);
    CHECK_FALSE(font.GetSubscript());
    REQUIRE(font);
    CHECK(font.GetSuperscript());
    REQUIRE(font);
}

TEST_CASE("wxExcelRange background and border properties round-trip",
          "[excel][range][format][fill][border]")
{
    ExcelTestWorkbook excel;
    REQUIRE(excel.IsReady());

    wxExcelRange range = excel.GetRange("B2:D4");

    const wxColour fillColor(78, 90, 123);
    wxExcelInterior interior = range.GetInterior();
    REQUIRE(interior);
    interior.SetPattern(xlPatternSolid);
    REQUIRE(interior);
    interior.SetColor(fillColor);
    REQUIRE(interior);
    CHECK(interior.GetPattern() == xlPatternSolid);
    REQUIRE(interior);
    CHECK(interior.GetColor() == fillColor);
    REQUIRE(interior);

    const wxColour borderColor(210, 45, 67);
    wxExcelBorder border = range.GetBorders()[xlEdgeBottom];
    REQUIRE(border);
    border.SetLineStyle(xlContinuous);
    REQUIRE(border);
    border.SetWeight(xlMedium);
    REQUIRE(border);
    border.SetColor(borderColor);
    REQUIRE(border);
    CHECK(border.GetLineStyle() == xlContinuous);
    REQUIRE(border);
    CHECK(border.GetWeight() == xlMedium);
    REQUIRE(border);
    CHECK(border.GetColor() == borderColor);
    REQUIRE(border);
}

TEST_CASE("wxExcelRange exposes every cell border side",
          "[excel][range][format][border][side]")
{
    ExcelTestWorkbook excel;
    REQUIRE(excel.IsReady());

    wxExcelBorders borders = excel.GetRange("B2:D4").GetBorders();
    REQUIRE(borders);

    struct BorderSideCase
    {
        const char* name;
        XlBordersIndex index;
    };

    const BorderSideCase borderSides[] = {
        { "diagonal down", xlDiagonalDown },
        { "diagonal up", xlDiagonalUp },
        { "bottom edge", xlEdgeBottom },
        { "left edge", xlEdgeLeft },
        { "right edge", xlEdgeRight },
        { "top edge", xlEdgeTop },
        { "inside horizontal", xlInsideHorizontal },
        { "inside vertical", xlInsideVertical },
    };

    for ( const BorderSideCase& borderSide : borderSides )
    {
        CAPTURE(borderSide.name);

        wxExcelBorder border = borders[borderSide.index];
        REQUIRE(border);
        border.SetLineStyle(xlContinuous);
        REQUIRE(border);
        border.SetWeight(xlThin);
        REQUIRE(border);
        CHECK(border.GetLineStyle() == xlContinuous);
        REQUIRE(border);
        CHECK(border.GetWeight() == xlThin);
        REQUIRE(border);
    }
}

TEST_CASE("wxExcelRange supports every border line style",
          "[excel][range][format][border][line-style]")
{
    ExcelTestWorkbook excel;
    REQUIRE(excel.IsReady());

    wxExcelBorder border =
        excel.GetRange("B2:D4").GetBorders()[xlEdgeBottom];
    REQUIRE(border);

    struct LineStyleCase
    {
        const char* name;
        long lineStyle;
    };

    const LineStyleCase lineStyles[] = {
        { "continuous", xlContinuous },
        { "dash", xlDash },
        { "dash dot", xlDashDot },
        { "dash dot dot", xlDashDotDot },
        { "dot", xlDot },
        { "double", xlDouble },
        { "none", xlLineStyleNone },
        { "slant dash dot", xlSlantDashDot },
    };

    for ( const LineStyleCase& lineStyle : lineStyles )
    {
        CAPTURE(lineStyle.name);

        border.SetWeight(xlThin);
        REQUIRE(border);
        border.SetLineStyle(lineStyle.lineStyle);
        REQUIRE(border);
        CHECK(border.GetLineStyle() == lineStyle.lineStyle);
        REQUIRE(border);
    }
}

TEST_CASE("wxExcelRange supports every border weight",
          "[excel][range][format][border][weight]")
{
    ExcelTestWorkbook excel;
    REQUIRE(excel.IsReady());

    wxExcelBorder border =
        excel.GetRange("B2:D4").GetBorders()[xlEdgeBottom];
    REQUIRE(border);

    struct BorderWeightCase
    {
        const char* name;
        XlBorderWeight weight;
    };

    const BorderWeightCase borderWeights[] = {
        { "hairline", xlHairline },
        { "thin", xlThin },
        { "medium", xlMedium },
        { "thick", xlThick },
    };

    for ( const BorderWeightCase& borderWeight : borderWeights )
    {
        CAPTURE(borderWeight.name);

        border.SetLineStyle(xlContinuous);
        REQUIRE(border);
        border.SetWeight(borderWeight.weight);
        REQUIRE(border);
        CHECK(border.GetWeight() == borderWeight.weight);
        REQUIRE(border);
    }
}

TEST_CASE("wxExcelRange supports every horizontal alignment",
          "[excel][range][format][alignment][horizontal]")
{
    ExcelTestWorkbook excel;
    REQUIRE(excel.IsReady());

    wxExcelRange range = excel.GetRange("B2:D4");
    range.SetValue("alignment");
    REQUIRE(range);

    struct HorizontalAlignmentCase
    {
        const char* name;
        XlHAlign alignment;
    };

    const HorizontalAlignmentCase alignments[] = {
        { "center", xlHAlignCenter },
        { "center across selection", xlHAlignCenterAcrossSelection },
        { "distributed", xlHAlignDistributed },
        { "fill", xlHAlignFill },
        { "general", xlHAlignGeneral },
        { "justify", xlHAlignJustify },
        { "left", xlHAlignLeft },
        { "right", xlHAlignRight },
    };

    for ( const HorizontalAlignmentCase& alignment : alignments )
    {
        CAPTURE(alignment.name);

        range.SetHorizontalAlignment(alignment.alignment);
        REQUIRE(range);
        CHECK(range.GetHorizontalAlignment() == alignment.alignment);
        REQUIRE(range);
    }
}

TEST_CASE("wxExcelRange supports every vertical alignment",
          "[excel][range][format][alignment][vertical]")
{
    ExcelTestWorkbook excel;
    REQUIRE(excel.IsReady());

    wxExcelRange range = excel.GetRange("B2:D4");
    range.SetValue("alignment");
    REQUIRE(range);

    struct VerticalAlignmentCase
    {
        const char* name;
        XlVAlign alignment;
    };

    const VerticalAlignmentCase alignments[] = {
        { "bottom", xlVAlignBottom },
        { "center", xlVAlignCenter },
        { "distributed", xlVAlignDistributed },
        { "justify", xlVAlignJustify },
        { "top", xlVAlignTop },
    };

    for ( const VerticalAlignmentCase& alignment : alignments )
    {
        CAPTURE(alignment.name);

        range.SetVerticalAlignment(alignment.alignment);
        REQUIRE(range);
        CHECK(range.GetVerticalAlignment() == alignment.alignment);
        REQUIRE(range);
    }
}

TEST_CASE("wxExcelRange number formats round-trip",
          "[excel][range][format][number-format]")
{
    ExcelTestWorkbook excel;
    REQUIRE(excel.IsReady());

    wxExcelRange range = excel.GetRange("B2:D4");

    struct NumberFormatCase
    {
        const char* name;
        const char* format;
    };

    const NumberFormatCase numberFormats[] = {
        { "general", "General" },
        { "integer", "0" },
        { "fixed decimal", "0.000" },
        { "thousands separator", "#,##0.00" },
        { "currency", "$#,##0.00" },
        { "percentage", "0.00%" },
        { "scientific", "0.00E+00" },
        { "date", "yyyy-mm-dd" },
        { "time", "hh:mm:ss" },
        { "text", "@" },
        { "fraction", "# ?/?" },
        { "positive, negative, and zero", "0.00;[Red]-0.00;0.00" },
    };

    for ( const NumberFormatCase& numberFormat : numberFormats )
    {
        CAPTURE(numberFormat.name);

        range.SetNumberFormat(numberFormat.format);
        REQUIRE(range);

        const wxString actual = range.GetNumberFormat();
        REQUIRE(range);
        CHECK(actual == numberFormat.format);
    }

    wxExcelRange firstCell = excel.GetRange("B6");
    wxExcelRange secondCell = excel.GetRange("C6");
    wxExcelRange mixedRange = excel.GetRange("B6:C6");

    firstCell.SetNumberFormat("0.00");
    REQUIRE(firstCell);
    secondCell.SetNumberFormat("0.00%");
    REQUIRE(secondCell);

    const wxString mixedFormat = mixedRange.GetNumberFormat();
    REQUIRE(mixedRange);
    CHECK(mixedFormat.empty());

    mixedRange.SetNumberFormat("#,##0");
    REQUIRE(mixedRange);
    CHECK(mixedRange.GetNumberFormat() == "#,##0");
    REQUIRE(mixedRange);
}

TEST_CASE("wxExcelRange layout properties round-trip",
          "[excel][range][format][layout]")
{
    ExcelTestWorkbook excel;
    REQUIRE(excel.IsReady());

    wxExcelRange range = excel.GetRange("B2:D4");
    range.SetHorizontalAlignment(xlCenter);
    REQUIRE(range);
    range.SetVerticalAlignment(xlVAlignCenter);
    REQUIRE(range);
    range.SetWrapText(true);
    REQUIRE(range);
    range.SetColumnWidth(18.0);
    REQUIRE(range);
    range.SetRowHeight(24.0);
    REQUIRE(range);

    CHECK(range.GetHorizontalAlignment() == xlCenter);
    REQUIRE(range);
    CHECK(range.GetVerticalAlignment() == xlVAlignCenter);
    REQUIRE(range);
    CHECK(range.GetWrapText().IsTrue());
    REQUIRE(range);
    CHECK(range.GetColumnWidth() == Catch::Approx(18.0).margin(0.5));
    REQUIRE(range);
    CHECK(range.GetRowHeight() == Catch::Approx(24.0).margin(0.1));
    REQUIRE(range);
}
