/////////////////////////////////////////////////////////////////////////////
// Purpose:     Tests for the objects of worksheets, workbooks, and the
//              application: custom properties, cell errors and error
//              checking, table styles, OLE objects, and multithreaded
//              calculation
// Author:      Claude Opus 5.5 (Anthropic), under the direction of PB
// Copyright:   (c) 2026 PB <pbfordev@gmail.com>
// License:     MIT license
/////////////////////////////////////////////////////////////////////////////

#include <catch2/catch_approx.hpp>
#include <catch2/catch_test_macros.hpp>

#include <type_traits>

#include "excel_test_fixture.h"
#include "temp_files.h"

using namespace wxAutoExcel;

namespace
{

// The OLEObjects collection is not updated when objects are added or
// deleted, so a new one must be obtained to count them.
long GetOLEObjectCount(ExcelTestWorkbook& book)
{
    wxExcelOLEObjects objects = book.GetWorksheet().OLEObjects();
    REQUIRE(objects);
    const long count = objects.GetCount();
    REQUIRE(objects);
    return count;
}

} // unnamed namespace

TEST_CASE("wxExcelCustomProperties adds, changes, and deletes properties",
          "[excel][worksheet][customproperties]")
{
    ExcelTestWorkbook book;
    REQUIRE(book.IsReady());

    wxExcelCustomProperties properties = book.GetWorksheet().GetCustomProperties();
    REQUIRE(properties);
    CHECK(properties.GetCount() == 0);

    wxExcelCustomProperty key = properties.Add("Key", "Value");
    REQUIRE(key);
    REQUIRE(properties.Add("Number", wxVariant(42L)));
    CHECK(properties.GetCount() == 2);
    REQUIRE(properties);

    CHECK(key.GetName() == "Key");
    CHECK(key.GetValue().GetString() == "Value");
    REQUIRE(key);

    key.SetValue("Changed");
    REQUIRE(key);
    CHECK(properties.GetItem(1).GetValue().GetString() == "Changed");

    wxExcelCustomProperty number = properties.GetItem(2);
    REQUIRE(number);
    CHECK(number.GetName() == "Number");
    CHECK(number.GetValue().GetLong() == 42);
    REQUIRE(number);

    long count = 0;
    for ( wxExcelCustomProperty property : properties )
    {
        CHECK(property);
        ++count;
    }
    CHECK(count == 2);

    key.Delete();
    REQUIRE(key);
    CHECK(properties.GetCount() == 1);
    CHECK(properties.GetItem(1).GetName() == "Number");
}

TEST_CASE("wxExcelErrors reports and ignores the errors of a cell",
          "[excel][range][errors]")
{
    ExcelTestWorkbook book;
    REQUIRE(book.IsReady());

    wxExcelErrorCheckingOptions options = wxAutoExcelTests::GetApplication().GetErrorCheckingOptions();
    REQUIRE(options);
    REQUIRE(options.GetNumberAsText());

    wxExcelRange text = book.GetRange("A1");
    text.SetValue("'123");
    REQUIRE(text);
    wxExcelRange number = book.GetRange("A2");
    number.SetValue(wxVariant(123L));
    REQUIRE(number);

    wxExcelError textError = text.GetErrors().GetItem(xlNumberAsText);
    REQUIRE(textError);
    CHECK(textError.GetValue());
    CHECK_FALSE(textError.GetIgnore());
    REQUIRE(textError);

    wxExcelError numberError = number.GetErrors().GetItem(xlNumberAsText);
    REQUIRE(numberError);
    CHECK_FALSE(numberError.GetValue());
    REQUIRE(numberError);

    textError.SetIgnore(true);
    REQUIRE(textError);
    CHECK(text.GetErrors().GetItem(xlNumberAsText).GetIgnore());

    // the checking can be turned off for the whole application
    options.SetNumberAsText(false);
    REQUIRE(options);
    CHECK_FALSE(options.GetNumberAsText());
    options.SetNumberAsText(true);
    REQUIRE(options);
    CHECK(options.GetNumberAsText());
}

TEST_CASE("wxExcelMultiThreadedCalculation changes the thread mode",
          "[excel][application][calculation]")
{
    wxExcelMultiThreadedCalculation calculation =
        wxAutoExcelTests::GetApplication().GetMultiThreadedCalculation();
    REQUIRE(calculation);

    const bool enabled = calculation.GetEnabled();
    const XlThreadMode mode = calculation.GetThreadMode();
    const long count = calculation.GetThreadCount();
    REQUIRE(calculation);
    CHECK(count >= 1);

    calculation.SetEnabled(!enabled);
    REQUIRE(calculation);
    CHECK(calculation.GetEnabled() == !enabled);

    calculation.SetThreadMode(xlThreadModeManual);
    REQUIRE(calculation);
    CHECK(calculation.GetThreadMode() == xlThreadModeManual);

    calculation.SetEnabled(enabled);
    calculation.SetThreadMode(mode);
    REQUIRE(calculation);
    CHECK(calculation.GetEnabled() == enabled);
    CHECK(calculation.GetThreadMode() == mode);
}

TEST_CASE("wxExcelTableStyles adds, formats, and deletes a custom table style",
          "[excel][workbook][tablestyles]")
{
    ExcelTestWorkbook book;
    REQUIRE(book.IsReady());

    wxExcelTableStyles styles = book.GetWorkbook().GetTableStyles();
    REQUIRE(styles);
    const long builtInCount = styles.GetCount();
    REQUIRE(styles);
    CHECK(builtInCount > 0);
    CHECK(styles.GetItem("TableStyleMedium2").GetBuiltIn());

    wxExcelTableStyle style = styles.Add("TestStyle");
    REQUIRE(style);
    CHECK(styles.GetCount() == builtInCount + 1);
    CHECK(style.GetName() == "TestStyle");
    CHECK_FALSE(style.GetBuiltIn());
    REQUIRE(style);

    wxExcelTableStyleElement header = style.GetTableStyleElements().GetItem(xlHeaderRow);
    REQUIRE(header);
    CHECK_FALSE(header.GetHasFormat());

    wxExcelInterior interior = header.GetInterior();
    REQUIRE(interior);
    interior.SetColor(*wxRED);
    REQUIRE(interior);
    CHECK(header.GetHasFormat());
    CHECK(header.GetInterior().GetColor() == *wxRED);
    REQUIRE(header);

    // the style can be used for a table
    book.GetRange("A1").SetValue("Header");
    book.GetRange("A2").SetValue(wxVariant(1L));
    wxExcelRange source = book.GetRange("A1:A2");
    XlYesNoGuess hasHeaders = xlYes;
    wxExcelListObject table = book.GetWorksheet().GetListObjects().Add(&source, &hasHeaders);
    REQUIRE(table);
    table.SetTableStyle(style);
    REQUIRE(table);
    CHECK(table.GetTableStyle().GetName() == "TestStyle");
    CHECK(book.GetRange("A1").GetDisplayFormat().GetInterior().GetColor() == *wxRED);

    header.Clear();
    REQUIRE(header);
    CHECK_FALSE(header.GetHasFormat());

    wxExcelTableStyle copy = style.Duplicate("CopiedStyle");
    REQUIRE(copy);
    CHECK(copy.GetName() == "CopiedStyle");
    CHECK(styles.GetCount() == builtInCount + 2);

    copy.Delete();
    REQUIRE(copy);
    CHECK(styles.GetCount() == builtInCount + 1);
}

TEST_CASE("wxExcelOLEObjects embeds and deletes a file",
          "[excel][worksheet][oleobjects]")
{
    ExcelTestWorkbook book;
    REQUIRE(book.IsReady());

    wxExcelOLEObjects objects = book.GetWorksheet().OLEObjects();
    REQUIRE(objects);
    CHECK(GetOLEObjectCount(book) == 0);

    TempTextFile file;
    double left = 100, top = 50;
    wxExcelOLEObject object = objects.Add(wxEmptyString, file.GetPath(), false, false,
                                          wxEmptyString, NULL, wxEmptyString, &left, &top);
    REQUIRE(object);
    CHECK(GetOLEObjectCount(book) == 1);
    CHECK(object.GetprogID() == "Package");
    // Excel rounds the position to whole pixels
    CHECK(object.GetLeft() == Catch::Approx(100.0).margin(1.0));
    CHECK(object.GetTop() == Catch::Approx(50.0).margin(1.0));
    REQUIRE(object);

    object.SetName("Embedded");
    REQUIRE(object);
    CHECK(objects.Item("Embedded").GetprogID() == "Package");

#if WXAUTOEXCEL_USE_SHAPES
    // the object is a shape too
    wxExcelShape shape = book.GetWorksheet().GetShapes().Item("Embedded");
    REQUIRE(shape);
    CHECK(shape.GetType() == msoEmbeddedOLEObject);
    CHECK(shape.GetOLEFormat().GetprogID() == "Package");
    REQUIRE(shape);
#endif // WXAUTOEXCEL_USE_SHAPES

    CHECK(object.Delete());
    CHECK(GetOLEObjectCount(book) == 0);
}

TEST_CASE("wxExcel collections are indexed with string literals",
          "[excel][workbook][collections]")
{
    ExcelTestWorkbook book;
    REQUIRE(book.IsReady());

    wxExcelWorksheet worksheet = book.GetWorksheet();
    worksheet.SetName("Data");
    REQUIRE(worksheet);

    // a string literal would be ambiguous if the collection converted to a
    // number, as the built-in "Data"[index] could be used then too
    static_assert(std::is_convertible<wxExcelWorksheets, bool>::value,
                  "objects convert to bool");
    static_assert(!std::is_convertible<wxExcelWorksheets, long>::value,
                  "objects do not convert to numbers");
    wxExcelWorksheets worksheets = book.GetWorkbook().GetWorksheets();
    REQUIRE(worksheets);
    CHECK(worksheets["Data"].GetName() == "Data");
    CHECK(worksheets[L"Data"].GetName() == "Data");
    CHECK(worksheets[wxS("Data")].GetName() == "Data");
    CHECK(worksheets[1].GetName() == "Data");
    // the literal 0 must still select the overload taking a number
    // (not evaluated, as Excel indices start at 1)
    static_cast<void>(sizeof(worksheets[0]));

    wxExcelStyles styles = book.GetWorkbook().GetStyles();
    REQUIRE(styles);
    CHECK(styles["Title"].GetFont().GetThemeFont() == xlThemeFontMajor);

#if WXAUTOEXCEL_USE_SHAPES
    wxExcelShapes shapes = worksheet.GetShapes();
    REQUIRE(shapes);
    wxExcelShape box = shapes.AddShape(msoShapeRectangle, 10, 10, 50, 50);
    REQUIRE(box);
    box.SetName("Box");
    REQUIRE(box);
    CHECK(shapes["Box"].GetName() == "Box");
#endif // WXAUTOEXCEL_USE_SHAPES
}
