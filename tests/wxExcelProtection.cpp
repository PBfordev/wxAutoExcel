/////////////////////////////////////////////////////////////////////////////
// Purpose:     Tests for worksheet protection: wxExcelProtection,
//              wxExcelAllowEditRanges, and wxExcelUserAccessList
// Author:      Claude Opus 5.5 (Anthropic), under the direction of PB
// Copyright:   (c) 2026 PB <pbfordev@gmail.com>
// License:     MIT license
/////////////////////////////////////////////////////////////////////////////

#include <catch2/catch_test_macros.hpp>

#include "excel_test_fixture.h"

using namespace wxAutoExcel;

namespace
{

wxExcelProtection GetProtection(ExcelTestWorkbook& book)
{
    wxExcelProtection protection = book.GetWorksheet().GetProtection();
    REQUIRE(protection);
    return protection;
}

wxExcelAllowEditRanges GetAllowEditRanges(ExcelTestWorkbook& book)
{
    wxExcelAllowEditRanges ranges = GetProtection(book).GetAllowEditRanges();
    REQUIRE(ranges);
    return ranges;
}

long GetCount(wxExcelAllowEditRanges& ranges)
{
    const long count = ranges.GetCount();
    REQUIRE(ranges);
    return count;
}

// Returns true if the value of the cell could be set.
bool CanSetValue(ExcelTestWorkbook& book, const wxString& address)
{
    wxAutoExcelObjectErrorModeOverrider noErrors(wxExcelObject::Err_DoNothing, true);
    wxExcelRange cell = book.GetRange(address);
    cell.SetValue(wxVariant(42L));
    return static_cast<bool>(cell);
}

} // unnamed namespace

TEST_CASE("wxExcelWorksheet protection allows only the given operations",
          "[excel][protection]")
{
    ExcelTestWorkbook book;
    REQUIRE(book.IsReady());
    wxExcelWorksheet& worksheet = book.GetWorksheet();

    CHECK_FALSE(worksheet.GetProtectContents());
    CHECK(CanSetValue(book, "A1"));

    worksheet.Protect("password", wxDefaultXlTribool, true, wxDefaultXlTribool, wxDefaultXlTribool,
                      true,                  // allowFormattingCells
                      false, false, false, false, false, false, false,
                      true,                  // allowSorting
                      true);                 // allowFiltering
    REQUIRE(worksheet);
    CHECK(worksheet.GetProtectContents());

    wxExcelProtection protection = GetProtection(book);
    CHECK(protection.GetAllowFormattingCells());
    CHECK_FALSE(protection.GetAllowFormattingColumns());
    CHECK_FALSE(protection.GetAllowFormattingRows());
    CHECK_FALSE(protection.GetAllowInsertingColumns());
    CHECK_FALSE(protection.GetAllowInsertingRows());
    CHECK_FALSE(protection.GetAllowInsertingHyperlinks());
    CHECK_FALSE(protection.GetAllowDeletingColumns());
    CHECK_FALSE(protection.GetAllowDeletingRows());
    CHECK(protection.GetAllowSorting());
    CHECK(protection.GetAllowFiltering());
    CHECK_FALSE(protection.GetAllowUsingPivotTables());
    REQUIRE(protection);

    // the cells are locked by default
    CHECK_FALSE(CanSetValue(book, "A1"));

    REQUIRE(worksheet.Unprotect("password"));
    CHECK_FALSE(worksheet.GetProtectContents());
    CHECK(CanSetValue(book, "A1"));
}

TEST_CASE("wxExcelAllowEditRanges adds, renames, and deletes ranges",
          "[excel][protection]")
{
    ExcelTestWorkbook book;
    REQUIRE(book.IsReady());

    wxExcelAllowEditRanges ranges = GetAllowEditRanges(book);
    CHECK(GetCount(ranges) == 0);

    wxExcelRange inputs = book.GetRange("B2:C5");
    wxExcelAllowEditRange range = ranges.Add("Inputs", inputs);
    REQUIRE(range);
    CHECK(GetCount(ranges) == 1);
    CHECK(range.GetTitle() == "Inputs");
    CHECK(range.GetRange().GetAddress(false, false) == "B2:C5");
    REQUIRE(range);

    range.SetTitle("Data");
    REQUIRE(range);
    wxExcelAllowEditRange found = ranges.GetItem("Data");
    REQUIRE(found);
    CHECK(found.GetRange().GetAddress(false, false) == "B2:C5");
    REQUIRE(found);

    range.SetRange(book.GetRange("D1:D3"));
    REQUIRE(range);
    CHECK(range.GetRange().GetAddress(false, false) == "D1:D3");
    REQUIRE(range);

    range.ChangePassword("secret");
    REQUIRE(range);

    wxExcelRange other = book.GetRange("E1");
    REQUIRE(ranges.Add("Other", other));
    CHECK(GetCount(ranges) == 2);
    CHECK(ranges.GetItem(2).GetTitle() == "Other");

    range.Delete();
    REQUIRE(range);
    CHECK(GetCount(ranges) == 1);
    CHECK(ranges.GetItem(1).GetTitle() == "Other");
}

TEST_CASE("wxExcelAllowEditRange can be edited on a protected worksheet",
          "[excel][protection]")
{
    ExcelTestWorkbook book;
    REQUIRE(book.IsReady());

    wxExcelAllowEditRanges ranges = GetAllowEditRanges(book);
    wxExcelRange inputs = book.GetRange("B2");
    REQUIRE(ranges.Add("Inputs", inputs));

    book.GetWorksheet().Protect("password");
    REQUIRE(book.GetWorksheet());

    CHECK(CanSetValue(book, "B2"));
    CHECK_FALSE(CanSetValue(book, "A1"));

    // the ranges cannot be changed while the worksheet is protected
    {
        wxAutoExcelObjectErrorModeOverrider noErrors(wxExcelObject::Err_DoNothing, true);
        wxExcelRange other = book.GetRange("C3");
        CHECK_FALSE(ranges.Add("Other", other));
    }

    REQUIRE(book.GetWorksheet().Unprotect("password"));
    CHECK(GetCount(ranges) == 1);
}

TEST_CASE("wxExcelUserAccessList adds and removes users of a range",
          "[excel][protection]")
{
    ExcelTestWorkbook book;
    REQUIRE(book.IsReady());

    wxExcelAllowEditRanges ranges = GetAllowEditRanges(book);
    wxExcelRange inputs = book.GetRange("B2:C5");
    wxExcelAllowEditRange range = ranges.Add("Inputs", inputs);
    REQUIRE(range);

    wxExcelUserAccessList users = range.GetUsers();
    REQUIRE(users);
    CHECK(users.GetCount() == 0);

    wxExcelUserAccess user = users.Add("Everyone", true);
    REQUIRE(user);
    CHECK(users.GetCount() == 1);
    CHECK(user.GetAllowEdit());
    CHECK_FALSE(user.GetName().empty());
    REQUIRE(user);

    user.SetAllowEdit(false);
    REQUIRE(user);
    CHECK_FALSE(users.GetItem(1).GetAllowEdit());

    users.DeleteAll();
    REQUIRE(users);
    CHECK(users.GetCount() == 0);
}
