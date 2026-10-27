/////////////////////////////////////////////////////////////////////////////
// Purpose:     Tests for checking the properties and methods of wxExcelObject
// Author:      Claude Opus 5.5 (Anthropic), under the direction of PB
// Copyright:   (c) 2026 PB <pbfordev@gmail.com>
// License:     MIT license
/////////////////////////////////////////////////////////////////////////////

#include <catch2/catch_test_macros.hpp>

#include "excel_test_fixture.h"

using namespace wxAutoExcel;

TEST_CASE("wxExcelObject HasProperty_ and HasMethod_ tell properties and methods apart",
          "[excel][object][members]")
{
    ExcelTestWorkbook excel;
    REQUIRE(excel.IsReady());

    wxExcelRange range = excel.GetRange("A1");

    SECTION("properties")
    {
        CHECK(range.HasProperty_("Value"));
        CHECK(range.HasProperty_("Formula"));
        // read-only
        CHECK(range.HasProperty_("Address"));
        // the case is ignored
        CHECK(range.HasProperty_("numberformat"));
        // hidden
        CHECK(range.HasProperty_("_Default"));

        CHECK_FALSE(range.HasProperty_("Clear"));
        CHECK_FALSE(range.HasProperty_("NoSuchProperty"));
        CHECK_FALSE(range.HasProperty_(""));
    }

    SECTION("methods")
    {
        CHECK(range.HasMethod_("Clear"));
        CHECK(range.HasMethod_("AutoFit"));
        CHECK(range.HasMethod_("clearcontents"));

        CHECK_FALSE(range.HasMethod_("Value"));
        CHECK_FALSE(range.HasMethod_("NoSuchMethod"));
    }

    SECTION("the result matches GetPropertyAndMethodNames_")
    {
        wxArrayString propertyNames, methodNames;
        REQUIRE(range.GetPropertyAndMethodNames_(propertyNames, methodNames, true));

        for ( const wxString& name : propertyNames )
        {
            INFO("property " << name);
            CHECK(range.HasProperty_(name));
            CHECK_FALSE(range.HasMethod_(name));
        }
        for ( const wxString& name : methodNames )
        {
            INFO("method " << name);
            CHECK(range.HasMethod_(name));
            CHECK_FALSE(range.HasProperty_(name));
        }
    }

    // the members of other objects are those of their type
    CHECK(excel.GetWorksheet().HasMethod_("Calculate"));
    CHECK_FALSE(excel.GetWorksheet().HasProperty_("Formula"));
}

TEST_CASE("wxExcelObject HasProperty_ and HasMethod_ return false without a dispatch",
          "[object][members]")
{
    wxExcelRange range;

    CHECK_FALSE(range.HasProperty_("Value"));
    CHECK_FALSE(range.HasMethod_("Clear"));
}
