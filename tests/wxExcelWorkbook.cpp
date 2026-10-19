/////////////////////////////////////////////////////////////////////////////
// Purpose:     Tests for wxExcelWorkbooks, wxExcelWorkbook, and document properties
// Author:      Claude Opus 5.5 (Anthropic), under the direction of PB
// Copyright:   (c) 2026 PB <pbfordev@gmail.com>
// License:     MIT license
/////////////////////////////////////////////////////////////////////////////

#include <catch2/catch_test_macros.hpp>

#include <wx/ffile.h>
#include <wx/filename.h>
#include <wx/utils.h>

#include "excel_test_fixture.h"

using namespace wxAutoExcel;

namespace
{

// A folder for the files saved by a test, deleted with its contents when
// the test ends.
class TempFolder
{
public:
    TempFolder()
    {
        static int counter = 0;
        path = wxFileName::GetTempDir() + wxFileName::GetPathSeparator()
               + wxString::Format("wxAutoExcel_tests_%lu_%d", wxGetProcessId(), ++counter);
        REQUIRE(wxFileName::Mkdir(path, wxS_DIR_DEFAULT, wxPATH_MKDIR_FULL));
    }

    ~TempFolder()
    {
        wxFileName::Rmdir(path, wxPATH_RMDIR_RECURSIVE);
    }

    wxString GetFilePath(const wxString& name) const
    {
        return path + wxFileName::GetPathSeparator() + name;
    }

    const wxString& GetPath() const
    {
        return path;
    }

private:
    wxString path;
};

// Closes the workbook without saving it when the test ends.
class WorkbookCloser
{
public:
    explicit WorkbookCloser(wxExcelWorkbook& workbook)
        : workbook(workbook)
    {}

    ~WorkbookCloser()
    {
        if ( workbook.IsOk_() )
        {
            workbook.SetSaved(true);
            workbook.Close(false);
        }
    }

private:
    wxExcelWorkbook& workbook;
};

long GetWorkbookCount()
{
    wxExcelWorkbooks workbooks = wxAutoExcelTests::GetApplication().GetWorkbooks();
    const long count = workbooks.GetCount();
    REQUIRE(workbooks);
    return count;
}

wxExcelWorkbook OpenWorkbook(const wxString& fileName)
{
    wxExcelWorkbooks workbooks = wxAutoExcelTests::GetApplication().GetWorkbooks();
    REQUIRE(workbooks);
    wxExcelWorkbook workbook = workbooks.Open(fileName);
    REQUIRE(workbook);
    REQUIRE(workbook.SetAutomationLCID_(wxExcelObject::lcidEnglishUS));
    return workbook;
}

wxVariant GetCellValue(wxExcelWorkbook& workbook, const wxString& address)
{
    wxExcelWorksheet worksheet = workbook.GetWorksheets()[1];
    REQUIRE(worksheet);
    wxExcelRange range = worksheet.GetRange(address);
    REQUIRE(range);
    const wxVariant value = range.GetValue();
    REQUIRE(range);
    return value;
}

} // unnamed namespace

TEST_CASE("wxExcelWorkbooks adds and closes workbooks",
          "[excel][workbook]")
{
    const long countBefore = GetWorkbookCount();

    wxExcelWorkbook workbook = wxAutoExcelTests::GetApplication().GetWorkbooks().Add();
    REQUIRE(workbook);
    WorkbookCloser closer(workbook);
    CHECK(GetWorkbookCount() == countBefore + 1);

    REQUIRE(workbook.Close(false));
    CHECK(GetWorkbookCount() == countBefore);
}

TEST_CASE("wxExcelWorkbook GetSaved reflects changes",
          "[excel][workbook]")
{
    ExcelTestWorkbook book;
    REQUIRE(book.IsReady());
    wxExcelWorkbook& workbook = book.GetWorkbook();

    workbook.SetSaved(true);
    REQUIRE(workbook);
    CHECK(workbook.GetSaved());

    book.GetRange("A1").SetValue(wxVariant(1L));
    const bool saved = workbook.GetSaved();
    REQUIRE(workbook);
    CHECK_FALSE(saved);
}

TEST_CASE("wxExcelWorkbook saves as a workbook which can be opened again",
          "[excel][workbook][file]")
{
    TempFolder folder;
    const wxString fileName = folder.GetFilePath("SavedWorkbook.xlsx");

    {
        ExcelTestWorkbook book;
        REQUIRE(book.IsReady());
        wxExcelWorkbook& workbook = book.GetWorkbook();

        book.GetRange("A1").SetValue(wxVariant("Saved text"));
        book.GetRange("B1").SetValue(wxVariant(42.5));

        XlFileFormat fileFormat = xlOpenXMLWorkbook;
        REQUIRE(workbook.SaveAs(fileName, &fileFormat));

        const bool saved = workbook.GetSaved();
        REQUIRE(workbook);
        CHECK(saved);

        CHECK(workbook.GetName() == "SavedWorkbook.xlsx");
        CHECK(wxFileName(workbook.GetFullName()).SameAs(wxFileName(fileName)));
        CHECK(wxFileName::DirName(workbook.GetPath()).SameAs(wxFileName::DirName(folder.GetPath())));

        const XlFileFormat savedFormat = workbook.GetFileFormat();
        REQUIRE(workbook);
        CHECK(savedFormat == xlOpenXMLWorkbook);
    }

    REQUIRE(wxFileName::FileExists(fileName));

    wxExcelWorkbook workbook = OpenWorkbook(fileName);
    WorkbookCloser closer(workbook);

    CHECK(GetCellValue(workbook, "A1").GetString() == "Saved text");
    CHECK(GetCellValue(workbook, "B1").GetDouble() == 42.5);
}

TEST_CASE("wxExcelWorkbook saves as CSV",
          "[excel][workbook][file]")
{
    TempFolder folder;
    const wxString fileName = folder.GetFilePath("SavedWorkbook.csv");

    {
        ExcelTestWorkbook book;
        REQUIRE(book.IsReady());

        book.GetRange("A1").SetValue(wxVariant("Name"));
        book.GetRange("B1").SetValue(wxVariant(42L));

        // CSV files are saved with the separators of the en-US locale, as
        // the Local argument of SaveAs is not used
        XlFileFormat fileFormat = xlCSV;
        REQUIRE(book.GetWorkbook().SaveAs(fileName, &fileFormat));
    }

    wxFFile file(fileName);
    REQUIRE(file.IsOpened());
    wxString content;
    REQUIRE(file.ReadAll(&content));
    content.Trim();
    CHECK(content == "Name,42");
}

TEST_CASE("wxExcelWorkbook ChangeFileAccess makes a saved workbook read-only",
          "[excel][workbook][file]")
{
    TempFolder folder;
    const wxString fileName = folder.GetFilePath("ReadOnly.xlsx");

    ExcelTestWorkbook book;
    REQUIRE(book.IsReady());
    wxExcelWorkbook& workbook = book.GetWorkbook();

    XlFileFormat fileFormat = xlOpenXMLWorkbook;
    REQUIRE(workbook.SaveAs(fileName, &fileFormat));
    CHECK_FALSE(workbook.GetReadOnly());

    workbook.ChangeFileAccess(xlReadOnly, wxEmptyString, false);
    REQUIRE(workbook);
    const bool readOnly = workbook.GetReadOnly();
    REQUIRE(workbook);
    CHECK(readOnly);
}

TEST_CASE("wxExcelWorkbook protects and unprotects its structure",
          "[excel][workbook]")
{
    ExcelTestWorkbook book;
    REQUIRE(book.IsReady());
    wxExcelWorkbook& workbook = book.GetWorkbook();

    CHECK_FALSE(workbook.GetProtectStructure());

    workbook.Protect("password", true);
    REQUIRE(workbook);
    CHECK(workbook.GetProtectStructure());

    // a protected structure prevents adding worksheets
    {
        wxAutoExcelObjectErrorModeOverrider noErrors(wxExcelObject::Err_DoNothing);
        wxExcelWorksheets worksheets = workbook.GetWorksheets();
        CHECK_FALSE(worksheets.Add());
    }

    REQUIRE(workbook.Unprotect("password"));
    CHECK_FALSE(workbook.GetProtectStructure());
    CHECK(book.GetWorksheets().Add());
}

TEST_CASE("wxExcelDocumentProperties built-in properties round-trip",
          "[excel][workbook][properties]")
{
    ExcelTestWorkbook book;
    REQUIRE(book.IsReady());

    wxExcelDocumentProperties properties = book.GetWorkbook().GetBuiltinDocumentProperties();
    REQUIRE(properties);

    const wxString names[] = { "Title", "Subject", "Author", "Keywords", "Comments" };
    for ( const wxString& name : names )
    {
        INFO("Property " << name.utf8_string());

        wxExcelDocumentProperty property = properties.GetItem(name);
        REQUIRE(property);

        const wxString value = "Test " + name;
        property.SetValue(wxVariant(value));
        REQUIRE(property);

        CHECK(property.GetName() == name);
        CHECK(property.GetType() == msoPropertyTypeString);

        const wxVariant read = properties.GetItem(name).GetValue();
        CHECK(read.GetString() == value);
    }
}

TEST_CASE("wxExcelDocumentProperties adds custom properties of different types",
          "[excel][workbook][properties]")
{
    ExcelTestWorkbook book;
    REQUIRE(book.IsReady());

    wxExcelDocumentProperties properties = book.GetWorkbook().GetCustomDocumentProperties();
    REQUIRE(properties);
    const long countBefore = properties.GetCount();
    REQUIRE(properties);

    wxVariant text("Custom text");
    wxExcelDocumentProperty textProperty =
        properties.Add("TextProperty", false, msoPropertyTypeString, &text);
    REQUIRE(textProperty);

    wxVariant number(123L);
    wxExcelDocumentProperty numberProperty =
        properties.Add("NumberProperty", false, msoPropertyTypeNumber, &number);
    REQUIRE(numberProperty);

    wxVariant flag(true);
    wxExcelDocumentProperty flagProperty =
        properties.Add("FlagProperty", false, msoPropertyTypeBoolean, &flag);
    REQUIRE(flagProperty);

    CHECK(properties.GetCount() == countBefore + 3);

    wxExcelDocumentProperty readText = properties.GetItem("TextProperty");
    REQUIRE(readText);
    CHECK(readText.GetType() == msoPropertyTypeString);
    CHECK(readText.GetValue().GetString() == "Custom text");

    wxExcelDocumentProperty readNumber = properties.GetItem("NumberProperty");
    REQUIRE(readNumber);
    CHECK(readNumber.GetType() == msoPropertyTypeNumber);
    CHECK(readNumber.GetValue().GetLong() == 123);

    wxExcelDocumentProperty readFlag = properties.GetItem("FlagProperty");
    REQUIRE(readFlag);
    CHECK(readFlag.GetType() == msoPropertyTypeBoolean);
    CHECK(readFlag.GetValue().GetBool());

    readNumber.Delete();
    REQUIRE(readNumber);
    CHECK(properties.GetCount() == countBefore + 2);
}
