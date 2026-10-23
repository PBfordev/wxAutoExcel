/////////////////////////////////////////////////////////////////////////////
// Purpose:     Tests for objects attached to ranges: wxExcelValidation,
//              wxExcelHyperlinks, wxExcelComments, and
//              wxExcelCommentsThreaded
// Author:      Claude Opus 5.5 (Anthropic), under the direction of PB
// Copyright:   (c) 2026 PB <pbfordev@gmail.com>
// License:     MIT license
/////////////////////////////////////////////////////////////////////////////

#include <catch2/catch_test_macros.hpp>

#include "excel_test_fixture.h"

using namespace wxAutoExcel;

namespace
{

void SetValue(ExcelTestWorkbook& book, const wxString& address, const wxVariant& value)
{
    wxExcelRange range = book.GetRange(address);
    range.SetValue(value);
    REQUIRE(range);
}

wxExcelValidation GetValidation(ExcelTestWorkbook& book, const wxString& address)
{
    wxExcelValidation validation = book.GetRange(address).GetValidation();
    REQUIRE(validation);
    return validation;
}

// Returns whether the value of the cell satisfies its validation.
bool IsValid(ExcelTestWorkbook& book, const wxString& address)
{
    wxExcelValidation validation = GetValidation(book, address);
    const bool valid = validation.GetValue();
    REQUIRE(validation);
    return valid;
}

wxString GetListSeparator()
{
    wxExcelApplication& application = wxAutoExcelTests::GetApplication();
    XlApplicationInternational index = xlListSeparator;
    const wxVariant separator = application.GetInternational(&index);
    REQUIRE(application);
    REQUIRE(separator.IsType("string"));
    return separator.GetString();
}

long GetHyperlinkCount(wxExcelHyperlinks hyperlinks)
{
    REQUIRE(hyperlinks);
    const long count = hyperlinks.GetCount();
    REQUIRE(hyperlinks);
    return count;
}

long GetCommentCount(ExcelTestWorkbook& book)
{
    wxExcelComments comments = book.GetWorksheet().GetComments();
    REQUIRE(comments);
    const long count = comments.GetCount();
    REQUIRE(comments);
    return count;
}

wxString GetText(wxExcelComment comment)
{
    REQUIRE(comment);
    const wxString text = comment.Text();
    REQUIRE(comment);
    return text;
}

} // unnamed namespace

TEST_CASE("wxExcelValidation restricts cells to a list",
          "[excel][validation]")
{
    ExcelTestWorkbook book;
    REQUIRE(book.IsReady());

    // the items of the list are separated with the list separator of the
    // system, not with a comma as in VBA
    const wxString separator = GetListSeparator();
    const wxString list = "Red" + separator + "Green" + separator + "Blue";

    wxExcelValidation validation = GetValidation(book, "A1:A5");
    XlDVAlertStyle alertStyle = xlValidAlertStop;
    XlFormatConditionOperator between = xlBetween;
    validation.Add(xlValidateList, &alertStyle, &between, list);
    REQUIRE(validation);

    CHECK(validation.GetType() == xlValidateList);
    REQUIRE(validation);
    CHECK(validation.GetFormula1() == list);
    REQUIRE(validation);
    CHECK(validation.GetAlertStyle() == xlValidAlertStop);
    REQUIRE(validation);
    CHECK(validation.GetInCellDropdown());
    REQUIRE(validation);

    // the validation applies to every cell of the range
    CHECK(GetValidation(book, "A4").GetType() == xlValidateList);

    // setting the value with automation is not validated, Value tells
    // whether the value is valid
    SetValue(book, "A2", "Green");
    CHECK(IsValid(book, "A2"));
    SetValue(book, "A3", "Purple");
    CHECK_FALSE(IsValid(book, "A3"));

    validation.SetInCellDropdown(false);
    REQUIRE(validation);
    CHECK_FALSE(validation.GetInCellDropdown());
    REQUIRE(validation);
}

TEST_CASE("wxExcelValidation restricts cells to numbers and is modified and deleted",
          "[excel][validation]")
{
    ExcelTestWorkbook book;
    REQUIRE(book.IsReady());

    wxExcelValidation validation = GetValidation(book, "B1:B3");
    XlDVAlertStyle alertStyle = xlValidAlertWarning;
    XlFormatConditionOperator between = xlBetween;
    validation.Add(xlValidateWholeNumber, &alertStyle, &between, "1", "10");
    REQUIRE(validation);

    CHECK(validation.GetType() == xlValidateWholeNumber);
    REQUIRE(validation);
    CHECK(validation.GetOperator() == xlBetween);
    REQUIRE(validation);
    CHECK(validation.GetFormula1() == "1");
    REQUIRE(validation);
    CHECK(validation.GetFormula2() == "10");
    REQUIRE(validation);

    SetValue(book, "B1", wxVariant(5.0));
    CHECK(IsValid(book, "B1"));
    SetValue(book, "B2", wxVariant(11.0));
    CHECK_FALSE(IsValid(book, "B2"));
    SetValue(book, "B3", wxVariant(2.5));
    CHECK_FALSE(IsValid(book, "B3"));

    validation.SetInputTitle("Number");
    validation.SetInputMessage("Enter a whole number.");
    validation.SetErrorTitle("Wrong number");
    validation.SetErrorMessage("The number must be greater than 10.");
    REQUIRE(validation);
    CHECK(validation.GetInputTitle() == "Number");
    CHECK(validation.GetInputMessage() == "Enter a whole number.");
    CHECK(validation.GetErrorTitle() == "Wrong number");
    REQUIRE(validation);

    XlDVType wholeNumber = xlValidateWholeNumber;
    XlFormatConditionOperator greater = xlGreater;
    validation.Modify(&wholeNumber, &alertStyle, &greater, "10");
    REQUIRE(validation);
    CHECK(validation.GetOperator() == xlGreater);
    REQUIRE(validation);
    CHECK(validation.GetFormula1() == "10");
    REQUIRE(validation);
    CHECK_FALSE(IsValid(book, "B1"));
    CHECK(IsValid(book, "B2"));

    validation.Delete();
    REQUIRE(validation);
    // any value of a cell without validation is valid, but it has no type
    CHECK(IsValid(book, "B2"));
    {
        wxLogNull noLog;
        wxExcelValidation deleted = GetValidation(book, "B1");
        deleted.GetType();
        CHECK_FALSE(deleted);
    }
}

TEST_CASE("wxExcelHyperlinks adds links to a web address and to a range",
          "[excel][hyperlinks]")
{
    ExcelTestWorkbook book;
    REQUIRE(book.IsReady());

    wxExcelWorksheet& worksheet = book.GetWorksheet();
    worksheet.SetName("Data");
    REQUIRE(worksheet);

    wxExcelHyperlinks hyperlinks = worksheet.GetHyperlinks();
    REQUIRE(hyperlinks);
    CHECK(GetHyperlinkCount(hyperlinks) == 0);

    wxExcelRange webCell = book.GetRange("A1");
    wxExcelHyperlink web = hyperlinks.Add(&webCell, "https://www.wxwidgets.org/",
                                          "", "The wxWidgets website", "wxWidgets");
    REQUIRE(web);
    CHECK(web.GetAddress() == "https://www.wxwidgets.org/");
    REQUIRE(web);
    CHECK(web.GetScreenTip() == "The wxWidgets website");
    REQUIRE(web);
    CHECK(web.GetTextToDisplay() == "wxWidgets");
    REQUIRE(web);
    CHECK(web.GetType() == msoHyperlinkRange);
    REQUIRE(web);

    wxExcelRange linkRange = web.GetRange();
    REQUIRE(linkRange);
    CHECK(linkRange.GetAddress() == "$A$1");
    REQUIRE(linkRange);

    // the text to display is the value of the cell
    const wxVariant value = webCell.GetValue();
    REQUIRE(webCell);
    REQUIRE(value.IsType("string"));
    CHECK(value.GetString() == "wxWidgets");

    wxExcelRange internalCell = book.GetRange("A2");
    wxExcelHyperlink internal = hyperlinks.Add(&internalCell, "", "Data!C10");
    REQUIRE(internal);
    CHECK(internal.GetAddress().empty());
    REQUIRE(internal);
    CHECK(internal.GetSubAddress() == "Data!C10");
    REQUIRE(internal);

    CHECK(GetHyperlinkCount(hyperlinks) == 2);
    CHECK(GetHyperlinkCount(webCell.GetHyperlinks()) == 1);

    internal.SetSubAddress("Data!D20");
    REQUIRE(internal);
    CHECK(internal.GetSubAddress() == "Data!D20");
    REQUIRE(internal);

    web.Delete();
    REQUIRE(web);
    CHECK(GetHyperlinkCount(hyperlinks) == 1);

    CHECK(internalCell.ClearHyperlinks());
    CHECK(GetHyperlinkCount(hyperlinks) == 0);
}

TEST_CASE("wxExcelComment is added, changed, and deleted",
          "[excel][comments]")
{
    ExcelTestWorkbook book;
    REQUIRE(book.IsReady());

    CHECK(GetCommentCount(book) == 0);

    wxExcelRange firstCell = book.GetRange("B2");
    wxExcelComment first = firstCell.AddComment("First note");
    REQUIRE(first);
    CHECK(GetText(first) == "First note");
    CHECK(GetText(firstCell.GetComment()) == "First note");

    wxExcelRange secondCell = book.GetRange("D4");
    wxExcelComment second = secondCell.AddComment("Second note");
    REQUIRE(second);
    CHECK(GetCommentCount(book) == 2);

    CHECK(GetText(first.Next()) == "Second note");
    CHECK(GetText(second.Previous()) == "First note");

    wxExcelComments comments = book.GetWorksheet().GetComments();
    REQUIRE(comments);
    CHECK(GetText(comments.GetItem(2)) == "Second note");

    // Text with a start position inserts the text, without it replaces it
    long start = 7;
    first.Text("long ", &start, false);
    REQUIRE(first);
    CHECK(GetText(first) == "First long note");
    first.Text("Changed");
    REQUIRE(first);
    CHECK(GetText(first) == "Changed");

    CHECK_FALSE(first.GetVisible());
    REQUIRE(first);
    first.SetVisible(true);
    REQUIRE(first);
    CHECK(first.GetVisible());
    REQUIRE(first);

    // a cell can have only one comment
    {
        wxLogNull noLog;
        wxExcelComment another = firstCell.AddComment("Another note");
        CHECK_FALSE(another);
        CHECK_FALSE(firstCell);
    }

    second.Delete();
    REQUIRE(second);
    CHECK(GetCommentCount(book) == 1);

    wxExcelRange used = book.GetRange("A1:E5");
    CHECK(used.ClearComments());
    CHECK(GetCommentCount(book) == 0);
}

TEST_CASE("wxExcelCommentThreaded is added with replies and deleted",
          "[excel][comments][threaded]")
{
    ExcelTestWorkbook book;
    REQUIRE(book.IsReady());

    wxExcelCommentsThreaded threads = book.GetWorksheet().GetCommentsThreaded();
    REQUIRE(threads);
    CHECK(threads.GetCount() == 0);
    REQUIRE(threads);

    wxExcelRange cell = book.GetRange("C3");
    wxExcelCommentThreaded comment = cell.AddCommentThreaded("Is this right?");
    REQUIRE(comment);
    CHECK(comment.Text() == "Is this right?");
    REQUIRE(comment);
    CHECK(threads.GetCount() == 1);
    REQUIRE(threads);

    wxExcelCommentThreaded fromCell = cell.GetCommentThreaded();
    REQUIRE(fromCell);
    CHECK(fromCell.Text() == "Is this right?");
    REQUIRE(fromCell);

    wxExcelCommentThreaded reply = comment.AddReply("Yes, it is.");
    REQUIRE(reply);
    CHECK(reply.Text() == "Yes, it is.");
    REQUIRE(reply);

    wxExcelCommentsThreaded replies = comment.GetReplies();
    REQUIRE(replies);
    CHECK(replies.GetCount() == 1);
    REQUIRE(replies);
    wxExcelCommentThreaded firstReply = replies.GetItem(1);
    REQUIRE(firstReply);
    CHECK(firstReply.Text() == "Yes, it is.");
    REQUIRE(firstReply);

    // the replies are not top-level comments
    CHECK(threads.GetCount() == 1);
    REQUIRE(threads);

    comment.Delete();
    REQUIRE(comment);
    CHECK(threads.GetCount() == 0);
    REQUIRE(threads);
}
