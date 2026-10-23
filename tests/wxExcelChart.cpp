/////////////////////////////////////////////////////////////////////////////
// Purpose:     Tests for embedded charts and chart sheets: wxExcelChartObjects,
//              wxExcelChart, wxExcelCharts, and wxExcelSeriesCollection
// Author:      Claude Opus 5.5 (Anthropic), under the direction of PB
// Copyright:   (c) 2026 PB <pbfordev@gmail.com>
// License:     MIT license
/////////////////////////////////////////////////////////////////////////////

#include <catch2/catch_test_macros.hpp>

#include "excel_test_fixture.h"

#if WXAUTOEXCEL_USE_CHARTS

using namespace wxAutoExcel;

namespace
{

void SetValue(ExcelTestWorkbook& book, const wxString& address, const wxVariant& value)
{
    wxExcelRange range = book.GetRange(address);
    range.SetValue(value);
    REQUIRE(range);
}

// Fills A1:C5 with a header row (Month, Income, Costs) and four months.
wxExcelRange FillData(ExcelTestWorkbook& book)
{
    SetValue(book, "A1", "Month");
    SetValue(book, "B1", "Income");
    SetValue(book, "C1", "Costs");

    static const char* const months[] = { "Jan", "Feb", "Mar", "Apr" };
    for ( long i = 0; i < 4; ++i )
    {
        const wxString row = wxString::Format("%ld", i + 2);
        SetValue(book, "A" + row, months[i]);
        SetValue(book, "B" + row, wxVariant(100.0 + 10 * i));
        SetValue(book, "C" + row, wxVariant(80.0 - 5 * i));
    }

    return book.GetRange("A1:C5");
}

wxExcelChartObjects GetChartObjects(ExcelTestWorkbook& book)
{
    wxExcelChartObjects chartObjects = book.GetWorksheet().ChartObjects();
    REQUIRE(chartObjects);
    return chartObjects;
}

long GetCount(wxExcelChartObjects chartObjects)
{
    const long count = chartObjects.GetCount();
    REQUIRE(chartObjects);
    return count;
}

long GetSeriesCount(wxExcelChart& chart)
{
    wxExcelSeriesCollection seriesCollection = chart.SeriesCollection();
    REQUIRE(seriesCollection);
    const long count = seriesCollection.GetCount();
    REQUIRE(seriesCollection);
    return count;
}

wxString GetSeriesName(wxExcelChart& chart, long index)
{
    wxExcelSeries series = chart.SeriesCollection(index);
    REQUIRE(series);
    const wxString name = series.GetName();
    REQUIRE(series);
    return name;
}

wxExcelChart AddChart(ExcelTestWorkbook& book, wxExcelChartObject* chartObjectOut = NULL)
{
    wxExcelRange data = FillData(book);

    wxExcelChartObject chartObject = GetChartObjects(book).Add(250, 20, 400, 250);
    REQUIRE(chartObject);
    wxExcelChart chart = chartObject.GetChart();
    REQUIRE(chart);

    chart.SetSourceData(data);
    REQUIRE(chart);

    if ( chartObjectOut )
        *chartObjectOut = chartObject;
    return chart;
}

} // unnamed namespace

TEST_CASE("wxExcelChartObjects adds an embedded chart of a given size",
          "[excel][chart]")
{
    ExcelTestWorkbook book;
    REQUIRE(book.IsReady());

    wxExcelChartObjects chartObjects = GetChartObjects(book);
    CHECK(GetCount(chartObjects) == 0);

    wxExcelChartObject chartObject;
    AddChart(book, &chartObject);
    CHECK(GetCount(chartObjects) == 1);

    CHECK(chartObject.GetLeft() == 250.0);
    CHECK(chartObject.GetTop() == 20.0);
    CHECK(chartObject.GetWidth() == 400.0);
    CHECK(chartObject.GetHeight() == 250.0);
    REQUIRE(chartObject);

    chartObject.SetLeft(10.0);
    chartObject.SetWidth(300.0);
    REQUIRE(chartObject);
    CHECK(chartObject.GetLeft() == 10.0);
    CHECK(chartObject.GetWidth() == 300.0);
    REQUIRE(chartObject);

    wxExcelRange topLeft = chartObject.GetTopLeftCell();
    REQUIRE(topLeft);
    CHECK(topLeft.GetAddress() == "$A$2");
    REQUIRE(topLeft);

    const wxString name = chartObject.GetName();
    REQUIRE(chartObject);
    wxExcelChartObject byName = book.GetWorksheet().ChartObjects(name);
    REQUIRE(byName);
    CHECK(byName.GetIndex() == 1);
    REQUIRE(byName);

    CHECK(chartObject.Delete());
    CHECK(GetCount(chartObjects) == 0);
}

TEST_CASE("wxExcelChart creates series from its source data",
          "[excel][chart][series]")
{
    ExcelTestWorkbook book;
    REQUIRE(book.IsReady());

    wxExcelChart chart = AddChart(book);

    // the series are in the columns, named after the header row
    CHECK(GetSeriesCount(chart) == 2);
    CHECK(GetSeriesName(chart, 1) == "Income");
    CHECK(GetSeriesName(chart, 2) == "Costs");

    wxExcelSeries income = chart.SeriesCollection("Income");
    REQUIRE(income);
    CHECK(income.GetFormula().Contains("$B$2:$B$5"));
    REQUIRE(income);

    // the series in the rows, named after the first column
    XlRowCol byRows = xlRows;
    chart.SetSourceData(book.GetRange("A1:C5"), &byRows);
    REQUIRE(chart);
    CHECK(GetSeriesCount(chart) == 4);
    CHECK(GetSeriesName(chart, 1) == "Jan");

    wxExcelSeries series = chart.SeriesCollection(1);
    REQUIRE(series);
    series.SetName("January");
    REQUIRE(series);
    CHECK(GetSeriesName(chart, 1) == "January");

    wxExcelSeries added = chart.SeriesCollection().NewSeries();
    REQUIRE(added);
    added.SetFormula("=SERIES(\"Extra\",,{1,2},5)");
    REQUIRE(added);
    CHECK(GetSeriesCount(chart) == 5);
    CHECK(GetSeriesName(chart, 5) == "Extra");
}

TEST_CASE("wxExcelChart changes its type and title",
          "[excel][chart]")
{
    ExcelTestWorkbook book;
    REQUIRE(book.IsReady());

    wxExcelChart chart = AddChart(book);

    chart.SetChartType(xlColumnClustered);
    REQUIRE(chart);
    CHECK(chart.GetChartType() == xlColumnClustered);
    REQUIRE(chart);

    chart.SetChartType(xlLine);
    REQUIRE(chart);
    CHECK(chart.GetChartType() == xlLine);
    REQUIRE(chart);

    chart.SetHasTitle(true);
    REQUIRE(chart);
    CHECK(chart.GetHasTitle());
    REQUIRE(chart);

    wxExcelChartTitle title = chart.GetChartTitle();
    REQUIRE(title);
    title.SetText("Income and costs");
    REQUIRE(title);
    CHECK(title.GetText() == "Income and costs");
    REQUIRE(title);

    chart.SetHasTitle(false);
    REQUIRE(chart);
    CHECK_FALSE(chart.GetHasTitle());
    REQUIRE(chart);
}

TEST_CASE("wxExcelChart moves from the worksheet to a chart sheet",
          "[excel][chart][charts]")
{
    ExcelTestWorkbook book;
    REQUIRE(book.IsReady());

    wxExcelChart chart = AddChart(book);
    wxExcelCharts charts = book.GetWorkbook().GetCharts();
    REQUIRE(charts);
    CHECK(charts.GetCount() == 0);
    REQUIRE(charts);

    wxExcelChart chartSheet = chart.Location(xlLocationAsNewSheet, "Overview");
    REQUIRE(chartSheet);
    CHECK(chartSheet.GetName() == "Overview");
    REQUIRE(chartSheet);
    CHECK(GetSeriesCount(chartSheet) == 2);

    CHECK(charts.GetCount() == 1);
    REQUIRE(charts);
    CHECK(GetCount(GetChartObjects(book)) == 0);

    wxExcelChart added = charts.Add();
    REQUIRE(added);
    CHECK(charts.GetCount() == 2);
    REQUIRE(charts);

    added.Delete();
    REQUIRE(added);
    CHECK(charts.GetCount() == 1);
    REQUIRE(charts);
}

#endif // WXAUTOEXCEL_USE_CHARTS
