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

    chartObject.SetName("Results");
    REQUIRE(chartObject);
    const wxString name = chartObject.GetName();
    REQUIRE(chartObject);
    CHECK(name == "Results");
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

TEST_CASE("wxExcelSeries adds error bars",
          "[excel][chart][series]")
{
    ExcelTestWorkbook book;
    REQUIRE(book.IsReady());

    wxExcelChart chart = AddChart(book);

    wxExcelSeries income = chart.SeriesCollection(1);
    REQUIRE(income);
    CHECK_FALSE(income.GetHasErrorBars());
    REQUIRE(income);

    double amount = 5;
    CHECK(income.ErrorBar(xlY, xlErrorBarIncludeBoth, xlErrorBarTypeFixedValue, &amount));
    REQUIRE(income);
    CHECK(income.GetHasErrorBars());
    REQUIRE(income);

    // the standard error does not need an amount
    wxExcelSeries costs = chart.SeriesCollection(2);
    REQUIRE(costs);
    CHECK(costs.ErrorBar(xlY, xlErrorBarIncludePlusValues, xlErrorBarTypeStError));
    REQUIRE(costs);
    CHECK(costs.GetHasErrorBars());
    REQUIRE(costs);
}

TEST_CASE("wxExcelSeries adds custom error bars",
          "[excel][chart][series]")
{
    ExcelTestWorkbook book;
    REQUIRE(book.IsReady());

    wxExcelChart chart = AddChart(book);
    wxExcelRange plus = book.GetRange("E2:E5");
    wxExcelRange minus = book.GetRange("F2:F5");
    plus.SetValue(wxVariant(3.0));
    REQUIRE(plus);
    minus.SetValue(wxVariant(1.0));
    REQUIRE(minus);

    wxExcelSeries series = chart.SeriesCollection(1);
    REQUIRE(series);

    SECTION("from ranges")
    {
        CHECK(series.ErrorBar(xlY, xlErrorBarIncludeBoth, xlErrorBarTypeCustom, plus, minus));
    }

    SECTION("from arrays")
    {
        const wxVariantList empty;
        wxVariant amount(empty), minusValues(empty);
        for ( long i = 0; i < 4; ++i )
        {
            amount.Append(wxVariant(2.0 * (i + 1)));
            minusValues.Append(wxVariant(1.0));
        }
        CHECK(series.ErrorBar(xlY, xlErrorBarIncludeBoth, xlErrorBarTypeCustom, amount, minusValues));
    }

    SECTION("from strings")
    {
        const wxString sheetName = book.GetWorksheet().GetName();
        REQUIRE_FALSE(sheetName.empty());
        CHECK(series.ErrorBar(xlY, xlErrorBarIncludeBoth, xlErrorBarTypeCustom,
                              wxVariant("={1,2,3,4}"), wxVariant("='" + sheetName + "'!$F$2:$F$5")));
    }

    SECTION("from the values of a range and a number")
    {
        CHECK(series.ErrorBar(xlY, xlErrorBarIncludeBoth, xlErrorBarTypeCustom, plus, wxVariant(1.0)));
    }

    REQUIRE(series);
    CHECK(series.GetHasErrorBars());
    REQUIRE(series);
}

TEST_CASE("wxExcelErrorBars changes and deletes the error bars of a series",
          "[excel][chart][series][errorbars]")
{
    ExcelTestWorkbook book;
    REQUIRE(book.IsReady());

    wxExcelChart chart = AddChart(book);
    wxExcelSeries series = chart.SeriesCollection(1);
    REQUIRE(series);

    double amount = 10;
    REQUIRE(series.ErrorBar(xlY, xlErrorBarIncludeBoth, xlErrorBarTypePercent, &amount));

    wxExcelErrorBars errorBars = series.GetErrorBars();
    REQUIRE(errorBars);
    CHECK_FALSE(errorBars.GetName().empty());
    CHECK(errorBars.GetEndStyle() == xlCap);
    REQUIRE(errorBars);

    errorBars.SetEndStyle(xlNoCap);
    REQUIRE(errorBars);
    CHECK(series.GetErrorBars().GetEndStyle() == xlNoCap);

    wxExcelBorder border = errorBars.GetBorder();
    REQUIRE(border);
    border.SetColor(*wxRED);
    REQUIRE(border);
    CHECK(errorBars.GetBorder().GetColor() == *wxRED);

    CHECK(errorBars.ClearFormats());
    CHECK(series.GetHasErrorBars());
    REQUIRE(series);

    CHECK(errorBars.Delete());
    CHECK_FALSE(series.GetHasErrorBars());
    REQUIRE(series);
}

TEST_CASE("wxExcelTrendlines adds trendlines of different types",
          "[excel][chart][series][trendlines]")
{
    ExcelTestWorkbook book;
    REQUIRE(book.IsReady());

    wxExcelChart chart = AddChart(book);
    wxExcelSeries series = chart.SeriesCollection(1);
    REQUIRE(series);

    wxExcelTrendlines trendlines = series.Trendlines();
    REQUIRE(trendlines);
    CHECK(trendlines.GetCount() == 0);
    REQUIRE(trendlines);

    // the type is linear by default
    wxExcelTrendline linear = trendlines.Add();
    REQUIRE(linear);
    CHECK(linear.GetType() == xlLinear);
    CHECK(linear.GetIndex() == 1);
    REQUIRE(linear);

    XlTrendlineType type = xlPolynomial;
    long order = 3;
    wxExcelTrendline polynomial = trendlines.Add(&type, &order);
    REQUIRE(polynomial);
    CHECK(polynomial.GetType() == xlPolynomial);
    CHECK(polynomial.GetOrder() == 3);
    REQUIRE(polynomial);

    type = xlMovingAvg;
    long period = 2;
    wxExcelTrendline movingAverage = trendlines.Add(&type, NULL, &period, NULL, NULL, NULL,
                                                    wxDefaultXlTribool, wxDefaultXlTribool, "Average");
    REQUIRE(movingAverage);
    CHECK(movingAverage.GetType() == xlMovingAvg);
    CHECK(movingAverage.GetPeriod() == 2);
    CHECK(movingAverage.GetName() == "Average");
    CHECK_FALSE(movingAverage.GetNameIsAuto());
    REQUIRE(movingAverage);

    CHECK(trendlines.GetCount() == 3);
    CHECK(trendlines.Item(2).GetType() == xlPolynomial);
    REQUIRE(trendlines);

    REQUIRE(polynomial.Delete());
    CHECK(trendlines.GetCount() == 2);
    CHECK(trendlines.Item(2).GetType() == xlMovingAvg);
    REQUIRE(trendlines);
}

TEST_CASE("wxExcelTrendline changes its forecast, intercept, and label",
          "[excel][chart][series][trendlines]")
{
    ExcelTestWorkbook book;
    REQUIRE(book.IsReady());

    wxExcelChart chart = AddChart(book);
    wxExcelSeries series = chart.SeriesCollection(1);
    REQUIRE(series);
    wxExcelTrendline trendline = series.Trendlines().Add();
    REQUIRE(trendline);

    trendline.SetForward2(1.5);
    trendline.SetBackward2(0.5);
    REQUIRE(trendline);
    CHECK(trendline.GetForward2() == 1.5);
    CHECK(trendline.GetBackward2() == 0.5);
    REQUIRE(trendline);

    CHECK(trendline.GetInterceptIsAuto());
    trendline.SetIntercept(10);
    REQUIRE(trendline);
    CHECK(trendline.GetIntercept() == 10.0);
    CHECK_FALSE(trendline.GetInterceptIsAuto());
    REQUIRE(trendline);

    trendline.SetType(xlExponential);
    REQUIRE(trendline);
    CHECK(trendline.GetType() == xlExponential);

    // showing the R-squared value adds the data label
    CHECK_FALSE(trendline.GetDisplayRSquared());
    trendline.SetDisplayRSquared(true);
    REQUIRE(trendline);
    CHECK(trendline.GetDisplayRSquared());
    // the text of the label is empty until the chart is drawn,
    // which it is not in the hidden Excel instance
    wxExcelDataLabel label = trendline.GetDataLabel();
    REQUIRE(label);

    CHECK(trendline.ClearFormats());
    CHECK(trendline.Delete());
    CHECK(series.Trendlines().GetCount() == 0);
}

#endif // WXAUTOEXCEL_USE_CHARTS
