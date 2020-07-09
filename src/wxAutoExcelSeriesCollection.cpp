/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Modified by:
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
/////////////////////////////////////////////////////////////////////////////


#include "wx/wxAutoExcel_prec.h"

#include "wx/wxAutoExcelSeriesCollection.h"

#if WXAUTOEXCEL_USE_CHARTS

#include "wx/wxAutoExcelRange.h"
#include "wx/wxAutoExcelSeries.h"

#include "wx/wxAutoExcel_private.h"

namespace wxAutoExcel {

// ***** class wxExcelSeriesCollection METHODS *****

wxExcelSeries wxExcelSeriesCollection::Add(wxExcelRange source, XlRowCol* rowcol,
                                           wxXlTribool seriesLabels, wxXlTribool categoryLabels,
                                           wxXlTribool replace)
{
    wxVariant vSource;
    wxExcelSeries series;

    if ( ObjectToVariant(&source, vSource, wxS("Source")) )
    {
        WXAUTOEXCEL_OPTIONALCPP_TO_OPTIONALVARIANT_NAME(Rowcol, ((long*)rowcol));
        WXAUTOEXCEL_OPTIONALCPPTBOOL_TO_OPTIONALVARIANT_NAME(SeriesLabels, seriesLabels);
        WXAUTOEXCEL_OPTIONALCPPTBOOL_TO_OPTIONALVARIANT_NAME(CategoryLabels, categoryLabels);
        WXAUTOEXCEL_OPTIONALCPPTBOOL_TO_OPTIONALVARIANT_NAME(Replace, replace);

        WXAUTOEXCEL_CALL_METHOD5("Add", vSource, vRowcol, vSeriesLabels, vCategoryLabels, vReplace, "void*", series);
        VariantToObject(vResult, &series);
    }
    return series;
}

bool wxExcelSeriesCollection::Extend(wxExcelRange source, XlRowCol* rowcol, wxXlTribool categoryLabels)
{
    wxVariant vSource;

    if ( ObjectToVariant(&source, vSource, wxS("Source")) )
    {
        WXAUTOEXCEL_OPTIONALCPP_TO_OPTIONALVARIANT_NAME(Rowcol, ((long*)rowcol));
        WXAUTOEXCEL_OPTIONALCPPTBOOL_TO_OPTIONALVARIANT_NAME(CategoryLabels, categoryLabels);

        WXAUTOEXCEL_CALL_METHOD3("Extend", vSource, vRowcol, vCategoryLabels, "bool", false);
        return vResult.GetBool();
    } else
        return false;
}

wxExcelSeries wxExcelSeriesCollection::Item(long index)
{
    wxExcelSeries series;
    WXAUTOEXCEL_CALL_METHOD1_OBJECT("Item", index, series);
}

wxExcelSeries wxExcelSeriesCollection::operator[](long index)
{
    return Item(index);
}

wxExcelSeries wxExcelSeriesCollection::NewSeries()
{
    wxExcelSeries series;
    WXAUTOEXCEL_CALL_METHOD0_OBJECT("NewSeries", series);
}

bool wxExcelSeriesCollection::Paste(XlRowCol* rowcol, wxXlTribool seriesLabels, wxXlTribool categoryLabels,
                                    wxXlTribool replace, wxXlTribool newSeries)
{

    WXAUTOEXCEL_OPTIONALCPP_TO_OPTIONALVARIANT_NAME(Rowcol, ((long*)rowcol));
    WXAUTOEXCEL_OPTIONALCPPTBOOL_TO_OPTIONALVARIANT_NAME(SeriesLabels, seriesLabels);
    WXAUTOEXCEL_OPTIONALCPPTBOOL_TO_OPTIONALVARIANT_NAME(CategoryLabels, categoryLabels);
    WXAUTOEXCEL_OPTIONALCPPTBOOL_TO_OPTIONALVARIANT_NAME(Replace, replace);
    WXAUTOEXCEL_OPTIONALCPPTBOOL_TO_OPTIONALVARIANT_NAME(NewSeries, newSeries);

    WXAUTOEXCEL_CALL_METHOD5("Paste", vRowcol, vSeriesLabels, vCategoryLabels, vReplace, vNewSeries, "bool", false);
    return vResult.GetBool();
}

// ***** class wxExcelSeriesCollection PROPERTIES *****


long wxExcelSeriesCollection::GetCount()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("Count");
}


// ***** class wxExcelFullSeriesCollection METHODS *****

wxExcelSeries wxExcelFullSeriesCollection::Item(long index)
{
    wxExcelSeries series;
    WXAUTOEXCEL_CALL_METHOD1_OBJECT("Item", index, series);
}

wxExcelSeries wxExcelFullSeriesCollection::operator[](long index)
{
    return Item(index);
}

// ***** class wxExcelFullSeriesCollection PROPERTIES *****

long wxExcelFullSeriesCollection::GetCount()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("Count");
}


} // namespace wxAutoExcel

#endif // #if WXAUTOEXCEL_USE_CHARTS
