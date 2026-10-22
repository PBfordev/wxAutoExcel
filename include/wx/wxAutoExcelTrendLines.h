/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// License:     MIT license
/////////////////////////////////////////////////////////////////////////////


#ifndef _WXAUTOEXCEL_TRENDLINES_H
#define _WXAUTOEXCEL_TRENDLINES_H

#include "wx/wxAutoExcel_defs.h"

#if WXAUTOEXCEL_USE_CHARTS

#include "wx/wxAutoExcel_object.h"
#include "wx/wxAutoExcel_enums.h"

namespace wxAutoExcel {


    /**
    @brief Represents Microsoft Excel Trendline object.
    */
   class WXDLLIMPEXP_WXAUTOEXCEL wxExcelTrendline : public wxExcelObject
    {
    public:
        // ***** METHODS *****

        /**
        Clears the formatting of the object.

        [MSDN documentation for Trendline.ClearFormats]().
        */
        bool ClearFormats();

        /**
        Deletes the object.

        [MSDN documentation for Trendline.Delete](https://learn.microsoft.com/en-us/office/vba/api/excel.trendline.delete).
        */
        bool Delete();

        /**
        Selects the object.

        [MSDN documentation for Trendline.Select](https://learn.microsoft.com/en-us/office/vba/api/excel.trendline.select).
        */
        bool Select();

        // ***** PROPERTIES *****


        /**
        Returns the number of periods (or units on a scatter chart) that the trendline extends backward. Read/write Long

        [MSDN documentation for Trendline.Backward](https://learn.microsoft.com/en-us/previous-versions/office/developer/office-2007/bb220877%28v=office.12%29).
        */
        long GetBackward();

        /**
        Sets the number of periods (or units on a scatter chart) that the trendline extends backward. Read/write Long

        [MSDN documentation for Trendline.Backward](https://learn.microsoft.com/en-us/previous-versions/office/developer/office-2007/bb220877%28v=office.12%29).
        */
        void SetBackward(long backward);

        /**
        Returns the number of periods (or units on a scatter chart) that the trendline extends backward.  Since Excel 2007.

        [MSDN documentation for Trendline.Backward2](https://learn.microsoft.com/en-us/office/vba/api/excel.trendline.backward2).
        */
        double GetBackward2();

        /**
        Sets the number of periods (or units on a scatter chart) that the trendline extends backward.  Since Excel 2007.

        [MSDN documentation for Trendline.Backward2](https://learn.microsoft.com/en-us/office/vba/api/excel.trendline.backward2).
        */
        void SetBackward2(double backward2);

        /**
        Returns a Border object that represents the border of the object.

        [MSDN documentation for Trendline.Border](https://learn.microsoft.com/en-us/office/vba/api/excel.trendline.border).
        */
        wxExcelBorder GetBorder();

        /**
        Returns a DataLabel object that represents the data label associated with the trendline.

        [MSDN documentation for Trendline.DataLabel](https://learn.microsoft.com/en-us/office/vba/api/excel.trendline.datalabel).
        */
        wxExcelDataLabel GetDataLabel();

        /**
        True if the equation for the trendline is displayed on the chart (in the same data label as the R-squared value). Setting this property to True automatically turns on data labels.

        [MSDN documentation for Trendline.DisplayEquation](https://learn.microsoft.com/en-us/office/vba/api/excel.trendline.displayequation).
        */
        bool GetDisplayEquation();

        /**
        True if the equation for the trendline is displayed on the chart (in the same data label as the R-squared value). Setting this property to True automatically turns on data labels.

        [MSDN documentation for Trendline.DisplayEquation](https://learn.microsoft.com/en-us/office/vba/api/excel.trendline.displayequation).
        */
        void SetDisplayEquation(bool displayEquation);

        /**
        True if the R-squared value of the trendline is displayed on the chart (in the same data label as the equation). Setting this property to True automatically turns on data labels.

        [MSDN documentation for Trendline.DisplayRSquared](https://learn.microsoft.com/en-us/office/vba/api/excel.trendline.displayrsquared).
        */
        bool GetDisplayRSquared();

        /**
        True if the R-squared value of the trendline is displayed on the chart (in the same data label as the equation). Setting this property to True automatically turns on data labels.

        [MSDN documentation for Trendline.DisplayRSquared](https://learn.microsoft.com/en-us/office/vba/api/excel.trendline.displayrsquared).
        */
        void SetDisplayRSquared(bool displayRSquared);

        /**
        Read-only Since Excel 2007.

        [MSDN documentation for Trendline.Format](https://learn.microsoft.com/en-us/office/vba/api/excel.trendline.format).
        */
        wxExcelChartFormat GetFormat();

        /**
        Returns the number of periods (or units on a scatter chart) that the trendline extends forward.

        [MSDN documentation for Trendline.Forward]().
        */
        long GetForward();

        /**
        Sets the number of periods (or units on a scatter chart) that the trendline extends forward.

        [MSDN documentation for Trendline.Forward]().
        */
        void SetForward(long forward);

        /**
        Returns the ChartFormat object.  Since Excel 2007.

        [MSDN documentation for Trendline.Forward2](https://learn.microsoft.com/en-us/office/vba/api/excel.trendline.forward2).
        */
        double GetForward2();

        /**
        Returns the ChartFormat object.  Since Excel 2007.

        [MSDN documentation for Trendline.Forward2](https://learn.microsoft.com/en-us/office/vba/api/excel.trendline.forward2).
        */
        void SetForward2(double forward2);

        /**
        Returns a Long value that represents the index number of the object within the collection of similar objects.

        [MSDN documentation for Trendline.Index](https://learn.microsoft.com/en-us/office/vba/api/excel.trendline.index).
        */
        long GetIndex();

        /**
        Returns the point where the trendline crosses the value axis. Read/write Double.

        [MSDN documentation for Trendline.Intercept](https://learn.microsoft.com/en-us/office/vba/api/excel.trendline.intercept).
        */
        double GetIntercept();

        /**
        Sets the point where the trendline crosses the value axis. Read/write Double.

        [MSDN documentation for Trendline.Intercept](https://learn.microsoft.com/en-us/office/vba/api/excel.trendline.intercept).
        */
        void SetIntercept(double intercept);

        /**
        True if the point where the trendline crosses the value axis is automatically determined by the regression.

        [MSDN documentation for Trendline.InterceptIsAuto](https://learn.microsoft.com/en-us/office/vba/api/excel.trendline.interceptisauto).
        */
        bool GetInterceptIsAuto();

        /**
        True if the point where the trendline crosses the value axis is automatically determined by the regression.

        [MSDN documentation for Trendline.InterceptIsAuto](https://learn.microsoft.com/en-us/office/vba/api/excel.trendline.interceptisauto).
        */
        void SetInterceptIsAuto(bool interceptIsAuto);

        /**
        Returns a String value representing the name of the object.

        [MSDN documentation for Trendline.Name](https://learn.microsoft.com/en-us/office/vba/api/excel.trendline.name).
        */
        wxString GetName();

        /**
        Sets a String value representing the name of the object.

        [MSDN documentation for Trendline.Name](https://learn.microsoft.com/en-us/office/vba/api/excel.trendline.name).
        */
        void SetName(const wxString& name);

        /**
        True if Microsoft Excel automatically determines the name of the trendline.

        [MSDN documentation for Trendline.NameIsAuto](https://learn.microsoft.com/en-us/office/vba/api/excel.trendline.nameisauto).
        */
        bool GetNameIsAuto();

        /**
        True if Microsoft Excel automatically determines the name of the trendline.

        [MSDN documentation for Trendline.NameIsAuto](https://learn.microsoft.com/en-us/office/vba/api/excel.trendline.nameisauto).
        */
        void SetNameIsAuto(bool nameIsAuto);

        /**
        Returns a Long value that represents the trendline order (an integer greater than 1) when the trendline type is xlPolynomial.

        [MSDN documentation for Trendline.Order](https://learn.microsoft.com/en-us/office/vba/api/excel.trendline.order).
        */
        long GetOrder();

        /**
        Sets a Long value that represents the trendline order (an integer greater than 1) when the trendline type is xlPolynomial.

        [MSDN documentation for Trendline.Order](https://learn.microsoft.com/en-us/office/vba/api/excel.trendline.order).
        */
        void SetOrder(long order);

        /**
        Returns the period for the moving-average trendline. Can be a value from 2 through 255.

        [MSDN documentation for Trendline.Period](https://learn.microsoft.com/en-us/office/vba/api/excel.trendline.period).
        */
        long GetPeriod();

        /**
        Sets the period for the moving-average trendline. Can be a value from 2 through 255.

        [MSDN documentation for Trendline.Period](https://learn.microsoft.com/en-us/office/vba/api/excel.trendline.period).
        */
        void SetPeriod(long period);

        /**
        Returns a XlTrendlineType value that represents the trendline type.

        [MSDN documentation for Trendline.Type](https://learn.microsoft.com/en-us/office/vba/api/excel.trendline.type).
        */
        XlTrendlineType GetType();

        /**
        Sets a XlTrendlineType value that represents the trendline type.

        [MSDN documentation for Trendline.Type](https://learn.microsoft.com/en-us/office/vba/api/excel.trendline.type).
        */
        void SetType(XlTrendlineType type);
        /**
        Returns "Trendline".
        */
        virtual wxString GetAutoExcelObjectName_() const { return wxS("Trendline"); }
    };

    /**
    @brief Represents Microsoft Excel Trendlines collection.
    */
   class WXDLLIMPEXP_WXAUTOEXCEL wxExcelTrendlines : public wxExcelObject
    {
    public:
        // ***** METHODS *****

        /**
        Creates a new trendline.

        [MSDN documentation for Trendlines.Add](https://learn.microsoft.com/en-us/office/vba/api/excel.trendlines.add).
        */
        wxExcelTrendline Add(XlTrendlineType* type = NULL, long* order = NULL,
                             long* period = NULL, long* forward = NULL, long* backward = NULL,
                             double* intercept = NULL,
                             wxXlTribool displayEquation = wxDefaultXlTribool, wxXlTribool displayRSquared = wxDefaultXlTribool,
                             const wxString& name = wxEmptyString);

        //@{
        /**
        Returns a single object from a collection.

        [MSDN documentation for Trendlines.Item](https://learn.microsoft.com/en-us/office/vba/api/excel.trendlines.item).
        */
        wxExcelTrendline Item(long index);
        wxExcelTrendline operator[](long index);
        //@}

        // ***** PROPERTIES *****


        /**
        Returns a Long value that represents the number of objects in the collection.

        [MSDN documentation for Trendlines.Count](https://learn.microsoft.com/en-us/office/vba/api/excel.trendlines.count).
        */
        long GetCount();

        /**
        Returns "Trendlines".
        */
        virtual wxString GetAutoExcelObjectName_() const { return wxS("Trendlines"); }
    };


} // namespace wxAutoExcel


#endif // #if WXAUTOEXCEL_USE_CHARTS

#endif // _WXAUTOEXCEL_TRENDLINES_H
