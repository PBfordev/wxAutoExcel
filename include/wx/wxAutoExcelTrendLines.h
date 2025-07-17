/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
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

        [MSDN documentation for Trendline.Delete](http://msdn.microsoft.com/en-us/library/bb179066).
        */
        bool Delete();

        /**
        Selects the object.

        [MSDN documentation for Trendline.Select](http://msdn.microsoft.com/en-us/library/bb214087).
        */
        bool Select();

        // ***** PROPERTIES *****


        /**
        Returns the number of periods (or units on a scatter chart) that the trendline extends backward. Read/write Long

        [MSDN documentation for Trendline.Backward](http://msdn.microsoft.com/en-us/library/bb220877).
        */
        long GetBackward();

        /**
        Sets the number of periods (or units on a scatter chart) that the trendline extends backward. Read/write Long

        [MSDN documentation for Trendline.Backward](http://msdn.microsoft.com/en-us/library/bb220877).
        */
        void SetBackward(long backward);

        /**
        Returns the number of periods (or units on a scatter chart) that the trendline extends backward.  Since Excel 2007.

        [MSDN documentation for Trendline.Backward2](http://msdn.microsoft.com/en-us/library/bb215954).
        */
        double GetBackward2();

        /**
        Sets the number of periods (or units on a scatter chart) that the trendline extends backward.  Since Excel 2007.

        [MSDN documentation for Trendline.Backward2](http://msdn.microsoft.com/en-us/library/bb215954).
        */
        void SetBackward2(double backward2);

        /**
        Returns a Border object that represents the border of the object.

        [MSDN documentation for Trendline.Border](http://msdn.microsoft.com/en-us/library/bb214057).
        */
        wxExcelBorder GetBorder();

        /**
        Returns a DataLabel object that represents the data label associated with the trendline.

        [MSDN documentation for Trendline.DataLabel](http://msdn.microsoft.com/en-us/library/bb214069).
        */
        wxExcelDataLabel GetDataLabel();

        /**
        True if the equation for the trendline is displayed on the chart (in the same data label as the R-squared value). Setting this property to True automatically turns on data labels.

        [MSDN documentation for Trendline.DisplayEquation](http://msdn.microsoft.com/en-us/library/bb177499).
        */
        bool GetDisplayEquation();

        /**
        True if the equation for the trendline is displayed on the chart (in the same data label as the R-squared value). Setting this property to True automatically turns on data labels.

        [MSDN documentation for Trendline.DisplayEquation](http://msdn.microsoft.com/en-us/library/bb177499).
        */
        void SetDisplayEquation(bool displayEquation);

        /**
        True if the R-squared value of the trendline is displayed on the chart (in the same data label as the equation). Setting this property to True automatically turns on data labels.

        [MSDN documentation for Trendline.DisplayRSquared](http://msdn.microsoft.com/en-us/library/bb220987).
        */
        bool GetDisplayRSquared();

        /**
        True if the R-squared value of the trendline is displayed on the chart (in the same data label as the equation). Setting this property to True automatically turns on data labels.

        [MSDN documentation for Trendline.DisplayRSquared](http://msdn.microsoft.com/en-us/library/bb220987).
        */
        void SetDisplayRSquared(bool displayRSquared);

        /**
        Read-only Since Excel 2007.

        [MSDN documentation for Trendline.Format](http://msdn.microsoft.com/en-us/library/bb242541).
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

        [MSDN documentation for Trendline.Forward2](http://msdn.microsoft.com/en-us/library/bb215956).
        */
        double GetForward2();

        /**
        Returns the ChartFormat object.  Since Excel 2007.

        [MSDN documentation for Trendline.Forward2](http://msdn.microsoft.com/en-us/library/bb215956).
        */
        void SetForward2(double forward2);

        /**
        Returns a Long value that represents the index number of the object within the collection of similar objects.

        [MSDN documentation for Trendline.Index](http://msdn.microsoft.com/en-us/library/bb214077).
        */
        long GetIndex();

        /**
        Returns the point where the trendline crosses the value axis. Read/write Double.

        [MSDN documentation for Trendline.Intercept](http://msdn.microsoft.com/en-us/library/bb242040).
        */
        double GetIntercept();

        /**
        Sets the point where the trendline crosses the value axis. Read/write Double.

        [MSDN documentation for Trendline.Intercept](http://msdn.microsoft.com/en-us/library/bb242040).
        */
        void SetIntercept(double intercept);

        /**
        True if the point where the trendline crosses the value axis is automatically determined by the regression.

        [MSDN documentation for Trendline.InterceptIsAuto](http://msdn.microsoft.com/en-us/library/bb177672).
        */
        bool GetInterceptIsAuto();

        /**
        True if the point where the trendline crosses the value axis is automatically determined by the regression.

        [MSDN documentation for Trendline.InterceptIsAuto](http://msdn.microsoft.com/en-us/library/bb177672).
        */
        void SetInterceptIsAuto(bool interceptIsAuto);

        /**
        Returns a String value representing the name of the object.

        [MSDN documentation for Trendline.Name](http://msdn.microsoft.com/en-us/library/bb214084).
        */
        wxString GetName();

        /**
        Sets a String value representing the name of the object.

        [MSDN documentation for Trendline.Name](http://msdn.microsoft.com/en-us/library/bb214084).
        */
        void SetName(const wxString& name);

        /**
        True if Microsoft Excel automatically determines the name of the trendline.

        [MSDN documentation for Trendline.NameIsAuto](http://msdn.microsoft.com/en-us/library/bb208815).
        */
        bool GetNameIsAuto();

        /**
        True if Microsoft Excel automatically determines the name of the trendline.

        [MSDN documentation for Trendline.NameIsAuto](http://msdn.microsoft.com/en-us/library/bb208815).
        */
        void SetNameIsAuto(bool nameIsAuto);

        /**
        Returns a Long value that represents the trendline order (an integer greater than 1) when the trendline type is xlPolynomial.

        [MSDN documentation for Trendline.Order](http://msdn.microsoft.com/en-us/library/bb214090).
        */
        long GetOrder();

        /**
        Sets a Long value that represents the trendline order (an integer greater than 1) when the trendline type is xlPolynomial.

        [MSDN documentation for Trendline.Order](http://msdn.microsoft.com/en-us/library/bb214090).
        */
        void SetOrder(long order);

        /**
        Returns the period for the moving-average trendline. Can be a value from 2 through 255.

        [MSDN documentation for Trendline.Period](http://msdn.microsoft.com/en-us/library/bb214096).
        */
        long GetPeriod();

        /**
        Sets the period for the moving-average trendline. Can be a value from 2 through 255.

        [MSDN documentation for Trendline.Period](http://msdn.microsoft.com/en-us/library/bb214096).
        */
        void SetPeriod(long period);

        /**
        Returns a XlTrendlineType value that represents the trendline type.

        [MSDN documentation for Trendline.Type](http://msdn.microsoft.com/en-us/library/bb215161).
        */
        XlTrendlineType GetType();

        /**
        Sets a XlTrendlineType value that represents the trendline type.

        [MSDN documentation for Trendline.Type](http://msdn.microsoft.com/en-us/library/bb215161).
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

        [MSDN documentation for Trendlines.Add](http://msdn.microsoft.com/en-us/library/bb179069).
        */
        wxExcelTrendline Add(XlTrendlineType* type = NULL, long* order = NULL,
                             long* period = NULL, long* forward = NULL, long* backward = NULL,
                             double* intercept = NULL,
                             wxXlTribool displayEquation = wxDefaultXlTribool, wxXlTribool displayRSquared = wxDefaultXlTribool,
                             const wxString& name = wxEmptyString);

        //@{
        /**
        Returns a single object from a collection.

        [MSDN documentation for Trendlines.Item](http://msdn.microsoft.com/en-us/library/bb179070).
        */
        wxExcelTrendline Item(long index);
        wxExcelTrendline operator[](long index);
        //@}

        // ***** PROPERTIES *****


        /**
        Returns a Long value that represents the number of objects in the collection.

        [MSDN documentation for Trendlines.Count](http://msdn.microsoft.com/en-us/library/bb214104).
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
