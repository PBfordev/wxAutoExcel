/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Modified by:
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
/////////////////////////////////////////////////////////////////////////////


#ifndef _WXAUTOEXCEL_ENUMS_H
#define _WXAUTOEXCEL_ENUMS_H

/** @file
Contains MS Excel enumerations.
*/

/**
@brief All wxAutoExcel classes and enumerations are declared in wxAutoExcel namespace.
*/

namespace wxAutoExcel {

/*************************************
    Microsoft Excel enumerations
*************************************/

/**
This enumeration groups together constants used with various Excel methods.

[Official VBA documentation for Constants](https://docs.microsoft.com/office/vba/api/excel.constants)
*/
enum Constants
{
    xl3DBar                 = -4099, /*!< 3D Bar */
    xl3DEffects1            =    13, /*!< 3D Effects1 */
    xl3DEffects2            =    14, /*!< 3D Effects2 */
    xl3DSurface             = -4103, /*!< 3D Surface */
    xlAbove                 =     0, /*!< Above */
    xlAccounting1           =     4, /*!< Accounting1 */
    xlAccounting2           =     5, /*!< Accounting2 */
    xlAccounting4           =    17, /*!< Accounting4 */
    xlAdd                   =     2, /*!< Add */
    xlAll                   = -4104, /*!< All */
    xlAccounting3           =     6, /*!< Accounting3 */
    xlAllExceptBorders      =     7, /*!< All Except Borders */
    xlAutomatic             = -4105, /*!< Automatic */
    xlBar                   =     2, /*!< Automatic */
    xlBelow                 =     1, /*!< Below */
    xlBidi                  = -5000, /*!< Bidi */
    xlBidiCalendar          =     3, /*!< BidiCalendar */
    xlBoth                  =     1, /*!< Both */
    xlBottom                = -4107, /*!< Bottom */
    xlCascade               =     7, /*!< Cascade */
    xlCenter                = -4108, /*!< Center */
    xlCenterAcrossSelection =     7, /*!< Center Across Selection */
    xlChart4                =     2, /*!< Chart 4 */
    xlChartSeries           =    17, /*!< Chart Series */
    xlChartShort            =     6, /*!< Chart Short */
    xlChartTitles           =    18, /*!< Chart Titles */
    xlChecker               =     9, /*!< Checker */
    xlCircle                =     8, /*!< Circle */
    xlClassic1              =     1, /*!< Classic1 */
    xlClassic2              =     2, /*!< Classic2 */
    xlClassic3              =     3, /*!< Classic3 */
    xlClosed                =     3, /*!< Closed */
    xlColor1                =     7, /*!< Color1 */
    xlColor2                =     8, /*!< Color2 */
    xlColor3                =     9, /*!< Color3 */
    xlColumn                =     3, /*!< Column */
    xlCombination           = -4111, /*!< Combination */
    xlComplete              =     4, /*!< Complete */
    xlConstants             =     2, /*!< Constants */
    xlContents              =     2, /*!< Contents */
    xlContext               = -5002, /*!< Context */
    xlCorner                =     2, /*!< Corner */
    xlCrissCross            =    16, /*!< CrissCross */
    xlCross                 =     4, /*!< Cross */
    xlCustom                = -4114, /*!< Custom */
    xlDebugCodePane         =    13, /*!< Debug Code Pane */
    xlDefaultAutoFormat     =    -1, /*!< Default Auto Format */
    xlDesktop               =     9, /*!< Desktop */
    xlDiamond               =     2, /*!< Diamond */
    xlDirect                =     1, /*!< Direct */
    xlDistributed           = -4117, /*!< Distributed */
    xlDivide                =     5, /*!< Divide */
    xlDoubleAccounting      =     5, /*!< Double Accounting */
    xlDoubleClosed          =     5, /*!< Double Closed */
    xlDoubleOpen            =     4, /*!< Double Open */
    xlDoubleQuote           =     1, /*!< Double Quote */
    xlDrawingObject         =    14, /*!< Drawing Object */
    xlEntireChart           =    20, /*!< Entire Chart */
    xlExcelMenus            =     1, /*!< Excel Menus */
    xlExtended              =     3, /*!< Extended */
    xlFill                  =     5, /*!< Fill */
    xlFirst                 =     0, /*!< First */
    xlFixedValue            =     1, /*!< Fixed Value */
    xlFloating              =     5, /*!< Floating */
    xlFormats               = -4122, /*!< Formats */
    xlFormula               =     5, /*!< Formula */
    xlFullScript            =     1, /*!< Full Script */
    xlGeneral               =     1, /*!< General */
    xlGray16                =    17, /*!< Gray16 */
    xlGray25                = -4124, /*!< Gray25 */
    xlGray50                = -4125, /*!< Gray50 */
    xlGray75                = -4126, /*!< Gray75 */
    xlGray8                 =    18, /*!< Gray8 */
    xlGregorian             =     2, /*!< Gregorian */
    xlGrid                  =    15, /*!< Grid */
    xlGridline              =    22, /*!< Gridline */
    xlHigh                  = -4127, /*!< High */
    xlHindiNumerals         =     3, /*!< Hindi Numerals */
    xlIcons                 =     1, /*!< Icons */
    xlImmediatePane         =    12, /*!< Immediate Pane */
    xlInside                =     2, /*!< Inside */
    xlInteger               =     2, /*!< Integer */
    xlJustify               = -4130, /*!< Justify */
    xlLast                  =     1, /*!< Last */
    xlLastCell              =    11, /*!< Last Cell */
    xlLatin                 = -5001, /*!< Latin */
    xlLeft                  = -4131, /*!< Left */
    xlLeftToRight           =     2, /*!< Left To Right */
    xlLightDown             =    13, /*!< Light Down */
    xlLightHorizontal       =    11, /*!< Light Horizontal */
    xlLightUp               =    14, /*!< Light Up */
    xlLightVertical         =    12, /*!< Light Vertical */
    xlList1                 =    10, /*!< List1 */
    xlList2                 =    11, /*!< List2 */
    xlList3                 =    12, /*!< List3 */
    xlLocalFormat1          =    15, /*!< Local Format1 */
    xlLocalFormat2          =    16, /*!< Local Format2 */
    xlLogicalCursor         =     1, /*!< Logical Cursor */
    xlLong                  =     3, /*!< Long */
    xlLotusHelp             =     2, /*!< Lotus Help */
    xlLow                   = -4134, /*!< Low */
    xlLTR                   = -5003, /*!< LTR */
    xlMacrosheetCell        =     7, /*!< MacrosheetCell */
// xlManual is commented out here to avoid conflict with XlSortOrder::xlManual
//    xlManual                = -4135, /*!< Manual */
    xlMaximum               =     2, /*!< Maximum */
    xlMinimum               =     4, /*!< Minimum */
    xlMinusValues           =     3, /*!< Minus Values */
    xlMixed                 =     2, /*!< Mixed */
    xlMixedAuthorizedScript =     4, /*!< Mixed Authorized Script */
    xlMixedScript           =     3, /*!< Mixed Script */
    xlModule                = -4141, /*!< Module */
    xlMultiply              =     4, /*!< Multiply */
    xlNarrow                =     1, /*!< Narrow */
    xlNextToAxis            =     4, /*!< Next To Axis */
    xlNoDocuments           =     3, /*!< No Documents */
    xlNone                  = -4142, /*!< None */
    xlNotes                 = -4144, /*!< Notes */
    xlOff                   = -4146, /*!< Off */
    xlOn                    =     1, /*!< On */
    xlOpaque                =     3, /*!< Opaque */
    xlOpen                  =     2, /*!< Open */
    xlOutside               =     3, /*!< Outside */
    xlPartial               =     3, /*!< Partial */
    xlPartialScript         =     2, /*!< Partial Script */
    xlPercent               =     2, /*!< Percent */
    xlPlus                  =     9, /*!< Plus */
    xlPlusValues            =     2, /*!< Plus Values */
    xlReference             =     4, /*!< Reference */
    xlRight                 = -4152, /*!< Right */
    xlRTL                   = -5004, /*!< RTL */
    xlScale                 =     3, /*!< Scale */
    xlSemiautomatic         =     2, /*!< Semiautomatic */
    xlSemiGray75            =    10, /*!< SemiGray75 */
    xlShort                 =     1, /*!< Short */
    xlShowLabel             =     4, /*!< Show Label */
    xlShowLabelAndPercent   =     5, /*!< Show Label and Percent */
    xlShowPercent           =     3, /*!< Show Percent */
    xlShowValue             =     2, /*!< Show Value */
    xlSimple                = -4154, /*!< Simple */
    xlSingle                =     2, /*!< Single */
    xlSingleAccounting      =     4, /*!< Single Accounting */
    xlSingleQuote           =     2, /*!< Single Quote */
    xlSolid                 =     1, /*!< Solid */
    xlSquare                =     1, /*!< Square */
    xlStar                  =     5, /*!< Star */
    xlStError               =     4, /*!< St Error */
    xlStrict                =     2, /*!< Strict */
    xlSubtract              =     3, /*!< Subtract */
    xlSystem                =     1, /*!< System */
    xlTextBox               =    16, /*!< Text Box */
    xlTiled                 =     1, /*!< Tiled */
    xlTitleBar              =     8, /*!< Title Bar */
    xlToolbar               =     1, /*!< Toolbar */
    xlToolbarButton         =     2, /*!< Toolbar Button */
    xlTop                   = -4160, /*!< Top */
    xlTopToBottom           =     1, /*!< Top To Bottom */
    xlTransparent           =     2, /*!< Transparent */
    xlTriangle              =     3, /*!< Triangle */
    xlVeryHidden            =     2, /*!< Very Hidden */
    xlVisible               =    12, /*!< Visible */
    xlVisualCursor          =     2, /*!< Visual Cursor */
    xlWatchPane             =    11, /*!< Watch Pane */
    xlWide                  =     3, /*!< Wide */
    xlWorkbookTab           =     6, /*!< Workbook Tab */
    xlWorksheet4            =     1, /*!< Worksheet4 */
    xlWorksheetCell         =     3, /*!< Worksheet Cell */
    xlWorksheetShort        =     5, /*!< Worksheet Short */
};

/**
Specifies if the values are above or below average.

[Official VBA documentation for XlAboveBelow](https://docs.microsoft.com/office/vba/api/excel.xlabovebelow)
*/
enum XlAboveBelow
{
    xlAboveAverage      = 0, /*!< Above average */
    xlAboveStdDev       = 4, /*!< Above standard deviation */
    xlBelowAverage      = 1, /*!< Below average */
    xlBelowStdDev       = 5, /*!< Below standard deviation */
    xlEqualAboveAverage = 2, /*!< Equal above average */
    xlEqualBelowAverage = 3, /*!< Equal below average */
};

/**
Specifies the action that should be performed.

[Official VBA documentation for XlActionType](https://docs.microsoft.com/office/vba/api/excel.xlactiontype)
*/
enum XlActionType
{
    xlActionTypeDrillthrough = 256, /*!< Drill through */
    xlActionTypeReport       = 128, /*!< Report */
    xlActionTypeRowset       =  16, /*!< Rowset */
    xlActionTypeUrl          =   1, /*!< URL */
};

/**
Specifies when to calculate changes when performing what-if analysis on a PivotTable based on an OLAP data source.

[Official VBA documentation for XlAllocation](https://docs.microsoft.com/office/vba/api/excel.xlallocation)
*/
enum XlAllocation
{
    xlAutomaticAllocation = 2, /*!< Calculate changes automatically after each value is changed. */
    xlManualAllocation    = 1, /*!< Calculate changes manually. */
};

/**
Specifies the method to use to allocate values when performing what-if analysis on a PivotTable report based on an OLAP data source.

[Official VBA documentation for XlAllocationMethod](https://docs.microsoft.com/office/vba/api/excel.xlallocationmethod)
*/
enum XlAllocationMethod
{
    xlEqualAllocation    = 1, /*!< Use equal allocation. */
    xlWeightedAllocation = 2, /*!< Use weighted allocation. */
};

/**
Specifies what value to allocate when performing what-if analysis on a PivotTable report based on an OLAP data source.

[Official VBA documentation for XlAllocationValue](https://docs.microsoft.com/office/vba/api/excel.xlallocationvalue)
*/
enum XlAllocationValue
{
    xlAllocateIncrement = 2, /*!< Increment based on the old value. */
    xlAllocateValue     = 1, /*!< The value entered divided by the number of allocations. */
};

/**
Specifies country/region and international settings.

[Official VBA documentation for XlApplicationInternational](https://docs.microsoft.com/office/vba/api/excel.xlapplicationinternational)
*/
enum XlApplicationInternational
{
    xl24HourClock             = 33, /*!< **True** if you are using 24-hour time; False if you are using 12-hour time. */
    xl4DigitYears             = 43, /*!< **True** if you are using four-digit years; False if you are using two-digit years. */
    xlAlternateArraySeparator = 16, /*!< Alternate array item separator to be used if the current array separator is the same as the decimal separator. */
    xlColumnSeparator         = 14, /*!< Character used to separate columns in array literals. */
    xlCountryCode             =  1, /*!< Country/Region version of Microsoft Excel. */
    xlCountrySetting          =  2, /*!< Current country/region setting in the Windows Control Panel. */
    xlCurrencyBefore          = 37, /*!< **True** if the currency symbol precedes the currency values; False if it follows them. */
    xlCurrencyCode            = 25, /*!< Currency symbol. */
    xlCurrencyDigits          = 27, /*!< Number of decimal digits to be used in currency formats. */
    xlCurrencyLeadingZeros    = 40, /*!< **True** if leading zeros are displayed for zero currency values. */
    xlCurrencyMinusSign       = 38, /*!< **True** if you are using a minus sign for negative numbers; False if you are using parentheses. */
    xlCurrencyNegative        = 28, /*!< Currency format for negative currency values:<br/>`0 = (symbolx) or (xsymbol)`, `1 = -symbolx or -xsymbol`, `2 = symbol-x or x-symbol`, or `3 = symbolx- or xsymbol-`, where symbol is the currency symbol of the country or region.<br/><br/>Note that the position of the currency symbol is determined by **xlCurrencyBefore**. */
    xlCurrencySpaceBefore     = 36, /*!< **True** if a space is added before the currency symbol. */
    xlCurrencyTrailingZeros   = 39, /*!< **True** if trailing zeros are displayed for zero currency values. */
    xlDateOrder               = 32, /*!< Order of date elements: `0 = month-day-year`, `1 = day-month-year`, `2 = year-month-day` */
    xlDateSeparator           = 17, /*!< Date separator (/). */
    xlDayCode                 = 21, /*!< Day symbol (d). */
    xlDayLeadingZero          = 42, /*!< **True** if a leading zero is displayed in days. */
    xlDecimalSeparator        =  3, /*!< Decimal separator. */
    xlGeneralFormatName       = 26, /*!< Name of the General number format. */
    xlHourCode                = 22, /*!< Hour symbol (h). */
    xlLeftBrace               = 12, /*!< Character used instead of the left brace ({) in array literals. */
    xlLeftBracket             = 10, /*!< Character used instead of the left bracket ([) in R1C1-style relative references. */
    xlListSeparator           =  5, /*!< List separator. */
    xlLowerCaseColumnLetter   =  9, /*!< Lowercase column letter. */
    xlLowerCaseRowLetter      =  8, /*!< Lowercase row letter. */
    xlMDY                     = 44, /*!< **True** if the date order is month-day-year for dates displayed in the long form; False if the date order is day-month-year. */
    xlMetric                  = 35, /*!< **True** if you are using the metric system; False if you are using the English measurement system. */
    xlMinuteCode              = 23, /*!< Minute symbol (m). */
    xlMonthCode               = 20, /*!< Month symbol (m). */
    xlMonthLeadingZero        = 41, /*!< **True** if a leading zero is displayed in months (when months are displayed as numbers). */
    xlMonthNameChars          = 30, /*!< Always returns three characters for backward compatibility. Abbreviated month names are read from Microsoft Windows and can be any length. */
    xlNoncurrencyDigits       = 29, /*!< Number of decimal digits to be used in noncurrency formats. */
    xlNonEnglishFunctions     = 34, /*!< **True** if you are not displaying functions in English. */
    xlRightBrace              = 13, /*!< Character used instead of the right brace (}) in array literals. */
    xlRightBracket            = 11, /*!< Character used instead of the right bracket (]) in R1C1-style references. */
    xlRowSeparator            = 15, /*!< Character used to separate rows in array literals. */
    xlSecondCode              = 24, /*!< Second symbol (s). */
    xlThousandsSeparator      =  4, /*!< Zero or thousands separator. */
    xlTimeLeadingZero         = 45, /*!< **True** if a leading zero is displayed in times. */
    xlTimeSeparator           = 18, /*!< Time separator (:). */
    xlUpperCaseColumnLetter   =  7, /*!< Uppercase column letter. */
    xlUpperCaseRowLetter      =  6, /*!< Uppercase row letter (for R1C1-style references). */
    xlWeekdayNameChars        = 31, /*!< Always returns three characters for backward compatibility. Abbreviated weekday names are read from Microsoft Windows and can be any length. */
    xlYearCode                = 19, /*!< Year symbol in number formats (y). */
};

/**
Specifies which range name is listed first when a cell reference is replaced by a row-oriented and column-oriented range name.

[Official VBA documentation for XlApplyNamesOrder](https://docs.microsoft.com/office/vba/api/excel.xlapplynamesorder)
*/
enum XlApplyNamesOrder
{
    xlColumnThenRow = 2, /*!< Columns listed before rows */
    xlRowThenColumn = 1, /*!< Rows listed before columns */
};

/**
Specifies spelling rules for the Arabic spelling checker.

[Official VBA documentation for XlArabicModes](https://docs.microsoft.com/office/vba/api/excel.xlarabicmodes)
*/
enum XlArabicModes
{
    xlArabicBothStrict      = 3, /*!< The spelling checker uses spelling rules regarding both Arabic words ending with the letter yaa and Arabic words beginning with an alef hamza. */
    xlArabicNone            = 0, /*!< The spelling checker ignores spelling rules regarding either Arabic words ending with the letter yaa or Arabic words beginning with an alef hamza. */
    xlArabicStrictAlefHamza = 1, /*!< The spelling checker uses spelling rules regarding Arabic words beginning with an alef hamza. */
    xlArabicStrictFinalYaa  = 2, /*!< The spelling checker uses spelling rules regarding Arabic words ending with the letter yaa. */
};

/**
Specifies how windows are arranged on the screen.

[Official VBA documentation for XlArrangeStyle](https://docs.microsoft.com/office/vba/api/excel.xlarrangestyle)
*/
enum XlArrangeStyle
{
    xlArrangeStyleCascade    =     7, /*!< Windows are cascaded. */
    xlArrangeStyleHorizontal = -4128, /*!< Windows are arranged horizontally. */
    xlArrangeStyleTiled      =     1, /*!< Default. Windows are tiled. */
    xlArrangeStyleVertical   = -4166, /*!< Windows are arranged vertically. */
};

/**
Specifies the length of the arrowhead at the end of a line.

[Official VBA documentation for XlArrowHeadLength](https://docs.microsoft.com/office/vba/api/excel.xlarrowheadlength)
*/
enum XlArrowHeadLength
{
    xlArrowHeadLengthLong   =     3, /*!< Longest arrowhead */
    xlArrowHeadLengthMedium = -4138, /*!< Medium-length arrowhead */
    xlArrowHeadLengthShort  =     1, /*!< Shortest arrowhead */
};

/**
Specifies the type of arrowhead to apply at the end of a line.

[Official VBA documentation for XlArrowHeadStyle](https://docs.microsoft.com/office/vba/api/excel.xlarrowheadstyle)
*/
enum XlArrowHeadStyle
{
    xlArrowHeadStyleClosed       =     3, /*!< Small arrowhead with curved edge at connection to line. */
    xlArrowHeadStyleDoubleClosed =     5, /*!< Large diamond-shaped arrowhead. */
    xlArrowHeadStyleDoubleOpen   =     4, /*!< Large arrowhead with curved edge at connection to line. */
    xlArrowHeadStyleNone         = -4142, /*!< No arrowhead. */
    xlArrowHeadStyleOpen         =     2, /*!< Large triangular arrowhead. */
};

/**
Specifies the width of the arrowhead at the end of a line.

[Official VBA documentation for XlArrowHeadWidth](https://docs.microsoft.com/office/vba/api/excel.xlarrowheadwidth)
*/
enum XlArrowHeadWidth
{
    xlArrowHeadWidthMedium = -4138, /*!< Medium-width arrowhead */
    xlArrowHeadWidthNarrow =     1, /*!< Narrowest arrowhead */
    xlArrowHeadWidthWide   =     3, /*!< Widest arrowhead */
};

/**
Specifies how the target range is to be filled, based on the contents of the source range.

[Official VBA documentation for XlAutoFillType](https://docs.microsoft.com/office/vba/api/excel.xlautofilltype)
*/
enum XlAutoFillType
{
    xlFillCopy     =  1, /*!< Copy the values and formats from the source range to the target range, repeating if necessary. */
    xlFillDays     =  5, /*!< Extend the names of the days of the week in the source range into the target range. Formats are copied from the source range to the target range, repeating if necessary. */
    xlFillDefault  =  0, /*!< Excel determines the values and formats used to fill the target range. */
    xlFillFormats  =  3, /*!< Copy only the formats from the source range to the target range, repeating if necessary. */
    xlFillMonths   =  7, /*!< Extend the names of the months in the source range into the target range. Formats are copied from the source range to the target range, repeating if necessary. */
    xlFillSeries   =  2, /*!< Extend the values in the source range into the target range as a series (for example, '1, 2' is extended as '3, 4, 5'). Formats are copied from the source range to the target range, repeating if necessary. */
    xlFillValues   =  4, /*!< Copy only the values from the source range to the target range, repeating if necessary. */
    xlFillWeekdays =  6, /*!< Extend the names of the days of the workweek in the source range into the target range. Formats are copied from the source range to the target range, repeating if necessary. */
    xlFillYears    =  8, /*!< Extend the years in the source range into the target range. Formats are copied from the source range to the target range, repeating if necessary. */
    xlGrowthTrend  = 10, /*!< Extend the numeric values from the source range into the target range, assuming that the relationships between the numbers in the source range are multiplicative (for example, '1, 2,' is extended as '4, 8, 16', assuming that each number is a result of multiplying the previous number by some value). Formats are copied from the source range to the target range, repeating if necessary. */
    xlLinearTrend  =  9, /*!< Extend the numeric values from the source range into the target range, assuming that the relationships between the numbers is additive (for example, '1, 2,' is extended as '3, 4, 5', assuming that each number is a result of adding some value to the previous number). Formats are copied from the source range to the target range, repeating if necessary. */
    xlFlashFill    = 11, /*!< Extend the values from the source range into the target range based on the detected pattern of previous user actions, repeating if necessary. */
};

/**
Specifies the operator to use to associate two criteria applied by a filter.

[Official VBA documentation for XlAutoFilterOperator](https://docs.microsoft.com/office/vba/api/excel.xlautofilteroperator)
*/
enum XlAutoFilterOperator
{
    xlAnd             =  1, /*!< Logical AND of Criteria1 and Criteria2 */
    xlBottom10Items   =  4, /*!< Lowest-valued items displayed (number of items specified in Criteria1) */
    xlBottom10Percent =  6, /*!< Lowest-valued items displayed (percentage specified in Criteria1) */
    xlFilterCellColor =  8, /*!< Color of the cell */
    xlFilterDynamic   = 11, /*!< Dynamic filter */
    xlFilterFontColor =  9, /*!< Color of the font */
    xlFilterIcon      = 10, /*!< Filter icon */
    xlFilterValues    =  7, /*!< Filter values */
    xlOr              =  2, /*!< Logical OR of Criteria1 or Criteria2 */
    xlTop10Items      =  3, /*!< Highest-valued items displayed (number of items specified in Criteria1) */
    xlTop10Percent    =  5, /*!< Highest-valued items displayed (percentage specified in Criteria1) */
};

/**
Specifies the point on the specified axis where the other axis crosses.

[Official VBA documentation for XlAxisCrosses](https://docs.microsoft.com/office/vba/api/excel.xlaxiscrosses)
*/
enum XlAxisCrosses
{
    xlAxisCrossesAutomatic = -4105, /*!< Microsoft Excel sets the axis crossing point. */
    xlAxisCrossesCustom    = -4114, /*!< The **[CrossesAt](Excel.CrossesAt.md)** property specifies the axis crossing point. */
    xlAxisCrossesMaximum   =     2, /*!< The axis crosses at the maximum value. */
    xlAxisCrossesMinimum   =     4, /*!< The axis crosses at the minimum value. */
};

/**
Specifies the type of axis group.

[Official VBA documentation for XlAxisGroup](https://docs.microsoft.com/office/vba/api/excel.xlaxisgroup)
*/
enum XlAxisGroup
{
    xlPrimary   = 1, /*!< Primary axis group */
    xlSecondary = 2, /*!< Secondary axis group */
};

/**
Specifies the axis type.

[Official VBA documentation for XlAxisType](https://docs.microsoft.com/office/vba/api/excel.xlaxistype)
*/
enum XlAxisType
{
    xlCategory   = 1, /*!< Axis displays categories. */
    xlSeriesAxis = 3, /*!< Axis displays data series. */
    xlValue      = 2, /*!< Axis displays values. */
};

/**
Specifies the background type for text in charts.

[Official VBA documentation for XlBackground](https://docs.microsoft.com/office/vba/api/excel.xlbackground)
*/
enum XlBackground
{
    xlBackgroundAutomatic   = -4105, /*!< Excel controls the background. */
    xlBackgroundOpaque      =     3, /*!< Opaque background. */
    xlBackgroundTransparent =     2, /*!< Transparent background. */
};

/**
Specifies the shape used with the 3D bar or column chart.

[Official VBA documentation for XlBarShape](https://docs.microsoft.com/office/vba/api/excel.xlbarshape)
*/
enum XlBarShape
{
    xlBox            = 0, /*!< Box. */
    xlConeToMax      = 5, /*!< Cone, truncated at value. */
    xlConeToPoint    = 4, /*!< Cone, coming to point at value. */
    xlCylinder       = 3, /*!< Cylinder. */
    xlPyramidToMax   = 2, /*!< Pyramid, truncated at value. */
    xlPyramidToPoint = 1, /*!< Pyramid, coming to point at value. */
};

/**
Constants passed to and returned by the [ChartGroup.BinsType](Excel.chartgroup.binstype.md) property.

[Official VBA documentation for XlBinsType](https://docs.microsoft.com/office/vba/api/excel.xlbinstype)
*/
enum XlBinsType
{
    xlBinsTypeAutomatic   = 0, /*!< Sets bins type automatically. */
    xlBinsTypeCategorical = 1, /*!< Sets bins type by category. */
    xlBinsTypeManual      = 2, /*!< Sets bins type manually. */
    xlBinsTypeBinSize     = 3, /*!< Sets bins type by size. */
    xlBinsTypeBinCount    = 4, /*!< Sets bins type by count. */
};

/**
Specifies the weight of the border around a range.

[Official VBA documentation for XlBorderWeight](https://docs.microsoft.com/office/vba/api/excel.xlborderweight)
*/
enum XlBorderWeight
{
    xlHairline =     1, /*!< Hairline (thinnest border). */
    xlMedium   = -4138, /*!< Medium. */
    xlThick    =     4, /*!< Thick (widest border). */
    xlThin     =     2, /*!< Thin. */
};

/**
Specifies the border to be retrieved.

[Official VBA documentation for XlBordersIndex](https://docs.microsoft.com/office/vba/api/excel.xlbordersindex)
*/
enum XlBordersIndex
{
    xlDiagonalDown     =  5, /*!< Border running from the upper-left corner to the lower-right of each cell in the range. */
    xlDiagonalUp       =  6, /*!< Border running from the lower-left corner to the upper-right of each cell in the range. */
    xlEdgeBottom       =  9, /*!< Border at the bottom of the range. */
    xlEdgeLeft         =  7, /*!< Border at the left edge of the range. */
    xlEdgeRight        = 10, /*!< Border at the right edge of the range. */
    xlEdgeTop          =  8, /*!< Border at the top of the range. */
    xlInsideHorizontal = 12, /*!< Horizontal borders for all cells in the range except borders on the outside of the range. */
    xlInsideVertical   = 11, /*!< Vertical borders for all the cells in the range except borders on the outside of the range. */
};

/**
Specifies which dialog box to display.

[Official VBA documentation for XlBuiltInDialog](https://docs.microsoft.com/office/vba/api/excel.xlbuiltindialog)
*/
enum XlBuiltInDialog
{
    xlDialogActivate                         =  103, /*!< **Activate** dialog box */
    xlDialogActiveCellFont                   =  476, /*!< **Active Cell Font** dialog box */
    xlDialogAddChartAutoformat               =  390, /*!< **Add Chart Autoformat** dialog box */
    xlDialogAddinManager                     =  321, /*!< **Addin Manager** dialog box */
    xlDialogAlignment                        =   43, /*!< **Alignment** dialog box */
    xlDialogApplyNames                       =  133, /*!< **Apply Names** dialog box */
    xlDialogApplyStyle                       =  212, /*!< **Apply Style** dialog box */
    xlDialogAppMove                          =  170, /*!< **AppMove** dialog box */
    xlDialogAppSize                          =  171, /*!< **AppSize** dialog box */
    xlDialogArrangeAll                       =   12, /*!< **Arrange All** dialog box */
    xlDialogAssignToObject                   =  213, /*!< **Assign To Object** dialog box */
    xlDialogAssignToTool                     =  293, /*!< **Assign To Tool** dialog box */
    xlDialogAttachText                       =   80, /*!< **Attach Text** dialog box */
    xlDialogAttachToolbars                   =  323, /*!< **Attach Toolbars** dialog box */
    xlDialogAutoCorrect                      =  485, /*!< **Auto Correct** dialog box */
    xlDialogAxes                             =   78, /*!< **Axes** dialog box */
    xlDialogBorder                           =   45, /*!< **Border** dialog box */
    xlDialogCalculation                      =   32, /*!< **Calculation** dialog box */
    xlDialogCellProtection                   =   46, /*!< **Cell Protection** dialog box */
    xlDialogChangeLink                       =  166, /*!< **Change Link** dialog box */
    xlDialogChartAddData                     =  392, /*!< **Chart Add Data** dialog box */
    xlDialogChartLocation                    =  527, /*!< **Chart Location** dialog box */
    xlDialogChartOptionsDataLabelMultiple    =  724, /*!< **Chart Options DataLabel Multiple** dialog box */
    xlDialogChartOptionsDataLabels           =  505, /*!< **Chart Options DataLabels** dialog box */
    xlDialogChartOptionsDataTable            =  506, /*!< **Chart Options DataTable** dialog box */
    xlDialogChartSourceData                  =  540, /*!< **Chart SourceData** dialog box */
    xlDialogChartTrend                       =  350, /*!< **Chart Trend** dialog box */
    xlDialogChartType                        =  526, /*!< **Chart Type** dialog box */
    xlDialogChartWizard                      =  288, /*!< **ChartWizard** dialog box */
    xlDialogCheckboxProperties               =  435, /*!< **Checkbox Properties** dialog box */
    xlDialogClear                            =   52, /*!< **Clear** dialog box */
    xlDialogColorPalette                     =  161, /*!< **Color Palette** dialog box */
    xlDialogColumnWidth                      =   47, /*!< **Column Width** dialog box */
    xlDialogCombination                      =   73, /*!< **Combination** dialog box */
    xlDialogConditionalFormatting            =  583, /*!< **Conditional Formatting** dialog box */
    xlDialogConsolidate                      =  191, /*!< **Consolidate** dialog box */
    xlDialogCopyChart                        =  147, /*!< **Copy Chart** dialog box */
    xlDialogCopyPicture                      =  108, /*!< **Copy Picture** dialog box */
    xlDialogCreateList                       =  796, /*!< **Create List** dialog box */
    xlDialogCreateNames                      =   62, /*!< **Create Names** dialog box */
    xlDialogCreatePublisher                  =  217, /*!< **Create Publisher** dialog box */
    xlDialogCreateRelationship               = 1272, /*!< **Create Relationship** dialog box */
    xlDialogCustomizeToolbar                 =  276, /*!< **Customize Toolbar** dialog box */
    xlDialogCustomViews                      =  493, /*!< **Custom Views** dialog box */
    xlDialogDataDelete                       =   36, /*!< **Data Delete** dialog box */
    xlDialogDataLabel                        =  379, /*!< **Data Label** dialog box */
    xlDialogDataLabelMultiple                =  723, /*!< **Data Label Multiple** dialog box */
    xlDialogDataSeries                       =   40, /*!< **Data Series** dialog box */
    xlDialogDataValidation                   =  525, /*!< **Data Validation** dialog box */
    xlDialogDefineName                       =   61, /*!< **Define Name** dialog box */
    xlDialogDefineStyle                      =  229, /*!< **Define Style** dialog box */
    xlDialogDeleteFormat                     =  111, /*!< **Delete Format** dialog box */
    xlDialogDeleteName                       =  110, /*!< **Delete Name** dialog box */
    xlDialogDemote                           =  203, /*!< **Demote** dialog box */
    xlDialogDisplay                          =   27, /*!< **Display** dialog box */
    xlDialogDocumentInspector                =  862, /*!< **Document Inspector** dialog box */
    xlDialogEditboxProperties                =  438, /*!< **Editbox Properties** dialog box */
    xlDialogEditColor                        =  223, /*!< **Edit Color** dialog box */
    xlDialogEditDelete                       =   54, /*!< **Edit Delete** dialog box */
    xlDialogEditionOptions                   =  251, /*!< **Edition Options** dialog box */
    xlDialogEditSeries                       =  228, /*!< **Edit Series** dialog box */
    xlDialogErrorbarX                        =  463, /*!< **Errorbar X** dialog box */
    xlDialogErrorbarY                        =  464, /*!< **Errorbar Y** dialog box */
    xlDialogErrorChecking                    =  732, /*!< **Error Checking** dialog box */
    xlDialogEvaluateFormula                  =  709, /*!< **Evaluate Formula** dialog box */
    xlDialogExternalDataProperties           =  530, /*!< **External Data Properties** dialog box */
    xlDialogExtract                          =   35, /*!< **Extract** dialog box */
    xlDialogFileDelete                       =    6, /*!< **File Delete** dialog box */
    xlDialogFileSharing                      =  481, /*!< **File Sharing** dialog box */
    xlDialogFillGroup                        =  200, /*!< **Fill Group** dialog box */
    xlDialogFillWorkgroup                    =  301, /*!< **Fill Workgroup** dialog box */
    xlDialogFilter                           =  447, /*!< **Dialog Filter** dialog box */
    xlDialogFilterAdvanced                   =  370, /*!< **Filter Advanced** dialog box */
    xlDialogFindFile                         =  475, /*!< **Find File** dialog box */
    xlDialogFont                             =   26, /*!< **Font** dialog box */
    xlDialogFontProperties                   =  381, /*!< **Font Properties** dialog box */
    xlDialogFormatAuto                       =  269, /*!< **Format Auto** dialog box */
    xlDialogFormatChart                      =  465, /*!< **Format Chart** dialog box */
    xlDialogFormatCharttype                  =  423, /*!< **Format Charttype** dialog box */
    xlDialogFormatFont                       =  150, /*!< **Format Font** dialog box */
    xlDialogFormatLegend                     =   88, /*!< **Format Legend** dialog box */
    xlDialogFormatMain                       =  225, /*!< **Format Main** dialog box */
    xlDialogFormatMove                       =  128, /*!< **Format Move** dialog box */
    xlDialogFormatNumber                     =   42, /*!< **Format Number** dialog box */
    xlDialogFormatOverlay                    =  226, /*!< **Format Overlay** dialog box */
    xlDialogFormatSize                       =  129, /*!< **Format Size** dialog box */
    xlDialogFormatText                       =   89, /*!< **Format Text** dialog box */
    xlDialogFormulaFind                      =   64, /*!< **Formula Find** dialog box */
    xlDialogFormulaGoto                      =   63, /*!< **Formula Goto** dialog box */
    xlDialogFormulaReplace                   =  130, /*!< **Formula Replace** dialog box */
    xlDialogFunctionWizard                   =  450, /*!< **Function Wizard** dialog box */
    xlDialogGallery3dArea                    =  193, /*!< **Gallery 3D Area** dialog box */
    xlDialogGallery3dBar                     =  272, /*!< **Gallery 3D Bar** dialog box */
    xlDialogGallery3dColumn                  =  194, /*!< **Gallery 3D Column** dialog box */
    xlDialogGallery3dLine                    =  195, /*!< **Gallery 3D Line** dialog box */
    xlDialogGallery3dPie                     =  196, /*!< **Gallery 3D Pie** dialog box */
    xlDialogGallery3dSurface                 =  273, /*!< **Gallery 3D Surface** dialog box */
    xlDialogGalleryArea                      =   67, /*!< **Gallery Area** dialog box */
    xlDialogGalleryBar                       =   68, /*!< **Gallery Bar** dialog box */
    xlDialogGalleryColumn                    =   69, /*!< **Gallery Column** dialog box */
    xlDialogGalleryCustom                    =  388, /*!< **Gallery Custom** dialog box */
    xlDialogGalleryDoughnut                  =  344, /*!< **Gallery Doughnut** dialog box */
    xlDialogGalleryLine                      =   70, /*!< **Gallery Line** dialog box */
    xlDialogGalleryPie                       =   71, /*!< **Gallery Pie** dialog box */
    xlDialogGalleryRadar                     =  249, /*!< **Gallery Radar** dialog box */
    xlDialogGalleryScatter                   =   72, /*!< **Gallery Scatter** dialog box */
    xlDialogGoalSeek                         =  198, /*!< **Goal Seek** dialog box */
    xlDialogGridlines                        =   76, /*!< **Gridlines** dialog box */
    xlDialogImportTextFile                   =  666, /*!< **Import Text File** dialog box */
    xlDialogInsert                           =   55, /*!< **Insert** dialog box */
    xlDialogInsertHyperlink                  =  596, /*!< **Insert Hyperlink** dialog box */
    xlDialogInsertObject                     =  259, /*!< **Insert Object** dialog box */
    xlDialogInsertPicture                    =  342, /*!< **Insert Picture** dialog box */
    xlDialogInsertTitle                      =  380, /*!< **Insert Title** dialog box */
    xlDialogLabelProperties                  =  436, /*!< **Label Properties** dialog box */
    xlDialogListboxProperties                =  437, /*!< **Listbox Properties** dialog box */
    xlDialogMacroOptions                     =  382, /*!< **Macro Options** dialog box */
    xlDialogMailEditMailer                   =  470, /*!< **Mail Edit Mailer** dialog box */
    xlDialogMailLogon                        =  339, /*!< **Mail Logon** dialog box */
    xlDialogMailNextLetter                   =  378, /*!< **Mail Next Letter** dialog box */
    xlDialogMainChart                        =   85, /*!< **Main Chart** dialog box */
    xlDialogMainChartType                    =  185, /*!< **Main Chart Type** dialog box */
    xlDialogManageRelationships              = 1271, /*!< **Manage Relationships** dialog box */
    xlDialogMenuEditor                       =  322, /*!< **Menu Editor** dialog box */
    xlDialogMove                             =  262, /*!< **Move** dialog box */
    xlDialogMyPermission                     =  834, /*!< **My Permission** dialog box */
    xlDialogNameManager                      =  977, /*!< **NameManager** dialog box */
    xlDialogNew                              =  119, /*!< **New** dialog box */
    xlDialogNewName                          =  978, /*!< **NewName** dialog box */
    xlDialogNewWebQuery                      =  667, /*!< **New Web Query** dialog box */
    xlDialogNote                             =  154, /*!< **Note** dialog box */
    xlDialogObjectProperties                 =  207, /*!< **Object Properties** dialog box */
    xlDialogObjectProtection                 =  214, /*!< **Object Protection** dialog box */
    xlDialogOpen                             =    1, /*!< **Open** dialog box */
    xlDialogOpenLinks                        =    2, /*!< **Open Links** dialog box */
    xlDialogOpenMail                         =  188, /*!< **Open Mail** dialog box */
    xlDialogOpenText                         =  441, /*!< **Open Text** dialog box */
    xlDialogOptionsCalculation               =  318, /*!< **Options Calculation** dialog box */
    xlDialogOptionsChart                     =  325, /*!< **Options Chart** dialog box */
    xlDialogOptionsEdit                      =  319, /*!< **Options Edit** dialog box */
    xlDialogOptionsGeneral                   =  356, /*!< **Options General** dialog box */
    xlDialogOptionsListsAdd                  =  458, /*!< **Options Lists Add** dialog box */
    xlDialogOptionsME                        =  647, /*!< **OptionsME** dialog box */
    xlDialogOptionsTransition                =  355, /*!< **Options Transition** dialog box */
    xlDialogOptionsView                      =  320, /*!< **Options View** dialog box */
    xlDialogOutline                          =  142, /*!< **Outline** dialog box */
    xlDialogOverlay                          =   86, /*!< **Overlay** dialog box */
    xlDialogOverlayChartType                 =  186, /*!< **Overlay ChartType** dialog box */
    xlDialogPageSetup                        =    7, /*!< **Page Setup** dialog box */
    xlDialogParse                            =   91, /*!< **Parse** dialog box */
    xlDialogPasteNames                       =   58, /*!< **Paste Names** dialog box */
    xlDialogPasteSpecial                     =   53, /*!< **Paste Special** dialog box */
    xlDialogPatterns                         =   84, /*!< **Patterns** dialog box */
    xlDialogPermission                       =  832, /*!< **Permission** dialog box */
    xlDialogPhonetic                         =  656, /*!< **Phonetic** dialog box */
    xlDialogPivotCalculatedField             =  570, /*!< **Pivot Calculated Field** dialog box */
    xlDialogPivotCalculatedItem              =  572, /*!< **Pivot Calculated Item** dialog box */
    xlDialogPivotClientServerSet             =  689, /*!< **Pivot Client Server Set** dialog box */
    xlDialogPivotFieldGroup                  =  433, /*!< **Pivot Field Group** dialog box */
    xlDialogPivotFieldProperties             =  313, /*!< **Pivot Field Properties** dialog box */
    xlDialogPivotFieldUngroup                =  434, /*!< **Pivot Field Ungroup** dialog box */
    xlDialogPivotShowPages                   =  421, /*!< **Pivot Show Pages** dialog box */
    xlDialogPivotSolveOrder                  =  568, /*!< **Pivot Solve Order** dialog box */
    xlDialogPivotTableOptions                =  567, /*!< **PivotTable Options** dialog box */
    xlDialogPivotTableSlicerConnections      = 1183, /*!< **PivotTable Slicer Connections** dialog box */
    xlDialogPivotTableWhatIfAnalysisSettings = 1153, /*!< **PivotTable What If Analysis Settings** dialog box */
    xlDialogPivotTableWizard                 =  312, /*!< **PivotTable Wizard** dialog box */
    xlDialogPlacement                        =  300, /*!< **Placement** dialog box */
    xlDialogPrint                            =    8, /*!< **Print** dialog box */
    xlDialogPrinterSetup                     =    9, /*!< **Printer Setup** dialog box */
    xlDialogPrintPreview                     =  222, /*!< **Print Preview** dialog box */
    xlDialogPromote                          =  202, /*!< **Promote** dialog box */
    xlDialogProperties                       =  474, /*!< **Properties** dialog box */
    xlDialogPropertyFields                   =  754, /*!< **Property Fields** dialog box */
    xlDialogProtectDocument                  =   28, /*!< **Protect Document** dialog box */
    xlDialogProtectSharing                   =  620, /*!< **Protect Sharing** dialog box */
    xlDialogPublishAsWebPage                 =  653, /*!< **Publish As WebPage** dialog box */
    xlDialogPushbuttonProperties             =  445, /*!< **Pushbutton Properties** dialog box */
    xlDialogRecommendedPivotTables           = 1258, /*!< **Recommended PivotTables** dialog box */
    xlDialogReplaceFont                      =  134, /*!< **Replace Font** dialog box */
    xlDialogRoutingSlip                      =  336, /*!< This object or member has been deprecated, but it remains part of the object model for backward compatibility. You should not use it in new applications. */
    xlDialogRowHeight                        =  127, /*!< **Row Height** dialog box */
    xlDialogRun                              =   17, /*!< **Run** dialog box */
    xlDialogSaveAs                           =    5, /*!< **SaveAs** dialog box */
    xlDialogSaveCopyAs                       =  456, /*!< **SaveCopyAs** dialog box */
    xlDialogSaveNewObject                    =  208, /*!< **Save New Object** dialog box */
    xlDialogSaveWorkbook                     =  145, /*!< **Save Workbook** dialog box */
    xlDialogSaveWorkspace                    =  285, /*!< **Save Workspace** dialog box */
    xlDialogScale                            =   87, /*!< **Scale** dialog box */
    xlDialogScenarioAdd                      =  307, /*!< **Scenario Add** dialog box */
    xlDialogScenarioCells                    =  305, /*!< **Scenario Cells** dialog box */
    xlDialogScenarioEdit                     =  308, /*!< **Scenario Edit** dialog box */
    xlDialogScenarioMerge                    =  473, /*!< **Scenario Merge** dialog box */
    xlDialogScenarioSummary                  =  311, /*!< **Scenario Summary** dialog box */
    xlDialogScrollbarProperties              =  420, /*!< **Scrollbar Properties** dialog box */
    xlDialogSearch                           =  731, /*!< **Search** dialog box */
    xlDialogSelectSpecial                    =  132, /*!< **Select Special** dialog box */
    xlDialogSendMail                         =  189, /*!< **Send Mail** dialog box */
    xlDialogSeriesAxes                       =  460, /*!< **Series Axes** dialog box */
    xlDialogSeriesOptions                    =  557, /*!< **Series Options** dialog box */
    xlDialogSeriesOrder                      =  466, /*!< **Series Order** dialog box */
    xlDialogSeriesShape                      =  504, /*!< **Series Shape** dialog box */
    xlDialogSeriesX                          =  461, /*!< **Series X** dialog box */
    xlDialogSeriesY                          =  462, /*!< **Series Y** dialog box */
    xlDialogSetBackgroundPicture             =  509, /*!< **Set Background Picture** dialog box */
    xlDialogSetManager                       = 1109, /*!< **Set Manager** dialog box */
    xlDialogSetMDXEditor                     = 1208, /*!< **Set MDX Editor** dialog box */
    xlDialogSetPrintTitles                   =   23, /*!< **Set Print Titles** dialog box */
    xlDialogSetTupleEditorOnColumns          = 1108, /*!< **Set Tuple Editor On Columns** dialog box */
    xlDialogSetTupleEditorOnRows             = 1107, /*!< **Set Tuple Editor On Rows** dialog box */
    xlDialogSetUpdateStatus                  =  159, /*!< **Set Update Status** dialog box */
    xlDialogShowDetail                       =  204, /*!< **Show Detail** dialog box */
    xlDialogShowToolbar                      =  220, /*!< **Show Toolbar** dialog box */
    xlDialogSize                             =  261, /*!< **Size** dialog box */
    xlDialogSlicerCreation                   = 1182, /*!< **Slicer Creation** dialog box */
    xlDialogSlicerPivotTableConnections      = 1184, /*!< **Slicer PivotTable Connections** dialog box */
    xlDialogSlicerSettings                   = 1179, /*!< **Slicer Settings** dialog box */
    xlDialogSort                             =   39, /*!< **Sort** dialog box */
    xlDialogSortSpecial                      =  192, /*!< **Sort Special** dialog box */
    xlDialogSparklineInsertColumn            = 1134, /*!< **Sparkline Insert Column** dialog box */
    xlDialogSparklineInsertLine              = 1133, /*!< **Sparkline Insert Line** dialog box */
    xlDialogSparklineInsertWinLoss           = 1135, /*!< **Sparkline Insert Win Loss** dialog box */
    xlDialogSplit                            =  137, /*!< **Split** dialog box */
    xlDialogStandardFont                     =  190, /*!< **Standard Font** dialog box */
    xlDialogStandardWidth                    =  472, /*!< **Standard Width** dialog box */
    xlDialogStyle                            =   44, /*!< **Style** dialog box */
    xlDialogSubscribeTo                      =  218, /*!< **Subscribe To** dialog box */
    xlDialogSubtotalCreate                   =  398, /*!< **Subtotal Create** dialog box */
    xlDialogSummaryInfo                      =  474, /*!< **Summary Info** dialog box */
    xlDialogTable                            =   41, /*!< **Table** dialog box */
    xlDialogTabOrder                         =  394, /*!< **Tab Order** dialog box */
    xlDialogTextToColumns                    =  422, /*!< **Text To Columns** dialog box */
    xlDialogUnhide                           =   94, /*!< **Unhide** dialog box */
    xlDialogUpdateLink                       =  201, /*!< **Update Link** dialog box */
    xlDialogVbaInsertFile                    =  328, /*!< **VBA Insert File** dialog box */
    xlDialogVbaMakeAddin                     =  478, /*!< **VBA Make Addin** dialog box */
    xlDialogVbaProcedureDefinition           =  330, /*!< **VBA Procedure Definition** dialog box */
    xlDialogView3d                           =  197, /*!< **View 3D** dialog box */
    xlDialogWebOptionsBrowsers               =  773, /*!< **Web Options Browsers** dialog box */
    xlDialogWebOptionsEncoding               =  686, /*!< **Web Options Encoding** dialog box */
    xlDialogWebOptionsFiles                  =  684, /*!< **Web Options Files** dialog box */
    xlDialogWebOptionsFonts                  =  687, /*!< **Web Options Fonts** dialog box */
    xlDialogWebOptionsGeneral                =  683, /*!< **Web Options General** dialog box */
    xlDialogWebOptionsPictures               =  685, /*!< **Web Options Pictures** dialog box */
    xlDialogWindowMove                       =   14, /*!< **Window Move** dialog box */
    xlDialogWindowSize                       =   13, /*!< **Window Size** dialog box */
    xlDialogWorkbookAdd                      =  281, /*!< **Workbook Add** dialog box */
    xlDialogWorkbookCopy                     =  283, /*!< **Workbook Copy** dialog box */
    xlDialogWorkbookInsert                   =  354, /*!< **Workbook Insert** dialog box */
    xlDialogWorkbookMove                     =  282, /*!< **Workbook Move** dialog box */
    xlDialogWorkbookName                     =  386, /*!< **Workbook Name** dialog box */
    xlDialogWorkbookNew                      =  302, /*!< **Workbook New** dialog box */
    xlDialogWorkbookOptions                  =  284, /*!< **Workbook Options** dialog box */
    xlDialogWorkbookProtect                  =  417, /*!< **Workbook Protect** dialog box */
    xlDialogWorkbookTabSplit                 =  415, /*!< **Workbook Tab Split** dialog box */
    xlDialogWorkbookUnhide                   =  384, /*!< **Workbook Unhide** dialog box */
    xlDialogWorkgroup                        =  199, /*!< **Workgroup** dialog box */
    xlDialogWorkspace                        =   95, /*!< **Workspace** dialog box */
    xlDialogZoom                             =  256, /*!< **Zoom** dialog box */
};

/**
Specifies the cell error number and value.

[Official VBA documentation for XlCVError](https://docs.microsoft.com/office/vba/api/excel.xlcverror)
*/
enum XlCVError
{
    xlErrDiv0  = 2007, /*!< Error number: 2007 */
    xlErrNA    = 2042, /*!< Error number: 2042 */
    xlErrName  = 2029, /*!< Error number: 2029 */
    xlErrNull  = 2000, /*!< Error number: 2000 */
    xlErrNum   = 2036, /*!< Error number: 2036 */
    xlErrRef   = 2023, /*!< Error number: 2023 */
    xlErrValue = 2015, /*!< Error number: 2015 */
    xlErrSpill = 2045, /*!< Error number: 2045 */
};

/**
Specifies what should be calculated.

[Official VBA documentation for XlCalcFor](https://docs.microsoft.com/office/vba/api/excel.xlcalcfor)
*/
enum XlCalcFor
{
    xlAllValues = 0, /*!< All values. */
    xlColGroups = 2, /*!< Column groups. */
    xlRowGroups = 1, /*!< Row groups. */
};

/**
Specifies the format of cell values that are based on the calculated member.

[Official VBA documentation for XlCalcMemNumberFormatType](https://docs.microsoft.com/office/vba/api/excel.xlcalcmemnumberformattype)
*/
enum XlCalcMemNumberFormatType
{
    xlNumberFormatTypeDefault = 0, /*!< Use the default format type of the calculated member for the cell value. */
    xlNumberFormatTypeNumber  = 1, /*!< Calculated member cell format is a number. */
    xlNumberFormatTypePercent = 2, /*!< Calculated member cell format is a percentage. */
};

/**
Specifies the type of a calculated member in a PivotTable.

[Official VBA documentation for XlCalculatedMemberType](https://docs.microsoft.com/office/vba/api/excel.xlcalculatedmembertype)
*/
enum XlCalculatedMemberType
{
    xlCalculatedMeasure = 2, /*!< The member is a Multidimensional Expressions (MDX) expression that defines the measure. */
    xlCalculatedMember  = 0, /*!< The member uses a Multidimensional Expression (MDX) formula. */
    xlCalculatedSet     = 1, /*!< The member contains an MDX formula for a set in a cube field. */
};

/**
Specifies the calculation mode.

[Official VBA documentation for XlCalculation](https://docs.microsoft.com/office/vba/api/excel.xlcalculation)
*/
enum XlCalculation
{
    xlCalculationAutomatic     = -4105, /*!< Excel controls recalculation. */
    xlCalculationManual        = -4135, /*!< Calculation is done when the user requests it. */
    xlCalculationSemiautomatic =     2, /*!< Excel controls recalculation but ignores changes in tables. */
};

/**
Specifies which key interrupts recalculation.

[Official VBA documentation for XlCalculationInterruptKey](https://docs.microsoft.com/office/vba/api/excel.xlcalculationinterruptkey)
*/
enum XlCalculationInterruptKey
{
    xlAnyKey = 2, /*!< Pressing any key interrupts recalculation. */
    xlEscKey = 1, /*!< Pressing the ESC key interrupts recalculation. */
    xlNoKey  = 0, /*!< No key press can interrupt recalculation. */
};

/**
Specifies the calculation state of the application.

[Official VBA documentation for XlCalculationState](https://docs.microsoft.com/office/vba/api/excel.xlcalculationstate)
*/
enum XlCalculationState
{
    xlCalculating = 1, /*!< Calculations in process. */
    xlDone        = 0, /*!< Calculations complete. */
    xlPending     = 2, /*!< Changes that trigger calculation have been made, but a recalculation has not yet been performed. */
};

/**
Specifies the category labels for the category label levels.

[Official VBA documentation for XlCategoryLabelLevel](https://docs.microsoft.com/office/vba/api/excel.xlcategorylabellevel)
*/
enum XlCategoryLabelLevel
{
    xlCategoryLabelLevelAll    = -1, /*!< Set category labels to all category label levels w/in range on the chart. */
    xlCategoryLabelLevelCustom = -2, /*!< Indicates literal data in the category labels. */
    xlCategoryLabelLevelNone   = -3, /*!< Set no category labels in the chart. Defaults to automatic indexed labels. */
};

/**
Specifies the type of the category axis.

[Official VBA documentation for XlCategoryType](https://docs.microsoft.com/office/vba/api/excel.xlcategorytype)
*/
enum XlCategoryType
{
    xlAutomaticScale = -4105, /*!< Excel controls the axis type. */
    xlCategoryScale  =     2, /*!< Axis groups data by an arbitrary set of categories. */
    xlTimeScale      =     3, /*!< Axis groups data on a time scale. */
};

/**
Specifies whether a PivotTable value cell has been edited or recalculated since the PivotTable report was created or the last commit operation was performed. 

[Official VBA documentation for XlCellChangedState](https://docs.microsoft.com/office/vba/api/excel.xlcellchangedstate)
*/
enum XlCellChangedState
{
    xlCellChangeApplied = 3, /*!< The value in the cell has been edited or recalculated, and that change has been applied to the data source. (Applies only PivotTable reports with OLAP data sources) */
    xlCellChanged       = 2, /*!< The value in the cell has been edited or recalculated. */
    xlCellNotChanged    = 1, /*!< The value in the cell has not been edited or recalculated. */
};

/**
Specifies the way that rows on the specified worksheet are added or deleted to accommodate the number of rows in a recordset returned by a query.

[Official VBA documentation for XlCellInsertionMode](https://docs.microsoft.com/office/vba/api/excel.xlcellinsertionmode)
*/
enum XlCellInsertionMode
{
    xlInsertDeleteCells = 1, /*!< Partial rows are inserted or deleted to match the exact number of rows required for the new recordset. */
    xlInsertEntireRows  = 2, /*!< Entire rows are inserted, if necessary, to accommodate any overflow. No cells or rows are deleted from the worksheet. */
    xlOverwriteCells    = 0, /*!< No new cells or rows are added to the worksheet. Data in surrounding cells is overwritten to accommodate any overflow. */
};

/**
Specifies the type of cells.

[Official VBA documentation for XlCellType](https://docs.microsoft.com/office/vba/api/excel.xlcelltype)
*/
enum XlCellType
{
    xlCellTypeAllFormatConditions  = -4172, /*!< Cells of any format. */
    xlCellTypeAllValidation        = -4174, /*!< Cells having validation criteria. */
    xlCellTypeBlanks               =     4, /*!< Empty cells. */
    xlCellTypeComments             = -4144, /*!< Cells containing notes. */
    xlCellTypeConstants            =     2, /*!< Cells containing constants. */
    xlCellTypeFormulas             = -4123, /*!< Cells containing formulas. */
    xlCellTypeLastCell             =    11, /*!< The last cell in the used range. */
    xlCellTypeSameFormatConditions = -4173, /*!< Cells having the same format. */
    xlCellTypeSameValidation       = -4175, /*!< Cells having the same validation criteria. */
    xlCellTypeVisible              =    12, /*!< All visible cells. */
};

/**
Specifies the position of the chart element.

[Official VBA documentation for XlChartElementPosition](https://docs.microsoft.com/office/vba/api/excel.xlchartelementposition)
*/
enum XlChartElementPosition
{
    xlChartElementPositionAutomatic = -4105, /*!< Automatically sets the position of the chart element. */
    xlChartElementPositionCustom    = -4114, /*!< Specifies a specific position for the chart element. */
};

/**
Specifies a chart gallery.

[Official VBA documentation for XlChartGallery](https://docs.microsoft.com/office/vba/api/excel.xlchartgallery)
*/
enum XlChartGallery
{
    xlAnyGallery  = 23, /*!< Either of the galleries. */
    xlBuiltIn     = 21, /*!< The built-in gallery. */
    xlUserDefined = 22, /*!< The user-defined gallery. */
};

/**
Specifies the type of the chart item.

[Official VBA documentation for XlChartItem](https://docs.microsoft.com/office/vba/api/excel.xlchartitem)
*/
enum XlChartItem
{
    xlAxis                  = 21, /*!< Axis. */
    xlAxisTitle             = 17, /*!< Axis title. */
    xlChartArea             =  2, /*!< Chart area. */
    xlChartTitle            =  4, /*!< Chart title. */
    xlCorners               =  6, /*!< Corners. */
    xlDataLabel             =  0, /*!< Data label. */
    xlDataTable             =  7, /*!< Data table. */
    xlDisplayUnitLabel      = 30, /*!< Display unit label. */
    xlDownBars              = 20, /*!< Down bars. */
    xlDropLines             = 26, /*!< Drop lines. */
    xlErrorBars             =  9, /*!< Error bars. */
    xlFloor                 = 23, /*!< Floor. */
    xlHiLoLines             = 25, /*!< HiLo lines. */
    xlLeaderLines           = 29, /*!< Leader lines. */
    xlLegend                = 24, /*!< Legend. */
    xlLegendEntry           = 12, /*!< Legend entry. */
    xlLegendKey             = 13, /*!< Legend key. */
    xlMajorGridlines        = 15, /*!< Major gridlines. */
    xlMinorGridlines        = 16, /*!< Minor gridlines. */
    xlNothing               = 28, /*!< Nothing. */
    xlPivotChartDropZone    = 32, /*!< PivotChart drop zone. */
    xlPivotChartFieldButton = 31, /*!< PivotChart field button. */
    xlPlotArea              = 19, /*!< Plot area. */
    xlRadarAxisLabels       = 27, /*!< Radar axis labels. */
    xlSeries                =  3, /*!< Series. */
    xlSeriesLines           = 22, /*!< Series lines. */
    xlShape                 = 14, /*!< Shape. */
    xlTrendline             =  8, /*!< Trend line. */
    xlUpBars                = 18, /*!< Up bars. */
    xlWalls                 =  5, /*!< Walls. */
    xlXErrorBars            = 10, /*!< X error bars. */
    xlYErrorBars            = 11, /*!< Y error bars. */
};

/**
Specifies where to relocate a chart.

[Official VBA documentation for XlChartLocation](https://docs.microsoft.com/office/vba/api/excel.xlchartlocation)
*/
enum XlChartLocation
{
    xlLocationAsNewSheet = 1, /*!< Chart is moved to a new sheet. */
    xlLocationAsObject   = 2, /*!< Chart is to be embedded in an existing sheet. */
    xlLocationAutomatic  = 3, /*!< Excel controls chart location. */
};

/**
Specifies the placement of a user-selected picture on a bar in a 3D bar or column.

[Official VBA documentation for XlChartPicturePlacement](https://docs.microsoft.com/office/vba/api/excel.xlchartpictureplacement)
*/
enum XlChartPicturePlacement
{
    xlAllFaces   = 7, /*!< Display on all faces. */
    xlEnd        = 2, /*!< Display on end. */
    xlEndSides   = 3, /*!< Display on end and sides. */
    xlFront      = 4, /*!< Display on front. */
    xlFrontEnd   = 6, /*!< Display on front and end. */
    xlFrontSides = 5, /*!< Display on front and sides. */
    xlSides      = 1, /*!< Display on sides. */
};

/**
Specifies how pictures are displayed on a column, bar picture chart, or legend key.

[Official VBA documentation for XlChartPictureType](https://docs.microsoft.com/office/vba/api/excel.xlchartpicturetype)
*/
enum XlChartPictureType
{
    xlStack      = 2, /*!< Picture is sized to repeat a maximum of 15 times in the longest stacked bar. */
    xlStackScale = 3, /*!< Picture is sized to a specified number of units and repeated the length of the bar. */
    xlStretch    = 1, /*!< Picture is stretched the full length of the stacked bar. */
};

/**
Specifies the values displayed in the second chart in a pie chart or a Bar of Pie chart.

[Official VBA documentation for XlChartSplitType](https://docs.microsoft.com/office/vba/api/excel.xlchartsplittype)
*/
enum XlChartSplitType
{
    xlSplitByCustomSplit  = 4, /*!< Arbitrary slides are displayed in the second chart. */
    xlSplitByPercentValue = 3, /*!< Second chart displays values less than some percentage of the total value. The percentage is specified by the **SplitValue** property. */
    xlSplitByPosition     = 1, /*!< Second chart displays the smallest values in the data series. The number of values to display is specified by the **SplitValue** property. */
    xlSplitByValue        = 2, /*!< Second chart displays values less than the value specified by the **SplitValue** property. */
};

/**
Specifies the chart type.

[Official VBA documentation for XlChartType](https://docs.microsoft.com/office/vba/api/excel.xlcharttype)
*/
enum XlChartType
{
    xl3DArea                   = -4098, /*!< 3D Area. */
    xl3DAreaStacked            =    78, /*!< 3D Stacked Area. */
    xl3DAreaStacked100         =    79, /*!< 100% Stacked Area. */
    xl3DBarClustered           =    60, /*!< 3D Clustered Bar. */
    xl3DBarStacked             =    61, /*!< 3D Stacked Bar. */
    xl3DBarStacked100          =    62, /*!< 3D 100% Stacked Bar. */
    xl3DColumn                 = -4100, /*!< 3D Column. */
    xl3DColumnClustered        =    54, /*!< 3D Clustered Column. */
    xl3DColumnStacked          =    55, /*!< 3D Stacked Column. */
    xl3DColumnStacked100       =    56, /*!< 3D 100% Stacked Column. */
    xl3DLine                   = -4101, /*!< 3D Line. */
    xl3DPie                    = -4102, /*!< 3D Pie. */
    xl3DPieExploded            =    70, /*!< Exploded 3D Pie. */
    xlArea                     =     1, /*!< Area */
    xlAreaStacked              =    76, /*!< Stacked Area. */
    xlAreaStacked100           =    77, /*!< 100% Stacked Area. */
    xlBarClustered             =    57, /*!< Clustered Bar. */
    xlBarOfPie                 =    71, /*!< Bar of Pie. */
    xlBarStacked               =    58, /*!< Stacked Bar. */
    xlBarStacked100            =    59, /*!< 100% Stacked Bar. */
    xlBoxwhisker               =   121, /*!< not officially documented */
    xlBubble                   =    15, /*!< Bubble. */
    xlBubble3DEffect           =    87, /*!< Bubble with 3D effects. */
    xlColumnClustered          =    51, /*!< Clustered Column. */
    xlColumnStacked            =    52, /*!< Stacked Column. */
    xlColumnStacked100         =    53, /*!< 100% Stacked Column. */
    xlConeBarClustered         =   102, /*!< Clustered Cone Bar. */
    xlConeBarStacked           =   103, /*!< Stacked Cone Bar. */
    xlConeBarStacked100        =   104, /*!< 100% Stacked Cone Bar. */
    xlConeCol                  =   105, /*!< 3D Cone Column. */
    xlConeColClustered         =    99, /*!< Clustered Cone Column. */
    xlConeColStacked           =   100, /*!< Stacked Cone Column. */
    xlConeColStacked100        =   101, /*!< 100% Stacked Cone Column. */
    xlCylinderBarClustered     =    95, /*!< Clustered Cylinder Bar. */
    xlCylinderBarStacked       =    96, /*!< Stacked Cylinder Bar. */
    xlCylinderBarStacked100    =    97, /*!< 100% Stacked Cylinder Bar. */
    xlCylinderCol              =    98, /*!< 3D Cylinder Column. */
    xlCylinderColClustered     =    92, /*!< Clustered Cone Column. */
    xlCylinderColStacked       =    93, /*!< Stacked Cone Column. */
    xlCylinderColStacked100    =    94, /*!< 100% Stacked Cylinder Column. */
    xlDoughnut                 = -4120, /*!< Doughnut. */
    xlDoughnutExploded         =    80, /*!< Exploded Doughnut. */
    xlFunnel                   =   123, /*!< not officially documented */
    xlHistogram                =   118, /*!< not officially documented */
    xlLine                     =     4, /*!< Line. */
    xlLineMarkers              =    65, /*!< Line with Markers. */
    xlLineMarkersStacked       =    66, /*!< Stacked Line with Markers. */
    xlLineMarkersStacked100    =    67, /*!< 100% Stacked Line with Markers. */
    xlLineStacked              =    63, /*!< Stacked Line. */
    xlLineStacked100           =    64, /*!< 100% Stacked Line. */
    xlPareto                   =   122, /*!< not officially documented */
    xlPie                      =     5, /*!< Pie. */
    xlPieExploded              =    69, /*!< Exploded Pie. */
    xlPieOfPie                 =    68, /*!< Pie of Pie. */
    xlPyramidBarClustered      =   109, /*!< Clustered Pyramid Bar. */
    xlPyramidBarStacked        =   110, /*!< Stacked Pyramid Bar. */
    xlPyramidBarStacked100     =   111, /*!< 100% Stacked Pyramid Bar. */
    xlPyramidCol               =   112, /*!< 3D Pyramid Column. */
    xlPyramidColClustered      =   106, /*!< Clustered Pyramid Column. */
    xlPyramidColStacked        =   107, /*!< Stacked Pyramid Column. */
    xlPyramidColStacked100     =   108, /*!< 100% Stacked Pyramid Column. */
    xlRadar                    = -4151, /*!< Radar. */
    xlRadarFilled              =    82, /*!< Filled Radar. */
    xlRadarMarkers             =    81, /*!< Radar with Data Markers. */
    xlRegionMap                =   140, /*!< Map chart. */
    xlStockHLC                 =    88, /*!< High-Low-Close. */
    xlStockOHLC                =    89, /*!< Open-High-Low-Close. */
    xlStockVHLC                =    90, /*!< Volume-High-Low-Close. */
    xlStockVOHLC               =    91, /*!< Volume-Open-High-Low-Close. */
    xlSunburst                 =   120, /*!< not officially documented */
    xlSurface                  =    83, /*!< 3D Surface. */
    xlSurfaceTopView           =    85, /*!< Surface (Top View). */
    xlSurfaceTopViewWireframe  =    86, /*!< Surface (Top View wireframe). */
    xlSurfaceWireframe         =    84, /*!< 3D Surface (wireframe). */
    xlTreemap                  =   117, /*!< not officially documented */
    xlWaterfall                =   119, /*!< not officially documented */
    xlXYScatter                = -4169, /*!< Scatter. */
    xlXYScatterLines           =    74, /*!< Scatter with Lines. */
    xlXYScatterLinesNoMarkers  =    75, /*!< Scatter with Lines and No Data Markers. */
    xlXYScatterSmooth          =    72, /*!< Scatter with Smoothed Lines. */
    xlXYScatterSmoothNoMarkers =    73, /*!< Scatter with Smoothed Lines and No Data Markers. */
};

/**
Specifies the type of version for the document checked in when using the **CheckIn** method. Applies to workbooks stored in a SharePoint library.

[Official VBA documentation for XlCheckInVersionType](https://docs.microsoft.com/office/vba/api/excel.xlcheckinversiontype)
*/
enum XlCheckInVersionType
{
    xlCheckInMajorVersion     = 1, /*!< Check in the major version. */
    xlCheckInMinorVersion     = 0, /*!< Check in the minor version. */
    xlCheckInOverwriteVersion = 2, /*!< Overwrite current version on the server. */
};

/**
Specifies the format of an item on the Microsoft Windows clipboard.

[Official VBA documentation for XlClipboardFormat](https://docs.microsoft.com/office/vba/api/excel.xlclipboardformat)
*/
enum XlClipboardFormat
{
    xlClipboardFormatBIFF           =  8, /*!< Binary Interchange file format for Excel version 2.x */
    xlClipboardFormatBIFF12         = 63, /*!< Binary Interchange file format 12 */
    xlClipboardFormatBIFF2          = 18, /*!< Binary Interchange file format 2 */
    xlClipboardFormatBIFF3          = 20, /*!< Binary Interchange file format 3 */
    xlClipboardFormatBIFF4          = 30, /*!< Binary Interchange file format 4 */
    xlClipboardFormatBinary         = 15, /*!< Binary format */
    xlClipboardFormatBitmap         =  9, /*!< Bitmap format */
    xlClipboardFormatCGM            = 13, /*!< CGM format */
    xlClipboardFormatCSV            =  5, /*!< CSV format */
    xlClipboardFormatDIF            =  4, /*!< DIF format */
    xlClipboardFormatDspText        = 12, /*!< Dsp Text format */
    xlClipboardFormatEmbeddedObject = 21, /*!< Embedded Object */
    xlClipboardFormatEmbedSource    = 22, /*!< Embedded Source */
    xlClipboardFormatLink           = 11, /*!< Link */
    xlClipboardFormatLinkSource     = 23, /*!< Link to the source file */
    xlClipboardFormatLinkSourceDesc = 32, /*!< Link to the source description */
    xlClipboardFormatMovie          = 24, /*!< Movie */
    xlClipboardFormatNative         = 14, /*!< Native */
    xlClipboardFormatObjectDesc     = 31, /*!< Object description */
    xlClipboardFormatObjectLink     = 19, /*!< Object link */
    xlClipboardFormatOwnerLink      = 17, /*!< Link to the owner */
    xlClipboardFormatPICT           =  2, /*!< Picture */
    xlClipboardFormatPrintPICT      =  3, /*!< Print picture */
    xlClipboardFormatRTF            =  7, /*!< RTF format */
    xlClipboardFormatScreenPICT     = 29, /*!< Screen Picture */
    xlClipboardFormatStandardFont   = 28, /*!< Standard Font */
    xlClipboardFormatStandardScale  = 27, /*!< Standard Scale */
    xlClipboardFormatSYLK           =  6, /*!< SYLK */
    xlClipboardFormatTable          = 16, /*!< Table */
    xlClipboardFormatText           =  0, /*!< Text */
    xlClipboardFormatToolFace       = 25, /*!< Tool Face */
    xlClipboardFormatToolFacePICT   = 26, /*!< Tool Face Picture */
    xlClipboardFormatVALU           =  1, /*!< Value */
    xlClipboardFormatWK1            = 10, /*!< Workbook */
};

/**
Specifies the value of the **CommandText** property.

[Official VBA documentation for XlCmdType](https://docs.microsoft.com/office/vba/api/excel.xlcmdtype)
*/
enum XlCmdType
{
    xlCmdCube            = 1, /*!< Contains a cube name for an OLAP data source. */
    xlCmdDAX             = 8, /*!< Contains a Data Analysis Expressions (DAX) formula. */
    xlCmdDefault         = 4, /*!< Contains command text that the OLE DB provider understands. */
    xlCmdExcel           = 7, /*!< Contains an Excel formula. */
    xlCmdList            = 5, /*!< Contains a pointer to list data. */
    xlCmdSql             = 2, /*!< Contains an SQL statement. */
    xlCmdTable           = 3, /*!< Contains a table name for accessing OLE DB data sources. */
    xlCmdTableCollection = 6, /*!< Contains the name of a table collection. */
};

/**
Specifies the color of a selected feature, such as a border, font, or fill.

[Official VBA documentation for XlColorIndex](https://docs.microsoft.com/office/vba/api/excel.xlcolorindex)
*/
enum XlColorIndex
{
    xlColorIndexAutomatic = -4105, /*!< Automatic color. */
    xlColorIndexNone      = -4142, /*!< No color. */
};

/**
Specifies how a column is to be parsed.

[Official VBA documentation for XlColumnDataType](https://docs.microsoft.com/office/vba/api/excel.xlcolumndatatype)
*/
enum XlColumnDataType
{
    xlDMYFormat     =  4, /*!< DMY date format. */
    xlDYMFormat     =  7, /*!< DYM date format. */
    xlEMDFormat     = 10, /*!< EMD date format. */
    xlGeneralFormat =  1, /*!< General. */
    xlMDYFormat     =  3, /*!< MDY date format. */
    xlMYDFormat     =  6, /*!< MYD date format. */
    xlSkipColumn    =  9, /*!< Column is not parsed. */
    xlTextFormat    =  2, /*!< Text. */
    xlYDMFormat     =  8, /*!< YDM date format. */
    xlYMDFormat     =  5, /*!< YMD date format. */
};

/**
Specifies the state of the command underlines in Microsoft Excel for the Macintosh.

[Official VBA documentation for XlCommandUnderlines](https://docs.microsoft.com/office/vba/api/excel.xlcommandunderlines)
*/
enum XlCommandUnderlines
{
    xlCommandUnderlinesAutomatic = -4105, /*!< Excel controls the display of command underlines. */
    xlCommandUnderlinesOff       = -4146, /*!< Command underlines are not displayed. */
    xlCommandUnderlinesOn        =     1, /*!< Command underlines are displayed. */
};

/**
Specifies the way that cells display comments and comment indicators.

[Official VBA documentation for XlCommentDisplayMode](https://docs.microsoft.com/office/vba/api/excel.xlcommentdisplaymode)
*/
enum XlCommentDisplayMode
{
    xlCommentAndIndicator  =  1, /*!< Display comment and indicator at all times. */
    xlCommentIndicatorOnly = -1, /*!< Display comment indicator only. Display comment when mouse pointer is moved over cell. */
    xlNoIndicator          =  0, /*!< Display neither the comment nor the comment indicator at any time. */
};

/**
Specifies the types of condition values that can be used.

[Official VBA documentation for XlConditionValueTypes](https://docs.microsoft.com/office/vba/api/excel.xlconditionvaluetypes)
*/
enum XlConditionValueTypes
{
    xlConditionValueAutomaticMax =  7, /*!< The longest data bar is proportional to the maximum value in the range. */
    xlConditionValueAutomaticMin =  6, /*!< The shortest data bar is proportional to the minimum value in the range. */
    xlConditionValueFormula      =  4, /*!< Formula is used. */
    xlConditionValueHighestValue =  2, /*!< Highest value from the list of values. */
    xlConditionValueLowestValue  =  1, /*!< Lowest value from the list of values. */
    xlConditionValueNone         = -1, /*!< No conditional value. */
    xlConditionValueNumber       =  0, /*!< Number is used. */
    xlConditionValuePercent      =  3, /*!< Percentage is used. */
    xlConditionValuePercentile   =  5, /*!< Percentile is used. */
};

/**
Specifies the type of database connection.

[Official VBA documentation for XlConnectionType](https://docs.microsoft.com/office/vba/api/excel.xlconnectiontype)
*/
enum XlConnectionType
{
    xlConnectionTypeDATAFEED  = 6, /*!< Data Feed */
    xlConnectionTypeMODEL     = 7, /*!< PowerPivot Model */
    xlConnectionTypeNOSOURCE  = 9, /*!< No source */
    xlConnectionTypeODBC      = 2, /*!< ODBC */
    xlConnectionTypeOLEDB     = 1, /*!< OLEDB */
    xlConnectionTypeTEXT      = 4, /*!< Text */
    xlConnectionTypeWEB       = 5, /*!< Web */
    xlConnectionTypeWORKSHEET = 8, /*!< Worksheet */
    xlConnectionTypeXMLMAP    = 3, /*!< XML MAP */
};

/**
Specifies the subtotal function.

[Official VBA documentation for XlConsolidationFunction](https://docs.microsoft.com/office/vba/api/excel.xlconsolidationfunction)
*/
enum XlConsolidationFunction
{
    xlAverage       = -4106, /*!< Average. */
    xlCount         = -4112, /*!< Count. */
    xlCountNums     = -4113, /*!< Count numerical values only. */
    xlDistinctCount =    11, /*!< Count using Distinct Count analysis. */
    xlMax           = -4136, /*!< Maximum. */
    xlMin           = -4139, /*!< Minimum. */
    xlProduct       = -4149, /*!< Multiply. */
    xlStDev         = -4155, /*!< Standard deviation, based on a sample. */
    xlStDevP        = -4156, /*!< Standard deviation, based on the whole population. */
    xlSum           = -4157, /*!< Sum. */
    xlUnknown       =  1000, /*!< No subtotal function specified. */
    xlVar           = -4164, /*!< Variation, based on a sample. */
    xlVarP          = -4165, /*!< Variation, based on the whole population. */
};

/**
Specifies the operator used in a function.

[Official VBA documentation for XlContainsOperator](https://docs.microsoft.com/office/vba/api/excel.xlcontainsoperator)
*/
enum XlContainsOperator
{
    xlBeginsWith     = 2, /*!< Begins with a specified value. */
    xlContains       = 0, /*!< Contains a specified value. */
    xlDoesNotContain = 1, /*!< Does not contain the specified value. */
    xlEndsWith       = 3, /*!< Endswith the specified value */
};

/**
Specifies the format of the picture being copied.

[Official VBA documentation for XlCopyPictureFormat](https://docs.microsoft.com/office/vba/api/excel.xlcopypictureformat)
*/
enum XlCopyPictureFormat
{
    xlBitmap  =     2, /*!< Picture copied in bitmap (raster) format: bmp, jpg, gif, png. */
    xlPicture = -4147, /*!< Picture copied in vector format: emf, wmf. */
};

/**
Specifies the processing for a file when it is opened.

[Official VBA documentation for XlCorruptLoad](https://docs.microsoft.com/office/vba/api/excel.xlcorruptload)
*/
enum XlCorruptLoad
{
    xlExtractData = 2, /*!< Workbook is opened in extract data mode. */
    xlNormalLoad  = 0, /*!< Workbook is opened normally. */
    xlRepairFile  = 1, /*!< Workbook is opened in repair mode. */
};

/**
Specifies the 32-bit creator code for Excel for Macintosh (decimal 1480803660, Hex 5843454C, string XCEL).

[Official VBA documentation for XlCreator](https://docs.microsoft.com/office/vba/api/excel.xlcreator)
*/
enum XlCreator
{
    xlCreatorCode = 1480803660, /*!< The Excel for Macintosh creator code. */
};

/**
Specifies the type of credentials method used.

[Official VBA documentation for XlCredentialsMethod](https://docs.microsoft.com/office/vba/api/excel.xlcredentialsmethod)
*/
enum XlCredentialsMethod
{
    CredentialsMethodIntegrated = 0, /*!< Integrated */
    CredentialsMethodNone       = 1, /*!< No credentials used */
    CredentialsMethodStored     = 2, /*!< Use stored credentials */
};

/**
Specifies the subtype of the CubeField.

[Official VBA documentation for XlCubeFieldSubType](https://docs.microsoft.com/office/vba/api/excel.xlcubefieldsubtype)
*/
enum XlCubeFieldSubType
{
    xlCubeAttribute         =  4, /*!< Attribute */
    xlCubeCalculatedMeasure =  5, /*!< Calculated Measure */
    xlCubeHierarchy         =  1, /*!< Hierarchy */
    xlCubeImplicitMeasure   = 11, /*!< An implicit measure */
    xlCubeKPIGoal           =  7, /*!< KPI Goal */
    xlCubeKPIStatus         =  8, /*!< KPI Status */
    xlCubeKPITrend          =  9, /*!< KPI Trend */
    xlCubeKPIValue          =  6, /*!< KPI Value */
    xlCubeKPIWeight         = 10, /*!< KPI Weight */
    xlCubeMeasure           =  2, /*!< Measure */
    xlCubeSet               =  3, /*!< Set */
};

/**
Specifies whether the OLAP field is a hierarchy, set, or measure field.

[Official VBA documentation for XlCubeFieldType](https://docs.microsoft.com/office/vba/api/excel.xlcubefieldtype)
*/
enum XlCubeFieldType
{
    xlHierarchy = 1, /*!< OLAP field is a hierarchy. */
    xlMeasure   = 2, /*!< OLAP field is a measure. */
    xlSet       = 3, /*!< OLAP field is a set. */
};

/**
Specifies whether status is Copy mode or Cut mode.

[Official VBA documentation for XlCutCopyMode](https://docs.microsoft.com/office/vba/api/excel.xlcutcopymode)
*/
enum XlCutCopyMode
{
    xlCopy = 1, /*!< In Copy mode */
    xlCut  = 2, /*!< In Cut mode */
};

/**
Specifies the icon used in message boxes displayed during validation.

[Official VBA documentation for XlDVAlertStyle](https://docs.microsoft.com/office/vba/api/excel.xldvalertstyle)
*/
enum XlDVAlertStyle
{
    xlValidAlertInformation = 3, /*!< Information icon. */
    xlValidAlertStop        = 1, /*!< Stop icon. */
    xlValidAlertWarning     = 2, /*!< Warning icon. */
};

/**
Specifies the type of validation test to be performed in conjunction with values.

[Official VBA documentation for XlDVType](https://docs.microsoft.com/office/vba/api/excel.xldvtype)
*/
enum XlDVType
{
    xlValidateCustom      = 7, /*!< Data is validated using an arbitrary formula. */
    xlValidateDate        = 4, /*!< Date values. */
    xlValidateDecimal     = 2, /*!< Numeric values. */
    xlValidateInputOnly   = 0, /*!< Validate only when user changes the value. */
    xlValidateList        = 3, /*!< Value must be present in a specified list. */
    xlValidateTextLength  = 6, /*!< Length of text. */
    xlValidateTime        = 5, /*!< Time values. */
    xlValidateWholeNumber = 1, /*!< Whole numeric values. */
};

/**
Specifies the axis position for a range of cells with conditional formatting as data bars.

[Official VBA documentation for XlDataBarAxisPosition](https://docs.microsoft.com/office/vba/api/excel.xldatabaraxisposition)
*/
enum XlDataBarAxisPosition
{
    xlDataBarAxisAutomatic = 0, /*!< Display the axis at a variable position based on the ratio of the minimum negative value to the maximum positive value in the range. Positive values are displayed in a left-to-right direction. Negative values are displayed in a right-to-left direction. When all values are positive or all values are negative, no axis is displayed. */
    xlDataBarAxisMidpoint  = 1, /*!< Display the axis at the midpoint of the cell regardless of the set of values in the range. Positive values are displayed in a left-to-right direction. Negative values are displayed in a right-to-left direction. */
    xlDataBarAxisNone      = 2, /*!< No axis is displayed, and both positive and negative values are displayed in the left-to-right direction. */
};

/**
Specifies the border of a data bar.

[Official VBA documentation for XlDataBarBorderType](https://docs.microsoft.com/office/vba/api/excel.xldatabarbordertype)
*/
enum XlDataBarBorderType
{
    xlDataBarBorderNone  = 0, /*!< The data bar has no border. */
    xlDataBarBorderSolid = 1, /*!< The data bar has a solid border. */
};

/**
Specifies how a data bar is filled with color.

[Official VBA documentation for XlDataBarFillType](https://docs.microsoft.com/office/vba/api/excel.xldatabarfilltype)
*/
enum XlDataBarFillType
{
    xlDataBarFillGradient = 1, /*!< The data bar is filled with a color gradient. */
    xlDataBarFillSolid    = 0, /*!< The data bar is filled with solid color. */
};

/**
Specifies whether to use the same border and fill color as positive data bars.

[Official VBA documentation for XlDataBarNegativeColorType](https://docs.microsoft.com/office/vba/api/excel.xldatabarnegativecolortype)
*/
enum XlDataBarNegativeColorType
{
    xlDataBarColor          = 0, /*!< Use the color specified in the **Negative Value and Axis Setting** dialog box or by using the **[ColorType](Excel.NegativeBarFormat.ColorType.md)** and **[BorderColorType](Excel.NegativeBarFormat.BorderColorType.md)** properties of the **[NegativeBarFormat](Excel.NegativeBarFormat.md)** object. */
    xlDataBarSameAsPositive = 1, /*!< Use the same color as positive data bars. */
};

/**
Specifies where the data label is positioned.

[Official VBA documentation for XlDataLabelPosition](https://docs.microsoft.com/office/vba/api/excel.xldatalabelposition)
*/
enum XlDataLabelPosition
{
    xlLabelPositionAbove      =     0, /*!< Data label is positioned above the data point. */
    xlLabelPositionBelow      =     1, /*!< Data label is positioned below the data point. */
    xlLabelPositionBestFit    =     5, /*!< Microsoft Office Excel 2007 sets the position of the data label. */
    xlLabelPositionCenter     = -4108, /*!< Data label is centered on the data point or is inside a bar or pie chart. */
    xlLabelPositionCustom     =     7, /*!< Data label is in a custom position. */
    xlLabelPositionInsideBase =     4, /*!< Data label is positioned inside the data point at the bottom edge. */
    xlLabelPositionInsideEnd  =     3, /*!< Data label is positioned inside the data point at the top edge. */
    xlLabelPositionLeft       = -4131, /*!< Data label is positioned to the left of the data point. */
    xlLabelPositionMixed      =     6, /*!< Data labels are in multiple positions. */
    xlLabelPositionOutsideEnd =     2, /*!< Data label is positioned outside the data point at the top edge. */
    xlLabelPositionRight      = -4152, /*!< Data label is positioned to the right of the data point. */
};

/**
Specifies the separator used with data labels.

[Official VBA documentation for XlDataLabelSeparator](https://docs.microsoft.com/office/vba/api/excel.xldatalabelseparator)
*/
enum XlDataLabelSeparator
{
    xlDataLabelSeparatorDefault = 1, /*!< Excel selects the separator. */
};

/**
Specifies the type of data label to apply.

[Official VBA documentation for XlDataLabelsType](https://docs.microsoft.com/office/vba/api/excel.xldatalabelstype)
*/
enum XlDataLabelsType
{
    xlDataLabelsShowBubbleSizes     =     6, /*!< Show the size of the bubble in reference to the absolute value. */
    xlDataLabelsShowLabel           =     4, /*!< Category for the point. */
    xlDataLabelsShowLabelAndPercent =     5, /*!< Percentage of the total, and category for the point. Available only for pie charts and doughnut charts. */
    xlDataLabelsShowNone            = -4142, /*!< No data labels. */
    xlDataLabelsShowPercent         =     3, /*!< Percentage of the total. Available only for pie charts and doughnut charts. */
    xlDataLabelsShowValue           =     2, /*!< Default value for the point (assumed if this argument is not specified). */
};

/**
Specifies the type of date to apply to a data series.

[Official VBA documentation for XlDataSeriesDate](https://docs.microsoft.com/office/vba/api/excel.xldataseriesdate)
*/
enum XlDataSeriesDate
{
    xlDay     = 1, /*!< Day */
    xlMonth   = 3, /*!< Month */
    xlWeekday = 2, /*!< Weekdays */
    xlYear    = 4, /*!< Year */
};

/**
Specifies the data series to create.

[Official VBA documentation for XlDataSeriesType](https://docs.microsoft.com/office/vba/api/excel.xldataseriestype)
*/
enum XlDataSeriesType
{
    xlAutoFill         =     4, /*!< Fill series according to AutoFill settings. */
    xlChronological    =     3, /*!< Fill with date values. */
    xlDataSeriesLinear = -4132, /*!< Extend values, assuming an additive progression (for example, '1, 2' is extended as '3, 4, 5'). */
    xlGrowth           =     2, /*!< Extend values, assuming a multiplicative progression (for example, '1, 2' is extended as '4, 8, 16'). */
};

/**
Specifies how to shift cells to replace deleted cells.

[Official VBA documentation for XlDeleteShiftDirection](https://docs.microsoft.com/office/vba/api/excel.xldeleteshiftdirection)
*/
enum XlDeleteShiftDirection
{
    xlShiftToLeft = -4159, /*!< Cells are shifted to the left. */
    xlShiftUp     = -4162, /*!< Cells are shifted up. */
};

/**
Specifies the direction in which to move.

[Official VBA documentation for XlDirection](https://docs.microsoft.com/office/vba/api/excel.xldirection)
*/
enum XlDirection
{
    xlDown    = -4121, /*!< Down. */
    xlToLeft  = -4159, /*!< To left. */
    xlToRight = -4161, /*!< To right. */
    xlUp      = -4162, /*!< Up. */
};

/**
Specifies how blank cells are plotted on a chart.

[Official VBA documentation for XlDisplayBlanksAs](https://docs.microsoft.com/office/vba/api/excel.xldisplayblanksas)
*/
enum XlDisplayBlanksAs
{
    xlInterpolated = 3, /*!< Values are interpolated into the chart. */
    xlNotPlotted   = 1, /*!< Blank cells are not plotted. */
    xlZero         = 2, /*!< Blanks are plotted as zero. */
};

/**
Specifies how shapes are displayed.

[Official VBA documentation for XlDisplayDrawingObjects](https://docs.microsoft.com/office/vba/api/excel.xldisplaydrawingobjects)
*/
enum XlDisplayDrawingObjects
{
    xlDisplayShapes = -4104, /*!< Show all shapes. */
    xlHide          =     3, /*!< Hide all shapes. */
    xlPlaceholders  =     2, /*!< Show only placeholders. */
};

/**
Specifies the display unit label for an axis.

[Official VBA documentation for XlDisplayUnit](https://docs.microsoft.com/office/vba/api/excel.xldisplayunit)
*/
enum XlDisplayUnit
{
    xlHundredMillions  =  -8, /*!< Hundreds of millions. */
    xlHundreds         =  -2, /*!< Hundreds. */
    xlHundredThousands =  -5, /*!< Hundreds of thousands. */
    xlMillionMillions  = -10, /*!< Millions of millions. */
    xlMillions         =  -6, /*!< Millions. */
    xlTenMillions      =  -7, /*!< Tens of millions. */
    xlTenThousands     =  -4, /*!< Tens of thousands. */
    xlThousandMillions =  -9, /*!< Thousands of millions. */
    xlThousands        =  -3, /*!< Thousands. */
};

/**
Specifies whether duplicate or unique values shoud be displayed.

[Official VBA documentation for XlDupeUnique](https://docs.microsoft.com/office/vba/api/excel.xldupeunique)
*/
enum XlDupeUnique
{
    xlDuplicate = 1, /*!< Display duplicate values. */
    xlUnique    = 0, /*!< Display unique values. */
};

/**
Specifies the filter criterion.

[Official VBA documentation for XlDynamicFilterCriteria](https://docs.microsoft.com/office/vba/api/excel.xldynamicfiltercriteria)
*/
enum XlDynamicFilterCriteria
{
    xlFilterAboveAverage              = 33, /*!< Filter all above-average values. */
    xlFilterAllDatesInPeriodApril     = 24, /*!< Filter all dates in April. */
    xlFilterAllDatesInPeriodAugust    = 28, /*!< Filter all dates in August. */
    xlFilterAllDatesInPeriodDecember  = 32, /*!< Filter all dates in December. */
    xlFilterAllDatesInPeriodFebruary  = 22, /*!< Filter all dates in February. */
    xlFilterAllDatesInPeriodJanuary   = 21, /*!< Filter all dates in January. */
    xlFilterAllDatesInPeriodJuly      = 27, /*!< Filter all dates in July. */
    xlFilterAllDatesInPeriodJune      = 26, /*!< Filter all dates in June. */
    xlFilterAllDatesInPeriodMarch     = 23, /*!< Filter all dates in March. */
    xlFilterAllDatesInPeriodMay       = 25, /*!< Filter all dates in May. */
    xlFilterAllDatesInPeriodNovember  = 31, /*!< Filter all dates in November. */
    xlFilterAllDatesInPeriodOctober   = 30, /*!< Filter all dates in October. */
    xlFilterAllDatesInPeriodQuarter1  = 17, /*!< Filter all dates in Quarter1. */
    xlFilterAllDatesInPeriodQuarter2  = 18, /*!< Filter all dates in Quarter2. */
    xlFilterAllDatesInPeriodQuarter3  = 19, /*!< Filter all dates in Quarter3. */
    xlFilterAllDatesInPeriodQuarter4  = 20, /*!< Filter all dates in Quarter4. */
    xlFilterAllDatesInPeriodSeptember = 29, /*!< Filter all dates in September. */
    xlFilterBelowAverage              = 34, /*!< Filter all below-average values. */
    xlFilterLastMonth                 =  8, /*!< Filter all values related to last month. */
    xlFilterLastQuarter               = 11, /*!< Filter all values related to last quarter. */
    xlFilterLastWeek                  =  5, /*!< Filter all values related to last week. */
    xlFilterLastYear                  = 14, /*!< Filter all values related to last year. */
    xlFilterNextMonth                 =  9, /*!< Filter all values related to next month. */
    xlFilterNextQuarter               = 12, /*!< Filter all values related to next quarter. */
    xlFilterNextWeek                  =  6, /*!< Filter all values related to next week. */
    xlFilterNextYear                  = 15, /*!< Filter all values related to next year. */
    xlFilterThisMonth                 =  7, /*!< Filter all values related to the current month. */
    xlFilterThisQuarter               = 10, /*!< Filter all values related to the current quarter. */
    xlFilterThisWeek                  =  4, /*!< Filter all values related to the current week. */
    xlFilterThisYear                  = 13, /*!< Filter all values related to the current year. */
    xlFilterToday                     =  1, /*!< Filter all values related to the current date. */
    xlFilterTomorrow                  =  3, /*!< Filter all values related to tomorrow. */
    xlFilterYearToDate                = 16, /*!< Filter all values from today until a year ago. */
    xlFilterYesterday                 =  2, /*!< Filter all values related to yesterday. */
};

/**
Specifies the format of the published edition. This enumeration is only for Macintosh and should not be used.

[Official VBA documentation for XlEditionFormat](https://docs.microsoft.com/office/vba/api/excel.xleditionformat)
*/
enum XlEditionFormat
{
    xlBIFF = 2, /*!< Binary Interchange file format. */
    xlPICT = 1, /*!< Metafile picture structure (.wmf). */
    xlRTF  = 4, /*!< Rich Text Format (.rtf). */
    xlVALU = 8, /*!< VALU. */
};

/**
This enumeration is only for Macintosh and should not be used.

[Official VBA documentation for XlEditionOptionsOption](https://docs.microsoft.com/office/vba/api/excel.xleditionoptionsoption)
*/
enum XlEditionOptionsOption
{
    xlAutomaticUpdate  = 4, /*!< Automatic update. */
    xlCancel           = 1, /*!< Cancel. */
    xlChangeAttributes = 6, /*!< Change attributes. */
    xlManualUpdate     = 5, /*!< Manual update. */
    xlOpenSource       = 3, /*!< Open source. */
    xlSelect           = 3, /*!< Select. */
    xlSendPublisher    = 2, /*!< Send to Microsoft Publisher. */
    xlUpdateSubscriber = 2, /*!< Update subscriber. */
};

/**
Specifies the type of edition to be changed.

[Official VBA documentation for XlEditionType](https://docs.microsoft.com/office/vba/api/excel.xleditiontype)
*/
enum XlEditionType
{
    xlPublisher  = 1, /*!< Publisher */
    xlSubscriber = 2, /*!< Subscriber */
};

/**
Specifies how Microsoft Office Excel 2007 handles CTRL+BREAK (or ESC or COMMAND+PERIOD) user interruptions to the running procedure.

[Official VBA documentation for XlEnableCancelKey](https://docs.microsoft.com/office/vba/api/excel.xlenablecancelkey)
*/
enum XlEnableCancelKey
{
    xlDisabled     = 0, /*!< Cancel key trapping is completely disabled. */
    xlErrorHandler = 2, /*!< The interrupt is sent to the running procedure as an error, trappable by an error handler set up with an On Error GoTo statement. The trappable error code is 18. */
    xlInterrupt    = 1, /*!< The current procedure is interrupted, and the user can debug or end the procedure. */
};

/**
Specifies what can be selected on the sheet.

[Official VBA documentation for XlEnableSelection](https://docs.microsoft.com/office/vba/api/excel.xlenableselection)
*/
enum XlEnableSelection
{
    xlNoRestrictions =     0, /*!< Anything can be selected. */
    xlNoSelection    = -4142, /*!< Nothing can be selected. */
    xlUnlockedCells  =     1, /*!< Only unlocked cells can be selected. */
};

/**
Specifies the end style for error bars.

[Official VBA documentation for XlEndStyleCap](https://docs.microsoft.com/office/vba/api/excel.xlendstylecap)
*/
enum XlEndStyleCap
{
    xlCap   = 1, /*!< Caps applied. */
    xlNoCap = 2, /*!< No caps applied. */
};

/**
Specifies which axis values are to receive error bars.

[Official VBA documentation for XlErrorBarDirection](https://docs.microsoft.com/office/vba/api/excel.xlerrorbardirection)
*/
enum XlErrorBarDirection
{
    xlX = -4168, /*!< Bars run parallel to the Y axis for X-axis values. */
    xlY =     1, /*!< Bars run parallel to the X axis for Y-axis values. */
};

/**
Specifies which error-bar parts to include.

[Official VBA documentation for XlErrorBarInclude](https://docs.microsoft.com/office/vba/api/excel.xlerrorbarinclude)
*/
enum XlErrorBarInclude
{
    xlErrorBarIncludeBoth        =     1, /*!< Both positive and negative error range. */
    xlErrorBarIncludeMinusValues =     3, /*!< Only negative error range. */
    xlErrorBarIncludeNone        = -4142, /*!< No error bar range. */
    xlErrorBarIncludePlusValues  =     2, /*!< Only positive error range. */
};

/**
Specifies the range marked by error bars.

[Official VBA documentation for XlErrorBarType](https://docs.microsoft.com/office/vba/api/excel.xlerrorbartype)
*/
enum XlErrorBarType
{
    xlErrorBarTypeCustom     = -4114, /*!< Range is set by fixed values or cell values. */
    xlErrorBarTypeFixedValue =     1, /*!< Fixed-length error bars. */
    xlErrorBarTypePercent    =     2, /*!< Percentage of range to be covered by the error bars. */
    xlErrorBarTypeStDev      = -4155, /*!< Shows range for specified number of standard deviations. */
    xlErrorBarTypeStError    =     4, /*!< Shows standard error range. */
};

/**
Specifies the type of error object to be retrieved from the **Errors** collection.

[Official VBA documentation for XlErrorChecks](https://docs.microsoft.com/office/vba/api/excel.xlerrorchecks)
*/
enum XlErrorChecks
{
    xlEmptyCellReferences     = 7, /*!< The cell contains a formula referring to empty cells. */
    xlEvaluateToError         = 1, /*!< The cell evaluates to an error value. */
    xlInconsistentFormula     = 4, /*!< The cell contains an inconsistent formula for a region. */
    xlInconsistentListFormula = 9, /*!< The cell contains an inconsistent formula for a list. */
    xlListDataValidation      = 8, /*!< Data in the list contains a validation error. */
    xlNumberAsText            = 3, /*!< Number entered as text. */
    xlOmittedCells            = 5, /*!< Cells omitted. */
    xlTextDate                = 2, /*!< Date entered as text. */
    xlUnlockedFormulaCells    = 6, /*!< Formula cells are unlocked. */
};

/**
Specifies the new access mode for the object.

[Official VBA documentation for XlFileAccess](https://docs.microsoft.com/office/vba/api/excel.xlfileaccess)
*/
enum XlFileAccess
{
    xlReadOnly  = 3, /*!< Read-only */
    xlReadWrite = 2, /*!< Read/write */
};

/**
Specifies the file format when saving the worksheet.

[Official VBA documentation for XlFileFormat](https://docs.microsoft.com/office/vba/api/excel.xlfileformat)
*/
enum XlFileFormat
{
    xlAddIn                       =    18, /*!< Microsoft Excel 97-2003 Add-In */
    xlAddIn8                      =    18, /*!< Microsoft Excel 97-2003 Add-In */
    xlCSV                         =     6, /*!< CSV */
    xlCSVMac                      =    22, /*!< Macintosh CSV */
    xlCSVMSDOS                    =    24, /*!< MSDOS CSV */
    xlCSVUTF8                     =    62, /*!< UTF8 CSV */
    xlCSVWindows                  =    23, /*!< Windows CSV */
    xlCurrentPlatformText         = -4158, /*!< Current Platform Text */
    xlDBF2                        =     7, /*!< Dbase 2 format */
    xlDBF3                        =     8, /*!< Dbase 3 format */
    xlDBF4                        =    11, /*!< Dbase 4 format */
    xlDIF                         =     9, /*!< Data Interchange format */
    xlExcel12                     =    50, /*!< Excel Binary Workbook */
    xlExcel2                      =    16, /*!< Excel version 2.0 (1987) */
    xlExcel2FarEast               =    27, /*!< Excel version 2.0 Asia (1987) */
    xlExcel3                      =    29, /*!< Excel version 3.0 (1990) */
    xlExcel4                      =    33, /*!< Excel version 4.0 (1992) */
    xlExcel4Workbook              =    35, /*!< Excel version 4.0. Workbook format (1992) */
    xlExcel5                      =    39, /*!< Excel version 5.0 (1994) */
    xlExcel7                      =    39, /*!< Excel 95 (version 7.0) */
    xlExcel8                      =    56, /*!< Excel 97-2003 Workbook */
    xlExcel9795                   =    43, /*!< Excel version 95 and 97 */
    xlHtml                        =    44, /*!< HTML format */
    xlIntlAddIn                   =    26, /*!< International Add-In */
    xlIntlMacro                   =    25, /*!< International Macro */
    xlOpenDocumentSpreadsheet     =    60, /*!< OpenDocument Spreadsheet */
    xlOpenXMLAddIn                =    55, /*!< Open XML Add-In */
    xlOpenXMLStrictWorkbook       =    61, /*!< Strict Open XML file */
    xlOpenXMLTemplate             =    54, /*!< Open XML Template */
    xlOpenXMLTemplateMacroEnabled =    53, /*!< Open XML Template Macro Enabled */
    xlOpenXMLWorkbook             =    51, /*!< Open XML Workbook */
    xlOpenXMLWorkbookMacroEnabled =    52, /*!< Open XML Workbook Macro Enabled */
    xlSYLK                        =     2, /*!< Symbolic Link format */
    xlTemplate                    =    17, /*!< Excel Template format */
    xlTemplate8                   =    17, /*!< Template 8 */
    xlTextMac                     =    19, /*!< Macintosh Text */
    xlTextMSDOS                   =    21, /*!< MSDOS Text */
    xlTextPrinter                 =    36, /*!< Printer Text */
    xlTextWindows                 =    20, /*!< Windows Text */
    xlUnicodeText                 =    42, /*!< Unicode Text */
    xlWebArchive                  =    45, /*!< Web Archive */
    xlWJ2WD1                      =    14, /*!< Japanese 1-2-3 */
    xlWJ3                         =    40, /*!< Japanese 1-2-3 */
    xlWJ3FJ3                      =    41, /*!< Japanese 1-2-3 format */
    xlWK1                         =     5, /*!< Lotus 1-2-3 format */
    xlWK1ALL                      =    31, /*!< Lotus 1-2-3 format */
    xlWK1FMT                      =    30, /*!< Lotus 1-2-3 format */
    xlWK3                         =    15, /*!< Lotus 1-2-3 format */
    xlWK3FM3                      =    32, /*!< Lotus 1-2-3 format */
    xlWK4                         =    38, /*!< Lotus 1-2-3 format */
    xlWKS                         =     4, /*!< Lotus 1-2-3 format */
    xlWorkbookDefault             =    51, /*!< Workbook default */
    xlWorkbookNormal              = -4143, /*!< Workbook normal */
    xlWorks2FarEast               =    28, /*!< Microsoft Works 2.0 Asian format */
    xlWQ1                         =    34, /*!< Quattro Pro format */
    xlXMLSpreadsheet              =    46, /*!< XML Spreadsheet */
};

/**
Specifies how to validate the data caches for PivotTable reports.

[Official VBA documentation for XlFileValidationPivotMode](https://docs.microsoft.com/office/vba/api/excel.xlfilevalidationpivotmode)
*/
enum XlFileValidationPivotMode
{
    xlFileValidationPivotDefault = 0, /*!< Validate the contents of data caches as specified by the **PivotOptions** registry setting (default). */
    xlFileValidationPivotRun     = 1, /*!< Validate the contents of all data caches regardless of the registry setting. */
    xlFileValidationPivotSkip    = 2, /*!< Do not validate the contents of data caches. */
};

/**
Specifies how to copy the range.

[Official VBA documentation for XlFillWith](https://docs.microsoft.com/office/vba/api/excel.xlfillwith)
*/
enum XlFillWith
{
    xlFillWithAll      = -4104, /*!< Copy contents and formats. */
    xlFillWithContents =     2, /*!< Copy contents only. */
    xlFillWithFormats  = -4122, /*!< Copy formats only. */
};

/**
Specifies whether data is to be copied or left in place during a filter operation.

[Official VBA documentation for XlFilterAction](https://docs.microsoft.com/office/vba/api/excel.xlfilteraction)
*/
enum XlFilterAction
{
    xlFilterCopy    = 2, /*!< Copy filtered data to new location. */
    xlFilterInPlace = 1, /*!< Leave data in place. */
};

/**
Specifies how dates should be filtered in the specified period.

[Official VBA documentation for XlFilterAllDatesInPeriod](https://docs.microsoft.com/office/vba/api/excel.xlfilteralldatesinperiod)
*/
enum XlFilterAllDatesInPeriod
{
    xlFilterAllDatesInPeriodDay    = 2, /*!< Filter all dates for the specified date. */
    xlFilterAllDatesInPeriodHour   = 3, /*!< Filter all dates for the specified hour. */
    xlFilterAllDatesInPeriodMinute = 4, /*!< Filter all dates until the specified minute. */
    xlFilterAllDatesInPeriodMonth  = 1, /*!< Filter all dates for the specified month. */
    xlFilterAllDatesInPeriodSecond = 5, /*!< Filter all dates until the specified second. */
    xlFilterAllDatesInPeriodYear   = 0, /*!< Filter all dates for the specified year. */
};

/**
Used to return a status from filter functions.

[Official VBA documentation for XlFilterStatus](https://docs.microsoft.com/office/vba/api/excel.xlfilterstatus)
*/
enum XlFilterStatus
{
    xlFilterStatusOK             = 0, /*!< Signifies OK or successful. */
    xlFilterStatusDateWrongOrder = 1, /*!< SetFilterDateRange(?): StartDate > EndDate */
    xlFilterStatusDateHasTime    = 2, /*!< SetFilterDateRange(?): StartDate or EndDate have a time portion. */
    xlFilterStatusInvalidDate    = 3, /*!< SetFilterDateRange(?): StartDate or EndDate are not valid dates. */
};

/**
Specifies the type of data to search.

[Official VBA documentation for XlFindLookIn](https://docs.microsoft.com/office/vba/api/excel.xlfindlookin)
*/
enum XlFindLookIn
{
    xlComments         = -4144, /*!< Comments */
    xlCommentsThreaded = -4184, /*!< Threaded comments */
    xlFormulas         = -4123, /*!< Formulas */
    xlValues           = -4163, /*!< Values */
};

/**
Specifies the quality of spreadsheets saved in different fixed formats.

[Official VBA documentation for XlFixedFormatQuality](https://docs.microsoft.com/office/vba/api/excel.xlfixedformatquality)
*/
enum XlFixedFormatQuality
{
    xlQualityMinimum  = 1, /*!< Minimum quality */
    xlQualityStandard = 0, /*!< Standard quality */
};

/**
Specifies the type of file format.

[Official VBA documentation for XlFixedFormatType](https://docs.microsoft.com/office/vba/api/excel.xlfixedformattype)
*/
enum XlFixedFormatType
{
    xlTypePDF = 0, /*!< "PDF" - Portable Document Format file (.pdf) */
    xlTypeXPS = 1, /*!< "XPS" - XPS Document (.xps) */
};

/**
Constants passed to various **WorksheetFunction** and **Workbook** statistical forecasting methods.

[Official VBA documentation for XlForecastAggregation](https://docs.microsoft.com/office/vba/api/excel.xlforecastaggregation)
*/
enum XlForecastAggregation
{
    xlForecastAggregationAveragexlForecastAggregationCountxlForecastAggregationCountAxlForecastAggregationMaxxlForecastAggregationMedianxlForecastAggregationMinimumxlForecastAggregationSum = 1234567, /*!< Average aggregationCount aggregationCountA aggregationMaximum aggregationMedian aggregationMinimum aggregationSum aggregation */
    xlForecastAggregationAverage                                                                                                                                                             =       1,
    xlForecastAggregationCount                                                                                                                                                               =       2,
    xlForecastAggregationCountA                                                                                                                                                              =       3,
    xlForecastAggregationMax                                                                                                                                                                 =       4,
    xlForecastAggregationMedian                                                                                                                                                              =       5,
    xlForecastAggregationMin                                                                                                                                                                 =       6,
    xlForecastAggregationSum                                                                                                                                                                 =       7,
};

/**
Constants passed to the [Workbook.CreateForecastSheet Method](Excel.workbook.createforecastsheet.md).

[Official VBA documentation for XlForecastChartType](https://docs.microsoft.com/office/vba/api/excel.xlforecastcharttype)
*/
enum XlForecastChartType
{
    xlForecastChartTypeColumnxlForecastChartTypeLine = 10, /*!< Column chartLine chart */
    xlForecastChartTypeColumn                        =  1,
    xlForecastChartTypeLine                          =  0,
};

/**
Constants passed to various **WorksheetFunction** and **Workbook** statistical forecasting methods.

[Official VBA documentation for XlForecastDataCompletion](https://docs.microsoft.com/office/vba/api/excel.xlforecastdatacompletion)
*/
enum XlForecastDataCompletion
{
    xlForecastDataCompletionInterpolatexlForecastDataCompletionZeros = 10, /*!< Data completion by interpolationData completion by zeroes */
    xlForecastDataCompletionInterpolate                              =  1,
    xlForecastDataCompletionZeros                                    =  0,
};

/**
Specifies the type of the form control.

[Official VBA documentation for XlFormControl](https://docs.microsoft.com/office/vba/api/excel.xlformcontrol)
*/
enum XlFormControl
{
    xlButtonControl = 0, /*!< Button. */
    xlCheckBox      = 1, /*!< Check box. */
    xlDropDown      = 2, /*!< Combo box. */
    xlEditBox       = 3, /*!< Text box. */
    xlGroupBox      = 4, /*!< Group box. */
    xlLabel         = 5, /*!< Label. */
    xlListBox       = 6, /*!< List box. */
    xlOptionButton  = 7, /*!< Option button. */
    xlScrollBar     = 8, /*!< Scroll bar. */
    xlSpinner       = 9, /*!< Spinner. */
};

/**
Specifies the operator to use to compare a formula against the value in a cell or, for **xlBetween** and **xlNotBetween**, to compare two formulas.

[Official VBA documentation for XlFormatConditionOperator](https://docs.microsoft.com/office/vba/api/excel.xlformatconditionoperator)
*/
enum XlFormatConditionOperator
{
    xlBetween      = 1, /*!< Between. Can be used only if two formulas are provided. */
    xlEqual        = 3, /*!< Equal. */
    xlGreater      = 5, /*!< Greater than. */
    xlGreaterEqual = 7, /*!< Greater than or equal to. */
    xlLess         = 6, /*!< Less than. */
    xlLessEqual    = 8, /*!< Less than or equal to. */
    xlNotBetween   = 2, /*!< Not between. Can be used only if two formulas are provided. */
    xlNotEqual     = 4, /*!< Not equal. */
};

/**
Specifies whether the conditional format is based on a cell value or an expression.

[Official VBA documentation for XlFormatConditionType](https://docs.microsoft.com/office/vba/api/excel.xlformatconditiontype)
*/
enum XlFormatConditionType
{
    xlAboveAverageCondition = 12, /*!< Above average condition */
    xlBlanksCondition       = 10, /*!< Blanks condition */
    xlCellValue             =  1, /*!< Cell value */
    xlColorScale            =  3, /*!< Color scale */
    xlDataBar               =  4, /*!< DataBar */
    xlErrorsCondition       = 16, /*!< Errors condition */
    xlExpression            =  2, /*!< Expression */
    xlIconSet               =  6, /*!< Icon set */
    xlNoBlanksCondition     = 13, /*!< No blanks condition */
    xlNoErrorsCondition     = 17, /*!< No errors condition */
    xlTextString            =  9, /*!< Text string */
    xlTimePeriod            = 11, /*!< Time period */
    xlTop10                 =  5, /*!< Top 10 values */
    xlUniqueValues          =  8, /*!< Unique values */
};

/**
Specifies the types of format filters.

[Official VBA documentation for XlFormatFilterTypes](https://docs.microsoft.com/office/vba/api/excel.xlformatfiltertypes)
*/
enum XlFormatFilterTypes
{
    FilterBottom        = 0, /*!< Bottom. */
    FilterBottomPercent = 2, /*!< Bottom Percent. */
    FilterTop           = 1, /*!< Top. */
    FilterTopPercent    = 3, /*!< Top Percent. */
};

/**
Specifies the formula label type for the specified range.

[Official VBA documentation for XlFormulaLabel](https://docs.microsoft.com/office/vba/api/excel.xlformulalabel)
*/
enum XlFormulaLabel
{
    xlColumnLabels =     2, /*!< Column labels only. */
    xlMixedLabels  =     3, /*!< Row and column labels. */
    xlNoLabels     = -4142, /*!< No labels. */
    xlRowLabels    =     1, /*!< Row labels only. */
};

/**
Specifies the type of table references.

[Official VBA documentation for XlGenerateTableRefs](https://docs.microsoft.com/office/vba/api/excel.xlgeneratetablerefs)
*/
enum XlGenerateTableRefs
{
    xlA1TableRefs = 0, /*!< A1 Table References. */
    xlTableNames  = 1, /*!< Table Names. */
};

/**
Constants passed to and returned by the **Series.GeoMappingLevel** property.

[Official VBA documentation for XlGeoMappingLevel](https://docs.microsoft.com/office/vba/api/excel.xlgeomappinglevel)
*/
enum XlGeoMappingLevel
{
    xlGeoMappingLevelAutomatic         = 0, /*!< Use highest resolution mapping level. */
    xlGeoMappingLevelDataOnly          = 1, /*!< Only map regions with data. */
    xlGeoMappingLevelPostalCode        = 2, /*!< Map by postcode. */
    xlGeoMappingLevelCounty            = 3, /*!< Map by county. */
    xlGeoMappingLevelState             = 4, /*!< Map by state. */
    xlGeoMappingLevelCountryRegion     = 5, /*!< Map by country. */
    xlGeoMappingLevelCountryRegionList = 6, /*!< Map by region (group of countries). */
    xlGeoMappingLevelWorld             = 7, /*!< Map the whole world. */
};

/**
Constants passed to and returned by the **Series.GeoProjectionType** property.

[Official VBA documentation for XlGeoProjectionType](https://docs.microsoft.com/office/vba/api/excel.xlgeoprojectiontype)
*/
enum XlGeoProjectionType
{
    xlGeoProjectionTypeAutomatic = 0, /*!< Heuristically choose best map projection. */
    xlGeoProjectionTypeMercator  = 1, /*!< Use the Mercator map projection. */
    xlGeoProjectionTypeMiller    = 2, /*!< Use the Miller map projection. */
    xlGeoProjectionTypeAlbers    = 3, /*!< Use the Albers map projection. */
    xlGeoProjectionTypeRobinson  = 4, /*!< Use the Robinson map projection. */
};

/**
Specifies the type of **gradient fill**.

[Official VBA documentation for XlGradientFillType](https://docs.microsoft.com/office/vba/api/excel.xlgradientfilltype)
*/
enum XlGradientFillType
{
    GradientFillLinear = 0, /*!< Gradient is filled in a straight line. */
    GradientFillPath   = 1, /*!< Gradient is filled in a non-linear or curved path. */
};

/**
Specifies the horizontal alignment for the object.

[Official VBA documentation for XlHAlign](https://docs.microsoft.com/office/vba/api/excel.xlhalign)
*/
enum XlHAlign
{
    xlHAlignCenter                = -4108, /*!< Center. */
    xlHAlignCenterAcrossSelection =     7, /*!< Center across selection. */
    xlHAlignDistributed           = -4117, /*!< Distribute. */
    xlHAlignFill                  =     5, /*!< Fill. */
    xlHAlignGeneral               =     1, /*!< Align according to data type. */
    xlHAlignJustify               = -4130, /*!< Justify. */
    xlHAlignLeft                  = -4131, /*!< Left. */
    xlHAlignRight                 = -4152, /*!< Right. */
};

/**
Specifies the mode for the Hebrew spelling checker.

[Official VBA documentation for XlHebrewModes](https://docs.microsoft.com/office/vba/api/excel.xlhebrewmodes)
*/
enum XlHebrewModes
{
    xlHebrewFullScript            = 0, /*!< The conventional script type as required by the Hebrew Language Academy when writing text without diacritics. */
    xlHebrewMixedAuthorizedScript = 3, /*!< The Hebrew traditional script. */
    xlHebrewMixedScript           = 2, /*!< In this mode the speller accepts any word recognized as Hebrew, whether in Full Script, Partial Script, or any unconventional spelling variation that is known to the speller. */
    xlHebrewPartialScript         = 1, /*!< In this mode the speller accepts words both in Full Script and Partial Script. Some words will be flagged since this spelling is not authorized in either Full script or Partial script. */
};

/**
Specifies which set of changes is shown in a shared workbook.

[Official VBA documentation for XlHighlightChangesTime](https://docs.microsoft.com/office/vba/api/excel.xlhighlightchangestime)
*/
enum XlHighlightChangesTime
{
    xlAllChanges      = 2, /*!< Show all changes. */
    xlNotYetReviewed  = 3, /*!< Show only changes not yet reviewed. */
    xlSinceMyLastSave = 1, /*!< Show changes made since last save by last user. */
};

/**
Specifies the description of the Japanese input rules.

[Official VBA documentation for XlIMEMode](https://docs.microsoft.com/office/vba/api/excel.xlimemode)
*/
enum XlIMEMode
{
    xlIMEModeAlpha        =  8, /*!< Half-width alphanumeric. */
    xlIMEModeAlphaFull    =  7, /*!< Full-width alphanumeric. */
    xlIMEModeDisable      =  3, /*!< Disable. */
    xlIMEModeHangul       = 10, /*!< Hangul. */
    xlIMEModeHangulFull   =  9, /*!< Full-width Hangul. */
    xlIMEModeHiragana     =  4, /*!< Hiragana. */
    xlIMEModeKatakana     =  5, /*!< Katakana. */
    xlIMEModeKatakanaHalf =  6, /*!< Half-width Katakana. */
    xlIMEModeNoControl    =  0, /*!< No control. */
    xlIMEModeOff          =  2, /*!< Off (English mode). */
    xlIMEModeOn           =  1, /*!< Mode on. */
};

/**
Specifies the icon for a criterion in an icon set conditional formatting rule.

[Official VBA documentation for XlIcon](https://docs.microsoft.com/office/vba/api/excel.xlicon)
*/
enum XlIcon
{
    xlIcon0Bars                        = 37, /*!< **Signal Meter With No Filled Bars** */
    xlIcon0FilledBoxes                 = 52, /*!< **0 Filled Boxes** */
    xlIcon1Bar                         = 38, /*!< **Signal Meter With One Filled Bar** */
    xlIcon1FilledBox                   = 51, /*!< **1 Filled Boxes** */
    xlIcon2Bars                        = 39, /*!< **Signal Meter With Two Filled Bars** */
    xlIcon2FilledBoxes                 = 50, /*!< **2 Filled Boxes** */
    xlIcon3Bars                        = 40, /*!< **Signal Meter With Three Filled Bars** */
    xlIcon3FilledBoxes                 = 49, /*!< **3 Filled Boxes** */
    xlIcon4Bars                        = 41, /*!< **Signal Meter With Four Filled Bars** */
    xlIcon4FilledBoxes                 = 48, /*!< **4 Filled Boxes** */
    xlIconBlackCircle                  = 32, /*!< **Black Circle** */
    xlIconBlackCircleWithBorder        = 13, /*!< **Black Circle With Border** */
    xlIconCircleWithOneWhiteQuarter    = 33, /*!< **Circle With One White Quarter** */
    xlIconCircleWithThreeWhiteQuarters = 35, /*!< **Circle With Three White Quarters** */
    xlIconCircleWithTwoWhiteQuarters   = 34, /*!< **Circle With Two White Quarters** */
    xlIconGoldStar                     = 42, /*!< **Gold Star** */
    xlIconGrayCircle                   = 31, /*!< **Gray Circle** */
    xlIconGrayDownArrow                =  6, /*!< **Gray Down Arrow** */
    xlIconGrayDownInclineArrow         = 28, /*!< **Gray Down Incline Arrow** */
    xlIconGraySideArrow                =  5, /*!< **Gray Side Arrow** */
    xlIconGrayUpArrow                  =  4, /*!< **Gray Up Arrow** */
    xlIconGrayUpInclineArrow           = 27, /*!< **Gray Up Incline Arrow** */
    xlIconGreenCheck                   = 22, /*!< **Green Check** */
    xlIconGreenCheckSymbol             = 19, /*!< **Green Check Symbol** */
    xlIconGreenCircle                  = 10, /*!< **Green Circle** */
    xlIconGreenFlag                    =  7, /*!< **Green Flag** */
    xlIconGreenTrafficLight            = 14, /*!< **Green Traffic Light** */
    xlIconGreenUpArrow                 =  1, /*!< **Green Up Arrow** */
    xlIconGreenUpTriangle              = 45, /*!< **Green Up Triangle** */
    xlIconHalfGoldStar                 = 43, /*!< **Half Gold Star** */
    xlIconNoCellIcon                   = -1, /*!< **No Cell Icon** */
    xlIconPinkCircle                   = 30, /*!< **Pink Circle** */
    xlIconRedCircle                    = 29, /*!< **Red Circle** */
    xlIconRedCircleWithBorder          = 12, /*!< **Red Circle With Border** */
    xlIconRedCross                     = 24, /*!< **Red Cross** */
    xlIconRedCrossSymbol               = 21, /*!< **Red Cross Symbol** */
    xlIconRedDiamond                   = 18, /*!< **Red Diamond** */
    xlIconRedDownArrow                 =  3, /*!< **Red Down Arrow** */
    xlIconRedDownTriangle              = 47, /*!< **Red Down Triangle** */
    xlIconRedFlag                      =  9, /*!< **Red Flag** */
    xlIconRedTrafficLight              = 16, /*!< **Red Traffic Light** */
    xlIconSilverStar                   = 44, /*!< **Silver Star** */
    xlIconWhiteCircleAllWhiteQuarters  = 36, /*!< **White Circle (All White Quarters)** */
    xlIconYellowCircle                 = 11, /*!< **Yellow Circle** */
    xlIconYellowDash                   = 46, /*!< **Yellow Dash** */
    xlIconYellowDownInclineArrow       = 26, /*!< **Yellow Down Incline Arrow** */
    xlIconYellowExclamation            = 23, /*!< **Yellow Exclamation** */
    xlIconYellowExclamationSymbol      = 20, /*!< **Yellow Exclamation Symbol** */
    xlIconYellowFlag                   =  8, /*!< **Yellow Flag** */
    xlIconYellowSideArrow              =  2, /*!< **Yellow Side Arrow** */
    xlIconYellowTrafficLight           = 15, /*!< **Yellow Traffic Light** */
    xlIconYellowTriangle               = 17, /*!< **Yellow Triangle** */
    xlIconYellowUpInclineArrow         = 25, /*!< **Yellow Up Incline Arrow** */
};

/**
Specifies the type of icon set.

[Official VBA documentation for XlIconSet](https://docs.microsoft.com/office/vba/api/excel.xliconset)
*/
enum XlIconSet
{
    xl3Arrows         =  1, /*!< 3 Arrows */
    xl3ArrowsGray     =  2, /*!< 3 Arrows Gray */
    xl3Flags          =  3, /*!< 3 Flags */
    xl3Signs          =  6, /*!< 3 Signs */
    xl3Symbols        =  7, /*!< 3 Symbols */
    xl3TrafficLights1 =  4, /*!< 3 Traffic Lights 1 */
    xl3TrafficLights2 =  5, /*!< 3 Traffic Lights 2 */
    xl4Arrows         =  8, /*!< 4 Arrows */
    xl4ArrowsGray     =  9, /*!< 4 Arrows Gray */
    xl4CRV            = 11, /*!< 4 CRV */
    xl4RedToBlack     = 10, /*!< 4 Red To Black */
    xl4TrafficLights  = 12, /*!< 4 Traffic Lights */
    xl5Arrows         = 13, /*!< 5 Arrows */
    xl5ArrowsGray     = 14, /*!< 5 Arrows Gray */
    xl5CRV            = 15, /*!< 5 CRV */
    xl5Quarters       = 16, /*!< 5 Quarters */
};

/**
Specifies the format in which to return data from a database.

[Official VBA documentation for XlImportDataAs](https://docs.microsoft.com/office/vba/api/excel.xlimportdataas)
*/
enum XlImportDataAs
{
    xlPivotTableReport = 1, /*!< Returns the data as a PivotTable. */
    xlQueryTable       = 0, /*!< Returns the data as a QueryTable. */
};

/**
Specifies from where to copy the format for inserted cells.

[Official VBA documentation for XlInsertFormatOrigin](https://docs.microsoft.com/office/vba/api/excel.xlinsertformatorigin)
*/
enum XlInsertFormatOrigin
{
    xlFormatFromLeftOrAbove  = 0, /*!< Copy the format from cells above and/or to the left. */
    xlFormatFromRightOrBelow = 1, /*!< Copy the format from cells below and/or to the right. */
};

/**
Specifies the direction in which to shift cells during an insertion.

[Official VBA documentation for XlInsertShiftDirection](https://docs.microsoft.com/office/vba/api/excel.xlinsertshiftdirection)
*/
enum XlInsertShiftDirection
{
    xlShiftDown    = -4121, /*!< Shift cells down. */
    xlShiftToRight = -4161, /*!< Shift cells to the right. */
};

/**
Specifies the way the specified PivotTable items appear&mdash;in table format or in outline format.

[Official VBA documentation for XlLayoutFormType](https://docs.microsoft.com/office/vba/api/excel.xllayoutformtype)
*/
enum XlLayoutFormType
{
    xlOutline = 1, /*!< The **[LayoutSubtotalLocation](Excel.CubeField.LayoutSubtotalLocation.md)** property specifies where the subtotal appears in the PivotTable report. */
    xlTabular = 0, /*!< Default. */
};

/**
Specifies the type of layout row.

[Official VBA documentation for XlLayoutRowType](https://docs.microsoft.com/office/vba/api/excel.xllayoutrowtype)
*/
enum XlLayoutRowType
{
    xlCompactRow = 0, /*!< Compact Row */
    xlOutlineRow = 2, /*!< Outline Row */
    xlTabularRow = 1, /*!< Tabular Row */
};

/**
Specifies the position of the legend on a chart.

[Official VBA documentation for XlLegendPosition](https://docs.microsoft.com/office/vba/api/excel.xllegendposition)
*/
enum XlLegendPosition
{
    xlLegendPositionBottom = -4107, /*!< Below the chart. */
    xlLegendPositionCorner =     2, /*!< In the upper-right corner of the chart border. */
    xlLegendPositionCustom = -4161, /*!< A custom position. */
    xlLegendPositionLeft   = -4131, /*!< Left of the chart. */
    xlLegendPositionRight  = -4152, /*!< Right of the chart. */
    xlLegendPositionTop    = -4160, /*!< Above the chart. */
};

/**
Specifies the line style for the border.

[Official VBA documentation for XlLineStyle](https://docs.microsoft.com/office/vba/api/excel.xllinestyle)
*/
enum XlLineStyle
{
    xlContinuous    =     1, /*!< Continuous line. */
    xlDash          = -4115, /*!< Dashed line. */
    xlDashDot       =     4, /*!< Alternating dashes and dots. */
    xlDashDotDot    =     5, /*!< Dash followed by two dots. */
    xlDot           = -4118, /*!< Dotted line. */
    xlDouble        = -4119, /*!< Double line. */
    xlLineStyleNone = -4142, /*!< No line. */
    xlSlantDashDot  =    13, /*!< Slanted dashes. */
};

/**
Specifies the type of link.

[Official VBA documentation for XlLink](https://docs.microsoft.com/office/vba/api/excel.xllink)
*/
enum XlLink
{
    xlExcelLinks  = 1, /*!< The link is to an Excel worksheet. */
    xlOLELinks    = 2, /*!< The link is to an OLE source. */
    xlPublishers  = 5, /*!< Macintosh only. */
    xlSubscribers = 6, /*!< Macintosh only. */
};

/**
Specifies the type of information the link will return.

[Official VBA documentation for XlLinkInfo](https://docs.microsoft.com/office/vba/api/excel.xllinkinfo)
*/
enum XlLinkInfo
{
    xlEditionDate    = 2, /*!< Applies only to editions in the Macintosh operating system. */
    xlLinkInfoStatus = 3, /*!< Returns the link status. */
    xlUpdateState    = 1, /*!< Specifies whether the link updates automatically or manually. */
};

/**
Specifies the type of link.

[Official VBA documentation for XlLinkInfoType](https://docs.microsoft.com/office/vba/api/excel.xllinkinfotype)
*/
enum XlLinkInfoType
{
    xlLinkInfoOLELinks    = 2, /*!< OLE or DDE server */
    xlLinkInfoPublishers  = 5, /*!< Publisher */
    xlLinkInfoSubscribers = 6, /*!< Subscriber */
};

/**
Specifies the status of a link.

[Official VBA documentation for XlLinkStatus](https://docs.microsoft.com/office/vba/api/excel.xllinkstatus)
*/
enum XlLinkStatus
{
    xlLinkStatusCopiedValues        = 10, /*!< Copied values. */
    xlLinkStatusIndeterminate       =  5, /*!< Unable to determine status. */
    xlLinkStatusInvalidName         =  7, /*!< Invalid name. */
    xlLinkStatusMissingFile         =  1, /*!< File missing. */
    xlLinkStatusMissingSheet        =  2, /*!< Sheet missing. */
    xlLinkStatusNotStarted          =  6, /*!< Not started. */
    xlLinkStatusOK                  =  0, /*!< No errors. */
    xlLinkStatusOld                 =  3, /*!< Status may be out of date. */
    xlLinkStatusSourceNotCalculated =  4, /*!< Not yet calculated. */
    xlLinkStatusSourceNotOpen       =  8, /*!< Not open. */
    xlLinkStatusSourceOpen          =  9, /*!< Source document is open. */
};

/**
Specifies the type of link.

[Official VBA documentation for XlLinkType](https://docs.microsoft.com/office/vba/api/excel.xllinktype)
*/
enum XlLinkType
{
    xlLinkTypeExcelLinks = 1, /*!< A link to a Microsoft Excel source. */
    xlLinkTypeOLELinks   = 2, /*!< A link to an OLE source. */
};

/**
Indicates the state of cells that may contain Linked data types such as [Stocks or Geography](https://support.office.com/article/stock-quotes-and-geographic-data-61a33056-9935-484f-8ac8-f1a89e210877). These are the possible values of the [Range.LinkedDataTypeState](Excel.Range.LinkedDataTypeState.md) property.

[Official VBA documentation for XlLinkedDataTypeState](https://docs.microsoft.com/office/vba/api/excel.xllinkeddatatypestate)
*/
enum XlLinkedDataTypeState
{
    xlLinkedDataTypeStateNone                 = 0, /*!< The cell does not contain any Linked data types. */
    xlLinkedDataTypeStateValidLinkedData      = 1, /*!< The cell contains a Linked data type. */
    xlLinkedDataTypeStateDisambiguationNeeded = 2, /*!< The cell needs to be disambiguated by the user before a Linked data type can be inserted. For example, if the user types "New York" into a cell and attempts to convert it to a "Geography" data type, they may need to select whether they meant New York State or New York City. Until they do so, the cell will be in this state. */
    xlLinkedDataTypeStateBrokenLinkedData     = 3, /*!< There is a valid Linked data type in the cell, but entity no longer exists on the service. */
    xlLinkedDataTypeStateFetchingData         = 4, /*!< The Linked data type in the cell is in the middle of refreshing new data from the service. */
};

/**
Specifies the conflict resolution options for updating a list on a Microsoft SharePoint Foundation site with the changes made to a list in a Microsoft Excel worksheet.

[Official VBA documentation for XlListConflict](https://docs.microsoft.com/office/vba/api/excel.xllistconflict)
*/
enum XlListConflict
{
    xlListConflictDialog              = 0, /*!< Display a dialog box that allows the user to choose how to resolve conflicts. */
    xlListConflictDiscardAllConflicts = 2, /*!< Accept the version of the data stored on the SharePoint site. */
    xlListConflictError               = 3, /*!< Raise an error if a conflict occurs. */
    xlListConflictRetryAllConflicts   = 1, /*!< Overwrite the version of the data stored on the SharePoint site. */
};

/**
Specifies the data type of a list column connected to a Microsoft SharePoint Foundation site.

[Official VBA documentation for XlListDataType](https://docs.microsoft.com/office/vba/api/excel.xllistdatatype)
*/
enum XlListDataType
{
    xlListDataTypeCheckbox          =  9, /*!< Check box. */
    xlListDataTypeChoice            =  6, /*!< Single-choice field. */
    xlListDataTypeChoiceMulti       =  7, /*!< Multiple-choice field. */
    xlListDataTypeCounter           = 11, /*!< Counter. */
    xlListDataTypeCurrency          =  4, /*!< Currency. */
    xlListDataTypeDateTime          =  5, /*!< Date/time. */
    xlListDataTypeHyperLink         = 10, /*!< Hyperlink. */
    xlListDataTypeListLookup        =  8, /*!< Lookup list. */
    xlListDataTypeMultiLineRichText = 12, /*!< Rich text format with multiple lines. */
    xlListDataTypeMultiLineText     =  2, /*!< Plain text with multiple lines. */
    xlListDataTypeNone              =  0, /*!< Type not specified. */
    xlListDataTypeNumber            =  3, /*!< Numerical. */
    xlListDataTypeText              =  1, /*!< Plain text. */
};

/**
Specifies the current source of the list.

[Official VBA documentation for XlListObjectSourceType](https://docs.microsoft.com/office/vba/api/excel.xllistobjectsourcetype)
*/
enum XlListObjectSourceType
{
    xlSrcExternal = 0, /*!< External data source (Microsoft SharePoint Foundation site). */
    xlSrcModel    = 4, /*!< PowerPivot Model */
    xlSrcQuery    = 3, /*!< Query */
    xlSrcRange    = 1, /*!< Range */
    xlSrcXml      = 2, /*!< XML */
};

/**
Specifies the part of the PivotTable report that contains the upper-left corner of a range.

[Official VBA documentation for XlLocationInTable](https://docs.microsoft.com/office/vba/api/excel.xllocationintable)
*/
enum XlLocationInTable
{
    xlColumnHeader = -4110, /*!< Column header */
    xlColumnItem   =     5, /*!< Column item */
    xlDataHeader   =     3, /*!< Data header */
    xlDataItem     =     7, /*!< Data item */
    xlPageHeader   =     2, /*!< Page header */
    xlPageItem     =     6, /*!< Page item */
    xlRowHeader    = -4153, /*!< Row header */
    xlRowItem      =     4, /*!< Row item */
    xlTableBody    =     8, /*!< Table body */
};

/**
Specifies whether a match is made against the whole of the search text or any part of the search text.

[Official VBA documentation for XlLookAt](https://docs.microsoft.com/office/vba/api/excel.xllookat)
*/
enum XlLookAt
{
    xlPart  = 2, /*!< Match against any part of the search text. */
    xlWhole = 1, /*!< Match against the whole of the search text. */
};

/**
Specifies what to look for in searches.

[Official VBA documentation for XlLookFor](https://docs.microsoft.com/office/vba/api/excel.xllookfor)
*/
enum XlLookFor
{
    LookForBlanks   = 0, /*!< Blanks */
    LookForErrors   = 1, /*!< Errors */
    LookForFormulas = 2, /*!< Formulas */
};

/**
Specifies a Microsoft application.

[Official VBA documentation for XlMSApplication](https://docs.microsoft.com/office/vba/api/excel.xlmsapplication)
*/
enum XlMSApplication
{
    xlMicrosoftAccess       = 4, /*!< Microsoft Office Access */
    xlMicrosoftFoxPro       = 5, /*!< Microsoft FoxPro */
    xlMicrosoftMail         = 3, /*!< Microsoft Office Outlook */
    xlMicrosoftPowerPoint   = 2, /*!< Microsoft Office PowerPoint */
    xlMicrosoftProject      = 6, /*!< Microsoft Office Project */
    xlMicrosoftSchedulePlus = 7, /*!< Microsoft Schedule Plus */
    xlMicrosoftWord         = 1, /*!< Microsoft Office Word */
};

/**
Specifies the mail system that is installed on the host computer.

[Official VBA documentation for XlMailSystem](https://docs.microsoft.com/office/vba/api/excel.xlmailsystem)
*/
enum XlMailSystem
{
    xlMAPI         = 1, /*!< MAPI-complaint system */
    xlNoMailSystem = 0, /*!< No mail system */
    xlPowerTalk    = 2, /*!< PowerTalk mail system */
};

/**
Specifies the marker style for a point or series in a line chart, scatter chart, or radar chart.

[Official VBA documentation for XlMarkerStyle](https://docs.microsoft.com/office/vba/api/excel.xlmarkerstyle)
*/
enum XlMarkerStyle
{
    xlMarkerStyleAutomatic = -4105, /*!< Automatic markers */
    xlMarkerStyleCircle    =     8, /*!< Circular markers */
    xlMarkerStyleDash      = -4115, /*!< Long bar markers */
    xlMarkerStyleDiamond   =     2, /*!< Diamond-shaped markers */
    xlMarkerStyleDot       = -4118, /*!< Short bar markers */
    xlMarkerStyleNone      = -4142, /*!< No markers */
    xlMarkerStylePicture   = -4147, /*!< Picture markers */
    xlMarkerStylePlus      =     9, /*!< Square markers with a plus sign */
    xlMarkerStyleSquare    =     1, /*!< Square markers */
    xlMarkerStyleStar      =     5, /*!< Square markers with an asterisk */
    xlMarkerStyleTriangle  =     3, /*!< Triangular markers */
    xlMarkerStyleX         = -4168, /*!< Square markers with an X */
};

/**
Specifies the measurement units.

[Official VBA documentation for XlMeasurementUnits](https://docs.microsoft.com/office/vba/api/excel.xlmeasurementunits)
*/
enum XlMeasurementUnits
{
    xlCentimeters = 1, /*!< Centimeters */
    xlInches      = 0, /*!< Inches */
    xlMillimeters = 2, /*!< Millimeters */
};

/**
Specifies the source of the change to the data model.

[Official VBA documentation for XlModelChangeSource](https://docs.microsoft.com/office/vba/api/excel.xlmodelchangesource)
*/
enum XlModelChangeSource
{
    xlChangeByExcel           = 0, /*!< Excel */
    xlChangeByPowerPivotAddIn = 1, /*!< PowerPivot add-in */
};

/**
Specifies which mouse button was pressed.

[Official VBA documentation for XlMouseButton](https://docs.microsoft.com/office/vba/api/excel.xlmousebutton)
*/
enum XlMouseButton
{
    xlNoButton        = 0, /*!< No button was pressed. */
    xlPrimaryButton   = 1, /*!< The primary button (normally the left mouse button) was pressed. */
    xlSecondaryButton = 2, /*!< The secondary button (normally the right mouse button) was pressed. */
};

/**
Specifies the appearance of the mouse pointer in Excel.

[Official VBA documentation for XlMousePointer](https://docs.microsoft.com/office/vba/api/excel.xlmousepointer)
*/
enum XlMousePointer
{
    xlDefault        = -4143, /*!< The default pointer. */
    xlIBeam          =     3, /*!< The I-beam pointer. */
    xlNorthwestArrow =     1, /*!< The northwest-arrow pointer. */
    xlWait           =     2, /*!< The hourglass pointer. */
};

/**
Specifies the OLE object type.

[Official VBA documentation for XlOLEType](https://docs.microsoft.com/office/vba/api/excel.xloletype)
*/
enum XlOLEType
{
    xlOLEControl = 2, /*!< ActiveX control */
    xlOLEEmbed   = 1, /*!< Embedded OLE object */
    xlOLELink    = 0, /*!< Linked OLE object */
};

/**
Specifies the verb on which the server of the OLE object should act.

[Official VBA documentation for XlOLEVerb](https://docs.microsoft.com/office/vba/api/excel.xloleverb)
*/
enum XlOLEVerb
{
    xlVerbOpen    = 2, /*!< Open the object. */
    xlVerbPrimary = 1, /*!< Perform the primary action for the server. */
};

/**
Specifies the horizontal overflow setting for a text frame.

[Official VBA documentation for XlOartHorizontalOverflow](https://docs.microsoft.com/office/vba/api/excel.xloarthorizontaloverflow)
*/
enum XlOartHorizontalOverflow
{
    xlOartHorizontalOverflowClip     = 1, /*!< Hide text that does not fit horizontally in the text frame. */
    xlOartHorizontalOverflowOverflow = 0, /*!< Allow text to overflow the text frame horizontally. */
};

/**
Specifies the vertical overflow setting for a text frame.

[Official VBA documentation for XlOartVerticalOverflow](https://docs.microsoft.com/office/vba/api/excel.xloartverticaloverflow)
*/
enum XlOartVerticalOverflow
{
    xlOartVerticalOverflowClip     = 1, /*!< Hide text that does not fit vertically within the text frame. */
    xlOartVerticalOverflowEllipsis = 2, /*!< Hide text that does not fit vertically within the text frame, and add an ellipsis (...) at the end of the visible text. */
    xlOartVerticalOverflowOverflow = 0, /*!< Allow text to overflow the text frame vertically (can be from the top, bottom, or both depending on the text alignment). */
};

/**
Specifies the way a chart is scaled to fit on a page.

[Official VBA documentation for XlObjectSize](https://docs.microsoft.com/office/vba/api/excel.xlobjectsize)
*/
enum XlObjectSize
{
    xlFitToPage  = 2, /*!< Print the chart as large as possible, while retaining the chart's height-to-width ratio as shown on the screen. */
    xlFullPage   = 3, /*!< Print the chart to fit the page, adjusting the height-to-width ratio as necessary. */
    xlScreenSize = 1, /*!< Print the chart the same size as it appears on the screen. */
};

/**
Specifies the order in which cells are processed.

[Official VBA documentation for XlOrder](https://docs.microsoft.com/office/vba/api/excel.xlorder)
*/
enum XlOrder
{
    xlDownThenOver = 1, /*!< Process down the rows before processing across pages or page fields to the right. */
    xlOverThenDown = 2, /*!< Process across pages or page fields to the right before moving down the rows. */
};

/**
Specifies the text orientation.

[Official VBA documentation for XlOrientation](https://docs.microsoft.com/office/vba/api/excel.xlorientation)
*/
enum XlOrientation
{
    xlDownward   = -4170, /*!< Text runs downward. */
    xlHorizontal = -4128, /*!< Text runs horizontally. */
    xlUpward     = -4171, /*!< Text runs upward. */
    xlVertical   = -4166, /*!< Text runs downward and is centered in the cell. */
};

/**
Specifies what can be selected in a PivotTable during a structured selection. These constants can be combined to select multiple types.

[Official VBA documentation for XlPTSelectionMode](https://docs.microsoft.com/office/vba/api/excel.xlptselectionmode)
*/
enum XlPTSelectionMode
{
    xlBlanks       =   4, /*!< Blanks */
    xlButton       =  15, /*!< Buttons */
    xlDataAndLabel =   0, /*!< Data and labels */
    xlDataOnly     =   2, /*!< Data */
    xlFirstRow     = 256, /*!< First row */
    xlLabelOnly    =   1, /*!< Label */
    xlOrigin       =   3, /*!< Origin */
};

/**
Specifies page break location on the worksheet.

[Official VBA documentation for XlPageBreak](https://docs.microsoft.com/office/vba/api/excel.xlpagebreak)
*/
enum XlPageBreak
{
    xlPageBreakAutomatic = -4105, /*!< Excel will automatically add page breaks. */
    xlPageBreakManual    = -4135, /*!< Page breaks are manually inserted. */
    xlPageBreakNone      = -4142, /*!< Page breaks are not inserted on the worksheet. */
};

/**
Specifies whether a page break is full screen or applies only within the print area.

[Official VBA documentation for XlPageBreakExtent](https://docs.microsoft.com/office/vba/api/excel.xlpagebreakextent)
*/
enum XlPageBreakExtent
{
    xlPageBreakFull    = 1, /*!< Full screen. */
    xlPageBreakPartial = 2, /*!< Only within print area. */
};

/**
Specifies the page orientation when the worksheet is printed.

[Official VBA documentation for XlPageOrientation](https://docs.microsoft.com/office/vba/api/excel.xlpageorientation)
*/
enum XlPageOrientation
{
    xlLandscape = 2, /*!< Landscape mode. */
    xlPortrait  = 1, /*!< Portrait mode. */
};

/**
Specifies the size of the paper.

[Official VBA documentation for XlPaperSize](https://docs.microsoft.com/office/vba/api/excel.xlpapersize)
*/
enum XlPaperSize
{
    xlPaper10x14              =  16, /*!< 10 in. x 14 in. */
    xlPaper11x17              =  17, /*!< 11 in. x 17 in. */
    xlPaperA3                 =   8, /*!< A3 (297 mm x 420 mm) */
    xlPaperA4                 =   9, /*!< A4 (210 mm x 297 mm) */
    xlPaperA4Small            =  10, /*!< A4 Small (210 mm x 297 mm) */
    xlPaperA5                 =  11, /*!< A5 (148 mm x 210 mm) */
    xlPaperB4                 =  12, /*!< B4 (250 mm x 354 mm) */
    xlPaperB5                 =  13, /*!< A5 (148 mm x 210 mm) */
    xlPaperCsheet             =  24, /*!< C size sheet */
    xlPaperDsheet             =  25, /*!< D size sheet */
    xlPaperEnvelope10         =  20, /*!< Envelope #10 (4-1/8 in. x 9-1/2 in.) */
    xlPaperEnvelope11         =  21, /*!< Envelope #11 (4-1/2 in. x 10-3/8 in.) */
    xlPaperEnvelope12         =  22, /*!< Envelope #12 (4-1/2 in. x 11 in.) */
    xlPaperEnvelope14         =  23, /*!< Envelope #14 (5 in. x 11-1/2 in.) */
    xlPaperEnvelope9          =  19, /*!< Envelope #9 (3-7/8 in. x 8-7/8 in.) */
    xlPaperEnvelopeB4         =  33, /*!< Envelope B4 (250 mm x 353 mm) */
    xlPaperEnvelopeB5         =  34, /*!< Envelope B5 (176 mm x 250 mm) */
    xlPaperEnvelopeB6         =  35, /*!< Envelope B6 (176 mm x 125 mm) */
    xlPaperEnvelopeC3         =  29, /*!< Envelope C3 (324 mm x 458 mm) */
    xlPaperEnvelopeC4         =  30, /*!< Envelope C4 (229 mm x 324 mm) */
    xlPaperEnvelopeC5         =  28, /*!< Envelope C5 (162 mm x 229 mm) */
    xlPaperEnvelopeC6         =  31, /*!< Envelope C6 (114 mm x 162 mm) */
    xlPaperEnvelopeC65        =  32, /*!< Envelope C65 (114 mm x 229 mm) */
    xlPaperEnvelopeDL         =  27, /*!< Envelope DL (110 mm x 220 mm) */
    xlPaperEnvelopeItaly      =  36, /*!< Envelope (110 mm x 230 mm) */
    xlPaperEnvelopeMonarch    =  37, /*!< Envelope Monarch (3-7/8 in. x 7-1/2 in.) */
    xlPaperEnvelopePersonal   =  38, /*!< Envelope (3-5/8 in. x 6-1/2 in.) */
    xlPaperEsheet             =  26, /*!< E size sheet */
    xlPaperExecutive          =   7, /*!< Executive (7-1/2 in. x 10-1/2 in.) */
    xlPaperFanfoldLegalGerman =  41, /*!< German Legal Fanfold (8-1/2 in. x 13 in.) */
    xlPaperFanfoldStdGerman   =  40, /*!< German Legal Fanfold (8-1/2 in. x 13 in.) */
    xlPaperFanfoldUS          =  39, /*!< U.S. Standard Fanfold (14-7/8 in. x 11 in.) */
    xlPaperFolio              =  14, /*!< Folio (8-1/2 in. x 13 in.) */
    xlPaperLedger             =   4, /*!< Ledger (17 in. x 11 in.) */
    xlPaperLegal              =   5, /*!< Legal (8-1/2 in. x 14 in.) */
    xlPaperLetter             =   1, /*!< Letter (8-1/2 in. x 11 in.) */
    xlPaperLetterSmall        =   2, /*!< Letter Small (8-1/2 in. x 11 in.) */
    xlPaperNote               =  18, /*!< Note (8-1/2 in. x 11 in.) */
    xlPaperQuarto             =  15, /*!< Quarto (215 mm x 275 mm) */
    xlPaperStatement          =   6, /*!< Statement (5-1/2 in. x 8-1/2 in.) */
    xlPaperTabloid            =   3, /*!< Tabloid (11 in. x 17 in.) */
    xlPaperUser               = 256, /*!< User-defined */
};

/**
Specifies the data type of a query parameter.

[Official VBA documentation for XlParameterDataType](https://docs.microsoft.com/office/vba/api/excel.xlparameterdatatype)
*/
enum XlParameterDataType
{
    xlParamTypeBigInt        = -5, /*!< Big integer. */
    xlParamTypeBinary        = -2, /*!< Binary. */
    xlParamTypeBit           = -7, /*!< Bit. */
    xlParamTypeChar          =  1, /*!< String. */
    xlParamTypeDate          =  9, /*!< Date. */
    xlParamTypeDecimal       =  3, /*!< Decimal. */
    xlParamTypeDouble        =  8, /*!< Double. */
    xlParamTypeFloat         =  6, /*!< Float. */
    xlParamTypeInteger       =  4, /*!< Integer. */
    xlParamTypeLongVarBinary = -4, /*!< Long binary. */
    xlParamTypeLongVarChar   = -1, /*!< Long string. */
    xlParamTypeNumeric       =  2, /*!< Numeric. */
    xlParamTypeReal          =  7, /*!< Real. */
    xlParamTypeSmallInt      =  5, /*!< Small integer. */
    xlParamTypeTime          = 10, /*!< Time. */
    xlParamTypeTimestamp     = 11, /*!< Time stamp. */
    xlParamTypeTinyInt       = -6, /*!< Tiny integer. */
    xlParamTypeUnknown       =  0, /*!< Type unknown. */
    xlParamTypeVarBinary     = -3, /*!< Variable-length binary. */
    xlParamTypeVarChar       = 12, /*!< Variable-length string. */
    xlParamTypeWChar         = -8, /*!< Unicode character string. */
};

/**
Specifies how to determine the value of the parameter for the specified query table.

[Official VBA documentation for XlParameterType](https://docs.microsoft.com/office/vba/api/excel.xlparametertype)
*/
enum XlParameterType
{
    xlConstant = 1, /*!< Uses the value specified by the _Value_ argument. */
    xlPrompt   = 0, /*!< Displays a dialog box that prompts the user for the value. The _Value_ argument specifies the text shown in the dialog box. */
    xlRange    = 2, /*!< Uses the value of the cell in the upper-left corner of the range. The _Value_ argument specifies a **[Range](Excel.Range(object).md)** object. */
};

/**
Constants passed to and returned by the **Series.ParentDataLabelOption** property.

[Official VBA documentation for XlParentDataLabelOptions](https://docs.microsoft.com/office/vba/api/excel.xlparentdatalabeloptions)
*/
enum XlParentDataLabelOptions
{
    xlParentDataLabelOptionsBanner      = 1, /*!< Banner parent data label */
    xlParentDataLabelOptionsNone        = 0, /*!< No parent data label */
    xlParentDataLabelOptionsOverlapping = 2, /*!< Overlapping parent data label */
};

/**
Specifies how numeric data will be calculated with the destinations cells on the worksheet.

[Official VBA documentation for XlPasteSpecialOperation](https://docs.microsoft.com/office/vba/api/excel.xlpastespecialoperation)
*/
enum XlPasteSpecialOperation
{
    xlPasteSpecialOperationAdd      =     2, /*!< Copied data will be added to the value in the destination cell. */
    xlPasteSpecialOperationDivide   =     5, /*!< Copied data will divide the value in the destination cell. */
    xlPasteSpecialOperationMultiply =     4, /*!< Copied data will multiply the value in the destination cell. */
    xlPasteSpecialOperationNone     = -4142, /*!< No calculation will be done in the paste operation. */
    xlPasteSpecialOperationSubtract =     3, /*!< Copied data will be subtracted from the value in the destination cell. */
};

/**
Specifies the part of the range to be pasted.

[Official VBA documentation for XlPasteType](https://docs.microsoft.com/office/vba/api/excel.xlpastetype)
*/
enum XlPasteType
{
    xlPasteAll                          = -4104, /*!< Everything will be pasted. */
    xlPasteAllExceptBorders             =     7, /*!< Everything except borders will be pasted. */
    xlPasteAllMergingConditionalFormats =    14, /*!< Everything will be pasted and conditional formats will be merged. */
    xlPasteAllUsingSourceTheme          =    13, /*!< Everything will be pasted using the source theme. */
    xlPasteColumnWidths                 =     8, /*!< Copied column width is pasted. */
    xlPasteComments                     = -4144, /*!< Comments are pasted. */
    xlPasteFormats                      = -4122, /*!< Copied source format is pasted. */
    xlPasteFormulas                     = -4123, /*!< Formulas are pasted. */
    xlPasteFormulasAndNumberFormats     =    11, /*!< Formulas and Number formats are pasted. */
    xlPasteValidation                   =     6, /*!< Validations are pasted. */
    xlPasteValues                       = -4163, /*!< Values are pasted. */
    xlPasteValuesAndNumberFormats       =    12, /*!< Values and Number formats are pasted. */
};

/**
Specifies the interior pattern of a chart or interior object.

[Official VBA documentation for XlPattern](https://docs.microsoft.com/office/vba/api/excel.xlpattern)
*/
enum XlPattern
{
    xlPatternAutomatic       = -4105, /*!< Excel controls the pattern. */
    xlPatternChecker         =     9, /*!< Checkerboard. */
    xlPatternCrissCross      =    16, /*!< Criss-cross lines. */
    xlPatternDown            = -4121, /*!< Dark diagonal lines running from the upper-left to the lower-right. */
    xlPatternGray16          =    17, /*!< 16% gray. */
    xlPatternGray25          = -4124, /*!< 25% gray. */
    xlPatternGray50          = -4125, /*!< 50% gray. */
    xlPatternGray75          = -4126, /*!< 75% gray. */
    xlPatternGray8           =    18, /*!< 8% gray. */
    xlPatternGrid            =    15, /*!< Grid. */
    xlPatternHorizontal      = -4128, /*!< Dark horizontal lines. */
    xlPatternLightDown       =    13, /*!< Light diagonal lines running from the upper-left to the lower-right. */
    xlPatternLightHorizontal =    11, /*!< Light horizontal lines. */
    xlPatternLightUp         =    14, /*!< Light diagonal lines running from the lower-left to the upper-right. */
    xlPatternLightVertical   =    12, /*!< Light vertical bars. */
    xlPatternNone            = -4142, /*!< No pattern. */
    xlPatternSemiGray75      =    10, /*!< 75% dark gray. */
    xlPatternSolid           =     1, /*!< Solid color. */
    xlPatternUp              = -4162, /*!< Dark diagonal lines running from the lower-left to the upper-right. */
    xlPatternVertical        = -4166, /*!< Dark vertical bars. */
};

/**
Specifies the alignment for phonetic text. Used with a **Phonetic** or a **Phonetics** object.

[Official VBA documentation for XlPhoneticAlignment](https://docs.microsoft.com/office/vba/api/excel.xlphoneticalignment)
*/
enum XlPhoneticAlignment
{
    xlPhoneticAlignCenter      = 2, /*!< Centered */
    xlPhoneticAlignDistributed = 3, /*!< Distributed */
    xlPhoneticAlignLeft        = 1, /*!< Left aligned */
    xlPhoneticAlignNoControl   = 0, /*!< Excel controls alignment */
};

/**
Specifies the type of phonetic text in a cell.

[Official VBA documentation for XlPhoneticCharacterType](https://docs.microsoft.com/office/vba/api/excel.xlphoneticcharactertype)
*/
enum XlPhoneticCharacterType
{
    xlHiragana     = 2, /*!< Hiragana */
    xlKatakana     = 1, /*!< Katakana */
    xlKatakanaHalf = 0, /*!< Half-size Katakana */
    xlNoConversion = 3, /*!< No conversion */
};

/**
Specifies how the picture should be copied.

[Official VBA documentation for XlPictureAppearance](https://docs.microsoft.com/office/vba/api/excel.xlpictureappearance)
*/
enum XlPictureAppearance
{
    xlPrinter = 2, /*!< The picture is copied as it will look when it is printed. */
    xlScreen  = 1, /*!< The picture is copied to resemble its display on the screen as closely as possible. */
};

/**
Specifies how to convert a graphic.

[Official VBA documentation for XlPictureConvertorType](https://docs.microsoft.com/office/vba/api/excel.xlpictureconvertortype)
*/
enum XlPictureConvertorType
{
    xlBMP =  1, /*!< Windows version 2.0 - compatible bitmap */
    xlCGM =  7, /*!< Computer Graphics Metafile */
    xlDRW =  4, /*!< DRW */
    xlDXF =  5, /*!< DXF */
    xlEPS =  8, /*!< Encapsulated Postscript */
    xlHGL =  6, /*!< HGL */
    xlPCT = 13, /*!< Bitmap Graphic (Apple PICT format) */
    xlPCX = 10, /*!< PC Paintbrush Bitmap Graphic */
    xlPIC = 11, /*!< PIC */
    xlPLT = 12, /*!< PLT */
    xlTIF =  9, /*!< Tagged Image Format File */
    xlWMF =  2, /*!< Windows Metafile */
    xlWPG =  3, /*!< WordPerfect/DrawPerfect Graphic */
};

/**
Specifies which position on the slice to return the coordinate of.

[Official VBA documentation for XlPieSliceIndex](https://docs.microsoft.com/office/vba/api/excel.xlpiesliceindex)
*/
enum XlPieSliceIndex
{
    xlCenterPoint                    = 5, /*!< The center point of a pie slice. */
    xlInnerCenterPoint               = 8, /*!< The innermost center point of a doughnut slice. */
    xlInnerClockwisePoint            = 7, /*!< The innermost point of the most clockwise radius of a doughnut slice. */
    xlInnerCounterClockwisePoint     = 9, /*!< The innermost point of the most counterclockwise radius of a doughnut slice. */
    xlMidClockwiseRadiusPoint        = 4, /*!< The midpoint of the most clockwise radius of a slice. */
    xlMidCounterClockwiseRadiusPoint = 6, /*!< The midpoint of the most counterclockwise radius of a slice. */
    xlOuterCenterPoint               = 2, /*!< The outer center point of the circumference of a slice. */
    xlOuterClockwisePoint            = 3, /*!< The outermost clockwise point of the circumference of a slice. */
    xlOuterCounterClockwisePoint     = 1, /*!< The outermost counterclockwise point of the circumference of a slice. */
};

/**
Specifies the horizontal or vertical position of a point on a pie chart, in [points](../language/glossary/vbe-glossary.md#point), from the top or left edge of the object to the top or left edge of the chart area.

[Official VBA documentation for XlPieSliceLocation](https://docs.microsoft.com/office/vba/api/excel.xlpieslicelocation)
*/
enum XlPieSliceLocation
{
    xlHorizontalCoordinate = 1, /*!< The horizontal coordinate (x) */
    xlVerticalCoordinate   = 2, /*!< The vertical coordinate (y) */
};

/**
Specifies the **PivotTable** entity to which the cell corresponds.

[Official VBA documentation for XlPivotCellType](https://docs.microsoft.com/office/vba/api/excel.xlpivotcelltype)
*/
enum XlPivotCellType
{
    xlPivotCellBlankCell      = 9, /*!< A structural blank cell in the PivotTable. */
    xlPivotCellCustomSubtotal = 7, /*!< A cell in the row or column area that is a custom subtotal. */
    xlPivotCellDataField      = 4, /*!< A data field label (not the **Data** button). */
    xlPivotCellDataPivotField = 8, /*!< The **Data** button. */
    xlPivotCellGrandTotal     = 3, /*!< A cell in a row or column area that is a grand total. */
    xlPivotCellPageFieldItem  = 6, /*!< The cell that shows the selected item of a Page field. */
    xlPivotCellPivotField     = 5, /*!< The button for a field (not the **Data** button). */
    xlPivotCellPivotItem      = 1, /*!< A cell in the row or column area that is not a subtotal, grand total, custom subtotal, or blank line. */
    xlPivotCellSubtotal       = 2, /*!< A cell in the row or column area that is a subtotal. */
    xlPivotCellValue          = 0, /*!< Any cell in the data area (except a blank row). */
};

/**
This enumeration specifies the conditional formatting applied for filtering values from the **PivotTable** object.

[Official VBA documentation for XlPivotConditionScope](https://docs.microsoft.com/office/vba/api/excel.xlpivotconditionscope)
*/
enum XlPivotConditionScope
{
    xlDataFieldScope = 2, /*!< Based on the data in the specified fields. */
    xlFieldsScope    = 1, /*!< Based on the specified fields. */
    xlSelectionScope = 0, /*!< Based on the specified selection criteria. */
};

/**
Specifies the type of calculation performed by a data PivotField when a custom calculation is used.

[Official VBA documentation for XlPivotFieldCalculation](https://docs.microsoft.com/office/vba/api/excel.xlpivotfieldcalculation)
*/
enum XlPivotFieldCalculation
{
    xlDifferenceFrom          =     2, /*!< The difference from the value of the Base item in the Base field. */
    xlIndex                   =     9, /*!< Data calculated as ((value in cell) x (Grand Total of Grand Totals)) / ((Grand Row Total) x (Grand Column Total)). */
    xlNoAdditionalCalculation = -4143, /*!< No calculation. */
    xlPercentDifferenceFrom   =     4, /*!< Percentage difference from the value of the Base item in the Base field. */
    xlPercentOf               =     3, /*!< Percentage of the value of the Base item in the Base field. */
    xlPercentOfColumn         =     7, /*!< Percentage of the total for the column or series. */
    xlPercentOfParent         =    12, /*!< Percentage of the total of the specified parent Base field. */
    xlPercentOfParentColumn   =    11, /*!< Percentage of the total of the parent column. */
    xlPercentOfParentRow      =    10, /*!< Percentage of the total of the parent row. */
    xlPercentOfRow            =     6, /*!< Percentage of the total for the row or category. */
    xlPercentOfTotal          =     8, /*!< Percentage of the grand total of all the data or data points in the report. */
    xlPercentRunningTotal     =    13, /*!< Percentage of the running total of the specified Base field. */
    xlRankAscending           =    14, /*!< Rank smallest to largest. */
    xlRankDecending           =    15, /*!< Rank largest to smallest. */
    xlRunningTotal            =     5, /*!< Data for successive items in the Base field as a running total. */
};

/**
Specifies the type of data in the **PivotTable** field.

[Official VBA documentation for XlPivotFieldDataType](https://docs.microsoft.com/office/vba/api/excel.xlpivotfielddatatype)
*/
enum XlPivotFieldDataType
{
    xlDate   =     2, /*!< Contains a date. */
    xlNumber = -4145, /*!< Contains a number. */
    xlText   = -4158, /*!< Contains text. */
};

/**
Specifies the location of the field in a PivotTable report.

[Official VBA documentation for XlPivotFieldOrientation](https://docs.microsoft.com/office/vba/api/excel.xlpivotfieldorientation)
*/
enum XlPivotFieldOrientation
{
    xlColumnField = 2, /*!< Column */
    xlDataField   = 4, /*!< Data */
    xlHidden      = 0, /*!< Hidden */
    xlPageField   = 3, /*!< Page */
    xlRowField    = 1, /*!< Row */
};

/**
Specifies whether to repeat all field item labels in a PivotTable report.

[Official VBA documentation for XlPivotFieldRepeatLabels](https://docs.microsoft.com/office/vba/api/excel.xlpivotfieldrepeatlabels)
*/
enum XlPivotFieldRepeatLabels
{
    xlDoNotRepeatLabels = 1, /*!< Do not repeat item labels. */
    xlRepeatLabels      = 2, /*!< Repeat all item labels. */
};

/**
The type of filter applied.

[Official VBA documentation for XlPivotFilterType](https://docs.microsoft.com/office/vba/api/excel.xlpivotfiltertype)
*/
enum XlPivotFilterType
{
    xlBefore                        = 31, /*!< Filters for all dates before a specified date */
    xlBeforeOrEqualTo               = 32, /*!< Filters for all dates on or before a specified date */
    xlAfter                         = 33, /*!< Filters for all dates after a specified date */
    xlAfterOrEqualTo                = 34, /*!< Filters for all dates on or after a specified date */
    xlAllDatesInPeriodJanuary       = 57, /*!< Filters for all dates in January */
    xlAllDatesInPeriodFebruary      = 58, /*!< Filters for all dates in February */
    xlAllDatesInPeriodMarch         = 59, /*!< Filters for all dates in March */
    xlAllDatesInPeriodApril         = 60, /*!< Filters for all dates in April */
    xlAllDatesInPeriodMay           = 61, /*!< Filters for all dates in May */
    xlAllDatesInPeriodJune          = 62, /*!< Filters for all dates in June */
    xlAllDatesInPeriodJuly          = 63, /*!< Filters for all dates in July */
    xlAllDatesInPeriodAugust        = 64, /*!< Filters for all dates in August */
    xlAllDatesInPeriodSeptember     = 65, /*!< Filters for all dates in September */
    xlAllDatesInPeriodOctober       = 66, /*!< Filters for all dates in October */
    xlAllDatesInPeriodNovember      = 67, /*!< Filters for all dates in November */
    xlAllDatesInPeriodDecember      = 68, /*!< Filters for all dates in December */
    xlAllDatesInPeriodQuarter1      = 53, /*!< Filters for all dates in Quarter1 */
    xlAllDatesInPeriodQuarter2      = 54, /*!< Filters for all dates in Quarter2 */
    xlAllDatesInPeriodQuarter3      = 55, /*!< Filters for all dates in Quarter3 */
    xlAllDatesInPeriodQuarter4      = 56, /*!< Filters for all dates in Quarter 4 */
    xlBottomCount                   =  2, /*!< Filters for the specified number of values from the bottom of a list */
    xlBottomPercent                 =  4, /*!< Filters for the specified percentage of values from the bottom of a list */
    xlBottomSum                     =  6, /*!< Sum of the values from the bottom of the list */
    xlCaptionBeginsWith             = 17, /*!< Filters for all captions beginning with the specified string */
    xlCaptionContains               = 21, /*!< Filters for all captions that contain the specified string */
    xlCaptionDoesNotBeginWith       = 18, /*!< Filters for all captions that don't begin with the specified string */
    xlCaptionDoesNotContain         = 22, /*!< Filters for all captions that don't contain the specified string */
    xlCaptionDoesNotEndWith         = 20, /*!< Filters for all captions that don't end with the specified string */
    xlCaptionDoesNotEqual           = 16, /*!< Filters for all captions that don't match the specified string */
    xlCaptionEndsWith               = 19, /*!< Filters for all captions that end with the specified string */
    xlCaptionEquals                 = 15, /*!< Filters for all captions that match the specified string */
    xlCaptionIsBetween              = 27, /*!< Filters for all captions that are between a specified range of values */
    xlCaptionIsGreaterThan          = 23, /*!< Filters for all captions that are greater than the specified value */
    xlCaptionIsGreaterThanOrEqualTo = 24, /*!< Filters for all captions that are greater than or match the specified value */
    xlCaptionIsLessThan             = 25, /*!< Filters for all captions that are less than the specified value */
    xlCaptionIsLessThanOrEqualTo    = 26, /*!< Filters for all captions that are less than or match the specified value */
    xlCaptionIsNotBetween           = 28, /*!< Filters for all captions that are not between a specified range of values */
    xlDateBetween                   = 35, /*!< Filters for all dates that are between a specified range of dates */
    xlDateLastMonth                 = 45, /*!< Filters for all dates that apply to the previous month */
    xlDateLastQuarter               = 48, /*!< Filters for all dates that apply to the previous quarter */
    xlDateLastWeek                  = 42, /*!< Filters for all dates that apply to the previous week */
    xlDateLastYear                  = 51, /*!< Filters for all dates that apply to the previous year */
    xlDateNextMonth                 = 43, /*!< Filters for all dates that apply to the next month */
    xlDateNextQuarter               = 46, /*!< Filters for all dates that apply to the next quarter */
    xlDateNextWeek                  = 40, /*!< Filters for all dates that apply to the next week */
    xlDateNextYear                  = 49, /*!< Filters for all dates that apply to the next year */
    xlDateThisMonth                 = 44, /*!< Filters for all dates that apply to the current month */
    xlDateThisQuarter               = 47, /*!< Filters for all dates that apply to the current quarter */
    xlDateThisWeek                  = 41, /*!< Filters for all dates that apply to the current week */
    xlDateThisYear                  = 50, /*!< Filters for all dates that apply to the current year */
    xlDateToday                     = 38, /*!< Filters for all dates that apply to the current date */
    xlDateTomorrow                  = 37, /*!< Filters for all dates that apply to the next day */
    xlDateYesterday                 = 39, /*!< Filters for all dates that apply to the previous day */
    xlNotSpecificDate               = 30, /*!< Filters for all dates that don't match a specified date */
    xlSpecificDate                  = 29, /*!< Filters for all dates that match a specified date */
    xlTopCount                      =  1, /*!< Filters for the specified number of values from the top of a list */
    xlTopPercent                    =  3, /*!< Filters for the specified percentage of values from a list */
    xlTopSum                        =  5, /*!< Sum of the values from the top of the list */
    xlValueDoesNotEqual             =  8, /*!< Filters for all values that don't match the specified value */
    xlValueEquals                   =  7, /*!< Filters for all values that match the specified value */
    xlValueIsBetween                = 13, /*!< Filters for all values that are between a specified range of values */
    xlValueIsGreaterThan            =  9, /*!< Filters for all values that are greater than the specified value */
    xlValueIsGreaterThanOrEqualTo   = 10, /*!< Filters for all values that are greater than or match the specified value */
    xlValueIsLessThan               = 11, /*!< Filters for all values that are less than the specified value */
    xlValueIsLessThanOrEqualTo      = 12, /*!< Filters for all values that are less than or match the specified value */
    xlValueIsNotBetween             = 14, /*!< Filters for all values that are not between a specified range of values */
    xlYearToDate                    = 52, /*!< Filters for all values that are within one year of a specified date */
};

/**
Specifies the type of report formatting to be applied to the specified PivotTable report.

[Official VBA documentation for XlPivotFormatType](https://docs.microsoft.com/office/vba/api/excel.xlpivotformattype)
*/
enum XlPivotFormatType
{
    xlPTClassic = 20, /*!< PivotTable classic format. */
    xlPTNone    = 21, /*!< Does not apply formatting to the PivotTable report. */
    xlReport1   =  0, /*!< Use the xlReport1 formatting for the PivotTable. */
    xlReport10  =  9, /*!< Use the xlReport10 formatting for the PivotTable. */
    xlReport2   =  1, /*!< Use the xlReport2 formatting for the PivotTable. */
    xlReport3   =  2, /*!< Use the xlReport3 formatting for the PivotTable. */
    xlReport4   =  3, /*!< Use the xlReport4 formatting for the PivotTable. */
    xlReport5   =  4, /*!< Use the xlReport5 formatting for the PivotTable. */
    xlReport6   =  5, /*!< Use the xlReport6 formatting for the PivotTable. */
    xlReport7   =  6, /*!< Use the xlReport7 formatting for the PivotTable. */
    xlReport8   =  7, /*!< Use the xlReport8 formatting for the PivotTable. */
    xlReport9   =  8, /*!< Use the xlReport9 formatting for the PivotTable. */
    xlTable1    = 10, /*!< Use the xlTable1 formatting for the PivotTable. */
    xlTable10   = 19, /*!< Use the xlTable10 formatting for the PivotTable. */
    xlTable2    = 11, /*!< Use the xlTable2 formatting for the PivotTable. */
    xlTable3    = 12, /*!< Use the xlTable3 formatting for the PivotTable. */
    xlTable4    = 13, /*!< Use the xlTable4 formatting for the PivotTable. */
    xlTable5    = 14, /*!< Use the xlTable5 formatting for the PivotTable. */
    xlTable6    = 15, /*!< Use the xlTable6 formatting for the PivotTable. */
    xlTable7    = 16, /*!< Use the xlTable7 formatting for the PivotTable. */
    xlTable8    = 17, /*!< Use the xlTable8 formatting for the PivotTable. */
    xlTable9    = 18, /*!< Use the xlTable9 formatting for the PivotTable. */
};

/**
Specifies the type of the PivotLine.

[Official VBA documentation for XlPivotLineType](https://docs.microsoft.com/office/vba/api/excel.xlpivotlinetype)
*/
enum XlPivotLineType
{
    xlPivotLineBlank      = 3, /*!< Blank line after each group. */
    xlPivotLineGrandTotal = 2, /*!< Grand Total line. */
    xlPivotLineRegular    = 0, /*!< Regular PivotLine with pivot items. */
    xlPivotLineSubtotal   = 1, /*!< Subtotal line. */
};

/**
Specifies the maximum number of unique items allowed per PivotField.

[Official VBA documentation for XlPivotTableMissingItems](https://docs.microsoft.com/office/vba/api/excel.xlpivottablemissingitems)
*/
enum XlPivotTableMissingItems
{
    xlMissingItemsDefault =      -1, /*!< The default number of unique items per PivotField allowed. */
    xlMissingItemsMax     =   32500, /*!< The maximum number of unique items per PivotField allowed (32,500) for a pre-Excel 2007 PivotTable. */
    xlMissingItemsMax2    = 1048576, /*!< The maximum number of unique items per PivotField allowed (1,048,576) for PivotTables in Excel 2007 and later. */
    xlMissingItemsNone    =       0, /*!< No unique items per PivotField allowed (zero). */
};

/**
Specifies the source of the report data.

[Official VBA documentation for XlPivotTableSourceType](https://docs.microsoft.com/office/vba/api/excel.xlpivottablesourcetype)
*/
enum XlPivotTableSourceType
{
    xlConsolidation =     3, /*!< Multiple consolidation ranges. */
    xlDatabase      =     1, /*!< Microsoft Excel list or database. */
    xlExternal      =     2, /*!< Data from another application. */
    xlPivotTable    = -4148, /*!< Same source as another PivotTable report. */
    xlScenario      =     4, /*!< Data is based on scenarios created using the Scenario Manager. */
};

/**
Specifies the version of a PivotTable or a PivotCache. Creating PivotTables with a specific version ensures that tables created in Excel behave in the same manner as they did in the corresponding version of Excel.

[Official VBA documentation for XlPivotTableVersionList](https://docs.microsoft.com/office/vba/api/excel.xlpivottableversionlist)
*/
enum XlPivotTableVersionList
{
    xlPivotTableVersion2000    =  0, /*!< Excel 2000 */
    xlPivotTableVersion10      =  1, /*!< Excel 2002 */
    xlPivotTableVersion11      =  2, /*!< Excel 2003 */
    xlPivotTableVersion12      =  3, /*!< Excel 2007 */
    xlPivotTableVersion14      =  4, /*!< Excel 2010 */
    xlPivotTableVersion15      =  5, /*!< Excel 2013 */
    xlPivotTableVersionCurrent = -1, /*!< Provided only for backward compatibility */
};

/**
Specifies the way that an object is attached to its underlying cells.

[Official VBA documentation for XlPlacement](https://docs.microsoft.com/office/vba/api/excel.xlplacement)
*/
enum XlPlacement
{
    xlFreeFloating = 3, /*!< Object is free floating. */
    xlMove         = 2, /*!< Object is moved with the cells. */
    xlMoveAndSize  = 1, /*!< Object is moved and sized with the cells. */
};

/**
Specifies the platform on which a text file originated.

[Official VBA documentation for XlPlatform](https://docs.microsoft.com/office/vba/api/excel.xlplatform)
*/
enum XlPlatform
{
    xlMacintosh = 1, /*!< Macintosh */
    xlMSDOS     = 3, /*!< MS-DOS */
    xlWindows   = 2, /*!< Microsoft Windows */
};

/**
Specifies the mode for checking the spelling of Portuguese.

[Official VBA documentation for XlPortugueseReform](https://docs.microsoft.com/office/vba/api/excel.xlportuguesereform)
*/
enum XlPortugueseReform
{
    xlPortugueseBoth       = 3, /*!< The spelling checker recognizes both pre-reform and post-reform spellings. */
    xlPortuguesePostReform = 2, /*!< The spelling checker recognizes only post-reform spellings. */
    xlPortuguesePreReform  = 1, /*!< The spelling checker recognizes only pre-reform spellings. */
};

/**
Specifies the type of print error displayed.

[Official VBA documentation for XlPrintErrors](https://docs.microsoft.com/office/vba/api/excel.xlprinterrors)
*/
enum XlPrintErrors
{
    xlPrintErrorsBlank     = 1, /*!< Print errors are blank. */
    xlPrintErrorsDash      = 2, /*!< Print errors are displayed as dashes. */
    xlPrintErrorsDisplayed = 0, /*!< All print errors are displayed. */
    xlPrintErrorsNA        = 3, /*!< Print errors are displayed as not available. */
};

/**
Specifies the way that comments are printed with the sheet.

[Official VBA documentation for XlPrintLocation](https://docs.microsoft.com/office/vba/api/excel.xlprintlocation)
*/
enum XlPrintLocation
{
    xlPrintInPlace    =    16, /*!< Comments will be printed where they were inserted on the worksheet. */
    xlPrintNoComments = -4142, /*!< Comments will not be printed. */
    xlPrintSheetEnd   =     1, /*!< Comments will be printed as end notes at the end of the worksheet. */
};

/**
Specifies the priority of a SendMailer message.

[Official VBA documentation for XlPriority](https://docs.microsoft.com/office/vba/api/excel.xlpriority)
*/
enum XlPriority
{
    xlPriorityHigh   = -4127, /*!< High */
    xlPriorityLow    = -4134, /*!< Low */
    xlPriorityNormal = -4143, /*!< Normal */
};

/**
Specifies where to display the property.

[Official VBA documentation for XlPropertyDisplayedIn](https://docs.microsoft.com/office/vba/api/excel.xlpropertydisplayedin)
*/
enum XlPropertyDisplayedIn
{
    xlDisplayPropertyInPivotTable           = 1, /*!< Displays member property in the PivotTable only. This is the default value. */
    xlDisplayPropertyInPivotTableAndTooltip = 3, /*!< Displays member property in the tooltip only. */
    xlDisplayPropertyInTooltip              = 2, /*!< Displays member property in both the tooltip and the PivotTable. */
};

/**
Specifies how the Protected View window was closed.

[Official VBA documentation for XlProtectedViewCloseReason](https://docs.microsoft.com/office/vba/api/excel.xlprotectedviewclosereason)
*/
enum XlProtectedViewCloseReason
{
    xlProtectedViewCloseEdit   = 1, /*!< The window was closed when the user clicked the **Enable Editing** button. */
    xlProtectedViewCloseForced = 2, /*!< The window was closed because the application shut it down forcefully or stopped responding. */
    xlProtectedViewCloseNormal = 0, /*!< The window was closed normally. */
};

/**
Specifies the state of the Protected View window.

[Official VBA documentation for XlProtectedViewWindowState](https://docs.microsoft.com/office/vba/api/excel.xlprotectedviewwindowstate)
*/
enum XlProtectedViewWindowState
{
    xlProtectedViewWindowMaximized = 2, /*!< Maximized */
    xlProtectedViewWindowMinimized = 1, /*!< Minimized */
    xlProtectedViewWindowNormal    = 0, /*!< Normal */
};

/**
Specifies the type of query used by Microsoft Excel to populate the query table or PivotTable cache.

[Official VBA documentation for XlQueryType](https://docs.microsoft.com/office/vba/api/excel.xlquerytype)
*/
enum XlQueryType
{
    xlADORecordset = 7, /*!< Based on an ADO recordset query */
    xlDAORecordset = 2, /*!< Based on a DAO recordset query, for query tables only */
    xlODBCQuery    = 1, /*!< Based on an ODBC data source */
    xlOLEDBQuery   = 5, /*!< Based on an OLE DB query, including OLAP data sources */
    xlTextImport   = 6, /*!< Based on a text file, for query tables only */
    xlWebQuery     = 4, /*!< Based on a webpage, for query tables only */
};

/**
Indicates for which top level button the callout user interface is displayed.

[Official VBA documentation for XlQuickAnalysisMode](https://docs.microsoft.com/office/vba/api/excel.xlquickanalysismode)
*/
enum XlQuickAnalysisMode
{
    xlLensOnly          = 0, /*!< Show the button but no callout user interface */
    xlFormatConditions  = 1, /*!< Conditional Formatting */
    xlRecommendedCharts = 2, /*!< Charts */
    xlTotals            = 3, /*!< Totals */
    xlTables            = 4, /*!< Tables */
    xlSparklines        = 5, /*!< Sparklines */
};

/**
Specifies the predefined format when a range is automatically formatted.

[Official VBA documentation for XlRangeAutoFormat](https://docs.microsoft.com/office/vba/api/excel.xlrangeautoformat)
*/
enum XlRangeAutoFormat
{
    xlRangeAutoFormat3DEffects1        =    13, /*!< 3D effects 1. */
    xlRangeAutoFormat3DEffects2        =    14, /*!< 3D effects 2. */
    xlRangeAutoFormatAccounting1       =     4, /*!< Accounting 1. */
    xlRangeAutoFormatAccounting2       =     5, /*!< Accounting 2. */
    xlRangeAutoFormatAccounting3       =     6, /*!< Accounting 3. */
    xlRangeAutoFormatAccounting4       =    17, /*!< Accounting 4. */
    xlRangeAutoFormatClassic1          =     1, /*!< Classic 1. */
    xlRangeAutoFormatClassic2          =     2, /*!< Classic 2. */
    xlRangeAutoFormatClassic3          =     3, /*!< Classic 3. */
    xlRangeAutoFormatClassicPivotTable =    31, /*!< Classic PivotTable. */
    xlRangeAutoFormatColor1            =     7, /*!< Color 1. */
    xlRangeAutoFormatColor2            =     8, /*!< Color 2. */
    xlRangeAutoFormatColor3            =     9, /*!< Color 3. */
    xlRangeAutoFormatList1             =    10, /*!< List 1. */
    xlRangeAutoFormatList2             =    11, /*!< List 2. */
    xlRangeAutoFormatList3             =    12, /*!< List 3. */
    xlRangeAutoFormatLocalFormat1      =    15, /*!< Local Format 1. */
    xlRangeAutoFormatLocalFormat2      =    16, /*!< Local Format 2. */
    xlRangeAutoFormatLocalFormat3      =    19, /*!< Local Format 3. */
    xlRangeAutoFormatLocalFormat4      =    20, /*!< Local Format 4. */
    xlRangeAutoFormatNone              = -4142, /*!< No specified format. */
    xlRangeAutoFormatPTNone            =    42, /*!< No specified PivotTable format. */
    xlRangeAutoFormatReport1           =    21, /*!< Report 1. */
    xlRangeAutoFormatReport10          =    30, /*!< Report 10. */
    xlRangeAutoFormatReport2           =    22, /*!< Report 2. */
    xlRangeAutoFormatReport3           =    23, /*!< Report 3. */
    xlRangeAutoFormatReport4           =    24, /*!< Report 4. */
    xlRangeAutoFormatReport5           =    25, /*!< Report 5. */
    xlRangeAutoFormatReport6           =    26, /*!< Report 6. */
    xlRangeAutoFormatReport7           =    27, /*!< Report 7. */
    xlRangeAutoFormatReport8           =    28, /*!< Report 8. */
    xlRangeAutoFormatReport9           =    29, /*!< Report 9. */
    xlRangeAutoFormatSimple            = -4154, /*!< Simple. */
    xlRangeAutoFormatTable1            =    32, /*!< Table 1. */
    xlRangeAutoFormatTable10           =    41, /*!< Table 10. */
    xlRangeAutoFormatTable2            =    33, /*!< Table 2. */
    xlRangeAutoFormatTable3            =    34, /*!< Table 3. */
    xlRangeAutoFormatTable4            =    35, /*!< Table 4. */
    xlRangeAutoFormatTable5            =    36, /*!< Table 5. */
    xlRangeAutoFormatTable6            =    37, /*!< Table 6. */
    xlRangeAutoFormatTable7            =    38, /*!< Table 7. */
    xlRangeAutoFormatTable8            =    39, /*!< Table 8. */
    xlRangeAutoFormatTable9            =    40, /*!< Table 9. */
};

/**
Specifies the range value data type.

[Official VBA documentation for XlRangeValueDataType](https://docs.microsoft.com/office/vba/api/excel.xlrangevaluedatatype)
*/
enum XlRangeValueDataType
{
    xlRangeValueDefault        = 10, /*!< Default. If the specified **Range** object is empty, returns the value Empty (use the IsEmpty function to test for this case). If the **Range** object contains more than one cell, returns an array of values (use the IsArray function to test for this case). */
    xlRangeValueMSPersistXML   = 12, /*!< Returns the recordset representation of the specified **Range** object in an XML format. */
    xlRangeValueXMLSpreadsheet = 11, /*!< Returns the values, formatting, formulas, and names of the specified **Range** object in the XML Spreadsheet format. */
};

/**
Specifies the reference style.

[Official VBA documentation for XlReferenceStyle](https://docs.microsoft.com/office/vba/api/excel.xlreferencestyle)
*/
enum XlReferenceStyle
{
    xlA1   =     1, /*!< Default. Use **xlA1** to return an A1-style reference. */
    xlR1C1 = -4150, /*!< Use **xlR1C1** to return an R1C1-style reference. */
};

/**
Specifies cell reference style when a formula is being converted.

[Official VBA documentation for XlReferenceType](https://docs.microsoft.com/office/vba/api/excel.xlreferencetype)
*/
enum XlReferenceType
{
    xlAbsolute        = 1, /*!< Convert to absolute row and column style. */
    xlAbsRowRelColumn = 2, /*!< Convert to absolute row and relative column style. */
    xlRelative        = 4, /*!< Convert to relative row and column style. */
    xlRelRowAbsColumn = 3, /*!< Convert to relative row and absolute column style. */
};

/**
Constants passed to and returned by the **Series.RegionLabelOptions** property.

[Official VBA documentation for XlRegionLabelOptions](https://docs.microsoft.com/office/vba/api/excel.xlregionlabeloptions)
*/
enum XlRegionLabelOptions
{
    xlRegionLabelOptionsNone        = 0, /*!< Don't show region labels. */
    xlRegionLabelOptionsBestFitOnly = 1, /*!< Only show labels that can be fitted within the bounds of their regions. */
    xlRegionLabelOptionsShowAll     = 2, /*!< Show all region labels. */
};

/**
Specifies the type information to be removed from the document information.

[Official VBA documentation for XlRemoveDocInfoType](https://docs.microsoft.com/office/vba/api/excel.xlremovedocinfotype)
*/
enum XlRemoveDocInfoType
{
    xlRDIAll                       = 99, /*!< Removes all documentation information. */
    xlRDIComments                  =  1, /*!< Removes comments from the document information. */
    xlRDIContentType               = 16, /*!< Removes content type data from the document information. */
    xlRDIDefinedNameComments       = 18, /*!< Removes defined name comments from the documentation information. */
    xlRDIDocumentManagementPolicy  = 15, /*!< Removes document management policy data from the document information. */
    xlRDIDocumentProperties        =  8, /*!< Removes document properties from the document information. */
    xlRDIDocumentServerProperties  = 14, /*!< Removes server properties from the document information. */
    xlRDIDocumentWorkspace         = 10, /*!< Removes workspace data from the document information. */
    xlRDIEmailHeader               =  5, /*!< Removes email headers from the document information. */
    xlRDIExcelDataModel            = 23, /*!< Removes Data Model data from the document information. */
    xlRDIInactiveDataConnections   = 19, /*!< Removes inactive data connection data from the document information. */
    xlRDIInkAnnotations            = 11, /*!< Removes ink annotations from the document information. */
    xlRDIInlineWebExtensions       = 21, /*!< Removes inline Web Extensions from the document information. */
    xlRDIPrinterPath               = 20, /*!< Removes printer paths from the document information. */
    xlRDIPublishInfo               = 13, /*!< Removes the publish information data from the document information. */
    xlRDIRemovePersonalInformation =  4, /*!< Removes personal information from the document information. */
    xlRDIRoutingSlip               =  6, /*!< Removes routing slip information from the document information. */
    xlRDIScenarioComments          = 12, /*!< Removes scenario comments from the document information. */
    xlRDISendForReview             =  7, /*!< Removes the send for review information from the document information. */
    xlRDITaskpaneWebExtensions     = 22, /*!< Removes task pane Web Extensions from the document information. */
};

/**
Specifies the RGB color.

[Official VBA documentation for XlRgbColor](https://docs.microsoft.com/office/vba/api/excel.xlrgbcolor)
*/
enum XlRgbColor
{
    rgbAliceBlue            = 16775408, /*!< Alice Blue */
    rgbAntiqueWhite         = 14150650, /*!< Antique White */
    rgbAqua                 = 16776960, /*!< Aqua */
    rgbAquamarine           = 13959039, /*!< Aquamarine */
    rgbAzure                = 16777200, /*!< Azure */
    rgbBeige                = 14480885, /*!< Beige */
    rgbBisque               = 12903679, /*!< Bisque */
    rgbBlack                =        0, /*!< Black */
    rgbBlanchedAlmond       = 13495295, /*!< Blanched Almond */
    rgbBlue                 = 16711680, /*!< Blue */
    rgbBlueViolet           = 14822282, /*!< Blue Violet */
    rgbBrown                =  2763429, /*!< Brown */
    rgbBurlyWood            =  8894686, /*!< Burly Wood */
    rgbCadetBlue            = 10526303, /*!< Cadet Blue */
    rgbChartreuse           =    65407, /*!< Chartreuse */
    rgbCoral                =  5275647, /*!< Coral */
    rgbCornflowerBlue       = 15570276, /*!< Cornflower Blue */
    rgbCornsilk             = 14481663, /*!< Cornsilk */
    rgbCrimson              =  3937500, /*!< Crimson */
    rgbDarkBlue             =  9109504, /*!< Dark Blue */
    rgbDarkCyan             =  9145088, /*!< Dark Cyan */
    rgbDarkGoldenrod        =   755384, /*!< Dark Goldenrod */
    rgbDarkGray             = 11119017, /*!< Dark Gray */
    rgbDarkGreen            =    25600, /*!< Dark Green */
    rgbDarkGrey             = 11119017, /*!< Dark Grey */
    rgbDarkKhaki            =  7059389, /*!< Dark Khaki */
    rgbDarkMagenta          =  9109643, /*!< Dark Magenta */
    rgbDarkOliveGreen       =  3107669, /*!< Dark Olive Green */
    rgbDarkOrange           =    36095, /*!< Dark Orange */
    rgbDarkOrchid           = 13382297, /*!< Dark Orchid */
    rgbDarkRed              =      139, /*!< Dark Red */
    rgbDarkSalmon           =  8034025, /*!< Dark Salmon */
    rgbDarkSeaGreen         =  9419919, /*!< Dark Sea Green */
    rgbDarkSlateBlue        =  9125192, /*!< Dark Slate Blue */
    rgbDarkSlateGray        =  5197615, /*!< Dark Slate Gray */
    rgbDarkSlateGrey        =  5197615, /*!< Dark Slate Grey */
    rgbDarkTurquoise        = 13749760, /*!< Dark Turquoise */
    rgbDarkViolet           = 13828244, /*!< Dark Violet */
    rgbDeepPink             =  9639167, /*!< Deep Pink */
    rgbDeepSkyBlue          = 16760576, /*!< Deep Sky Blue */
    rgbDimGray              =  6908265, /*!< Dim Gray */
    rgbDimGrey              =  6908265, /*!< Dim Grey */
    rgbDodgerBlue           = 16748574, /*!< Dodger Blue */
    rgbFireBrick            =  2237106, /*!< Fire Brick */
    rgbFloralWhite          = 15792895, /*!< Floral White */
    rgbForestGreen          =  2263842, /*!< Forest Green */
    rgbFuchsia              = 16711935, /*!< Fuchsia */
    rgbGainsboro            = 14474460, /*!< Gainsboro */
    rgbGhostWhite           = 16775416, /*!< Ghost White */
    rgbGold                 =    55295, /*!< Gold */
    rgbGoldenrod            =  2139610, /*!< Goldenrod */
    rgbGray                 =  8421504, /*!< Gray */
    rgbGreen                =    32768, /*!< Green */
    rgbGreenYellow          =  3145645, /*!< Green Yellow */
    rgbGrey                 =  8421504, /*!< Grey */
    rgbHoneydew             = 15794160, /*!< Honeydew */
    rgbHotPink              = 11823615, /*!< Hot Pink */
    rgbIndianRed            =  6053069, /*!< Indian Red */
    rgbIndigo               =  8519755, /*!< Indigo */
    rgbIvory                = 15794175, /*!< Ivory */
    rgbKhaki                =  9234160, /*!< Khaki */
    rgbLavender             = 16443110, /*!< Lavender */
    rgbLavenderBlush        = 16118015, /*!< Lavender Blush */
    rgbLawnGreen            =    64636, /*!< Lawn Green */
    rgbLemonChiffon         = 13499135, /*!< Lemon Chiffon */
    rgbLightBlue            = 15128749, /*!< Light Blue */
    rgbLightCoral           =  8421616, /*!< Light Coral */
    rgbLightCyan            =  9145088, /*!< Light Cyan */
    rgbLightGoldenrodYellow = 13826810, /*!< LightGoldenrodYellow */
    rgbLightGray            = 13882323, /*!< Light Gray */
    rgbLightGreen           =  9498256, /*!< Light Green */
    rgbLightGrey            = 13882323, /*!< Light Grey */
    rgbLightPink            = 12695295, /*!< Light Pink */
    rgbLightSalmon          =  8036607, /*!< Light Salmon */
    rgbLightSeaGreen        = 11186720, /*!< Light Sea Green */
    rgbLightSkyBlue         = 16436871, /*!< Light Sky Blue */
    rgbLightSlateGray       = 10061943, /*!< Light Slate Gray */
    rgbLightSteelBlue       = 14599344, /*!< Light Steel Blue */
    rgbLightYellow          = 14745599, /*!< Light Yellow */
    rgbLime                 =    65280, /*!< Lime */
    rgbLimeGreen            =  3329330, /*!< Lime Green */
    rgbLinen                = 15134970, /*!< Linen */
    rgbMaroon               =      128, /*!< Maroon */
    rgbMediumAquamarine     = 11206502, /*!< Medium Aquamarine */
    rgbMediumBlue           = 13434880, /*!< Medium Blue */
    rgbMediumOrchid         = 13850042, /*!< Medium Orchid */
    rgbMediumPurple         = 14381203, /*!< Medium Purple */
    rgbMediumSeaGreen       =  7451452, /*!< Medium Sea Green */
    rgbMediumSlateBlue      = 15624315, /*!< Medium Slate Blue */
    rgbMediumSpringGreen    = 10156544, /*!< Medium Spring Green */
    rgbMediumTurquoise      = 13422920, /*!< Medium Turquoise */
    rgbMediumVioletRed      =  8721863, /*!< Medium Violet Red */
    rgbMidnightBlue         =  7346457, /*!< Midnight Blue */
    rgbMintCream            = 16449525, /*!< Mint Cream */
    rgbMistyRose            = 14804223, /*!< Misty Rose */
    rgbMoccasin             = 11920639, /*!< Moccasin */
    rgbNavajoWhite          = 11394815, /*!< Navajo White */
    rgbNavy                 =  8388608, /*!< Navy */
    rgbNavyBlue             =  8388608, /*!< Navy Blue */
    rgbOldLace              = 15136253, /*!< Old Lace */
    rgbOlive                =    32896, /*!< Olive */
    rgbOliveDrab            =  2330219, /*!< Olive Drab */
    rgbOrange               =    42495, /*!< Orange */
    rgbOrangeRed            =    17919, /*!< Orange Red */
    rgbOrchid               = 14053594, /*!< Orchid */
    rgbPaleGoldenrod        =  7071982, /*!< Pale Goldenrod */
    rgbPaleGreen            = 10025880, /*!< Pale Green */
    rgbPaleTurquoise        = 15658671, /*!< Pale Turquoise */
    rgbPaleVioletRed        =  9662683, /*!< Pale Violet Red */
    rgbPapayaWhip           = 14020607, /*!< Papaya Whip */
    rgbPeachPuff            = 12180223, /*!< Peach Puff */
    rgbPeru                 =  4163021, /*!< Peru */
    rgbPink                 = 13353215, /*!< Pink */
    rgbPlum                 = 14524637, /*!< Plum */
    rgbPowderBlue           = 15130800, /*!< Powder Blue */
    rgbPurple               =  8388736, /*!< Purple */
    rgbRed                  =      255, /*!< Red */
    rgbRosyBrown            =  9408444, /*!< Rosy Brown */
    rgbRoyalBlue            = 14772545, /*!< Royal Blue */
    rgbSalmon               =  7504122, /*!< Salmon */
    rgbSandyBrown           =  6333684, /*!< Sandy Brown */
    rgbSeaGreen             =  5737262, /*!< Sea Green */
    rgbSeashell             = 15660543, /*!< Seashell */
    rgbSienna               =  2970272, /*!< Sienna */
    rgbSilver               = 12632256, /*!< Silver */
    rgbSkyBlue              = 15453831, /*!< Sky Blue */
    rgbSlateBlue            = 13458026, /*!< Slate Blue */
    rgbSlateGray            =  9470064, /*!< Slate Gray */
    rgbSnow                 = 16448255, /*!< Snow */
    rgbSpringGreen          =  8388352, /*!< Spring Green */
    rgbSteelBlue            = 11829830, /*!< Steel Blue */
    rgbTan                  =  9221330, /*!< Tan */
    rgbTeal                 =  8421376, /*!< Teal */
    rgbThistle              = 14204888, /*!< Thistle */
    rgbTomato               =  4678655, /*!< Tomato */
    rgbTurquoise            = 13688896, /*!< Turquoise */
    rgbViolet               = 15631086, /*!< Violet */
    rgbWheat                = 11788021, /*!< Wheat */
    rgbWhite                = 16777215, /*!< White */
    rgbWhiteSmoke           = 16119285, /*!< White Smoke */
    rgbYellow               =    65535, /*!< Yellow */
    rgbYellowGreen          =  3329434, /*!< Yellow Green */
};

/**
Specifies how the PivotTable cache or a [query table](excel.querytable.md) connects to its data source.

[Official VBA documentation for XlRobustConnect](https://docs.microsoft.com/office/vba/api/excel.xlrobustconnect)
*/
enum XlRobustConnect
{
    xlAlways     = 1, /*!< The PivotTable cache or query table always uses external source information (as defined by the **[SourceConnectionFile](Excel.PivotCache.SourceConnectionFile.md)** or **[SourceDataFile](Excel.PivotCache.SourceDataFile.md)** property) to reconnect. */
    xlAsRequired = 0, /*!< The PivotTable cache or query table uses external source information to reconnect by using the **[Connection](Excel.PivotCache.Connection.md)** property. */
    xlNever      = 2, /*!< The PivotTable cache or query table never uses source information to reconnect. */
};

/**
Specifies whether the values corresponding to a particular data series are in rows or columns.

[Official VBA documentation for XlRowCol](https://docs.microsoft.com/office/vba/api/excel.xlrowcol)
*/
enum XlRowCol
{
    xlColumns = 2, /*!< Data series is in a row. */
    xlRows    = 1, /*!< Data series is in a column. */
};

/**
Specifies the automatic macro to run.

[Official VBA documentation for XlRunAutoMacro](https://docs.microsoft.com/office/vba/api/excel.xlrunautomacro)
*/
enum XlRunAutoMacro
{
    xlAutoActivate   = 3, /*!< Auto_Activate macros */
    xlAutoClose      = 2, /*!< Auto_Close macros */
    xlAutoDeactivate = 4, /*!< Auto_Deactivate macros */
    xlAutoOpen       = 1, /*!< Auto_Open macros */
};

/**
Specifies during file close if the file will be saved.

[Official VBA documentation for XlSaveAction](https://docs.microsoft.com/office/vba/api/excel.xlsaveaction)
*/
enum XlSaveAction
{
    xlDoNotSaveChanges = 2, /*!< Changes will not be saved. */
    xlSaveChanges      = 1, /*!< Changes will be saved. */
};

/**
Specifies the access mode for the Save As function.

[Official VBA documentation for XlSaveAsAccessMode](https://docs.microsoft.com/office/vba/api/excel.xlsaveasaccessmode)
*/
enum XlSaveAsAccessMode
{
    xlExclusive = 3, /*!< Exclusive mode */
    xlNoChange  = 1, /*!< Default (does not change the access mode) */
    xlShared    = 2, /*!< Share list */
};

/**
Specifies the way that conflicts are to be resolved whenever a shared workbook is updated.

[Official VBA documentation for XlSaveConflictResolution](https://docs.microsoft.com/office/vba/api/excel.xlsaveconflictresolution)
*/
enum XlSaveConflictResolution
{
    xlLocalSessionChanges = 2, /*!< The local user's changes are always accepted. */
    xlOtherSessionChanges = 3, /*!< The local user's changes are always rejected. */
    xlUserResolution      = 1, /*!< A dialog box asks the user to resolve the conflict. */
};

/**
Specifies the scale type of the value axis.

[Official VBA documentation for XlScaleType](https://docs.microsoft.com/office/vba/api/excel.xlscaletype)
*/
enum XlScaleType
{
    xlScaleLinear      = -4132, /*!< Linear */
    xlScaleLogarithmic = -4133, /*!< Logarithmic */
};

/**
Specifies the search direction when searching a range.

[Official VBA documentation for XlSearchDirection](https://docs.microsoft.com/office/vba/api/excel.xlsearchdirection)
*/
enum XlSearchDirection
{
    xlNext     = 1, /*!< Search for next matching value in range. */
    xlPrevious = 2, /*!< Search for previous matching value in range. */
};

/**
Specifies the order in which to search the range.

[Official VBA documentation for XlSearchOrder](https://docs.microsoft.com/office/vba/api/excel.xlsearchorder)
*/
enum XlSearchOrder
{
    xlByColumns = 2, /*!< Searches down through a column, then moves to the next column. */
    xlByRows    = 1, /*!< Searches across a row, then moves to the next row. */
};

/**
Specifies the extent of the search for the range.

[Official VBA documentation for XlSearchWithin](https://docs.microsoft.com/office/vba/api/excel.xlsearchwithin)
*/
enum XlSearchWithin
{
    xlWithinSheet    = 1, /*!< Limit search to current sheet. */
    xlWithinWorkbook = 2, /*!< Search whole workbook. */
};

/**
Specifies the series labels for the series label levels.

[Official VBA documentation for XlSeriesNameLevel](https://docs.microsoft.com/office/vba/api/excel.xlseriesnamelevel)
*/
enum XlSeriesNameLevel
{
    xlSeriesNameLevelAll    = -1, /*!< Set series names to all series name levels w/in range on the chart. */
    xlSeriesNameLevelCustom = -2, /*!< Indicates literal data in the series names. */
    xlSeriesNameLevelNone   = -3, /*!< Set no category labels in the chart. Defaults to automatic indexed labels. */
};

/**
Specifies the worksheet type.

[Official VBA documentation for XlSheetType](https://docs.microsoft.com/office/vba/api/excel.xlsheettype)
*/
enum XlSheetType
{
    xlChart                = -4109, /*!< Chart */
    xlDialogSheet          = -4116, /*!< Dialog sheet */
    xlExcel4IntlMacroSheet =     4, /*!< Excel version 4 international macro sheet */
    xlExcel4MacroSheet     =     3, /*!< Excel version 4 macro sheet */
    xlWorksheet            = -4167, /*!< Worksheet */
};

/**
Specifies whether the object is visible.

[Official VBA documentation for XlSheetVisibility](https://docs.microsoft.com/office/vba/api/excel.xlsheetvisibility)
*/
enum XlSheetVisibility
{
    xlSheetHidden     =  0, /*!< Hides the worksheet which the user can unhide via menu. */
    xlSheetVeryHidden =  2, /*!< Hides the object so that the only way for you to make it visible again is by setting this property to True (the user cannot make the object visible). */
    xlSheetVisible    = -1, /*!< Displays the sheet. */
};

/**
Specifies what the bubble size represents on a bubble chart.

[Official VBA documentation for XlSizeRepresents](https://docs.microsoft.com/office/vba/api/excel.xlsizerepresents)
*/
enum XlSizeRepresents
{
    xlSizeIsArea  = 1, /*!< Area of the bubble. */
    xlSizeIsWidth = 2, /*!< Width of the bubble. */
};

/**
Designates the type of slicer or slicer cache.

[Official VBA documentation for XlSlicerCacheType](https://docs.microsoft.com/office/vba/api/excel.xlslicercachetype)
*/
enum XlSlicerCacheType
{
    xlSlicer   = 1, /*!< Slicer cache represents a Slicer. */
    xlTimeline = 2, /*!< Slicer cache represents a Timeline. */
};

/**
Specifies the type of cross filtering used by the specified slicer cache and how it is visualized.

[Official VBA documentation for XlSlicerCrossFilterType](https://docs.microsoft.com/office/vba/api/excel.xlslicercrossfiltertype)
*/
enum XlSlicerCrossFilterType
{
    xlSlicerCrossFilterHideButtonsWithNoData  = 4, /*!< Cross filtering is turned on for this slicer cache, any tile with no data for a filtering selection in other slicers connected to the same data source will be dimmed. Additionally, buttons will be hidden. */
    xlSlicerCrossFilterShowItemsWithDataAtTop = 2, /*!< Cross filtering is turned on for this slicer cache, any tile with no data for a filtering selection in other slicers connected to the same data source will be dimmed. Additionally, tiles with data are moved to the top in the slicer. (Default) */
    xlSlicerCrossFilterShowItemsWithNoData    = 3, /*!< Cross filtering is turned on for this slicer cache, any tile with no data for a filtering selection in other slicers connected to the same data source will be dimmed. */
    xlSlicerNoCrossFilter                     = 1, /*!< Cross filtering is turned off entirely, so all tiles are displayed and active (not dimmed) regardless of filtering selections in other slicers. */
};

/**
Specifies whether items displayed in the slicer are sorted, and if they are sorted, whether they are sorted in ascending or descending order by item captions.

[Official VBA documentation for XlSlicerSort](https://docs.microsoft.com/office/vba/api/excel.xlslicersort)
*/
enum XlSlicerSort
{
    xlSlicerSortAscending       = 2, /*!< Slicer items are sorted in ascending order by item captions. */
    xlSlicerSortDataSourceOrder = 1, /*!< Slicer items are displayed in the order provided by the data source. */
    xlSlicerSortDescending      = 3, /*!< Slicer items are sorted in descending order by item captions. */
};

/**
Specifies how to sort text.

[Official VBA documentation for XlSortDataOption](https://docs.microsoft.com/office/vba/api/excel.xlsortdataoption)
*/
enum XlSortDataOption
{
    xlSortNormal        = 0, /*!< default. Sorts numeric and text data separately. */
    xlSortTextAsNumbers = 1, /*!< Treat text as numeric data for the sort. */
};

/**
Specifies the type of sort.

[Official VBA documentation for XlSortMethod](https://docs.microsoft.com/office/vba/api/excel.xlsortmethod)
*/
enum XlSortMethod
{
    xlPinYin = 1, /*!< Phonetic Chinese sort order for characters. This is the default value. */
    xlStroke = 2, /*!< Sort by the quantity of strokes in each character. */
};

/**
Specifies how to sort when using East Asian sorting methods.

[Official VBA documentation for XlSortMethodOld](https://docs.microsoft.com/office/vba/api/excel.xlsortmethodold)
*/
enum XlSortMethodOld
{
    xlCodePage  = 2, /*!< Sort by code page. */
    xlSyllabary = 1, /*!< Sort phonetically. */
};

/**
Specifies the parameter on which the data should be sorted.

[Official VBA documentation for XlSortOn](https://docs.microsoft.com/office/vba/api/excel.xlsorton)
*/
enum XlSortOn
{
    SortOnCellColor = 1, /*!< Cell color. */
    SortOnFontColor = 2, /*!< Font color. */
    SortOnIcon      = 3, /*!< Icon. */
    SortOnValues    = 0, /*!< Values. */
};

/**
Specifies the sort order for the specified field or range.

[Official VBA documentation for XlSortOrder](https://docs.microsoft.com/office/vba/api/excel.xlsortorder)
*/
enum XlSortOrder
{
    xlAscending  =     1, /*!< Sorts the specified field in ascending order. This is the default value. */
    xlDescending =     2, /*!< Sorts the specified field in descending order. */
    xlManual     = -4135, /*!< Manual sort (you can drag items to rearrange them). */
};

/**
Specifies the sort orientation.

[Official VBA documentation for XlSortOrientation](https://docs.microsoft.com/office/vba/api/excel.xlsortorientation)
*/
enum XlSortOrientation
{
    xlSortColumns = 1, /*!< Sorts by column. */
    xlSortRows    = 2, /*!< Sorts by row. This is the default value. */
};

/**
Specifies which elements are to be sorted. Use this argument only when sorting PivotTable reports.

[Official VBA documentation for XlSortType](https://docs.microsoft.com/office/vba/api/excel.xlsorttype)
*/
enum XlSortType
{
    xlSortLabels = 2, /*!< Sorts the PivotTable report by labels. */
    xlSortValues = 1, /*!< Sorts the PivotTable report by values. */
};

/**
Identifies the source object.

[Official VBA documentation for XlSourceType](https://docs.microsoft.com/office/vba/api/excel.xlsourcetype)
*/
enum XlSourceType
{
    xlSourceAutoFilter = 3, /*!< An AutoFilter range */
    xlSourceChart      = 5, /*!< A chart */
    xlSourcePivotTable = 6, /*!< A PivotTable report */
    xlSourcePrintArea  = 2, /*!< A range of cells selected for printing */
    xlSourceQuery      = 7, /*!< A query table (external data range) */
    xlSourceRange      = 4, /*!< A range of cells */
    xlSourceSheet      = 1, /*!< An entire worksheet */
    xlSourceWorkbook   = 0, /*!< A workbook */
};

/**
Specifies the mode for checking the spelling of Spanish.

[Official VBA documentation for XlSpanishModes](https://docs.microsoft.com/office/vba/api/excel.xlspanishmodes)
*/
enum XlSpanishModes
{
    xlSpanishTuteoAndVoseo = 1, /*!< Tuteo and Voseo verb forms. */
    xlSpanishTuteoOnly     = 0, /*!< Tuteo verb forms only. */
    xlSpanishVoseoOnly     = 2, /*!< Voseo verb forms only. */
};

/**
Specifies how the minimum or maximum value of the vertical axis of the sparkline is scaled relative to other sparklines in the group.

[Official VBA documentation for XlSparkScale](https://docs.microsoft.com/office/vba/api/excel.xlsparkscale)
*/
enum XlSparkScale
{
    xlSparkScaleCustom = 3, /*!< The minimum or maximum value for the vertical axis of the sparkline has a user-defined value. */
    xlSparkScaleGroup  = 1, /*!< The minimum or maximum value for the vertical axes of all of the sparklines in the group have the same value. */
    xlSparkScaleSingle = 2, /*!< The minimum or maximum value for the vertical axis of each sparkline in the group is automatically set to its own calculated value. */
};

/**
Specifies the type of sparkline.

[Official VBA documentation for XlSparkType](https://docs.microsoft.com/office/vba/api/excel.xlsparktype)
*/
enum XlSparkType
{
    xlSparkColumn           = 2, /*!< A column chart sparkline. */
    xlSparkColumnStacked100 = 3, /*!< A win/loss chart sparkline. */
    xlSparkLine             = 1, /*!< A line chart sparkline. */
};

/**
Specifies how to plot the sparkline when the data on which it is based is in a square-shaped range.

[Official VBA documentation for XlSparklineRowCol](https://docs.microsoft.com/office/vba/api/excel.xlsparklinerowcol)
*/
enum XlSparklineRowCol
{
    xlSparklineColumnsSquare = 2, /*!< Plot the data by columns. */
    xlSparklineNonSquare     = 0, /*!< The sparkline is not bound to data in a square-shaped range. */
    xlSparklineRowsSquare    = 1, /*!< Plot the data by rows. */
};

/**
Specifies the order in which the cells are spoken.

[Official VBA documentation for XlSpeakDirection](https://docs.microsoft.com/office/vba/api/excel.xlspeakdirection)
*/
enum XlSpeakDirection
{
    xlSpeakByColumns = 1, /*!< Reads down a column, then moves to the next column. */
    xlSpeakByRows    = 0, /*!< Reads across a row, then moves to the next row. */
};

/**
Specifies cells with a particular type of value to include in the result.

[Official VBA documentation for XlSpecialCellsValue](https://docs.microsoft.com/office/vba/api/excel.xlspecialcellsvalue)
*/
enum XlSpecialCellsValue
{
    xlErrors     = 16, /*!< Cells with errors. */
    xlLogical    =  4, /*!< Cells with logical values. */
    xlNumbers    =  1, /*!< Cells with numeric values. */
    xlTextValues =  2, /*!< Cells with text. */
};

/**
Specifies the standard color scale.

[Official VBA documentation for XlStdColorScale](https://docs.microsoft.com/office/vba/api/excel.xlstdcolorscale)
*/
enum XlStdColorScale
{
    ColorScaleBlackWhite = 3, /*!< Black over White. */
    ColorScaleGYR        = 2, /*!< GYR. */
    ColorScaleRYG        = 1, /*!< RYG. */
    ColorScaleWhiteBlack = 4, /*!< White over Black. */
};

/**
Specifies the format to use when subscribing to a published edition.

[Official VBA documentation for XlSubscribeToFormat](https://docs.microsoft.com/office/vba/api/excel.xlsubscribetoformat)
*/
enum XlSubscribeToFormat
{
    xlSubscribeToPicture = -4147, /*!< Picture */
    xlSubscribeToText    = -4158, /*!< Text */
};

/**
Specifies where the subtotal will be displayed on the worksheet.

[Official VBA documentation for XlSubtotalLocationType](https://docs.microsoft.com/office/vba/api/excel.xlsubtotallocationtype)
*/
enum XlSubtotalLocationType
{
    xlAtBottom = 2, /*!< Subtotal will be at the bottom. */
    xlAtTop    = 1, /*!< Subtotal will be at the top. */
};

/**
Specifies the location of the summary columns in the outline.

[Official VBA documentation for XlSummaryColumn](https://docs.microsoft.com/office/vba/api/excel.xlsummarycolumn)
*/
enum XlSummaryColumn
{
    xlSummaryOnLeft  = -4131, /*!< The summary column will be positioned to the left of the detail columns in the outline. */
    xlSummaryOnRight = -4152, /*!< The summary column will be positioned to the right of the detail columns in the outline. */
};

/**
Specifies the type of summary to be created for scenarios.

[Official VBA documentation for XlSummaryReportType](https://docs.microsoft.com/office/vba/api/excel.xlsummaryreporttype)
*/
enum XlSummaryReportType
{
    xlStandardSummary   =     1, /*!< List scenarios side by side. */
    xlSummaryPivotTable = -4148, /*!< Display scenarios in a PivotTable report. */
};

/**
Specifies the location of the summary rows in the outline.

[Official VBA documentation for XlSummaryRow](https://docs.microsoft.com/office/vba/api/excel.xlsummaryrow)
*/
enum XlSummaryRow
{
    xlSummaryAbove = 0, /*!< The summary row will be positioned above the detail rows in the outline. */
    xlSummaryBelow = 1, /*!< The summary row will be positioned below the detail rows in the outline. */
};

/**
Specifies the first or last tab position.

[Official VBA documentation for XlTabPosition](https://docs.microsoft.com/office/vba/api/excel.xltabposition)
*/
enum XlTabPosition
{
    xlTabPositionFirst = 0, /*!< First tab position. */
    xlTabPositionLast  = 1, /*!< Last tab position. */
};

/**
Specifies the table style element used.

[Official VBA documentation for XlTableStyleElementType](https://docs.microsoft.com/office/vba/api/excel.xltablestyleelementtype)
*/
enum XlTableStyleElementType
{
    xlBlankRow                              = 19, /*!< Blank row */
    xlColumnStripe1                         =  7, /*!< Column Stripe1 */
    xlColumnStripe2                         =  8, /*!< Column Stripe2 */
    xlColumnSubheading1                     = 20, /*!< Column Subheading1 */
    xlColumnSubheading2                     = 21, /*!< Column Subheading2 */
    xlColumnSubheading3                     = 22, /*!< Column Subheading3 */
    xlFirstColumn                           =  3, /*!< First column */
    xlFirstHeaderCell                       =  9, /*!< First header cell */
    xlFirstTotalCell                        = 11, /*!< First total cell */
    xlGrandTotalColumn                      =  4, /*!< Grand total column */
    xlGrandTotalRow                         =  2, /*!< Grand total row */
    xlHeaderRow                             =  1, /*!< Header row */
    xlLastColumn                            =  4, /*!< Last column */
    xlLastHeaderCell                        = 10, /*!< Last header cell */
    xlLastTotalCell                         = 12, /*!< Last total cell */
    xlPageFieldLabels                       = 26, /*!< Page field labels */
    xlPageFieldValues                       = 27, /*!< Page field values */
    xlRowStripe1                            =  5, /*!< Row Stripe1 */
    xlRowStripe2                            =  6, /*!< Row Stripe2 */
    xlRowSubheading1                        = 23, /*!< Row Subheading1 */
    xlRowSubheading2                        = 24, /*!< Row Subheading2 */
    xlRowSubheading3                        = 25, /*!< Row Subheading3 */
    xlSlicerHoveredSelectedItemWithData     = 33, /*!< A selected item, hovered over by the user, that contains data. */
    xlSlicerHoveredSelectedItemWithNoData   = 35, /*!< A selected item, hovered over by the user, that does not contain data. */
    xlSlicerHoveredUnselectedItemWithData   = 32, /*!< An item, hovered over by the user, that is not selected and that contains data. */
    xlSlicerHoveredUnselectedItemWithNoData = 34, /*!< A selected item, hovered over by the user, that is not selected and that does not contain data. */
    xlSlicerSelectedItemWithData            = 30, /*!< A selected item that contains data. */
    xlSlicerSelectedItemWithNoData          = 31, /*!< A selected item that does not contain data. */
    xlSlicerUnselectedItemWithData          = 28, /*!< An item that is not selected that contains data. */
    xlSlicerUnselectedItemWithNoData        = 29, /*!< An item that is not selected that does not contain data. */
    xlSubtotalColumn1                       = 13, /*!< Subtotal Column1 */
    xlSubtotalColumn2                       = 14, /*!< Subtotal Column2 */
    xlSubtotalColumn3                       = 15, /*!< Subtotal Column3 */
    xlSubtotalRow1                          = 16, /*!< Subtotal Row1 */
    xlSubtotalRow2                          = 17, /*!< Subtotal Row2 */
    xlSubtotalRow3                          = 18, /*!< Subtotal Row3 */
    xlTimelinePeriodLabels1                 = 38, /*!< Timeline Period Label */
    xlTimelinePeriodLabels2                 = 39, /*!< Additional Timeline Period Label */
    xlTimelineSelectedTimeBlock             = 40, /*!< Selected Timeline Time Block */
    xlTimelineSelectedTimeBlockSpace        = 42, /*!< Selected Timeline Time Block space */
    xlTimelineSelectionLabel                = 36, /*!< Timeline Selection Label */
    xlTimelineTimeLevel                     = 37, /*!< Timeline Level */
    xlTimelineUnselectedTimeBlock           = 41, /*!< Unselected Timeline Time Block */
    xlTotalRow                              =  2, /*!< Total Row */
    xlWholeTable                            =  0, /*!< Whole Table */
};

/**
Specifies the column format for the data in the text file that you are importing into a query table.

[Official VBA documentation for XlTextParsingType](https://docs.microsoft.com/office/vba/api/excel.xltextparsingtype)
*/
enum XlTextParsingType
{
    xlDelimited  = 1, /*!< Default. Indicates that the file is delimited by delimiter characters. */
    xlFixedWidth = 2, /*!< Indicates that the data in the file is arranged in columns of fixed widths. */
};

/**
Specifies the delimiter to use to specify text.

[Official VBA documentation for XlTextQualifier](https://docs.microsoft.com/office/vba/api/excel.xltextqualifier)
*/
enum XlTextQualifier
{
    xlTextQualifierDoubleQuote =     1, /*!< Double quotation mark ("). */
    xlTextQualifierNone        = -4142, /*!< No delimiter. */
    xlTextQualifierSingleQuote =     2, /*!< Single quotation mark ('). */
};

/**
Specifies whether the visual layout of the text being imported is left-to-right or right-to-left.

[Official VBA documentation for XlTextVisualLayoutType](https://docs.microsoft.com/office/vba/api/excel.xltextvisuallayouttype)
*/
enum XlTextVisualLayoutType
{
    xlTextVisualLTR = 1, /*!< Left-to-right */
    xlTextVisualRTL = 2, /*!< Right-to-left */
};

/**
Specifies the theme color to be used.

[Official VBA documentation for XlThemeColor](https://docs.microsoft.com/office/vba/api/excel.xlthemecolor)
*/
enum XlThemeColor
{
    xlThemeColorAccent1           =  5, /*!< Accent1 */
    xlThemeColorAccent2           =  6, /*!< Accent2 */
    xlThemeColorAccent3           =  7, /*!< Accent3 */
    xlThemeColorAccent4           =  8, /*!< Accent4 */
    xlThemeColorAccent5           =  9, /*!< Accent5 */
    xlThemeColorAccent6           = 10, /*!< Accent6 */
    xlThemeColorDark1             =  1, /*!< Dark1 */
    xlThemeColorDark2             =  3, /*!< Dark2 */
    xlThemeColorFollowedHyperlink = 12, /*!< Followed hyperlink */
    xlThemeColorHyperlink         = 11, /*!< Hyperlink */
    xlThemeColorLight1            =  2, /*!< Light1 */
    xlThemeColorLight2            =  4, /*!< Light2 */
};

/**
Specifies the theme font to use.

[Official VBA documentation for XlThemeFont](https://docs.microsoft.com/office/vba/api/excel.xlthemefont)
*/
enum XlThemeFont
{
    xlThemeFontMajor = 2, /*!< Major. */
    xlThemeFontMinor = 1, /*!< Minor. */
    xlThemeFontNone  = 0, /*!< Do not use any theme font. */
};

/**
Specifies the control over the multi-threaded calculation mode.

[Official VBA documentation for XlThreadMode](https://docs.microsoft.com/office/vba/api/excel.xlthreadmode)
*/
enum XlThreadMode
{
    xlThreadModeAutomatic = 0, /*!< Multi-threaded calculation mode is automatic. */
    xlThreadModeManual    = 1, /*!< Multi-threaded calculation mode is manual. */
};

/**
Specifies the text orientation for tick-mark labels.

[Official VBA documentation for XlTickLabelOrientation](https://docs.microsoft.com/office/vba/api/excel.xlticklabelorientation)
*/
enum XlTickLabelOrientation
{
    xlTickLabelOrientationAutomatic  = -4105, /*!< Text orientation set by Excel. */
    xlTickLabelOrientationDownward   = -4170, /*!< Text runs down. */
    xlTickLabelOrientationHorizontal = -4128, /*!< Characters run horizontally. */
    xlTickLabelOrientationUpward     = -4171, /*!< Text runs up. */
    xlTickLabelOrientationVertical   = -4166, /*!< Characters run vertically. */
};

/**
Specifies the position of tick-mark labels on the specified axis.

[Official VBA documentation for XlTickLabelPosition](https://docs.microsoft.com/office/vba/api/excel.xlticklabelposition)
*/
enum XlTickLabelPosition
{
    xlTickLabelPositionHigh       = -4127, /*!< Top or right side of the chart. */
    xlTickLabelPositionLow        = -4134, /*!< Bottom or left side of the chart. */
    xlTickLabelPositionNextToAxis =     4, /*!< Next to axis (where axis is not at either side of the chart). */
    xlTickLabelPositionNone       = -4142, /*!< No tick marks. */
};

/**
Specifies the position of major and minor tick marks for an axis.

[Official VBA documentation for XlTickMark](https://docs.microsoft.com/office/vba/api/excel.xltickmark)
*/
enum XlTickMark
{
    xlTickMarkCross   =     4, /*!< Crosses the axis */
    xlTickMarkInside  =     2, /*!< Inside the axis */
    xlTickMarkNone    = -4142, /*!< No mark */
    xlTickMarkOutside =     3, /*!< Outside the axis */
};

/**
Specifies the time period.

[Official VBA documentation for XlTimePeriods](https://docs.microsoft.com/office/vba/api/excel.xltimeperiods)
*/
enum XlTimePeriods
{
    xlLast7Days = 2, /*!< Last 7 days */
    xlLastMonth = 5, /*!< Last month */
    xlLastWeek  = 4, /*!< Last week */
    xlNextMonth = 8, /*!< Next month */
    xlNextWeek  = 7, /*!< Next week */
    xlThisMonth = 9, /*!< This month */
    xlThisWeek  = 3, /*!< This week */
    xlToday     = 0, /*!< Today */
    xlTomorrow  = 6, /*!< Tomorrow */
    xlYesterday = 1, /*!< Yesterday */
};

/**
Specifies the unit of time for chart axes and data series.

[Official VBA documentation for XlTimeUnit](https://docs.microsoft.com/office/vba/api/excel.xltimeunit)
*/
enum XlTimeUnit
{
    xlDays   = 0, /*!< Days */
    xlMonths = 1, /*!< Months */
    xlYears  = 2, /*!< Years */
};

/**
One of the built-in hierarchy levels that Timeline supports.

[Official VBA documentation for XlTimelineLevel](https://docs.microsoft.com/office/vba/api/excel.xltimelinelevel)
*/
enum XlTimelineLevel
{
    xlTimelineLevelYears    = 0, /*!< Years level */
    xlTimelineLevelQuarters = 1, /*!< Quarters level */
    xlTimelineLevelMonths   = 2, /*!< Months level */
    xlTimelineLevelDays     = 3, /*!< Days level */
};

/**
Specifies which properties of a toolbar are restricted. Options can be combined using Or.

[Official VBA documentation for XlToolbarProtection](https://docs.microsoft.com/office/vba/api/excel.xltoolbarprotection)
*/
enum XlToolbarProtection
{
    xlNoButtonChanges       =     1, /*!< No button changes permitted. */
    xlNoChanges             =     4, /*!< No changes of any kind. */
    xlNoDockingChanges      =     3, /*!< No changes to toolbar's docking position. */
    xlNoShapeChanges        =     2, /*!< No changes to toolbar shape. */
    xlToolbarProtectionNone = -4143, /*!< All changes permitted. */
};

/**
Specifies the top 10 values from the top or bottom of a series of values.

[Official VBA documentation for XlTopBottom](https://docs.microsoft.com/office/vba/api/excel.xltopbottom)
*/
enum XlTopBottom
{
    xlTop10Bottom = 0, /*!< Top 10 bottom values */
    xlTop10Top    = 1, /*!< Top 10 values */
};

/**
Specifies the type of calculation in the Totals row of a list column.

[Official VBA documentation for XlTotalsCalculation](https://docs.microsoft.com/office/vba/api/excel.xltotalscalculation)
*/
enum XlTotalsCalculation
{
    xlTotalsCalculationAverage   = 2, /*!< Average */
    xlTotalsCalculationCount     = 3, /*!< Count of non-empty cells */
    xlTotalsCalculationCountNums = 4, /*!< Count of cells with numeric values */
    xlTotalsCalculationCustom    = 9, /*!< Custom calculation */
    xlTotalsCalculationMax       = 6, /*!< Maximum value in the list */
    xlTotalsCalculationMin       = 5, /*!< Minimum value in the list */
    xlTotalsCalculationNone      = 0, /*!< No calculation */
    xlTotalsCalculationStdDev    = 7, /*!< Standard deviation value */
    xlTotalsCalculationSum       = 1, /*!< Sum of all values in the list column */
    xlTotalsCalculationVar       = 8, /*!< Variable */
};

/**
Specifies how the trendline that smooths out fluctuations in the data is calculated.

[Official VBA documentation for XlTrendlineType](https://docs.microsoft.com/office/vba/api/excel.xltrendlinetype)
*/
enum XlTrendlineType
{
    xlExponential =     5, /*!< Uses an equation to calculate the least squares fit through points, for example, y=ab^x . */
    xlLinear      = -4132, /*!< Uses the linear equation y = mx + b to calculate the least squares fit through points. */
    xlLogarithmic = -4133, /*!< Uses the equation y = c ln x + b to calculate the least squares fit through points. */
    xlMovingAvg   =     6, /*!< Uses a sequence of averages computed from parts of the data series. The number of points equals the total number of points in the series less the number specified for the period. */
    xlPolynomial  =     3, /*!< Uses an equation to calculate the least squares fit through points, for example, y = ax^6 + bx^5 + cx^4 + dx^3 + ex^2 + fx + g. */
    xlPower       =     4, /*!< Uses an equation to calculate the least squares fit through points, for example, y = ax^b. */
};

/**
Specifies the type of underline applied to a font.

[Official VBA documentation for XlUnderlineStyle](https://docs.microsoft.com/office/vba/api/excel.xlunderlinestyle)
*/
enum XlUnderlineStyle
{
    xlUnderlineStyleDouble           = -4119, /*!< Double thick underline. */
    xlUnderlineStyleDoubleAccounting =     5, /*!< Two thin underlines placed close together. */
    xlUnderlineStyleNone             = -4142, /*!< No underlining. */
    xlUnderlineStyleSingle           =     2, /*!< Single underlining. */
    xlUnderlineStyleSingleAccounting =     4, /*!< Not supported. */
};

/**
Specifies a workbook's setting for updating embedded OLE links.

[Official VBA documentation for XlUpdateLinks](https://docs.microsoft.com/office/vba/api/excel.xlupdatelinks)
*/
enum XlUpdateLinks
{
    xlUpdateLinksAlways      = 3, /*!< Embedded OLE links are always updated for the specified workbook. */
    xlUpdateLinksNever       = 2, /*!< Embedded OLE links are never updated for the specified workbook. */
    xlUpdateLinksUserSetting = 1, /*!< Embedded OLE links are updated according to the user's settings for the specified workbook. */
};

/**
Specifies the vertical alignment for the object.

[Official VBA documentation for XlVAlign](https://docs.microsoft.com/office/vba/api/excel.xlvalign)
*/
enum XlVAlign
{
    xlVAlignBottom      = -4107, /*!< Bottom */
    xlVAlignCenter      = -4108, /*!< Center */
    xlVAlignDistributed = -4117, /*!< Distributed */
    xlVAlignJustify     = -4130, /*!< Justify */
    xlVAlignTop         = -4160, /*!< Top */
};

/**
Specifies the type of workbook to create. The new workbook contains a single sheet of the specified type.

[Official VBA documentation for XlWBATemplate](https://docs.microsoft.com/office/vba/api/excel.xlwbatemplate)
*/
enum XlWBATemplate
{
    xlWBATChart                = -4109, /*!< Chart */
    xlWBATExcel4IntlMacroSheet =     4, /*!< Excel version 4 macro */
    xlWBATExcel4MacroSheet     =     3, /*!< Excel version 4 international macro */
    xlWBATWorksheet            = -4167, /*!< Worksheet */
};

/**
Specifies how much formatting from a webpage, if any, is applied when a webpage is imported into a query table.

[Official VBA documentation for XlWebFormatting](https://docs.microsoft.com/office/vba/api/excel.xlwebformatting)
*/
enum XlWebFormatting
{
    xlWebFormattingAll  = 1, /*!< All formatting is imported. */
    xlWebFormattingNone = 3, /*!< No formatting is imported. */
    xlWebFormattingRTF  = 2, /*!< Rich Text Format - compatible formatting is imported. */
};

/**
Specifies whether an entire webpage, all tables on the webpage, or only a specific table is imported into a query table.

[Official VBA documentation for XlWebSelectionType](https://docs.microsoft.com/office/vba/api/excel.xlwebselectiontype)
*/
enum XlWebSelectionType
{
    xlAllTables       = 2, /*!< All tables */
    xlEntirePage      = 1, /*!< Entire page */
    xlSpecifiedTables = 3, /*!< Specified tables */
};

/**
Specifies the state of the window.

[Official VBA documentation for XlWindowState](https://docs.microsoft.com/office/vba/api/excel.xlwindowstate)
*/
enum XlWindowState
{
    xlMaximized = -4137, /*!< Maximized */
    xlMinimized = -4140, /*!< Minimized */
    xlNormal    = -4143, /*!< Normal */
};

/**
Specifies how the chart is displayed.

[Official VBA documentation for XlWindowType](https://docs.microsoft.com/office/vba/api/excel.xlwindowtype)
*/
enum XlWindowType
{
    xlChartAsWindow =     5, /*!< The chart will open in a new window. */
    xlChartInPlace  =     4, /*!< The chart will be displayed on the current worksheet. */
    xlClipboard     =     3, /*!< The chart is copied to the clipboard. */
    xlInfo          = -4129, /*!< This constant has been deprecated. */
    xlWorkbook      =     1, /*!< This constant applies to Macintosh only. */
};

/**
Specifies the view showing in the window.

[Official VBA documentation for XlWindowView](https://docs.microsoft.com/office/vba/api/excel.xlwindowview)
*/
enum XlWindowView
{
    xlNormalView       = 1, /*!< Normal. */
    xlPageBreakPreview = 2, /*!< Page break preview. */
    xlPageLayoutView   = 3, /*!< Page layout view. */
};

/**
Specifies, in a Microsoft Excel version 4 macro worksheet, what type of macro a name refers to or whether the name refers to a macro.

[Official VBA documentation for XlXLMMacroType](https://docs.microsoft.com/office/vba/api/excel.xlxlmmacrotype)
*/
enum XlXLMMacroType
{
    xlCommand  = 2, /*!< Custom command. */
    xlFunction = 1, /*!< Custom function. */
    xlNotXLM   = 3, /*!< Not a macro. */
};

/**
Specifies the results of the save or export operation.

[Official VBA documentation for XlXmlExportResult](https://docs.microsoft.com/office/vba/api/excel.xlxmlexportresult)
*/
enum XlXmlExportResult
{
    xlXmlExportSuccess          = 0, /*!< The XML data file was successfully exported. */
    xlXmlExportValidationFailed = 1, /*!< The contents of the XML data file don't match the specified schema map. */
};

/**
Specifies the results of the refresh or import operation.

[Official VBA documentation for XlXmlImportResult](https://docs.microsoft.com/office/vba/api/excel.xlxmlimportresult)
*/
enum XlXmlImportResult
{
    xlXmlImportElementsTruncated = 1, /*!< The contents of the specified XML data file have been truncated because the XML data file is too large for the worksheet. */
    xlXmlImportSuccess           = 0, /*!< The XML data file was successfully imported. */
    xlXmlImportValidationFailed  = 2, /*!< The contents of the XML data file don't match the specified schema map. */
};

/**
Specifies how Excel opens the XML data file.

[Official VBA documentation for XlXmlLoadOption](https://docs.microsoft.com/office/vba/api/excel.xlxmlloadoption)
*/
enum XlXmlLoadOption
{
    xlXmlLoadImportToList = 2, /*!< Places the contents of the XML data file in an XML table. */
    xlXmlLoadMapXml       = 3, /*!< Displays the schema of the XML data file in the **XML Structure** task pane. */
    xlXmlLoadOpenXml      = 1, /*!< Opens the XML data file. The contents of the file will be flattened. */
    xlXmlLoadPromptUser   = 0, /*!< Prompts the user to choose how to open the file. */
};

/**
Specifies whether or not the first row contains headers. Cannot be used when sorting PivotTable reports.

[Official VBA documentation for XlYesNoGuess](https://docs.microsoft.com/office/vba/api/excel.xlyesnoguess)
*/
enum XlYesNoGuess
{
    xlGuess = 0, /*!< Excel determines whether there is a header, and where it is, if there is one. */
    xlNo    = 2, /*!< Default. The entire range should be sorted. */
    xlYes   = 1, /*!< The entire range should not be sorted. */
};



/*************************************
    Microsoft Office enumerations
*************************************/

/**
Specifies constants that define the styles of the groups on the **File** tab.

[Official VBA documentation for BackstageGroupStyle](https://docs.microsoft.com/office/vba/api/office.backstagegroupstyle)
*/
enum BackstageGroupStyle
{
    BackstageGroupStyleError   = 2, /*!< Error style */
    BackstageGroupStyleNormal  = 0, /*!< Normal style */
    BackstageGroupStyleWarning = 1, /*!< Warning style */
};

/**
Provides information about the digital certificate.

[Official VBA documentation for CertificateDetail](https://docs.microsoft.com/office/vba/api/office.certificatedetail)
*/
enum CertificateDetail
{
    certdetAvailable      = 0, /*!< Specifies that the digital certificate is available for signing. */
    certdetExpirationDate = 3, /*!< The expiration date of the certificate. */
    certdetIssuer         = 2, /*!< The issuing authority of the certification. */
    certdetSubject        = 1, /*!< The holder of a Private Key corresponding to a Public Key. */
    certdetThumbprint     = 4, /*!< A hash of the certificate's complete contents. */
};

/**
Provides the results of verifying a digital certificate.

[Official VBA documentation for CertificateVerificationResults](https://docs.microsoft.com/office/vba/api/office.certificateverificationresults)
*/
enum CertificateVerificationResults
{
    certverresError      = 0, /*!< The verification resulted in an error. */
    certverresExpired    = 5, /*!< The certification has expired. */
    certverresInvalid    = 4, /*!< The certification is invalid. */
    certverresRevoked    = 6, /*!< The certification has been revoked. */
    certverresUntrusted  = 7, /*!< The certification is from an untrusted source. */
    certverresUnverified = 2, /*!< The certification is currently unverified. */
    certverresValid      = 3, /*!< The certification is valid. */
    certverresVerifying  = 1, /*!< The certificate is currently being verified. */
};

/**
Provides the status of verifying whether the content of a document has changed.

[Official VBA documentation for ContentVerificationResults](https://docs.microsoft.com/office/vba/api/office.contentverificationresults)
*/
enum ContentVerificationResults
{
    contverresError      = 0, /*!< The verification resulted in an error. */
    contverresModified   = 4, /*!< The content of the document has been modified since it was digitally signed. */
    contverresUnverified = 2, /*!< The document has not been verified. */
    contverresValid      = 3, /*!< The content of the document has been verified and is valid. */
    contverresVerifying  = 1, /*!< The content of the document is currently being verified. */
};

/**
Specifies the mode for encryption ciphers.

[Official VBA documentation for EncryptionCipherMode](https://docs.microsoft.com/office/vba/api/office.encryptionciphermode)
*/
enum EncryptionCipherMode
{
    cipherModeECB = 0, /*!< ECB cipher mode */
    cipherModeCBC = 1, /*!< CBC cipher mode */
};

/**
Specifies details about encryption providers.

[Official VBA documentation for EncryptionProviderDetail](https://docs.microsoft.com/office/vba/api/office.encryptionproviderdetail)
*/
enum EncryptionProviderDetail
{
    encprovdetURL             = 0, /*!< A URL encryption provider. */
    encprovdetAlgorithm       = 1, /*!< An algorithm encryption provider. */
    encprovdetBlockCipher     = 2, /*!< A block cipher encryption provider. */
    encprovdetCipherBlockSize = 3, /*!< A cipher block size encryption provider. */
    encprovdetCipherMode      = 4, /*!< A cipher mode encryption provider. */
};

/**
Specifies how the body of the email is displayed.

[Official VBA documentation for MailFormat](https://docs.microsoft.com/office/vba/api/office.mailformat)
*/
enum MailFormat
{
    mfHTML      = 2, /*!< The email is displayed as HyperText Markup Language (HTML). */
    mfPlainText = 1, /*!< The email is displayed as plain text. */
    mfRTF       = 3, /*!< The email is displayed as Rich Text Format (RTF). */
};

/**
Specifies behavior when the user cancels an alert. Only **msoAlertCancelDefault** is currently supported.

[Official VBA documentation for MsoAlertCancelType](https://docs.microsoft.com/office/vba/api/office.msoalertcanceltype)
*/
enum MsoAlertCancelType
{
    msoAlertCancelDefault = -1, /*!< Default behavior for canceling an alert. */
    msoAlertCancelFifth   =  4, /*!< Not supported. */
    msoAlertCancelFirst   =  0, /*!< Not supported. */
    msoAlertCancelFourth  =  3, /*!< Not supported. */
    msoAlertCancelSecond  =  1, /*!< Not supported. */
    msoAlertCancelThird   =  2, /*!< Not supported. */
};

/**
Specifies which icon, if any, to display with an alert. 

[Official VBA documentation for MsoAlertIconType](https://docs.microsoft.com/office/vba/api/office.msoalerticontype)
*/
enum MsoAlertIconType
{
    msoAlertIconCritical = 1, /*!< Displays the **Critical** icon. */
    msoAlertIconInfo     = 4, /*!< Displays the **Info** icon. */
    msoAlertIconNoIcon   = 0, /*!< Displays no icon with the alert message. */
    msoAlertIconQuery    = 2, /*!< Displays the **Query** icon. */
    msoAlertIconWarning  = 3, /*!< Displays the **Warning** icon. */
};

/**
Defines how to align specified objects relative to one another.

[Official VBA documentation for MsoAlignCmd](https://docs.microsoft.com/office/vba/api/office.msoaligncmd)
*/
enum MsoAlignCmd
{
    msoAlignBottoms = 5, /*!< Align bottoms of specified objects. */
    msoAlignCenters = 1, /*!< Align centers of specified objects. */
    msoAlignLefts   = 0, /*!< Align left sides of specified objects. */
    msoAlignMiddles = 4, /*!< Align middles of specified objects. */
    msoAlignRights  = 2, /*!< Align right sides of specified objects. */
    msoAlignTops    = 3, /*!< Align tops of specified objects. */
};

/**
Specifies a language setting in a Microsoft Office application. The **MsoAppLanguageID** enumeration is used with the **LanguageSettings** member of the **Application** object to determine the language used for the install language, the user interface language, or the Help language.

[Official VBA documentation for MsoAppLanguageID](https://docs.microsoft.com/office/vba/api/office.msoapplanguageid)
*/
enum MsoAppLanguageID
{
    msoLanguageIDExeMode    = 4, /*!< Execution mode language. */
    msoLanguageIDHelp       = 3, /*!< Help language. */
    msoLanguageIDInstall    = 1, /*!< Install language. */
    msoLanguageIDUI         = 2, /*!< User interface language. */
    msoLanguageIDUIPrevious = 5, /*!< User interface language used prior to the current user interface language. */
};

/**
Specifies the length of the arrowhead at the end of a line.

[Official VBA documentation for MsoArrowheadLength](https://docs.microsoft.com/office/vba/api/office.msoarrowheadlength)
*/
enum MsoArrowheadLength
{
    msoArrowheadLengthMedium =  2, /*!< Medium */
    msoArrowheadLengthMixed  = -2, /*!< Return value only; indicates a combination of the other states in the specified shape range. */
    msoArrowheadLong         =  3, /*!< Long */
    msoArrowheadShort        =  1, /*!< Short */
};

/**
Specifies the style of the arrowhead at the end of a line.

[Official VBA documentation for MsoArrowheadStyle](https://docs.microsoft.com/office/vba/api/office.msoarrowheadstyle)
*/
enum MsoArrowheadStyle
{
    msoArrowheadDiamond    =  5, /*!< Diamond-shaped */
    msoArrowheadNone       =  1, /*!< No arrowhead */
    msoArrowheadOpen       =  3, /*!< Open */
    msoArrowheadOval       =  6, /*!< Oval-shaped */
    msoArrowheadStealth    =  4, /*!< Stealth-shaped */
    msoArrowheadStyleMixed = -2, /*!< Return value only; indicates a combination of the other states. */
    msoArrowheadTriangle   =  2, /*!< Triangular */
};

/**
Specifies the width of the arrowhead at the end of a line.

[Official VBA documentation for MsoArrowheadWidth](https://docs.microsoft.com/office/vba/api/office.msoarrowheadwidth)
*/
enum MsoArrowheadWidth
{
    msoArrowheadNarrow      =  1, /*!< Narrow */
    msoArrowheadWide        =  3, /*!< Wide */
    msoArrowheadWidthMedium =  2, /*!< Medium */
    msoArrowheadWidthMixed  = -2, /*!< Return value only; indicates a combination of the other states. */
};

/**
Specifies the security mode an application uses when programmatically opening files.

[Official VBA documentation for MsoAutomationSecurity](https://docs.microsoft.com/office/vba/api/office.msoautomationsecurity)
*/
enum MsoAutomationSecurity
{
    msoAutomationSecurityByUI         = 2, /*!< Uses the security setting specified in the **Security** dialog box. */
    msoAutomationSecurityForceDisable = 3, /*!< Disables all macros in all files opened programmatically without showing any security alerts. */
    msoAutomationSecurityLow          = 1, /*!< Enables all macros. This is the default value when the application is started. */
};

/**
Specifies the shape type for an AutoShape object.

[Official VBA documentation for MsoAutoShapeType](https://docs.microsoft.com/office/vba/api/office.msoautoshapetype)
*/
enum MsoAutoShapeType
{
    msoShape10pointStar                      = 149, /*!< 10-point star */
    msoShape12pointStar                      = 150, /*!< 12-point star */
    msoShape16pointStar                      =  94, /*!< 16-point star */
    msoShape24pointStar                      =  95, /*!< 24-point star */
    msoShape32pointStar                      =  96, /*!< 32-point star */
    msoShape4pointStar                       =  91, /*!< 4-point star */
    msoShape5pointStar                       =  92, /*!< 5-point star */
    msoShape6pointStar                       = 147, /*!< 6-point star */
    msoShape7pointStar                       = 148, /*!< 7-point star */
    msoShape8pointStar                       =  93, /*!< 8-point star */
    msoShapeActionButtonBackorPrevious       = 129, /*!< **Back** or **Previous** button. Supports mouse-click and mouse-over actions. */
    msoShapeActionButtonBeginning            = 131, /*!< **Beginning** button. Supports mouse-click and mouse-over actions. */
    msoShapeActionButtonCustom               = 125, /*!< Button with no default picture or text. Supports mouse-click and mouse-over actions. */
    msoShapeActionButtonDocument             = 134, /*!< **Document** button. Supports mouse-click and mouse-over actions. */
    msoShapeActionButtonEnd                  = 132, /*!< **End** button. Supports mouse-click and mouse-over actions. */
    msoShapeActionButtonForwardorNext        = 130, /*!< **Forward** or **Next** button. Supports mouse-click and mouse-over actions. */
    msoShapeActionButtonHelp                 = 127, /*!< **Help** button. Supports mouse-click and mouse-over actions. */
    msoShapeActionButtonHome                 = 126, /*!< **Home** button. Supports mouse-click and mouse-over actions. */
    msoShapeActionButtonInformation          = 128, /*!< **Information** button. Supports mouse-click and mouse-over actions. */
    msoShapeActionButtonMovie                = 136, /*!< **Movie** button. Supports mouse-click and mouse-over actions. */
    msoShapeActionButtonReturn               = 133, /*!< **Return** button. Supports mouse-click and mouse-over actions. */
    msoShapeActionButtonSound                = 135, /*!< **Sound** button. Supports mouse-click and mouse-over actions. */
    msoShapeArc                              =  25, /*!< Arc */
    msoShapeBalloon                          = 137, /*!< Balloon */
    msoShapeBentArrow                        =  41, /*!< Block arrow that follows a curved 90-degree angle. */
    msoShapeBentUpArrow                      =  44, /*!< Block arrow that follows a sharp 90-degree angle. Points up by default. */
    msoShapeBevel                            =  15, /*!< Bevel */
    msoShapeBlockArc                         =  20, /*!< Block arc */
    msoShapeCan                              =  13, /*!< Can */
    msoShapeChartPlus                        = 182, /*!< Square divided vertically and horizontally into four quarters */
    msoShapeChartStar                        = 181, /*!< Square divided into six parts along vertical and diagonal lines */
    msoShapeChartX                           = 180, /*!< Square divided into four parts along diagonal lines */
    msoShapeChevron                          =  52, /*!< Chevron */
    msoShapeChord                            = 161, /*!< Circle with a line connecting two points on the perimeter through the interior of the circle; a circle with a chord */
    msoShapeCircularArrow                    =  60, /*!< Block arrow that follows a curved 180-degree angle */
    msoShapeCloud                            = 179, /*!< Cloud shape */
    msoShapeCloudCallout                     = 108, /*!< Cloud callout */
    msoShapeCorner                           = 162, /*!< Rectangle with rectangular-shaped hole. */
    msoShapeCornerTabs                       = 169, /*!< Four right triangles aligning along a rectangular path; four 'snipped' corners. */
    msoShapeCross                            =  11, /*!< Cross */
    msoShapeCube                             =  14, /*!< Cube */
    msoShapeCurvedDownArrow                  =  48, /*!< Block arrow that curves down */
    msoShapeCurvedDownRibbon                 = 100, /*!< Ribbon banner that curves down */
    msoShapeCurvedLeftArrow                  =  46, /*!< Block arrow that curves left */
    msoShapeCurvedRightArrow                 =  45, /*!< Block arrow that curves right */
    msoShapeCurvedUpArrow                    =  47, /*!< Block arrow that curves up */
    msoShapeCurvedUpRibbon                   =  99, /*!< Ribbon banner that curves up */
    msoShapeDecagon                          = 144, /*!< Decagon */
    msoShapeDiagonalStripe                   = 141, /*!< Rectangle with two triangles-shapes removed; a diagonal stripe */
    msoShapeDiamond                          =   4, /*!< Diamond */
    msoShapeDodecagon                        = 146, /*!< Dodecagon */
    msoShapeDonut                            =  18, /*!< Donut */
    msoShapeDoubleBrace                      =  27, /*!< Double brace */
    msoShapeDoubleBracket                    =  26, /*!< Double bracket */
    msoShapeDoubleWave                       = 104, /*!< Double wave */
    msoShapeDownArrow                        =  36, /*!< Block arrow that points down */
    msoShapeDownArrowCallout                 =  56, /*!< Callout with arrow that points down */
    msoShapeDownRibbon                       =  98, /*!< Ribbon banner with center area below ribbon ends */
    msoShapeExplosion1                       =  89, /*!< Explosion */
    msoShapeExplosion2                       =  90, /*!< Explosion */
    msoShapeFlowchartAlternateProcess        =  62, /*!< Alternate process flowchart symbol */
    msoShapeFlowchartCard                    =  75, /*!< Card flowchart symbol */
    msoShapeFlowchartCollate                 =  79, /*!< Collate flowchart symbol */
    msoShapeFlowchartConnector               =  73, /*!< Connector flowchart symbol */
    msoShapeFlowchartData                    =  64, /*!< Data flowchart symbol */
    msoShapeFlowchartDecision                =  63, /*!< Decision flowchart symbol */
    msoShapeFlowchartDelay                   =  84, /*!< Delay flowchart symbol */
    msoShapeFlowchartDirectAccessStorage     =  87, /*!< Direct access storage flowchart symbol */
    msoShapeFlowchartDisplay                 =  88, /*!< Display flowchart symbol */
    msoShapeFlowchartDocument                =  67, /*!< Document flowchart symbol */
    msoShapeFlowchartExtract                 =  81, /*!< Extract flowchart symbol */
    msoShapeFlowchartInternalStorage         =  66, /*!< Internal storage flowchart symbol */
    msoShapeFlowchartMagneticDisk            =  86, /*!< Magnetic disk flowchart symbol */
    msoShapeFlowchartManualInput             =  71, /*!< Manual input flowchart symbol */
    msoShapeFlowchartManualOperation         =  72, /*!< Manual operation flowchart symbol */
    msoShapeFlowchartMerge                   =  82, /*!< Merge flowchart symbol */
    msoShapeFlowchartMultidocument           =  68, /*!< Multi-document flowchart symbol */
    msoShapeFlowchartOfflineStorage          = 139, /*!< Offline storage flowchart symbol */
    msoShapeFlowchartOffpageConnector        =  74, /*!< Off-page connector flowchart symbol */
    msoShapeFlowchartOr                      =  78, /*!< "Or" flowchart symbol */
    msoShapeFlowchartPredefinedProcess       =  65, /*!< Predefined process flowchart symbol */
    msoShapeFlowchartPreparation             =  70, /*!< Preparation flowchart symbol */
    msoShapeFlowchartProcess                 =  61, /*!< Process flowchart symbol */
    msoShapeFlowchartPunchedTape             =  76, /*!< Punched tape flowchart symbol */
    msoShapeFlowchartSequentialAccessStorage =  85, /*!< Sequential access storage flowchart symbol */
    msoShapeFlowchartSort                    =  80, /*!< Sort flowchart symbol */
    msoShapeFlowchartStoredData              =  83, /*!< Stored data flowchart symbol */
    msoShapeFlowchartSummingJunction         =  77, /*!< Summing junction flowchart symbol */
    msoShapeFlowchartTerminator              =  69, /*!< Terminator flowchart symbol */
    msoShapeFoldedCorner                     =  16, /*!< Folded corner */
    msoShapeFrame                            = 158, /*!< Rectangular picture frame */
    msoShapeFunnel                           = 174, /*!< Funnel */
    msoShapeGear6                            = 172, /*!< Gear with six teeth */
    msoShapeGear9                            = 173, /*!< Gear with nine teeth */
    msoShapeHalfFrame                        = 159, /*!< Half of a rectangular picture frame */
    msoShapeHeart                            =  21, /*!< Heart */
    msoShapeHeptagon                         = 145, /*!< Heptagon */
    msoShapeHexagon                          =  10, /*!< Hexagon */
    msoShapeHorizontalScroll                 = 102, /*!< Horizontal scroll */
    msoShapeIsoscelesTriangle                =   7, /*!< Isosceles triangle */
    msoShapeLeftArrow                        =  34, /*!< Block arrow that points left */
    msoShapeLeftArrowCallout                 =  54, /*!< Callout with arrow that points left */
    msoShapeLeftBrace                        =  31, /*!< Left brace */
    msoShapeLeftBracket                      =  29, /*!< Left bracket */
    msoShapeLeftCircularArrow                = 176, /*!< Circular arrow pointing counter-clockwise */
    msoShapeLeftRightArrow                   =  37, /*!< Block arrow with arrowheads that point both left and right */
    msoShapeLeftRightArrowCallout            =  57, /*!< Callout with arrowheads that point both left and right */
    msoShapeLeftRightCircularArrow           = 177, /*!< Circular arrow pointing clockwise and counter-clockwise; a curved arrow with points at both ends */
    msoShapeLeftRightRibbon                  = 140, /*!< Ribbon with an arrow at both ends */
    msoShapeLeftRightUpArrow                 =  40, /*!< Block arrow with arrowheads that point left, right, and up */
    msoShapeLeftUpArrow                      =  43, /*!< Block arrow with arrowheads that point left and up */
    msoShapeLightningBolt                    =  22, /*!< Lightning bolt */
    msoShapeLineCallout1                     = 109, /*!< Callout with border and horizontal callout line */
    msoShapeLineCallout1AccentBar            = 113, /*!< Callout with horizontal accent bar */
    msoShapeLineCallout1BorderandAccentBar   = 121, /*!< Callout with border and horizontal accent bar */
    msoShapeLineCallout1NoBorder             = 117, /*!< Callout with horizontal line */
    msoShapeLineCallout2                     = 110, /*!< Callout with diagonal straight line */
    msoShapeLineCallout2AccentBar            = 114, /*!< Callout with diagonal callout line and accent bar */
    msoShapeLineCallout2BorderandAccentBar   = 122, /*!< Callout with border, diagonal straight line, and accent bar */
    msoShapeLineCallout2NoBorder             = 118, /*!< Callout with no border and diagonal callout line */
    msoShapeLineCallout3                     = 111, /*!< Callout with angled line */
    msoShapeLineCallout3AccentBar            = 115, /*!< Callout with angled callout line and accent bar */
    msoShapeLineCallout3BorderandAccentBar   = 123, /*!< Callout with border, angled callout line, and accent bar */
    msoShapeLineCallout3NoBorder             = 119, /*!< Callout with no border and angled callout line */
    msoShapeLineCallout4                     = 112, /*!< Callout with callout line segments forming a U-shape */
    msoShapeLineCallout4AccentBar            = 116, /*!< Callout with accent bar and callout line segments forming a U-shape */
    msoShapeLineCallout4BorderandAccentBar   = 124, /*!< Callout with border, accent bar, and callout line segments forming a U-shape */
    msoShapeLineCallout4NoBorder             = 120, /*!< Callout with no border and callout line segments forming a U-shape */
    msoShapeLineInverse                      = 183, /*!< Line inverse */
    msoShapeMathDivide                       = 166, /*!< Division symbol `÷` */
    msoShapeMathEqual                        = 167, /*!< Equivalence symbol `=` */
    msoShapeMathMinus                        = 164, /*!< Subtraction symbol `-` */
    msoShapeMathMultiply                     = 165, /*!< Multiplication symbol `x` */
    msoShapeMathNotEqual                     = 168, /*!< Non-equivalence symbol `≠` */
    msoShapeMathPlus                         = 163, /*!< Addition symbol `+` */
    msoShapeMixed                            =  -2, /*!< Return value only; indicates a combination of the other states. */
    msoShapeMoon                             =  24, /*!< Moon */
    msoShapeNonIsoscelesTrapezoid            = 143, /*!< Trapezoid with asymmetrical non-parallel sides */
    msoShapeNoSymbol                         =  19, /*!< "No" symbol */
    msoShapeNotchedRightArrow                =  50, /*!< Notched block arrow that points right */
    msoShapeNotPrimitive                     = 138, /*!< Not supported */
    msoShapeOctagon                          =   6, /*!< Octagon */
    msoShapeOval                             =   9, /*!< Oval */
    msoShapeOvalCallout                      = 107, /*!< Oval-shaped callout */
    msoShapeParallelogram                    =   2, /*!< Parallelogram */
    msoShapePentagon                         =  51, /*!< Pentagon */
    msoShapePie                              = 142, /*!< Circle ('pie') with a portion missing */
    msoShapePieWedge                         = 175, /*!< Quarter of a circular shape */
    msoShapePlaque                           =  28, /*!< Plaque */
    msoShapePlaqueTabs                       = 171, /*!< Four quarter-circles defining a rectangular shape */
    msoShapeQuadArrow                        =  39, /*!< Block arrows that point up, down, left, and right */
    msoShapeQuadArrowCallout                 =  59, /*!< Callout with arrows that point up, down, left, and right */
    msoShapeRectangle                        =   1, /*!< Rectangle */
    msoShapeRectangularCallout               = 105, /*!< Rectangular callout */
    msoShapeRegularPentagon                  =  12, /*!< Pentagon */
    msoShapeRightArrow                       =  33, /*!< Block arrow that points right */
    msoShapeRightArrowCallout                =  53, /*!< Callout with arrow that points right */
    msoShapeRightBrace                       =  32, /*!< Right brace */
    msoShapeRightBracket                     =  30, /*!< Right bracket */
    msoShapeRightTriangle                    =   8, /*!< Right triangle */
    msoShapeRound1Rectangle                  = 151, /*!< Rectangle with one rounded corner */
    msoShapeRound2DiagRectangle              = 157, /*!< Rectangle with two rounded corners, diagonally-opposed */
    msoShapeRound2SameRectangle              = 152, /*!< Rectangle with two-rounded corners that share a side */
    msoShapeRoundedRectangle                 =   5, /*!< Rounded rectangle */
    msoShapeRoundedRectangularCallout        = 106, /*!< Rounded rectangle-shaped callout */
    msoShapeSmileyFace                       =  17, /*!< Smiley face */
    msoShapeSnip1Rectangle                   = 155, /*!< Rectangle with one snipped corner */
    msoShapeSnip2DiagRectangle               = 157, /*!< Rectangle with two snipped corners, diagonally-opposed */
    msoShapeSnip2SameRectangle               = 156, /*!< Rectangle with two snipped corners that share a side */
    msoShapeSnipRoundRectangle               = 154, /*!< Rectangle with one snipped corner and one rounded corner */
    msoShapeSquareTabs                       = 170, /*!< Four small squares that define a rectangular shape */
    msoShapeStripedRightArrow                =  49, /*!< Block arrow that points right with stripes at the tail */
    msoShapeSun                              =  23, /*!< Sun */
    msoShapeSwooshArrow                      = 178, /*!< Curved arrow */
    msoShapeTear                             = 160, /*!< Water droplet */
    msoShapeTrapezoid                        =   3, /*!< Trapezoid */
    msoShapeUpArrow                          =  35, /*!< Block arrow that points up */
    msoShapeUpArrowCallout                   =  55, /*!< Callout with arrow that points up */
    msoShapeUpDownArrow                      =  38, /*!< Block arrow that points up and down */
    msoShapeUpDownArrowCallout               =  58, /*!< Callout with arrows that point up and down */
    msoShapeUpRibbon                         =  97, /*!< Ribbon banner with center area above ribbon ends */
    msoShapeUTurnArrow                       =  42, /*!< Block arrow forming a U shape */
    msoShapeVerticalScroll                   = 101, /*!< Vertical scroll */
    msoShapeWave                             = 103, /*!< Wave */
};

/**
Determines the type of automatic sizing allowed.

[Official VBA documentation for MsoAutoSize](https://docs.microsoft.com/office/vba/api/office.msoautosize)
*/
enum MsoAutoSize
{
    msoAutoSizeMixed          = -2, /*!< A combination of automatic sizing schemes are used. */
    msoAutoSizeNone           =  0, /*!< No autosizing. */
    msoAutoSizeShapeToFitText =  1, /*!< The shape is adjusted to fit the text. */
    msoAutoSizeTextToFitShape =  2, /*!< The text is adjusted to fit the shape. */
};

/**
Indicates the background style for an object.

[Official VBA documentation for MsoBackgroundStyleIndex](https://docs.microsoft.com/office/vba/api/office.msobackgroundstyleindex)
*/
enum MsoBackgroundStyleIndex
{
    msoBackgroundStyle1     =  1, /*!< Specifies Style1. */
    msoBackgroundStyle10    = 10, /*!< Specifies Style10. */
    msoBackgroundStyle11    = 11, /*!< Specifies Style11. */
    msoBackgroundStyle12    = 12, /*!< Specifies Style12. */
    msoBackgroundStyle2     =  2, /*!< Specifies Style2. */
    msoBackgroundStyle3     =  3, /*!< Specifies Style3. */
    msoBackgroundStyle4     =  4, /*!< Specifies Style4. */
    msoBackgroundStyle5     =  5, /*!< Specifies Style5. */
    msoBackgroundStyle6     =  6, /*!< Specifies Style6. */
    msoBackgroundStyle7     =  7, /*!< Specifies Style7. */
    msoBackgroundStyle8     =  8, /*!< Specifies Style8. */
    msoBackgroundStyle9     =  9, /*!< Specifies Style9. */
    msoBackgroundStyleMixed = -2, /*!< Specifies a combination of styles. */
    msoBackgroundStyleNone  =  0, /*!< Specifies no styles. */
};

/**
Specifies the position or behavior of a command bar.

[Official VBA documentation for MsoBarPosition](https://docs.microsoft.com/office/vba/api/office.msobarposition)
*/
enum MsoBarPosition
{
    msoBarBottom   = 3, /*!< Command bar is docked at the bottom of the application window. */
    msoBarFloating = 4, /*!< Command bar floats on top of the application window. */
    msoBarLeft     = 0, /*!< Command bar is docked on the left side of the application window. */
    msoBarMenuBar  = 6, /*!< Command bar will be a menu bar (Macintosh only). */
    msoBarPopup    = 5, /*!< Command bar will be a shortcut menu. */
    msoBarRight    = 2, /*!< Command bar is docked on the right side of the application window. */
    msoBarTop      = 1, /*!< Command bar is docked at the top of the application window. */
};

/**
Specifies how a command bar is protected from user customization.

[Official VBA documentation for MsoBarProtection](https://docs.microsoft.com/office/vba/api/office.msobarprotection)
*/
enum MsoBarProtection
{
    msoBarNoChangeDock     = 16, /*!< Docking setting cannot be changed. */
    msoBarNoChangeVisible  =  8, /*!< Command bar cannot be hidden. */
    msoBarNoCustomize      =  1, /*!< Command bar cannot be customized. */
    msoBarNoHorizontalDock = 64, /*!< Command bar cannot be docked to the top or bottom. */
    msoBarNoMove           =  4, /*!< Command bar cannot be moved. */
    msoBarNoProtection     =  0, /*!< All aspects of the command bar can be customized by user. */
    msoBarNoResize         =  2, /*!< Command bar cannot be resized. */
    msoBarNoVerticalDock   = 32, /*!< Command bar cannot be docked to the left or right. */
};

/**
Specifies whether a command bar is in the first row or last row relative to other command bars in the same docking area.

[Official VBA documentation for MsoBarRow](https://docs.microsoft.com/office/vba/api/office.msobarrow)
*/
enum MsoBarRow
{
    msoBarRowFirst =  0, /*!< First row of docking area. */
    msoBarRowLast  = -1, /*!< Last row of docking area. */
};

/**
Specifies the type of the command bar.

[Official VBA documentation for MsoBarType](https://docs.microsoft.com/office/vba/api/office.msobartype)
*/
enum MsoBarType
{
    msoBarTypeMenuBar = 1, /*!< Menu bar */
    msoBarTypeNormal  = 0, /*!< Default command bar */
    msoBarTypePopup   = 2, /*!< Shortcut menu */
};

/**
Specifies baseline text alignment.

[Official VBA documentation for MsoBaselineAlignment](https://docs.microsoft.com/office/vba/api/office.msobaselinealignment)
*/
enum MsoBaselineAlignment
{
    msoBaselineAlignAuto      =  5, /*!< Automatic alignment. */
    msoBaselineAlignBaseline  =  1, /*!< Baseline alignment. */
    msoBaselineAlignCenter    =  3, /*!< Center alignment. */
    msoBaselineAlignFarEast50 =  4, /*!< East Asia 50 alignment. */
    msoBaselineAlignMixed     = -2, /*!< Return value only; indicates a combination of the other states. */
    msoBaselineAlignTop       =  2, /*!< Top alignment. */
};

/**
Indicates the bevel type of a **ThreeDFormat** object.

[Official VBA documentation for MsoBevelType](https://docs.microsoft.com/office/vba/api/office.msobeveltype)
*/
enum MsoBevelType
{
    msoBevelAngle        =  6, /*!< Specifies an Angle bevel. */
    msoBevelArtDeco      = 13, /*!< Specifies an ArtDeco bevel. */
    msoBevelCircle       =  3, /*!< Specifies a Circle bevel. */
    msoBevelConvex       =  8, /*!< Specifies a Convex bevel. */
    msoBevelCoolSlant    =  9, /*!< Specifies a CoolSlant bevel. */
    msoBevelCross        =  5, /*!< Specifies a Cross bevel. */
    msoBevelDivot        = 10, /*!< Specifies a Divot bevel. */
    msoBevelHardEdge     = 12, /*!< Specifies a HardEdge bevel. */
    msoBevelNone         =  1, /*!< Specifies no bevel. */
    msoBevelRelaxedInset =  2, /*!< Specifies a RelaxedInset bevel. */
    msoBevelRiblet       = 11, /*!< Specifies a Riblet bevel. */
    msoBevelSlope        =  4, /*!< Specifies a Slope bevel. */
    msoBevelSoftRound    =  7, /*!< Specifies a SoftRound bevel. */
    msoBevelTypeMixed    = -2, /*!< Specifies a mixed type bevel. */
};

/**
Specifies how a shape appears when viewed in black-and-white mode.

[Official VBA documentation for MsoBlackWhiteMode](https://docs.microsoft.com/office/vba/api/office.msoblackwhitemode)
*/
enum MsoBlackWhiteMode
{
    msoBlackWhiteAutomatic        =  1, /*!< Default behavior */
    msoBlackWhiteBlack            =  8, /*!< Black */
    msoBlackWhiteBlackTextAndLine =  6, /*!< White with grayscale fill */
    msoBlackWhiteDontShow         = 10, /*!< Not shown */
    msoBlackWhiteGrayOutline      =  5, /*!< Gray with white fill */
    msoBlackWhiteGrayScale        =  2, /*!< Grayscale */
    msoBlackWhiteHighContrast     =  7, /*!< Black with white fill */
    msoBlackWhiteInverseGrayScale =  4, /*!< Inverse grayscale */
    msoBlackWhiteLightGrayScale   =  3, /*!< Light grayscale */
    msoBlackWhiteMixed            = -2, /*!< Not supported */
    msoBlackWhiteWhite            =  9, /*!< White */
};

/**
Specifies how many categories are supported by the provider.

[Official VBA documentation for MsoBlogCategorySupport](https://docs.microsoft.com/office/vba/api/office.msoblogcategorysupport)
*/
enum MsoBlogCategorySupport
{
    msoBlogMultipleCategories = 2, /*!< Multiple categories are supported. */
    msoBlogNoCategories       = 0, /*!< No categories are supported. */
    msoBlogOneCategory        = 1, /*!< One category is supported. */
};

/**
Specifies blog image types.

[Official VBA documentation for MsoBlogImageType](https://docs.microsoft.com/office/vba/api/office.msoblogimagetype)
*/
enum MsoBlogImageType
{
    msoBlogImageTypeGIF  = 2, /*!< GIF image */
    msoBlogImageTypeJPEG = 1, /*!< JPEG image */
    msoBlogImageTypePNG  = 3, /*!< PNG image */
};

/**
Specifies the application capabilities available for a document or presentation broadcasting session.

[Official VBA documentation for MsoBroadcastCapabilities](https://docs.microsoft.com/office/vba/api/office.msobroadcastcapabilities)
*/
enum MsoBroadcastCapabilities
{
    BroadcastCapFileSizeLimited      = 1, /*!< The size of the file being broadcasted is limited. */
    BroadcastCapSupportsMeetingNotes = 2, /*!< The presenters and attendees can take shared notes. */
    BroadcastCapSupportsUpdateDoc    = 4, /*!< The presenters and attendees can make updates to the file during the broadcast. */
};

/**
Specifies the current state of a document or presentation broadcast.

[Official VBA documentation for MsoBroadcastState](https://docs.microsoft.com/office/vba/api/office.msobroadcaststate)
*/
enum MsoBroadcastState
{
    BroadcastPaused  = 2, /*!< The broadcast is paused. */
    BroadcastStarted = 1, /*!< The broadcast has been started. */
    NoBroadcast      = 0, /*!< The file is not being broadcasted. */
};

/**
Specifies the bullet type.

[Official VBA documentation for MsoBulletType](https://docs.microsoft.com/office/vba/api/office.msobullettype)
*/
enum MsoBulletType
{
    msoBulletMixed      = -2, /*!< Return value only; indicates a combination of the other states. */
    msoBulletNone       =  0, /*!< No bullets. */
    msoBulletNumbered   =  2, /*!< Numbered bullets. */
    msoBulletPicture    =  3, /*!< Picture bullets. */
    msoBulletUnnumbered =  1, /*!< Unnumbered bullets. */
};

/**
Specifies the appearance of a command bar button control.

[Official VBA documentation for MsoButtonState](https://docs.microsoft.com/office/vba/api/office.msobuttonstate)
*/
enum MsoButtonState
{
    msoButtonDown  = -1, /*!< Button is pressed down. */
    msoButtonMixed =  2, /*!< Button is pressed down. */
    msoButtonUp    =  0, /*!< Button is not pressed down. */
};

/**
Specifies the style of a command bar button.

[Official VBA documentation for MsoButtonStyle](https://docs.microsoft.com/office/vba/api/office.msobuttonstyle)
*/
enum MsoButtonStyle
{
    msoButtonAutomatic               =  0, /*!< Default behavior. */
    msoButtonCaption                 =  2, /*!< Text only. */
    msoButtonIcon                    =  1, /*!< Image only. */
    msoButtonIconAndCaption          =  3, /*!< Image and text, with text to the right of image. */
    msoButtonIconAndCaptionBelow     = 11, /*!< Image with text below. */
    msoButtonIconAndWrapCaption      =  7, /*!< Image with text wrapped and to the right of the image. */
    msoButtonIconAndWrapCaptionBelow = 15, /*!< Image with text wrapped below image. */
    msoButtonWrapCaption             = 14, /*!< Text only, centered and wrapped. */
};

/**
Specifies the size of the angle between the callout line and the side of the callout text box.

[Official VBA documentation for MsoCalloutAngleType](https://docs.microsoft.com/office/vba/api/office.msocalloutangletype)
*/
enum MsoCalloutAngleType
{
    msoCalloutAngle30        =  2, /*!< 30° angle */
    msoCalloutAngle45        =  3, /*!< 45° angle */
    msoCalloutAngle60        =  4, /*!< 60° angle */
    msoCalloutAngle90        =  5, /*!< 90° angle */
    msoCalloutAngleAutomatic =  1, /*!< Default angle. Angle can be changed as you drag the object. */
    msoCalloutAngleMixed     = -2, /*!< Return value only; indicates a combination of the other states. */
};

/**
Specifies the starting position of the callout line relative to the text bounding box. Used with the **PresetDrop** method of the **CalloutFormat** object.

[Official VBA documentation for MsoCalloutDropType](https://docs.microsoft.com/office/vba/api/office.msocalloutdroptype)
*/
enum MsoCalloutDropType
{
    msoCalloutDropBottom =  4, /*!< Bottom */
    msoCalloutDropCenter =  3, /*!< Center */
    msoCalloutDropCustom =  1, /*!< Custom. If this value is used as the value for the **PresetDrop** property, the **Drop** and **AutoAttach** properties of the **CalloutFormat** object are used to determine where the callout line attaches to the text box. */
    msoCalloutDropMixed  = -2, /*!< Return value only; indicates a combination of the other states. */
    msoCalloutDropTop    =  2, /*!< Top */
};

/**
Specifies the type of callout line.

[Official VBA documentation for MsoCalloutType](https://docs.microsoft.com/office/vba/api/office.msocallouttype)
*/
enum MsoCalloutType
{
    msoCalloutFour  =  4, /*!< Callout line made up of two line segments. Callout line is attached on the right side of the text bounding box. */
    msoCalloutMixed = -2, /*!< Return value only; indicates a combination of the other states. */
    msoCalloutOne   =  1, /*!< Single, horizontal callout line. */
    msoCalloutThree =  3, /*!< Callout line made up of two line segments. Callout line is attached on the left side of the text bounding box. */
    msoCalloutTwo   =  2, /*!< Single, angled callout line. */
};

/**
Specifies the character set to be used when rendering text.

[Official VBA documentation for MsoCharacterSet](https://docs.microsoft.com/office/vba/api/office.msocharacterset)
*/
enum MsoCharacterSet
{
    msoCharacterSetArabic                                 =  1, /*!< Arabic character set */
    msoCharacterSetCyrillic                               =  2, /*!< Cyrillic character set */
    msoCharacterSetEnglishWesternEuropeanOtherLatinScript =  3, /*!< English, Western European, and other Latin script character set */
    msoCharacterSetGreek                                  =  4, /*!< Greek character set */
    msoCharacterSetHebrew                                 =  5, /*!< Hebrew character set */
    msoCharacterSetJapanese                               =  6, /*!< Japanese character set */
    msoCharacterSetKorean                                 =  7, /*!< Korean character set */
    msoCharacterSetMultilingualUnicode                    =  8, /*!< Multilingual Unicode character set */
    msoCharacterSetSimplifiedChinese                      =  9, /*!< Simplified Chinese character set */
    msoCharacterSetThai                                   = 10, /*!< Thai character set */
    msoCharacterSetTraditionalChinese                     = 11, /*!< Traditional Chinese character set */
    msoCharacterSetVietnamese                             = 12, /*!< Vietnamese character set */
};

/**
Specifies whether and how to display chart elements.

[Official VBA documentation for MsoChartElementType](https://docs.microsoft.com/office/vba/api/office.msochartelementtype)
*/
enum MsoChartElementType
{
    msoElementChartFloorNone                           = 1200, /*!< Do not display chart floor. */
    msoElementChartFloorShow                           = 1201, /*!< Display chart floor. */
    msoElementChartTitleAboveChart                     =    2, /*!< Display title above chart. */
    msoElementChartTitleCenteredOverlay                =    1, /*!< Display title as centered overlay. */
    msoElementChartTitleNone                           =    0, /*!< Do not display chart title. */
    msoElementChartWallNone                            = 1100, /*!< Do not display chart wall. */
    msoElementChartWallShow                            = 1101, /*!< Display chart wall. */
    msoElementDataLabelBestFit                         =  210, /*!< Use best fit for data label. */
    msoElementDataLabelBottom                          =  209, /*!< Display data label at bottom. */
    msoElementDataLabelCallout                         =  211, /*!< Display data label as a callout. */
    msoElementDataLabelCenter                          =  202, /*!< Display data label in center. */
    msoElementDataLabelInsideBase                      =  204, /*!< Display data label inside at the base. */
    msoElementDataLabelInsideEnd                       =  203, /*!< Display data label inside at the end. */
    msoElementDataLabelLeft                            =  206, /*!< Display data label to the left. */
    msoElementDataLabelNone                            =  200, /*!< Do not display data label. */
    msoElementDataLabelOutSideEnd                      =  205, /*!< Display data label outside at the end. */
    msoElementDataLabelRight                           =  207, /*!< Display data label to the right. */
    msoElementDataLabelShow                            =  201, /*!< Display data label. */
    msoElementDataLabelTop                             =  208, /*!< Display data label at the top. */
    msoElementDataTableNone                            =  500, /*!< Do not display data table. */
    msoElementDataTableShow                            =  501, /*!< Display data table. */
    msoElementDataTableWithLegendKeys                  =  502, /*!< Display data table with legend keys. */
    msoElementErrorBarNone                             =  700, /*!< Do not display error bar. */
    msoElementErrorBarPercentage                       =  702, /*!< Display percentage error bar. */
    msoElementErrorBarStandardDeviation                =  703, /*!< Display standard deviation error bar. */
    msoElementErrorBarStandardError                    =  701, /*!< Display standard error bar. */
    msoElementLegendBottom                             =  104, /*!< Display legend at the bottom. */
    msoElementLegendLeft                               =  103, /*!< Display legend at the left. */
    msoElementLegendLeftOverlay                        =  106, /*!< Overlay legend at the left. */
    msoElementLegendNone                               =  100, /*!< Do not display legend. */
    msoElementLegendRight                              =  101, /*!< Display legend at the right. */
    msoElementLegendRightOverlay                       =  105, /*!< Overlay legend at the right. */
    msoElementLegendTop                                =  102, /*!< Display legend at the top. */
    msoElementLineDropHiLoLine                         =  804, /*!< Display drop high/low line. */
    msoElementLineDropLine                             =  801, /*!< Display drop line. */
    msoElementLineHiLoLine                             =  802, /*!< Display high/low line. */
    msoElementLineNone                                 =  800, /*!< Do not display line. */
    msoElementLineSeriesLine                           =  803, /*!< Display series line. */
    msoElementPlotAreaNone                             = 1000, /*!< Do not display plot area. */
    msoElementPlotAreaShow                             = 1001, /*!< Display plot area. */
    msoElementPrimaryCategoryAxisBillions              =  374, /*!< Use billions for primary category axis units. */
    msoElementPrimaryCategoryAxisLogScale              =  375, /*!< Use log scale for primary category axis. */
    msoElementPrimaryCategoryAxisMillions              =  373, /*!< Use millions for primary category axis units. */
    msoElementPrimaryCategoryAxisNone                  =  348, /*!< Do not display primary category axis. */
    msoElementPrimaryCategoryAxisReverse               =  351, /*!< Reverse primary category axis. */
    msoElementPrimaryCategoryAxisShow                  =  349, /*!< Show primary category axis. */
    msoElementPrimaryCategoryAxisThousands             =  372, /*!< Use thousands for primary category axis units. */
    msoElementPrimaryCategoryAxisTitleAdjacentToAxis   =  301, /*!< Display primary category axis title adjacent to the axis. */
    msoElementPrimaryCategoryAxisTitleBelowAxis        =  302, /*!< Display primary category axis title below the axis. */
    msoElementPrimaryCategoryAxisTitleHorizontal       =  305, /*!< Display primary category axis title horizontally. */
    msoElementPrimaryCategoryAxisTitleNone             =  300, /*!< Do not display primary category axis title. */
    msoElementPrimaryCategoryAxisTitleRotated          =  303, /*!< Rotate primary category axis title. */
    msoElementPrimaryCategoryAxisTitleVertical         =  304, /*!< Display primary category axis title vertically. */
    msoElementPrimaryCategoryAxisWithoutLabels         =  350, /*!< Display primary category axis without labels. */
    msoElementPrimaryCategoryGridLinesMajor            =  334, /*!< Display major gridlines along primary category axis. */
    msoElementPrimaryCategoryGridLinesMinor            =  333, /*!< Display minor gridlines along primary category axis. */
    msoElementPrimaryCategoryGridLinesMinorMajor       =  335, /*!< Display both major and minor gridlines along primary category axis. */
    msoElementPrimaryCategoryGridLinesNone             =  332, /*!< Do not display grid lines along primary category axis. */
    msoElementPrimaryValueAxisBillions                 =  356, /*!< Use billions for primary value axis units. */
    msoElementPrimaryValueAxisLogScale                 =  357, /*!< Use log scale for primary value axis. */
    msoElementPrimaryValueAxisMillions                 =  355, /*!< Use millions for primary value axis units. */
    msoElementPrimaryValueAxisNone                     =  352, /*!< Do not display primary value axis. */
    msoElementPrimaryValueAxisShow                     =  353, /*!< Show primary value axis */
    msoElementPrimaryValueAxisThousands                =  354, /*!< Use thousands for primary value axis units. */
    msoElementPrimaryValueAxisTitleAdjacentToAxis      =  307, /*!< Place primary value axis title adjacent to the axis. */
    msoElementPrimaryValueAxisTitleBelowAxis           =  308, /*!< Place primary value axis title below the axis. */
    msoElementPrimaryValueAxisTitleHorizontal          =  311, /*!< Display primary value axis title horizontally. */
    msoElementPrimaryValueAxisTitleNone                =  306, /*!< Do not display primary value axis title. */
    msoElementPrimaryValueAxisTitleRotated             =  309, /*!< Rotate primary value axis title. */
    msoElementPrimaryValueAxisTitleVertical            =  310, /*!< Display primary value axis title vertically. */
    msoElementPrimaryValueGridLinesMajor               =  330, /*!< Display major gridlines along primary value axis. */
    msoElementPrimaryValueGridLinesMinor               =  329, /*!< Display minor gridlines along primary value axis. */
    msoElementPrimaryValueGridLinesMinorMajor          =  331, /*!< Display both major and minor gridlines along primary value axis. */
    msoElementPrimaryValueGridLinesNone                =  328, /*!< Do not display grid lines along primary value axis. */
    msoElementSecondaryCategoryAxisBillions            =  378, /*!< Use billions for secondary category axis units. */
    msoElementSecondaryCategoryAxisLogScale            =  379, /*!< Use log scale for secondary category axis. */
    msoElementSecondaryCategoryAxisMillions            =  377, /*!< Use millions for secondary category axis units. */
    msoElementSecondaryCategoryAxisNone                =  358, /*!< Do not display secondary category axis. */
    msoElementSecondaryCategoryAxisReverse             =  361, /*!< Reverse secondary category axis. */
    msoElementSecondaryCategoryAxisShow                =  359, /*!< Display secondary category axis. */
    msoElementSecondaryCategoryAxisThousands           =  376, /*!< Use thousands for secondary category axis units. */
    msoElementSecondaryCategoryAxisTitleAdjacentToAxis =  313, /*!< Display secondary category axis title adjacent to axis. */
    msoElementSecondaryCategoryAxisTitleBelowAxis      =  314, /*!< Display secondary category axis title below axis. */
    msoElementSecondaryCategoryAxisTitleHorizontal     =  317, /*!< Display secondary category axis title horizontally. */
    msoElementSecondaryCategoryAxisTitleNone           =  312, /*!< Do not display secondary category axis title. */
    msoElementSecondaryCategoryAxisTitleRotated        =  315, /*!< Rotate secondary category axis title. */
    msoElementSecondaryCategoryAxisTitleVertical       =  316, /*!< Display secondary category axis title vertically. */
    msoElementSecondaryCategoryAxisWithoutLabels       =  360, /*!< Display secondary category axis without labels. */
    msoElementSecondaryCategoryGridLinesMajor          =  342, /*!< Display major gridlines along secondary category axis. */
    msoElementSecondaryCategoryGridLinesMinor          =  341, /*!< Display minor gridlines along secondary category axis. */
    msoElementSecondaryCategoryGridLinesMinorMajor     =  343, /*!< Display both major and minor gridlines along secondary category axis. */
    msoElementSecondaryCategoryGridLinesNone           =  340, /*!< Do not display grid lines along secondary category axis. */
    msoElementSecondaryValueAxisBillions               =  366, /*!< Use billions for secondary value axis units. */
    msoElementSecondaryValueAxisLogScale               =  367, /*!< Use log scale for secondary value axis. */
    msoElementSecondaryValueAxisMillions               =  365, /*!< Use millions for secondary value axis units. */
    msoElementSecondaryValueAxisNone                   =  362, /*!< Do not display secondary value axis. */
    msoElementSecondaryValueAxisShow                   =  363, /*!< Display secondary value axis. */
    msoElementSecondaryValueAxisThousands              =  364, /*!< Use thousands for secondary value axis units. */
    msoElementSecondaryValueAxisTitleAdjacentToAxis    =  319, /*!< Display secondary value axis title adjacent to axis. */
    msoElementSecondaryValueAxisTitleBelowAxis         =  320, /*!< Display secondary value axis title below axis. */
    msoElementSecondaryValueAxisTitleHorizontal        =  323, /*!< Display secondary value axis title horizontally. */
    msoElementSecondaryValueAxisTitleNone              =  318, /*!< Do not display secondary value axis title. */
    msoElementSecondaryValueAxisTitleRotated           =  321, /*!< Rotate secondary value axis title. */
    msoElementSecondaryValueAxisTitleVertical          =  322, /*!< Display secondary value axis title vertically. */
    msoElementSecondaryValueGridLinesMajor             =  338, /*!< Display major gridlines along secondary value axis. */
    msoElementSecondaryValueGridLinesMinor             =  337, /*!< Display minor gridlines along secondary value axis. */
    msoElementSecondaryValueGridLinesMinorMajor        =  339, /*!< Display both major and minor gridlines along secondary value axis. */
    msoElementSecondaryValueGridLinesNone              =  336, /*!< Do not display gridlines along secondary value axis. */
    msoElementSeriesAxisGridLinesMajor                 =  346, /*!< Display major gridlines along series axis. */
    msoElementSeriesAxisGridLinesMinor                 =  345, /*!< Display minor gridlines along series axis. */
    msoElementSeriesAxisGridLinesMinorMajor            =  347, /*!< Display both major and minor gridlines along series axis. */
    msoElementSeriesAxisGridLinesNone                  =  344, /*!< Do not display gridlines along series axis. */
    msoElementSeriesAxisNone                           =  368, /*!< Do not display series axis. */
    msoElementSeriesAxisReverse                        =  371, /*!< Reverse series axis. */
    msoElementSeriesAxisShow                           =  369, /*!< Display series axis. */
    msoElementSeriesAxisTitleHorizontal                =  327, /*!< Display series axis title horizontally. */
    msoElementSeriesAxisTitleNone                      =  324, /*!< Do not display series axis title. */
    msoElementSeriesAxisTitleRotated                   =  325, /*!< Rotate series axis title. */
    msoElementSeriesAxisTitleVertical                  =  326, /*!< Display series axis title vertically. */
    msoElementSeriesAxisWithoutLabeling                =  370, /*!< Display series axis title without labeling. */
    msoElementTrendlineAddExponential                  =  602, /*!< Add an exponential trendline. */
    msoElementTrendlineAddLinear                       =  601, /*!< Add a linear trendline. */
    msoElementTrendlineAddLinearForecast               =  603, /*!< Add a linear forecast. */
    msoElementTrendlineAddTwoPeriodMovingAverage       =  604, /*!< Add a two-period moving average. */
    msoElementTrendlineNone                            =  600, /*!< Do not display trendline. */
    msoElementUpDownBarsNone                           =  900, /*!< Do not display up/down bars. */
    msoElementUpDownBarsShow                           =  901, /*!< Display up/down bars. */
};

/**
Specifies the type of data field to be inserted into a data label in a chart.

[Official VBA documentation for MsoChartFieldType](https://docs.microsoft.com/office/vba/api/office.msochartfieldtype)
*/
enum MsoChartFieldType
{
    msoChartFieldBubbleSize   = 1, /*!< Specifies the Bubble size of the data point. */
    msoChartFieldCategoryName = 2, /*!< Specifies the category name size of the data point. */
    msoChartFieldFormula      = 6, /*!< Specifies the formula used in the data point. */
    msoChartFieldPercentage   = 3, /*!< Specifies a percentage of the values. */
    msoChartFieldSeriesName   = 4, /*!< Specifies the data series name. */
    msoChartFieldValue        = 5, /*!< Specifies the value of the data field. */
    msoChartFieldRange        = 7, /*!< Specifies the value of a range of data. */
};

/**
Specifies Clipboard formats.

[Official VBA documentation for MsoClipboardFormat](https://docs.microsoft.com/office/vba/api/office.msoclipboardformat)
*/
enum MsoClipboardFormat
{
    msoClipboardFormatHTML      =  2, /*!< HTML format */
    msoClipboardFormatMixed     = -2, /*!< Return value only; indicates a combination of the other states. */
    msoClipboardFormatNative    =  1, /*!< Native format */
    msoClipboardFormatPlainText =  4, /*!< Plain text format */
    msoClipboardFormatRTF       =  3, /*!< RTF format */
};

/**
Specifies the color type.

[Official VBA documentation for MsoColorType](https://docs.microsoft.com/office/vba/api/office.msocolortype)
*/
enum MsoColorType
{
    msoColorTypeCMS    =  4, /*!< Color Management System color type. */
    msoColorTypeCMYK   =  3, /*!< Color is determined by values of cyan, magenta, yellow, and black. */
    msoColorTypeInk    =  5, /*!< Not supported. */
    msoColorTypeMixed  = -2, /*!< Not supported. */
    msoColorTypeRGB    =  1, /*!< Color is determined by values of red, green, and blue. */
    msoColorTypeScheme =  2, /*!< Color is defined by an application-specific scheme. */
};

/**
Specifies whether the command bar combo box includes a label or not.

[Official VBA documentation for MsoComboStyle](https://docs.microsoft.com/office/vba/api/office.msocombostyle)
*/
enum MsoComboStyle
{
    msoComboLabel  = 1, /*!< Combo box includes a label, specified by the **Caption** property of the combo box. */
    msoComboNormal = 0, /*!< Combo box does not include a label. */
};

/**
Specifies whether the command bar button is a hyperlink. If the command bar button is a hyperlink, further specifies whether the hyperlink should launch another application such as the browser or insert a picture at the active selection point.

[Official VBA documentation for MsoCommandBarButtonHyperlinkType](https://docs.microsoft.com/office/vba/api/office.msocommandbarbuttonhyperlinktype)
*/
enum MsoCommandBarButtonHyperlinkType
{
    msoCommandBarButtonHyperlinkInsertPicture = 2, /*!< Clicking the command bar button inserts a picture at the active selection point. */
    msoCommandBarButtonHyperlinkNone          = 0, /*!< The command bar button is not a hyperlink. */
    msoCommandBarButtonHyperlinkOpen          = 1, /*!< Clicking the command bar button opens the link specified in the command bar button's **TooltipText** property. */
};

/**
Specifies a type of connector.

[Official VBA documentation for MsoConnectorType](https://docs.microsoft.com/office/vba/api/office.msoconnectortype)
*/
enum MsoConnectorType
{
    msoConnectorCurve     =  3, /*!< Curved connector */
    msoConnectorElbow     =  2, /*!< Elbow connector */
    msoConnectorStraight  =  1, /*!< Straight line connector */
    msoConnectorTypeMixed = -2, /*!< Return value only; indicates a combination of the other states. */
};

/**
Specifies the address type for a contact card.

[Official VBA documentation for MsoContactCardAddressType](https://docs.microsoft.com/office/vba/api/office.msocontactcardaddresstype)
*/
enum MsoContactCardAddressType
{
    msoContactCardAddressTypeUnknown = 0, /*!< An unknown identifier for an address. */
    msoContactCardAddressTypeOutlook = 1, /*!< A unique identifier for an Outlook address. */
    msoContactCardAddressTypeSMTP    = 2, /*!< A unique identifier for an SMTP address. */
    msoContactCardAddressTypeIM      = 1, /*!< A unique identifier for an IM address. */
};

/**
Specifies how the contact card is displayed.

[Official VBA documentation for MsoContactCardStyle](https://docs.microsoft.com/office/vba/api/office.msocontactcardstyle)
*/
enum MsoContactCardStyle
{
    msoContactCardFull  = 1, /*!< The contact card is displayed as a full card. */
    msoContactCardHover = 0, /*!< The contact card is displayed as a hover card. */
};

/**
Specifies the contact card type.

[Official VBA documentation for MsoContactCardType](https://docs.microsoft.com/office/vba/api/office.msocontactcardtype)
*/
enum MsoContactCardType
{
    msoContactCardTypeEnterpriseContact        = 0, /*!< A contact card for an enterprise contact address. */
    msoContactCardTypePersonalContact          = 1, /*!< A contact card for a personal contact address. */
    msoContactCardTypeUnknownContact           = 2, /*!< A contact card for an unknown contact address. */
    msoContactCardTypeEnterpriseGroup          = 3, /*!< A contact card for an enterprise distribution list contact address. */
    msoContactCardTypePersonalDistributionList = 4, /*!< A contact card for a personal distribution list contact address. */
};

/**
Specifies the OLE client and OLE server roles in which a command bar control is used when two Microsoft Office applications are merged.

[Official VBA documentation for MsoControlOLEUsage](https://docs.microsoft.com/office/vba/api/office.msocontrololeusage)
*/
enum MsoControlOLEUsage
{
    msoControlOLEUsageBoth    = 3, /*!< Control runs on both client and server. */
    msoControlOLEUsageClient  = 2, /*!< Client-only control. */
    msoControlOLEUsageNeither = 0, /*!< Control runs on neither client nor server. */
    msoControlOLEUsageServer  = 1, /*!< Server-only control. */
};

/**
Specifies the type of the command bar control. Used with the **Add** method of the **CommandBarControls** object. Only a limited set of the control types can be created via the **CommandBars** object model: **msoControlButton**, **msoControlEdit**, **msoControlDropdown**, **msoControlComboBox**, **msoControlPopup**, and **msoControlActiveX**. Other control types may exist on built-in or add-in command bars, but cannot be created via the object model.

[Official VBA documentation for MsoControlType](https://docs.microsoft.com/office/vba/api/office.msocontroltype)
*/
enum MsoControlType
{
    msoControlActiveX             = 22, /*!< ActiveX control. */
    msoControlAutoCompleteCombo   = 26, /*!< Combo box in which the first matching choice is automatically filled in as the user types. Cannot be created through the object model. */
    msoControlButton              =  1, /*!< Command button. */
    msoControlButtonDropdown      =  5, /*!< Drop-down button. Cannot be created through the object model. */
    msoControlButtonPopup         = 12, /*!< Pop-up button. Cannot be created through the object model. */
    msoControlComboBox            =  4, /*!< Combo box. */
    msoControlCustom              =  0, /*!< Custom control. Cannot be created through the object model. */
    msoControlDropdown            =  3, /*!< Drop-down list. */
    msoControlEdit                =  2, /*!< Text box. */
    msoControlExpandingGrid       = 16, /*!< Expanding grid. Cannot be created through the object model. */
    msoControlGauge               = 19, /*!< Gauge control. Cannot be created through the object model. */
    msoControlGenericDropdown     =  8, /*!< Generic drop-down list. Cannot be created through the object model. */
    msoControlGraphicCombo        = 20, /*!< Graphic combo box. Cannot be created through the object model. */
    msoControlGraphicDropdown     =  9, /*!< Graphic drop-down list. Cannot be created through the object model. */
    msoControlGraphicPopup        = 11, /*!< Graphic pop-up menu. Cannot be created through the object model. */
    msoControlGrid                = 18, /*!< Grid. Cannot be created through the object model. */
    msoControlLabel               = 15, /*!< Label. Cannot be created through the object model. */
    msoControlLabelEx             = 24, /*!< Extended label. Cannot be created through the object model. */
    msoControlOCXDropdown         =  7, /*!< OCX drop-down list. Cannot be created through the object model. */
    msoControlPane                = 21, /*!< Pane. Cannot be created through the object model. */
    msoControlPopup               = 10, /*!< Pop-up. */
    msoControlSpinner             = 23, /*!< Spinner. Cannot be created through the object model. */
    msoControlSplitButtonMRUPopup = 14, /*!< Most Recently Used (MRU) pop-up. Cannot be created through the object model. */
    msoControlSplitButtonPopup    = 13, /*!< Split button pop-up. Cannot be created through the object model. */
    msoControlSplitDropdown       =  6, /*!< Split drop-down list. Cannot be created through the object model. */
    msoControlSplitExpandingGrid  = 17, /*!< Split expanding grid. Cannot be created through the object model. */
    msoControlWorkPane            = 25, /*!< Work pane. Cannot be created through the object model. */
};

/**
Specifies the docking behavior of the custom task pane.

[Official VBA documentation for MsoCTPDockPosition](https://docs.microsoft.com/office/vba/api/office.msoctpdockposition)
*/
enum MsoCTPDockPosition
{
    msoCTPDockPositionBottom   = 3, /*!< Dock the task pane at the bottom of the document window. */
    msoCTPDockPositionFloating = 4, /*!< Don't dock the task pane. */
    msoCTPDockPositionLeft     = 0, /*!< Dock the task pane on the left side of the document window. */
    msoCTPDockPositionRight    = 2, /*!< Dock the task pane on the right side of the document window. */
    msoCTPDockPositionTop      = 1, /*!< Dock the task pane at the top of the document window. */
};

/**
Specifies restrictions on the docking behavior of the custom task pane.

[Official VBA documentation for MsoCTPDockPositionRestrict](https://docs.microsoft.com/office/vba/api/office.msoctpdockpositionrestrict)
*/
enum MsoCTPDockPositionRestrict
{
    msoCTPDockPositionRestrictNoChange     = 1, /*!< There is no change from the current restriction setting for the task pane. */
    msoCTPDockPositionRestrictNoHorizontal = 2, /*!< Task pane can't be docked to either the right or the left side of the document window. */
    msoCTPDockPositionRestrictNone         = 0, /*!< No restrictions on docking the task pane. */
    msoCTPDockPositionRestrictNoVertical   = 3, /*!< Task pane can't be docked to either the top or the bottom of the document window. */
};

/**
Specifies the node type.

[Official VBA documentation for MsoCustomXMLNodeType](https://docs.microsoft.com/office/vba/api/office.msocustomxmlnodetype)
*/
enum MsoCustomXMLNodeType
{
    msoCustomXMLNodeAttribute             = 2, /*!< The node is an attribute. */
    msoCustomXMLNodeCData                 = 4, /*!< The node is a CData type. */
    msoCustomXMLNodeComment               = 8, /*!< The node is a comment. */
    msoCustomXMLNodeDocument              = 9, /*!< The node is a Document node. */
    msoCustomXMLNodeElement               = 1, /*!< The node is an element. */
    msoCustomXMLNodeProcessingInstruction = 7, /*!< The node is a processing instruction. */
    msoCustomXMLNodeText                  = 3, /*!< The node is a text node. */
};

/**
Indicates how validation errors will be cleared or generated.

[Official VBA documentation for MsoCustomXMLValidationErrorType](https://docs.microsoft.com/office/vba/api/office.msocustomxmlvalidationerrortype)
*/
enum MsoCustomXMLValidationErrorType
{
    msoCustomXMLValidationErrorAutomaticallyCleared = 1, /*!< Specifies that the error will clear itself whenever any change is made to the node it is bound to. */
    msoCustomXMLValidationErrorManual               = 2, /*!< Specifies that the error will not be cleared until the **Delete** method is called. */
    msoCustomXMLValidationErrorSchemaGenerated      = 0, /*!< Specifies that where there is a non-empty schema collection available for the custom XML part and validation is in effect, any changes to the part will cause validation errors. */
};

/**
Specifies the format of a date/time data type.

[Official VBA documentation for MsoDateTimeFormat](https://docs.microsoft.com/office/vba/api/office.msodatetimeformat)
*/
enum MsoDateTimeFormat
{
    msoDateTimeddddMMMMddyyyy =  2, /*!< Specifies a ddddMMMMddyyyy format. */
    msoDateTimedMMMMyyyy      =  3, /*!< Specifies a MMMMyyyy format. */
    msoDateTimedMMMyy         =  5, /*!< Specifies a MMMyy format. */
    msoDateTimeFigureOut      = 14, /*!< Specifies that the Office application will determine the format. */
    msoDateTimeFormatMixed    = -2, /*!< Specifies a mixed format. */
    msoDateTimeHmm            = 10, /*!< Specifies a Hmm format. */
    msoDateTimehmmAMPM        = 12, /*!< Specifies a hmmAMPM format. */
    msoDateTimeHmmss          = 11, /*!< Specifies a Hmmss format. */
    msoDateTimehmmssAMPM      = 13, /*!< Specifies a hmmssAMPM format. */
    msoDateTimeMdyy           =  1, /*!< Specifies a Mdyy format. */
    msoDateTimeMMddyyHmm      =  8, /*!< Specifies a MMddyyHmm format. */
    msoDateTimeMMddyyhmmAMPM  =  9, /*!< Specifies a MMddyyhmmAMPM format. */
    msoDateTimeMMMMdyyyy      =  4, /*!< Specifies a MMMMdyyyy format. */
    msoDateTimeMMMMyy         =  6, /*!< Specifies a MMMMyy format. */
    msoDateTimeMMyy           =  7, /*!< Specifies a MMyy format. */
};

/**
Specifies how to evenly distribute a collection of shapes. Used with the **Distribute** method of the **ShapeRange** collection.

[Official VBA documentation for MsoDistributeCmd](https://docs.microsoft.com/office/vba/api/office.msodistributecmd)
*/
enum MsoDistributeCmd
{
    msoDistributeHorizontally = 0, /*!< Distribute horizontally. */
    msoDistributeVertically   = 1, /*!< Distribute vertically. */
};

/**
Represents the results of running a Document Inspector module.

[Official VBA documentation for MsoDocInspectorStatus](https://docs.microsoft.com/office/vba/api/office.msodocinspectorstatus)
*/
enum MsoDocInspectorStatus
{
    msoDocInspectorStatusDocOk      = 0, /*!< Indicates that the Document Inspector module returned no issues or errors. */
    msoDocInspectorStatusError      = 2, /*!< Indicates that the Document Inspector module returned an error. */
    msoDocInspectorStatusIssueFound = 1, /*!< Indicates that the Document Inspector module found one or more occurrences of the search criteria. */
};

/**
Specifies the data type for a document property.

[Official VBA documentation for MsoDocProperties](https://docs.microsoft.com/office/vba/api/office.msodocproperties)
*/
enum MsoDocProperties
{
    msoPropertyTypeBoolean = 2, /*!< Boolean value. */
    msoPropertyTypeDate    = 3, /*!< Date value. */
    msoPropertyTypeFloat   = 5, /*!< Floating point value. */
    msoPropertyTypeNumber  = 1, /*!< Integer value. */
    msoPropertyTypeString  = 4, /*!< String value. */
};

/**
Specifies the editing type of a node.

[Official VBA documentation for MsoEditingType](https://docs.microsoft.com/office/vba/api/office.msoeditingtype)
*/
enum MsoEditingType
{
    msoEditingAuto      = 0, /*!< Changes the node to a type appropriate to the segments being connected. */
    msoEditingCorner    = 1, /*!< Changes the node to a corner node. */
    msoEditingSmooth    = 2, /*!< Changes the node to a smooth curve node. */
    msoEditingSymmetric = 3, /*!< Changes the node to a symmetric curve node. */
};

/**
Specifies the document encoding (code page or character set) for the web browser to use when a user views a saved document.

[Official VBA documentation for MsoEncoding](https://docs.microsoft.com/office/vba/api/office.msoencoding)
*/
enum MsoEncoding
{
    msoEncodingArabic                                              =  1256, /*!< Arabic */
    msoEncodingArabicASMO                                          =   708, /*!< Arabic ASMO */
    msoEncodingArabicAutoDetect                                    = 51256, /*!< Web browser auto-detects type of Arabic encoding to use. */
    msoEncodingArabicTransparentASMO                               =   720, /*!< Transparent Arabic */
    msoEncodingAutoDetect                                          = 50001, /*!< Web browser auto-detects type of encoding to use. */
    msoEncodingBaltic                                              =  1257, /*!< Baltic */
    msoEncodingCentralEuropean                                     =  1250, /*!< Central European */
    msoEncodingCyrillic                                            =  1251, /*!< Cyrillic */
    msoEncodingCyrillicAutoDetect                                  = 51251, /*!< Web browser auto-detects type of Cyrillic encoding to use. */
    msoEncodingEBCDICArabic                                        = 20420, /*!< Extended Binary Coded Decimal Interchange Code (EBCDIC) Arabic */
    msoEncodingEBCDICDenmarkNorway                                 = 20277, /*!< EBCDIC as used in Denmark and Norway */
    msoEncodingEBCDICFinlandSweden                                 = 20278, /*!< EBCDIC as used in Finland and Sweden */
    msoEncodingEBCDICFrance                                        = 20297, /*!< EBCDIC as used in France */
    msoEncodingEBCDICGermany                                       = 20273, /*!< EBCDIC as used in Germany */
    msoEncodingEBCDICGreek                                         = 20423, /*!< EBCDIC as used in the Greek language */
    msoEncodingEBCDICGreekModern                                   =   875, /*!< EBCDIC as used in the Modern Greek language */
    msoEncodingEBCDICHebrew                                        = 20424, /*!< EBCDIC as used in the Hebrew language */
    msoEncodingEBCDICIcelandic                                     = 20871, /*!< EBCDIC as used in Iceland */
    msoEncodingEBCDICInternational                                 =   500, /*!< International EBCDIC */
    msoEncodingEBCDICItaly                                         = 20280, /*!< EBCDIC as used in Italy */
    msoEncodingEBCDICJapaneseKatakanaExtended                      = 20290, /*!< EBCDIC as used with Japanese Katakana (extended) */
    msoEncodingEBCDICJapaneseKatakanaExtendedAndJapanese           = 50930, /*!< EBCDIC as used with Japanese Katakana (extended) and Japanese */
    msoEncodingEBCDICJapaneseLatinExtendedAndJapanese              = 50939, /*!< EBCDIC as used with Japanese Latin (extended) and Japanese */
    msoEncodingEBCDICKoreanExtended                                = 20833, /*!< EBCDIC as used with Korean (extended) */
    msoEncodingEBCDICKoreanExtendedAndKorean                       = 50933, /*!< EBCDIC as used with Korean (extended) and Korean */
    msoEncodingEBCDICLatinAmericaSpain                             = 20284, /*!< EBCDIC as used in Latin America and Spain */
    msoEncodingEBCDICMultilingualROECELatin2                       =   870, /*!< EBCDIC Multilingual ROECE (Latin 2) */
    msoEncodingEBCDICRussian                                       = 20880, /*!< EBCDIC as used with Russian */
    msoEncodingEBCDICSerbianBulgarian                              = 21025, /*!< EBCDIC as used with Serbian and Bulgarian */
    msoEncodingEBCDICSimplifiedChineseExtendedAndSimplifiedChinese = 50935, /*!< EBCDIC as used with Simplified Chinese (extended) and Simplified Chinese */
    msoEncodingEBCDICThai                                          = 20838, /*!< EBCDIC as used with Thai */
    msoEncodingEBCDICTurkish                                       = 20905, /*!< EBCDIC as used with Turkish */
    msoEncodingEBCDICTurkishLatin5                                 =  1026, /*!< EBCDIC as used with Turkish (Latin 5) */
    msoEncodingEBCDICUnitedKingdom                                 = 20285, /*!< EBCDIC as used in the United Kingdom */
    msoEncodingEBCDICUSCanada                                      =    37, /*!< EBCDIC as used in the United States and Canada */
    msoEncodingEBCDICUSCanadaAndJapanese                           = 50931, /*!< EBCDIC as used in the United States and Canada, and with Japanese */
    msoEncodingEBCDICUSCanadaAndTraditionalChinese                 = 50937, /*!< EBCDIC as used in the United States and Canada, and with Traditional Chinese */
    msoEncodingEUCChineseSimplifiedChinese                         = 51936, /*!< Extended Unix Code (EUC) as used with Chinese and Simplified Chinese */
    msoEncodingEUCJapanese                                         = 51932, /*!< EUC as used with Japanese */
    msoEncodingEUCKorean                                           = 51949, /*!< EUC as used with Korean */
    msoEncodingEUCTaiwaneseTraditionalChinese                      = 51950, /*!< EUC as used with Taiwanese and Traditional Chinese */
    msoEncodingEuropa3                                             = 29001, /*!< Europa */
    msoEncodingExtAlphaLowercase                                   = 21027, /*!< Extended Alpha lowercase */
    msoEncodingGreek                                               =  1253, /*!< Greek */
    msoEncodingGreekAutoDetect                                     = 51253, /*!< Web browser auto-detects type of Greek encoding to use. */
    msoEncodingHebrew                                              =  1255, /*!< Hebrew */
    msoEncodingHZGBSimplifiedChinese                               = 52936, /*!< Simplified Chinese (HZGB) */
    msoEncodingIA5German                                           = 20106, /*!< German (International Alphabet No. 5, or IA5) */
    msoEncodingIA5IRV                                              = 20105, /*!< IA5, International Reference Version (IRV) */
    msoEncodingIA5Norwegian                                        = 20108, /*!< IA5 as used with Norwegian */
    msoEncodingIA5Swedish                                          = 20107, /*!< IA5 as used with Swedish */
    msoEncodingISCIIAssamese                                       = 57006, /*!< Indian Script Code for Information Interchange (ISCII) as used with Assamese */
    msoEncodingISCIIBengali                                        = 57003, /*!< ISCII as used with Bengali */
    msoEncodingISCIIDevanagari                                     = 57002, /*!< ISCII as used with Devanagari */
    msoEncodingISCIIGujarati                                       = 57010, /*!< ISCII as used with Gujarati */
    msoEncodingISCIIKannada                                        = 57008, /*!< ISCII as used with Kannada */
    msoEncodingISCIIMalayalam                                      = 57009, /*!< ISCII as used with Malayalam */
    msoEncodingISCIIOriya                                          = 57007, /*!< ISCII as used with Odia (Oriya) */
    msoEncodingISCIIPunjabi                                        = 57011, /*!< ISCII as used with Punjabi */
    msoEncodingISCIITamil                                          = 57004, /*!< ISCII as used with Tamil */
    msoEncodingISCIITelugu                                         = 57005, /*!< ISCII as used with Telugu */
    msoEncodingISO2022CNSimplifiedChinese                          = 50229, /*!< ISO 2022-CN encoding as used with Simplified Chinese */
    msoEncodingISO2022CNTraditionalChinese                         = 50227, /*!< ISO 2022-CN encoding as used with Traditional Chinese */
    msoEncodingISO2022JPJISX02011989                               = 50222, /*!< ISO 2022-JP */
    msoEncodingISO2022JPJISX02021984                               = 50221, /*!< ISO 2022-JP */
    msoEncodingISO2022JPNoHalfwidthKatakana                        = 50220, /*!< ISO 2022-JP with no half-width Katakana */
    msoEncodingISO2022KR                                           = 50225, /*!< ISO 2022-KR */
    msoEncodingISO6937NonSpacingAccent                             = 20269, /*!< ISO 6937 Non-Spacing Accent */
    msoEncodingISO885915Latin9                                     = 28605, /*!< ISO 8859-15 with Latin 9 */
    msoEncodingISO88591Latin1                                      = 28591, /*!< ISO 8859-1 Latin 1 */
    msoEncodingISO88592CentralEurope                               = 28592, /*!< ISO 8859-2 Central Europe */
    msoEncodingISO88593Latin3                                      = 28593, /*!< ISO 8859-3 Latin 3 */
    msoEncodingISO88594Baltic                                      = 28594, /*!< ISO 8859-4 Baltic */
    msoEncodingISO88595Cyrillic                                    = 28595, /*!< ISO 8859-5 Cyrillic */
    msoEncodingISO88596Arabic                                      = 28596, /*!< ISA 8859-6 Arabic */
    msoEncodingISO88597Greek                                       = 28597, /*!< ISO 8859-7 Greek */
    msoEncodingISO88598Hebrew                                      = 28598, /*!< ISO 8859-8 Hebrew */
    msoEncodingISO88598HebrewLogical                               = 38598, /*!< ISO 8859-8 Hebrew (Logical) */
    msoEncodingISO88599Turkish                                     = 28599, /*!< ISO 8859-9 Turkish */
    msoEncodingJapaneseAutoDetect                                  = 50932, /*!< Web browser auto-detects type of Japanese encoding to use. */
    msoEncodingJapaneseShiftJIS                                    =   932, /*!< Japanese (Shift-JIS) */
    msoEncodingKOI8R                                               = 20866, /*!< KOI8-R */
    msoEncodingKOI8U                                               = 21866, /*!< K0I8-U */
    msoEncodingKorean                                              =   949, /*!< Korean */
    msoEncodingKoreanAutoDetect                                    = 50949, /*!< Web browser auto-detects type of Korean encoding to use. */
    msoEncodingKoreanJohab                                         =  1361, /*!< Korean (Johab) */
    msoEncodingMacArabic                                           = 10004, /*!< Macintosh Arabic */
    msoEncodingMacCroatia                                          = 10082, /*!< Macintosh Croatian */
    msoEncodingMacCyrillic                                         = 10007, /*!< Macintosh Cyrillic */
    msoEncodingMacGreek1                                           = 10006, /*!< Macintosh Greek */
    msoEncodingMacHebrew                                           = 10005, /*!< Macintosh Hebrew */
    msoEncodingMacIcelandic                                        = 10079, /*!< Macintosh Icelandic */
    msoEncodingMacJapanese                                         = 10001, /*!< Macintosh Japanese */
    msoEncodingMacKorean                                           = 10003, /*!< Macintosh Korean */
    msoEncodingMacLatin2                                           = 10029, /*!< Macintosh Latin 2 */
    msoEncodingMacRoman                                            = 10000, /*!< Macintosh Roman */
    msoEncodingMacRomania                                          = 10010, /*!< Macintosh Romanian */
    msoEncodingMacSimplifiedChineseGB2312                          = 10008, /*!< Macintosh Simplified Chinese (GB 2312) */
    msoEncodingMacTraditionalChineseBig5                           = 10002, /*!< Macintosh Traditional Chinese (Big 5) */
    msoEncodingMacTurkish                                          = 10081, /*!< Macintosh Turkish */
    msoEncodingMacUkraine                                          = 10017, /*!< Macintosh Ukrainian */
    msoEncodingOEMArabic                                           =   864, /*!< OEM as used with Arabic */
    msoEncodingOEMBaltic                                           =   775, /*!< OEM as used with Baltic */
    msoEncodingOEMCanadianFrench                                   =   863, /*!< OEM as used with Canadian French */
    msoEncodingOEMCyrillic                                         =   855, /*!< OEM as used with Cyrillic */
    msoEncodingOEMCyrillicII                                       =   866, /*!< OEM as used with Cyrillic II */
    msoEncodingOEMGreek437G                                        =   737, /*!< OEM as used with Greek 437G */
    msoEncodingOEMHebrew                                           =   862, /*!< OEM as used with Hebrew */
    msoEncodingOEMIcelandic                                        =   861, /*!< OEM as used with Icelandic */
    msoEncodingOEMModernGreek                                      =   869, /*!< OEM as used with Modern Greek */
    msoEncodingOEMMultilingualLatinI                               =   850, /*!< OEM as used with multi-lingual Latin I */
    msoEncodingOEMMultilingualLatinII                              =   852, /*!< OEM as used with multi-lingual Latin II */
    msoEncodingOEMNordic                                           =   865, /*!< OEM as used with Nordic languages */
    msoEncodingOEMPortuguese                                       =   860, /*!< OEM as used with Portuguese */
    msoEncodingOEMTurkish                                          =   857, /*!< OEM as used with Turkish */
    msoEncodingOEMUnitedStates                                     =   437, /*!< OEM as used in the United States */
    msoEncodingSimplifiedChineseAutoDetect                         = 50936, /*!< Web browser auto-detects type of Simplified Chinese encoding to use. */
    msoEncodingSimplifiedChineseGB18030                            = 54936, /*!< Simplified Chinese GB 18030 */
    msoEncodingSimplifiedChineseGBK                                =   936, /*!< Simplified Chinese GBK */
    msoEncodingT61                                                 = 20261, /*!< T61 */
    msoEncodingTaiwanCNS                                           = 20000, /*!< Taiwan CNS */
    msoEncodingTaiwanEten                                          = 20002, /*!< Taiwan Eten */
    msoEncodingTaiwanIBM5550                                       = 20003, /*!< Taiwan IBM 5550 */
    msoEncodingTaiwanTCA                                           = 20001, /*!< Taiwan TCA */
    msoEncodingTaiwanTeleText                                      = 20004, /*!< Taiwan Teletext */
    msoEncodingTaiwanWang                                          = 20005, /*!< Taiwan Wang */
    msoEncodingThai                                                =   874, /*!< Thai */
    msoEncodingTraditionalChineseAutoDetect                        = 50950, /*!< Web browser auto-detects type of Traditional Chinese encoding to use. */
    msoEncodingTraditionalChineseBig5                              =   950, /*!< Traditional Chinese Big 5 */
    msoEncodingTurkish                                             =  1254, /*!< Turkish */
    msoEncodingUnicodeBigEndian                                    =  1201, /*!< Unicode big endian */
    msoEncodingUnicodeLittleEndian                                 =  1200, /*!< Unicode little endian */
    msoEncodingUSASCII                                             = 20127, /*!< United States ASCII */
    msoEncodingUTF7                                                = 65000, /*!< UTF-7 encoding */
    msoEncodingUTF8                                                = 65001, /*!< UTF-8 encoding */
    msoEncodingVietnamese                                          =  1258, /*!< Vietnamese */
    msoEncodingWestern                                             =  1252, /*!< Western */
};

/**
Specifies how to use the value specified in the _ExtraInfo_ parameter of the **FollowHyperlink** method.

[Official VBA documentation for MsoExtraInfoMethod](https://docs.microsoft.com/office/vba/api/office.msoextrainfomethod)
*/
enum MsoExtraInfoMethod
{
    msoMethodGet  = 0, /*!< The value specified in the _ExtraInfo_ parameter is a string that is appended to the address. */
    msoMethodPost = 1, /*!< The value specified in the _ExtraInfo_ parameter is posted as a string or byte array. */
};

/**
Specifies whether the extrusion color is based on the extruded shape's fill (the front face of the extrusion) and automatically changes when the shape's fill changes, or whether the extrusion color is independent of the shape's fill. Used with the **ExtrusionColorType** property of the **ThreeDFormat** object.

[Official VBA documentation for MsoExtrusionColorType](https://docs.microsoft.com/office/vba/api/office.msoextrusioncolortype)
*/
enum MsoExtrusionColorType
{
    msoExtrusionColorAutomatic =  1, /*!< Extrusion color is based on shape fill. */
    msoExtrusionColorCustom    =  2, /*!< Extrusion color is independent of shape fill. */
    msoExtrusionColorTypeMixed = -2, /*!< Return value only; indicates a combination of the other states. */
};

/**
Specifies the language to use to determine which line break level is used when the line break control option is turned on.

[Official VBA documentation for MsoFarEastLineBreakLanguageID](https://docs.microsoft.com/office/vba/api/office.msofareastlinebreaklanguageid)
*/
enum MsoFarEastLineBreakLanguageID
{
    msoFarEastLineBreakLanguageJapanese           = 1041, /*!< Japanese */
    msoFarEastLineBreakLanguageKorean             = 1042, /*!< Korean */
    msoFarEastLineBreakLanguageSimplifiedChinese  = 2052, /*!< SimplifiedChinese */
    msoFarEastLineBreakLanguageTraditionalChinese = 1028, /*!< TraditionalChinese */
};

/**
Specifies how the application handles calls to methods and properties that require features not yet installed.

[Official VBA documentation for MsoFeatureInstall](https://docs.microsoft.com/office/vba/api/office.msofeatureinstall)
*/
enum MsoFeatureInstall
{
    msoFeatureInstallNone           = 0, /*!< Generates a generic automation error at run time when uninstalled features are called. */
    msoFeatureInstallOnDemand       = 1, /*!< Prompts the user to install new features. */
    msoFeatureInstallOnDemandWithUI = 2, /*!< Displays a progress meter during installation; does not prompt the user to install new features. */
};

/**
Specifies the type of a **FileDialog** object.

[Official VBA documentation for MsoFileDialogType](https://docs.microsoft.com/office/vba/api/office.msofiledialogtype)
*/
enum MsoFileDialogType
{
    msoFileDialogFilePicker   = 3, /*!< **File Picker** dialog box. */
    msoFileDialogFolderPicker = 4, /*!< **Folder Picker** dialog box. */
    msoFileDialogOpen         = 1, /*!< **Open** dialog box. */
    msoFileDialogSaveAs       = 2, /*!< **Save As** dialog box. */
};

/**
Specifies the view presented to the user in a file dialog box.

[Official VBA documentation for MsoFileDialogView](https://docs.microsoft.com/office/vba/api/office.msofiledialogview)
*/
enum MsoFileDialogView
{
    msoFileDialogViewDetails    = 2, /*!< Files displayed in a list with detail information. */
    msoFileDialogViewLargeIcons = 6, /*!< Files displayed as large icons. */
    msoFileDialogViewList       = 1, /*!< Files displayed in a list without details. */
    msoFileDialogViewPreview    = 4, /*!< Files displayed in a list with a preview pane showing the selected file. */
    msoFileDialogViewProperties = 3, /*!< Files displayed in a list with a pane showing the selected file's properties. */
    msoFileDialogViewSmallIcons = 7, /*!< Files displayed as small icons. */
    msoFileDialogViewThumbnail  = 5, /*!< Files displayed as thumbnails. */
    msoFileDialogViewTiles      = 9, /*!< Files displayed as tiled icons. */
    msoFileDialogViewWebView    = 8, /*!< Files displayed in web view. */
};

/**
Specifies the action to take when a user clicks an item in the task pane. Used with the **Add** method of the **NewFile** object.

[Official VBA documentation for MsoFileNewAction](https://docs.microsoft.com/office/vba/api/office.msofilenewaction)
*/
enum MsoFileNewAction
{
    msoCreateNewFile = 1, /*!< Create a new file. */
    msoEditFile      = 0, /*!< Edit a file. */
    msoOpenFile      = 2, /*!< Open a file. */
};

/**
Specifies the task pane section to which to add a file or where the file reference exists. Used with the **Add** method and the **Remove** method of the **NewFile** object.

[Official VBA documentation for MsoFileNewSection](https://docs.microsoft.com/office/vba/api/office.msofilenewsection)
*/
enum MsoFileNewSection
{
    msoBottomSection       = 4, /*!< **Bottom** section */
    msoNew                 = 1, /*!< **New** section */
    msoNewfromExistingFile = 2, /*!< **New from Existing File** section */
    msoNewfromTemplate     = 3, /*!< **New from Template** section */
    msoOpenDocument        = 0, /*!< **Open Document** section */
};

/**
Specifies the file validation mode.

[Official VBA documentation for MsoFileValidationMode](https://docs.microsoft.com/office/vba/api/office.msofilevalidationmode)
*/
enum MsoFileValidationMode
{
    msoFileValidationDefault = 0, /*!< Validate the file (default). */
    msoFileValidationSkip    = 1, /*!< Do not validate the file. */
};

/**
Specifies a shape's fill type.

[Official VBA documentation for MsoFillType](https://docs.microsoft.com/office/vba/api/office.msofilltype)
*/
enum MsoFillType
{
    msoFillBackground =  5, /*!< Fill is the same as the background. */
    msoFillGradient   =  3, /*!< Gradient fill */
    msoFillMixed      = -2, /*!< Mixed fill */
    msoFillPatterned  =  2, /*!< Patterned fill */
    msoFillPicture    =  6, /*!< Picture fill */
    msoFillSolid      =  1, /*!< Solid fill */
    msoFillTextured   =  4, /*!< Textured fill */
};

/**
Specifies how the **Column** and **CompareTo** properties are compared for an **ODSOFilter** object.

[Official VBA documentation for MsoFilterComparison](https://docs.microsoft.com/office/vba/api/office.msofiltercomparison)
*/
enum MsoFilterComparison
{
    msoFilterComparisonContains         = 8, /*!< Column matches CompareTo if any part of the CompareTo string is contained in the Column value. */
    msoFilterComparisonEqual            = 0, /*!< Column matches CompareTo if the CompareTo value is the same as the Column value. */
    msoFilterComparisonGreaterThan      = 3, /*!< Column matches CompareTo if the Column value is greater than the CompareTo value. */
    msoFilterComparisonGreaterThanEqual = 5, /*!< Column matches CompareTo if the Column value is greater than or equal to the CompareTo value. */
    msoFilterComparisonIsBlank          = 6, /*!< Column passes filter if Column is blank. */
    msoFilterComparisonIsNotBlank       = 7, /*!< Column passes filter if Column is blank. */
    msoFilterComparisonLessThan         = 2, /*!< Column matches CompareTo if the Column value is less than the CompareTo value. */
    msoFilterComparisonLessThanEqual    = 4, /*!< Column matches CompareTo if the Column value is less than or equal to the CompareTo value. */
    msoFilterComparisonNotContains      = 9, /*!< Column matches CompareTo if any part of the CompareTo string is not contained in the Column value. */
    msoFilterComparisonNotEqual         = 1, /*!< Column matches CompareTo if the CompareTo value is not equal to the Column value. */
};

/**
Specifies how a filter criterion relates to other filter criteria. Used with the **Conjunction** property of the **ODSOFilters** object and with the **Add** method of the **MailMergeFilters** object.

[Official VBA documentation for MsoFilterConjunction](https://docs.microsoft.com/office/vba/api/office.msofilterconjunction)
*/
enum MsoFilterConjunction
{
    msoFilterConjunctionAnd = 0, /*!< And conjunction */
    msoFilterConjunctionOr  = 1, /*!< Or conjunction */
};

/**
Specifies whether a shape should be flipped horizontally or vertically.

[Official VBA documentation for MsoFlipCmd](https://docs.microsoft.com/office/vba/api/office.msoflipcmd)
*/
enum MsoFlipCmd
{
    msoFlipHorizontal = 0, /*!< Flip horizontally. */
    msoFlipVertical   = 1, /*!< Flip vertically. */
};

/**
Represents one of the three language fonts contained in the **ThemeFonts** collection.

[Official VBA documentation for MsoFontLanguageIndex](https://docs.microsoft.com/office/vba/api/office.msofontlanguageindex)
*/
enum MsoFontLanguageIndex
{
    msoThemeComplexScript = 2, /*!< Represents the font face for Complex Script languages. The Complex Script language collection supports Arabic, Georgian, Hebrew, Indian, Thai, and Vietnamese alphabets. */
    msoThemeEastAsian     = 3, /*!< Represents the East Asian font face. East Asian Languages include Simplified Chinese, Traditional Chinese, Japanese, and Korean. */
    msoThemeLatin         = 1, /*!< Represents the Latin font face. */
};

/**
Specifies the type of gradient used in a shape's fill.

[Official VBA documentation for MsoGradientColorType](https://docs.microsoft.com/office/vba/api/office.msogradientcolortype)
*/
enum MsoGradientColorType
{
    msoGradientColorMixed   = -2, /*!< Mixed gradient */
    msoGradientMultiColor   =  4, /*!< More than two colors */
    msoGradientOneColor     =  1, /*!< One-color gradient */
    msoGradientPresetColors =  3, /*!< Gradient colors set according to a built-in gradient of the set defined by the **MsoPresetGradientType** constant. */
    msoGradientTwoColors    =  2, /*!< Two-color gradient */
};

/**
Specifies the style for a gradient fill.

[Official VBA documentation for MsoGradientStyle](https://docs.microsoft.com/office/vba/api/office.msogradientstyle)
*/
enum MsoGradientStyle
{
    msoGradientDiagonalDown =  4, /*!< Diagonal gradient moving from a top corner down to the opposite corner. */
    msoGradientDiagonalUp   =  3, /*!< Diagonal gradient moving from a bottom corner up to the opposite corner. */
    msoGradientFromCenter   =  7, /*!< Gradient running from the center out to the corners. */
    msoGradientFromCorner   =  5, /*!< Gradient running from a corner to the other three corners. */
    msoGradientFromTitle    =  6, /*!< Gradient running from the title outward. */
    msoGradientHorizontal   =  1, /*!< Gradient running horizontally across the shape. */
    msoGradientMixed        = -2, /*!< Gradient is mixed. */
    msoGradientVertical     =  2, /*!< Gradient running vertically down the shape. */
};

/**
Preset graphic styles.

[Official VBA documentation for MsoGraphicStyleIndex](https://docs.microsoft.com/office/vba/api/office.msographicstyleindex)
*/
enum MsoGraphicStyleIndex
{
    msoGraphicStylePreset1    =  1, /*!< Graphic style 1 */
    msoGraphicStylePreset10   = 10, /*!< Graphic style 10 */
    msoGraphicStylePreset11   = 11, /*!< Graphic style 11 */
    msoGraphicStylePreset12   = 12, /*!< Graphic style 12 */
    msoGraphicStylePreset13   = 13, /*!< Graphic style 13 */
    msoGraphicStylePreset14   = 14, /*!< Graphic style 14 */
    msoGraphicStylePreset15   = 15, /*!< Graphic style 15 */
    msoGraphicStylePreset16   = 16, /*!< Graphic style 16 */
    msoGraphicStylePreset17   = 17, /*!< Graphic style 17 */
    msoGraphicStylePreset18   = 18, /*!< Graphic style 18 */
    msoGraphicStylePreset19   = 19, /*!< Graphic style 19 */
    msoGraphicStylePreset2    =  2, /*!< Graphic style 2 */
    msoGraphicStylePreset20   = 20, /*!< Graphic style 20 */
    msoGraphicStylePreset21   = 21, /*!< Graphic style 21 */
    msoGraphicStylePreset22   = 22, /*!< Graphic style 22 */
    msoGraphicStylePreset23   = 23, /*!< Graphic style 23 */
    msoGraphicStylePreset24   = 24, /*!< Graphic style 24 */
    msoGraphicStylePreset25   = 25, /*!< Graphic style 25 */
    msoGraphicStylePreset26   = 26, /*!< Graphic style 26 */
    msoGraphicStylePreset27   = 27, /*!< Graphic style 27 */
    msoGraphicStylePreset28   = 28, /*!< Graphic style 28 */
    msoGraphicStylePreset3    =  3, /*!< Graphic style 3 */
    msoGraphicStylePreset4    =  4, /*!< Graphic style 4 */
    msoGraphicStylePreset5    =  5, /*!< Graphic style 5 */
    msoGraphicStylePreset6    =  6, /*!< Graphic style 6 */
    msoGraphicStylePreset7    =  7, /*!< Graphic style 7 */
    msoGraphicStylePreset8    =  8, /*!< Graphic style 8 */
    msoGraphicStylePreset9    =  9, /*!< Graphic style 9 */
    msoGraphicStyleMixed      = -2, /*!< A mix of graphic styles */
    msoGraphicStyleNotAPreset =  0, /*!< No graphic style */
};

/**
Specifies the horizontal alignment of text in a text frame. Used with the **HorizontalAnchor** property of the **TextFrame** object.

[Official VBA documentation for MsoHorizontalAnchor](https://docs.microsoft.com/office/vba/api/office.msohorizontalanchor)
*/
enum MsoHorizontalAnchor
{
    msoAnchorCenter          =  2, /*!< Text is centered horizontally. */
    msoAnchorNone            =  1, /*!< No alignment. */
    msoHorizontalAnchorMixed = -2, /*!< Return value only; indicates a combination of the other states. */
};

/**
Specifies the type of hyperlink.

[Official VBA documentation for MsoHyperlinkType](https://docs.microsoft.com/office/vba/api/office.msohyperlinktype)
*/
enum MsoHyperlinkType
{
    msoHyperlinkInlineShape = 2, /*!< Hyperlink applies to an inline shape. Used only with Microsoft Word. */
    msoHyperlinkRange       = 0, /*!< Hyperlink applies to a **Range** object. */
    msoHyperlinkShape       = 1, /*!< Hyperlink applies to a **Shape** object. */
};

/**
Specifies constants that define the IOD (install on demand) groups.

[Official VBA documentation for MsoIodGroup](https://docs.microsoft.com/office/vba/api/office.msoiodgroup)
*/
enum MsoIodGroup
{
    msoIodGroupPIAs       = 0, /*!< PIAs group */
    msoIodGroupVSTOR35Mgd = 1, /*!< VSTO 3.5 managed group */
    msoIodGroupVSTOR40Mgd = 2, /*!< VSTO 4.0 managed group */
};

/**
Specifies the language identifier.

[Official VBA documentation for MsoLanguageID](https://docs.microsoft.com/office/vba/api/office.msolanguageid)
*/
enum MsoLanguageID
{
    msoLanguageIDAfrikaans                        =  1078, /*!< The Afrikaans language */
    msoLanguageIDAlbanian                         =  1052, /*!< The Albanian language */
    msoLanguageIDAmharic                          =  1118, /*!< The Amharic language */
    msoLanguageIDArabic                           =  1025, /*!< The Arabic language */
    msoLanguageIDArabicAlgeria                    =  5121, /*!< The Arabic Algeria language */
    msoLanguageIDArabicBahrain                    = 15361, /*!< The Arabic Bahrain language */
    msoLanguageIDArabicEgypt                      =  3073, /*!< The Arabic Egypt language */
    msoLanguageIDArabicIraq                       =  2049, /*!< The Arabic Iraq language */
    msoLanguageIDArabicJordan                     = 11265, /*!< The Arabic Jordan language */
    msoLanguageIDArabicKuwait                     = 13313, /*!< The Arabic Kuwait language */
    msoLanguageIDArabicLebanon                    = 12289, /*!< The Arabic Lebanon language */
    msoLanguageIDArabicLibya                      =  4097, /*!< The Arabic Libya language */
    msoLanguageIDArabicMorocco                    =  6145, /*!< The Arabic Morocco language */
    msoLanguageIDArabicOman                       =  8193, /*!< The Arabic Oman language */
    msoLanguageIDArabicQatar                      = 16385, /*!< The Arabic Qatar language */
    msoLanguageIDArabicSyria                      = 10241, /*!< The Arabic Syria language */
    msoLanguageIDArabicTunisia                    =  7169, /*!< The Arabic Tunisia language */
    msoLanguageIDArabicUAE                        = 14337, /*!< The Arabic UAE language */
    msoLanguageIDArabicYemen                      =  9217, /*!< The Arabic Yemen language */
    msoLanguageIDArmenian                         =  1067, /*!< The Armenian language */
    msoLanguageIDAssamese                         =  1101, /*!< The Assamese language */
    msoLanguageIDAzeriCyrillic                    =  2092, /*!< The Azerbaijani Cyrillic language */
    msoLanguageIDAzeriLatin                       =  1068, /*!< The Azerbaijani Latin language */
    msoLanguageIDBasque                           =  1069, /*!< Basque (Basque) */
    msoLanguageIDBelgianDutch                     =  2067, /*!< The Belgian Dutch language */
    msoLanguageIDBelgianFrench                    =  2060, /*!< The Belgian French language */
    msoLanguageIDBengali                          =  1093, /*!< The Bengali language */
    msoLanguageIDBosnian                          =  4122, /*!< The Bosnian language */
    msoLanguageIDBosnianBosniaHerzegovinaCyrillic =  8218, /*!< The Bosnian Bosnia Herzegovina Cyrillic language */
    msoLanguageIDBosnianBosniaHerzegovinaLatin    =  5146, /*!< The Bosnian Bosnia Herzegovina Latin language */
    msoLanguageIDBrazilianPortuguese              =  1046, /*!< The Portuguese (Brazil) language */
    msoLanguageIDBulgarian                        =  1026, /*!< The Bulgarian language */
    msoLanguageIDBurmese                          =  1109, /*!< The Burmese language */
    msoLanguageIDByelorussian                     =  1059, /*!< The Belarusian language */
    msoLanguageIDCatalan                          =  1027, /*!< The Catalan language */
    msoLanguageIDCherokee                         =  1116, /*!< The Cherokee language */
    msoLanguageIDChineseHongKongSAR               =  3076, /*!< The Chinese Hong Kong SAR language */
    msoLanguageIDChineseMacaoSAR                  =  5124, /*!< The Chinese Macao SAR language */
    msoLanguageIDChineseSingapore                 =  4100, /*!< The Chinese Singapore language */
    msoLanguageIDCroatian                         =  1050, /*!< The Croatian language */
    msoLanguageIDCzech                            =  1029, /*!< The Czech language */
    msoLanguageIDDanish                           =  1030, /*!< The Danish language */
    msoLanguageIDDivehi                           =  1125, /*!< The Divehi language */
    msoLanguageIDDutch                            =  1043, /*!< The Dutch language */
    msoLanguageIDEdo                              =  1126, /*!< The Edo language */
    msoLanguageIDEnglishAUS                       =  3081, /*!< The English AUS language */
    msoLanguageIDEnglishBelize                    = 10249, /*!< The English Belize language */
    msoLanguageIDEnglishCanadian                  =  4105, /*!< The English Canadian language */
    msoLanguageIDEnglishCaribbean                 =  9225, /*!< The English Caribbean language */
    msoLanguageIDEnglishIndonesia                 = 14345, /*!< The English Indonesia language */
    msoLanguageIDEnglishIreland                   =  6153, /*!< The English Ireland language */
    msoLanguageIDEnglishJamaica                   =  8201, /*!< The English Jamaica language */
    msoLanguageIDEnglishNewZealand                =  5129, /*!< The English NewZealand language */
    msoLanguageIDEnglishPhilippines               = 13321, /*!< The English Philippines language */
    msoLanguageIDEnglishSouthAfrica               =  7177, /*!< The English South Africa language */
    msoLanguageIDEnglishTrinidadTobago            = 11273, /*!< The English Trinidad Tobago language */
    msoLanguageIDEnglishUK                        =  2057, /*!< The English UK language */
    msoLanguageIDEnglishUS                        =  1033, /*!< The English US language */
    msoLanguageIDEnglishZimbabwe                  = 12297, /*!< The English Zimbabwe language */
    msoLanguageIDEstonian                         =  1061, /*!< The Estonian language */
    msoLanguageIDFaeroese                         =  1080, /*!< The Faeroese language */
    msoLanguageIDFarsi                            =  1065, /*!< The Farsi language */
    msoLanguageIDFilipino                         =  1124, /*!< The Filipino language */
    msoLanguageIDFinnish                          =  1035, /*!< The Finnish language */
    msoLanguageIDFrench                           =  1036, /*!< The French language */
    msoLanguageIDFrenchCameroon                   = 11276, /*!< The French Cameroon language */
    msoLanguageIDFrenchCanadian                   =  3084, /*!< The French Canadian language */
    msoLanguageIDFrenchCotedIvoire                = 12300, /*!< The French Coted Ivoire language */
    msoLanguageIDFrenchHaiti                      = 15372, /*!< The French Haiti language */
    msoLanguageIDFrenchLuxembourg                 =  5132, /*!< The French Luxembourg language */
    msoLanguageIDFrenchMali                       = 13324, /*!< The French Mali language */
    msoLanguageIDFrenchMonaco                     =  6156, /*!< The French Monaco language */
    msoLanguageIDFrenchMorocco                    = 14348, /*!< The French Morocco language */
    msoLanguageIDFrenchReunion                    =  8204, /*!< The French Reunion language */
    msoLanguageIDFrenchSenegal                    = 10252, /*!< The French Senegal language */
    msoLanguageIDFrenchWestIndies                 =  7180, /*!< The French West Indies language */
    msoLanguageIDFrenchCongoDRC                   =  9228, /*!< The French Congo DRC language */
    msoLanguageIDFrisianNetherlands               =  1122, /*!< The Frisian Netherlands language */
    msoLanguageIDFulfulde                         =  1127, /*!< The Fulfulde language */
    msoLanguageIDGaelicIreland                    =  2108, /*!< The Irish (Ireland) language */
    msoLanguageIDGaelicScotland                   =  1084, /*!< The Scottish Gaelic language */
    msoLanguageIDGalician                         =  1110, /*!< The Galician language */
    msoLanguageIDGeorgian                         =  1079, /*!< The Georgian language */
    msoLanguageIDGerman                           =  1031, /*!< The German language */
    msoLanguageIDGermanAustria                    =  3079, /*!< The German Austria language */
    msoLanguageIDGermanLiechtenstein              =  5127, /*!< The German Liechtenstein language */
    msoLanguageIDGermanLuxembourg                 =  4103, /*!< The German Luxembourg language */
    msoLanguageIDGreek                            =  1032, /*!< The Greek language */
    msoLanguageIDGuarani                          =  1140, /*!< The Guarani language */
    msoLanguageIDGujarati                         =  1095, /*!< The Gujarati language */
    msoLanguageIDHausa                            =  1128, /*!< The Hausa language */
    msoLanguageIDHawaiian                         =  1141, /*!< The Hawaiian language */
    msoLanguageIDHebrew                           =  1037, /*!< The Hebrew language */
    msoLanguageIDHindi                            =  1081, /*!< The Hindi language */
    msoLanguageIDHungarian                        =  1038, /*!< The Hungarian language */
    msoLanguageIDIbibio                           =  1129, /*!< The Ibibio language */
    msoLanguageIDIcelandic                        =  1039, /*!< The Icelandic language */
    msoLanguageIDIgbo                             =  1136, /*!< The Igbo language */
    msoLanguageIDIndonesian                       =  1057, /*!< The Indonesian language */
    msoLanguageIDInuktitut                        =  1117, /*!< The Inuktitut language */
    msoLanguageIDItalian                          =  1040, /*!< The Italian language */
    msoLanguageIDJapanese                         =  1041, /*!< The Japanese language */
    msoLanguageIDKannada                          =  1099, /*!< The Kannada language */
    msoLanguageIDKanuri                           =  1137, /*!< The Kanuri language */
    msoLanguageIDKashmiri                         =  1120, /*!< The Kashmiri language */
    msoLanguageIDKashmiriDevanagari               =  2144, /*!< The Kashmiri Devanagari language */
    msoLanguageIDKazakh                           =  1087, /*!< The Kazakh language */
    msoLanguageIDKhmer                            =  1107, /*!< The Khmer language */
    msoLanguageIDKirghiz                          =  1088, /*!< The Kirghiz language */
    msoLanguageIDKonkani                          =  1111, /*!< The Konkani language */
    msoLanguageIDKorean                           =  1042, /*!< The Korean language */
    msoLanguageIDKyrgyz                           =  1088, /*!< The Kyrgyz language */
    msoLanguageIDLao                              =  1108, /*!< The Lao language */
    msoLanguageIDLatin                            =  1142, /*!< The Latin language */
    msoLanguageIDLatvian                          =  1062, /*!< The Latvian language */
    msoLanguageIDLithuanian                       =  1063, /*!< The Lithuanian language */
    msoLanguageIDMacedonianFYROM                  =  1071, /*!< The Macedonian language */
    msoLanguageIDMalayalam                        =  1100, /*!< The Malayalam language */
    msoLanguageIDMalayBruneiDarussalam            =  2110, /*!< The Malay Brunei Darussalam language */
    msoLanguageIDMalaysian                        =  1086, /*!< The Malaysian language */
    msoLanguageIDMaltese                          =  1082, /*!< The Maltese language */
    msoLanguageIDManipuri                         =  1112, /*!< The Manipuri language */
    msoLanguageIDMaori                            =  1153, /*!< The Maori language */
    msoLanguageIDMarathi                          =  1102, /*!< The Marathi language */
    msoLanguageIDMexicanSpanish                   =  2058, /*!< The Mexican Spanish language */
    msoLanguageIDMixed                            =    -2, /*!< The Mixed language */
    msoLanguageIDMongolian                        =  1104, /*!< The Mongolian language */
    msoLanguageIDNepali                           =  1121, /*!< The Nepali language */
    msoLanguageIDNone                             =     0, /*!< No language specified */
    msoLanguageIDNoProofing                       =  1024, /*!< No proofing */
    msoLanguageIDNorwegianBokmol                  =  1044, /*!< The Norwegian Bokmol language */
    msoLanguageIDNorwegianNynorsk                 =  2068, /*!< The Norwegian Nynorsk language */
    msoLanguageIDOriya                            =  1096, /*!< The Odia language */
    msoLanguageIDOromo                            =  1138, /*!< The Oromo language */
    msoLanguageIDPashto                           =  1123, /*!< The Pashto language */
    msoLanguageIDPolish                           =  1045, /*!< The Polish language */
    msoLanguageIDPortuguese                       =  2070, /*!< The Portuguese language */
    msoLanguageIDPunjabi                          =  1094, /*!< The Punjabi language */
    msoLanguageIDQuechuaBolivia                   =  1131, /*!< The Quechua Bolivia language */
    msoLanguageIDQuechuaEcuador                   =  2155, /*!< The Quechua Ecuador language */
    msoLanguageIDQuechuaPeru                      =  3179, /*!< The Quechua Peru language */
    msoLanguageIDRhaetoRomanic                    =  1047, /*!< The Rhaeto Romanic language */
    msoLanguageIDRomanian                         =  1048, /*!< The Romanian language */
    msoLanguageIDRomanianMoldova                  =  2072, /*!< The Romanian Moldova language */
    msoLanguageIDRussian                          =  1049, /*!< The Russian language */
    msoLanguageIDRussianMoldova                   =  2073, /*!< The Russian Moldova language */
    msoLanguageIDSamiLappish                      =  1083, /*!< The Sami Lappish language */
    msoLanguageIDSanskrit                         =  1103, /*!< The Sanskrit language */
    msoLanguageIDSepedi                           =  1132, /*!< The Sepedi language */
    msoLanguageIDSerbianBosniaHerzegovinaCyrillic =  7194, /*!< The Serbian Bosnia Herzegovina Cyrillic language */
    msoLanguageIDSerbianBosniaHerzegovinaLatin    =  6170, /*!< The Serbian Bosnia Herzegovina Latin language */
    msoLanguageIDSerbianCyrillic                  =  3098, /*!< The Serbian Cyrillic language */
    msoLanguageIDSerbianLatin                     =  2074, /*!< The Serbian Latin language */
    msoLanguageIDSesotho                          =  1072, /*!< The Sesotho language */
    msoLanguageIDSimplifiedChinese                =  2052, /*!< The Simplified Chinese language */
    msoLanguageIDSindhi                           =  1113, /*!< The Sindhi language */
    msoLanguageIDSindhiPakistan                   =  2137, /*!< The Sindhi Pakistan language */
    msoLanguageIDSinhalese                        =  1115, /*!< The Sinhalese language */
    msoLanguageIDSlovak                           =  1051, /*!< The Slovak language */
    msoLanguageIDSlovenian                        =  1060, /*!< The Slovenian language */
    msoLanguageIDSomali                           =  1143, /*!< The Somali language */
    msoLanguageIDSorbian                          =  1070, /*!< The Sorbian language */
    msoLanguageIDSpanish                          =  1034, /*!< The Spanish language */
    msoLanguageIDSpanishArgentina                 = 11274, /*!< The Spanish Argentina language */
    msoLanguageIDSpanishBolivia                   = 16394, /*!< The Spanish Bolivia language */
    msoLanguageIDSpanishChile                     = 13322, /*!< The Spanish Chile language */
    msoLanguageIDSpanishColombia                  =  9226, /*!< The Spanish Colombia language */
    msoLanguageIDSpanishCostaRica                 =  5130, /*!< The Spanish Costa Rica language */
    msoLanguageIDSpanishDominicanRepublic         =  7178, /*!< The Spanish Dominican Republic language */
    msoLanguageIDSpanishEcuador                   = 12298, /*!< The Spanish Ecuador language */
    msoLanguageIDSpanishElSalvador                = 17418, /*!< The Spanish El Salvador language */
    msoLanguageIDSpanishGuatemala                 =  4106, /*!< The Spanish Guatemala language */
    msoLanguageIDSpanishHonduras                  = 18442, /*!< The Spanish Honduras language */
    msoLanguageIDSpanishModernSort                =  3082, /*!< The Spanish Modern Sort language */
    msoLanguageIDSpanishNicaragua                 = 19466, /*!< The Spanish Nicaragua language */
    msoLanguageIDSpanishPanama                    =  6154, /*!< The Spanish Panama language */
    msoLanguageIDSpanishParaguay                  = 15370, /*!< The Spanish Paraguay language */
    msoLanguageIDSpanishPeru                      = 10250, /*!< The Spanish Peru language */
    msoLanguageIDSpanishPuertoRico                = 20490, /*!< The Spanish Puerto Rico language */
    msoLanguageIDSpanishUruguay                   = 14346, /*!< The Spanish Uruguay language */
    msoLanguageIDSpanishVenezuela                 =  8202, /*!< The Spanish Venezuela language */
    msoLanguageIDSutu                             =  1072, /*!< The Sutu language */
    msoLanguageIDSwahili                          =  1089, /*!< The Swahili language */
    msoLanguageIDSwedish                          =  1053, /*!< The Swedish language */
    msoLanguageIDSwedishFinland                   =  2077, /*!< The Swedish Finland language */
    msoLanguageIDSwissFrench                      =  4108, /*!< The Swiss French language */
    msoLanguageIDSwissGerman                      =  2055, /*!< The Swiss German language */
    msoLanguageIDSwissItalian                     =  2064, /*!< The Swiss Italian language */
    msoLanguageIDSyriac                           =  1114, /*!< The Syriac language */
    msoLanguageIDTajik                            =  1064, /*!< The Tajik language */
    msoLanguageIDTamazight                        =  1119, /*!< The Tamazight language */
    msoLanguageIDTamazightLatin                   =  2143, /*!< The Tamazight Latin language */
    msoLanguageIDTamil                            =  1097, /*!< The Tamil language */
    msoLanguageIDTatar                            =  1092, /*!< The Tatar language */
    msoLanguageIDTelugu                           =  1098, /*!< The Telugu language */
    msoLanguageIDThai                             =  1054, /*!< The Thai language */
    msoLanguageIDTibetan                          =  1105, /*!< The Tibetan language */
    msoLanguageIDTigrignaEritrea                  =  2163, /*!< The Tigrigna Eritrea language */
    msoLanguageIDTigrignaEthiopic                 =  1139, /*!< The Tigrigna Ethiopic language */
    msoLanguageIDTraditionalChinese               =  1028, /*!< The Traditional Chinese language */
    msoLanguageIDTsonga                           =  1073, /*!< The Tsonga language */
    msoLanguageIDTswana                           =  1074, /*!< The Tswana language */
    msoLanguageIDTurkish                          =  1055, /*!< The Turkish language */
    msoLanguageIDTurkmen                          =  1090, /*!< The Turkmen language */
    msoLanguageIDUkrainian                        =  1058, /*!< The Ukrainian language */
    msoLanguageIDUrdu                             =  1056, /*!< The Urdu language */
    msoLanguageIDUzbekCyrillic                    =  2115, /*!< The Uzbek Cyrillic language */
    msoLanguageIDUzbekLatin                       =  1091, /*!< The Uzbek Latin language */
    msoLanguageIDVenda                            =  1075, /*!< The Venda language */
    msoLanguageIDVietnamese                       =  1066, /*!< The Vietnamese language */
    msoLanguageIDWelsh                            =  1106, /*!< The Welsh language */
    msoLanguageIDXhosa                            =  1076, /*!< The Xhosa language */
    msoLanguageIDYi                               =  1144, /*!< The Yi language */
    msoLanguageIDYiddish                          =  1085, /*!< The Yiddish language */
    msoLanguageIDYoruba                           =  1130, /*!< The Yoruba language */
    msoLanguageIDZulu                             =  1077, /*!< The Zulu language */
};

/**
Indicates the effects lighting for an object.

[Official VBA documentation for MsoLightRigType](https://docs.microsoft.com/office/vba/api/office.msolightrigtype)
*/
enum MsoLightRigType
{
    msoLightRigBalanced      = 14, /*!< Specifies the Balanced effect. */
    msoLightRigBrightRoom    = 27, /*!< Specifies the BrightRoom effect. */
    msoLightRigChilly        = 22, /*!< Specifies the Chilly effect. */
    msoLightRigContrasting   = 18, /*!< Specifies the Contrasting effect. */
    msoLightRigFlat          = 24, /*!< Specifies the Flat effect. */
    msoLightRigFlood         = 17, /*!< Specifies the Flood effect. */
    msoLightRigFreezing      = 23, /*!< Specifies the Freezing effect. */
    msoLightRigGlow          = 26, /*!< Specifies the Glow effect. */
    msoLightRigHarsh         = 16, /*!< Specifies the Harsh effect. */
    msoLightRigLegacyFlat1   =  1, /*!< Specifies the LegacyFlat1 effect. */
    msoLightRigLegacyFlat2   =  2, /*!< Specifies the LegacyFlat2 effect. */
    msoLightRigLegacyFlat3   =  3, /*!< Specifies the LegacyFlat3 effect. */
    msoLightRigLegacyFlat4   =  4, /*!< Specifies the LegacyFlat4 effect. */
    msoLightRigLegacyHarsh1  =  9, /*!< Specifies the LegacyHarsh1 effect. */
    msoLightRigLegacyHarsh2  = 10, /*!< Specifies the LegacyHarsh2 effect. */
    msoLightRigLegacyHarsh3  = 11, /*!< Specifies the LegacyHarsh3 effect. */
    msoLightRigLegacyHarsh4  = 12, /*!< Specifies the LegacyHarsh4 effect. */
    msoLightRigLegacyNormal1 =  5, /*!< Specifies the LegacyNormal1 effect. */
    msoLightRigLegacyNormal2 =  6, /*!< Specifies the LegacyNormal2 effect. */
    msoLightRigLegacyNormal3 =  7, /*!< Specifies the LegacyNormal3 effect. */
    msoLightRigLegacyNormal4 =  8, /*!< Specifies the LegacyNormal4 effect. */
    msoLightRigMixed         = -2, /*!< Specifies the Mixed effect. */
    msoLightRigMorning       = 19, /*!< Specifies the Morning effect. */
    msoLightRigSoft          = 15, /*!< Specifies the Soft effect. */
    msoLightRigSunrise       = 20, /*!< Specifies the Sunrise effect. */
    msoLightRigSunset        = 21, /*!< Specifies the Sunset effect. */
    msoLightRigThreePoint    = 13, /*!< Specifies the ThreePoint effect. */
    msoLightRigTwoPoint      = 25, /*!< Specifies the TwoPoint effect. */
};

/**
Specifies the type of line cap.

[Official VBA documentation for MsoLineCapStyle](https://docs.microsoft.com/office/vba/api/office.msolinecapstyle)
*/
enum MsoLineCapStyle
{
    msoLineCapFlat   =  3, /*!< Specifies a flat line cap. */
    msoLineCapMixed  = -2, /*!< Specifies a mixture of line caps. */
    msoLineCapRound  =  2, /*!< Specifies a rounded line cap. */
    msoLineCapSquare =  1, /*!< Specifies a squared-off line cap. */
};

/**
Specifies the dash style for a line. 

[Official VBA documentation for MsoLineDashStyle](https://docs.microsoft.com/office/vba/api/office.msolinedashstyle)
*/
enum MsoLineDashStyle
{
    msoLineDash           =  4, /*!< Line consists of dashes only. */
    msoLineDashDot        =  5, /*!< Line is a dash-dot pattern. */
    msoLineDashDotDot     =  6, /*!< Line is a dash-dot-dot pattern. */
    msoLineDashStyleMixed = -2, /*!< Not supported. */
    msoLineLongDash       =  7, /*!< Line consists of long dashes. */
    msoLineLongDashDot    =  8, /*!< Line is a long dash-dot pattern. */
    msoLineRoundDot       =  3, /*!< Line is made up of round dots. */
    msoLineSolid          =  1, /*!< Line is solid. */
    msoLineSquareDot      =  2, /*!< Line is made up of square dots. */
};

/**
Specifies the type of fill for a line.

[Official VBA documentation for MsoLineFillType](https://docs.microsoft.com/office/vba/api/office.msolinefilltype)
*/
enum MsoLineFillType
{
    msoLineFillBackground =  5, /*!< Specifies the same fill as the background. */
    msoLineFillGradient   =  3, /*!< Specifies a gradient line fill. */
    msoLineFillMixed      = -2, /*!< Specifies a mixture of line fill types. */
    msoLineFillNone       =  0, /*!< No line fill. */
    msoLineFillPatterned  =  2, /*!< Specifies a pattern line fill. */
    msoLineFillPicture    =  6, /*!< Specifies a picture line fill. */
    msoLineFillSolid      =  1, /*!< Specifies a solid color line fill. */
    msoLineFillTextured   =  4, /*!< Specifies a textured line fill. */
};

/**
Specifies the type of join where two lines connect.

[Official VBA documentation for MsoLineJoinStyle](https://docs.microsoft.com/office/vba/api/office.msolinejoinstyle)
*/
enum MsoLineJoinStyle
{
    msoLineJoinBevel =  2, /*!< Specifies a beveled join. */
    msoLineJoinMiter =  3, /*!< Specifies a mitered join. */
    msoLineJoinMixed = -2, /*!< Specifies a mixture of join types. */
    msoLineJoinRound =  1, /*!< Specifies a rounded join. */
};

/**
Specifies the style for a line.

[Official VBA documentation for MsoLineStyle](https://docs.microsoft.com/office/vba/api/office.msolinestyle)
*/
enum MsoLineStyle
{
    msoLineSingle           =  1, /*!< Single line. */
    msoLineStyleMixed       = -2, /*!< Not supported. */
    msoLineThickBetweenThin =  5, /*!< Thick line with a thin line on each side. */
    msoLineThickThin        =  4, /*!< Thick line next to thin line. For horizontal lines, the thick line is above the thin line. For vertical lines, the thick line is to the left of the thin line. */
    msoLineThinThick        =  3, /*!< Thick line next to thin line. For horizontal lines, the thick line is below the thin line. For vertical lines, the thick line is to the right of the thin line. */
    msoLineThinThin         =  2, /*!< Two thin lines. */
};

/**
Specifies the animation style for Microsoft Office command bars.

[Official VBA documentation for MsoMenuAnimation](https://docs.microsoft.com/office/vba/api/office.msomenuanimation)
*/
enum MsoMenuAnimation
{
    msoMenuAnimationNone   = 0, /*!< No animation */
    msoMenuAnimationRandom = 1, /*!< Random animation */
    msoMenuAnimationSlide  = 3, /*!< Menus slide into view */
    msoMenuAnimationUnfold = 2, /*!< Menus unfold into view */
};

/**
Specifies the output of a merge shapes operation.

[Official VBA documentation for MsoMergeCmd](https://docs.microsoft.com/office/vba/api/office.msomergecmd)
*/
enum MsoMergeCmd
{
    msoMergeCombine   = 2, /*!< Creates a new shape from selected shapes. If the selected shapes overlap, the area where they overlap is cut out, or discarded. */
    msoMergeFragment  = 5, /*!< Breaks a shape into smaller parts or creates new shapes from intersecting lines or from shapes that overlap. */
    msoMergeIntersect = 3, /*!< Forms a new closed shape from the area where selected shapes overlap, eliminating non-overlapping areas. */
    msoMergeSubtract  = 4, /*!< Creates a new shape by subtracting from the primary selection the areas where subsequent selections overlap. */
    msoMergeUnion     = 1, /*!< Creates a new shape from the perimeter of two or more overlapping shapes. The new shape is a set of all the points from the original shapes. */
};

/**
Specifies the metadata property type.

[Official VBA documentation for MsoMetaPropertyType](https://docs.microsoft.com/office/vba/api/office.msometapropertytype)
*/
enum MsoMetaPropertyType
{
    msoMetaPropertyTypeBoolean           =  1, /*!< Represents a Boolean value. */
    msoMetaPropertyTypeCalculated        =  3, /*!< Represents a calculated value. */
    msoMetaPropertyTypeChoice            =  2, /*!< Represents a value from one or more choices. */
    msoMetaPropertyTypeComputed          =  4, /*!< Represents a computed value. */
    msoMetaPropertyTypeCurrency          =  5, /*!< Represents a Currency value */
    msoMetaPropertyTypeDateTime          =  6, /*!< Represents a DateTime value. */
    msoMetaPropertyTypeFillInChoice      =  7, /*!< Represents a value from two or more choices that is written-in by the user. */
    msoMetaPropertyTypeGuid              =  8, /*!< Represents a GUID value. */
    msoMetaPropertyTypeInteger           =  9, /*!< Represents an Integer value. */
    msoMetaPropertyTypeLookup            = 10, /*!< Represents a value used to lookup another value. */
    msoMetaPropertyTypeMax               = 19, /*!< Represents the maximum value for a range. */
    msoMetaPropertyTypeMultiChoice       = 12, /*!< Represents a collection of choices. */
    msoMetaPropertyTypeMultiChoiceFillIn = 13, /*!< Represents a collection of choices that require the user to write-in a value. */
    msoMetaPropertyTypeMultiChoiceLookup = 11, /*!< Represents a collection of choices used to look up another value. */
    msoMetaPropertyTypeNote              = 14, /*!< Represents a value of one or more sentences. */
    msoMetaPropertyTypeNumber            = 15, /*!< Represents a generic number data type. */
    msoMetaPropertyTypeText              = 16, /*!< Represents a text value. */
    msoMetaPropertyTypeUnknown           =  0, /*!< Represents an unknown value. */
    msoMetaPropertyTypeUrl               = 17, /*!< Represents a URL. */
    msoMetaPropertyTypeUser              = 18, /*!< Represents a category of user. */
};

/**
Specifies numbered bullet styles.

[Official VBA documentation for MsoNumberedBulletStyle](https://docs.microsoft.com/office/vba/api/office.msonumberedbulletstyle)
*/
enum MsoNumberedBulletStyle
{
    msoBulletAlphaLCParenBoth      =  8, /*!< Lowercase alphabetical bullet with opening and closing parentheses. */
    msoBulletAlphaLCParenRight     =  9, /*!< Lowercase alphabetical bullet with closing parenthesis. */
    msoBulletAlphaLCPeriod         =  0, /*!< Lowercase alphabetical bullet with period. */
    msoBulletAlphaUCParenBoth      = 10, /*!< Uppercase alphabetical bullet with opening and closing parentheses. */
    msoBulletAlphaUCParenRight     = 11, /*!< Uppercase alphabetical bullet with closing parenthesis. */
    msoBulletAlphaUCPeriod         =  1, /*!< Uppercase alphabetical bullet with period. */
    msoBulletArabicAbjadDash       = 24, /*!< Arabic Abjad bullet with a dash. */
    msoBulletArabicAlphaDash       = 23, /*!< Arabic alphabetical bullet with a dash. */
    msoBulletArabicDBPeriod        = 29, /*!< Arabic DB bullet with period. */
    msoBulletArabicDBPlain         = 28, /*!< Plain Arabic DB bullet. */
    msoBulletArabicParenBoth       = 12, /*!< Arabic bullet with opening and closing parentheses. */
    msoBulletArabicParenRight      =  2, /*!< Arabic bullet with closing parenthesis. */
    msoBulletArabicPeriod          =  3, /*!< Arabic bullet with period. */
    msoBulletArabicPlain           = 13, /*!< Plain Arabic bullet. */
    msoBulletCircleNumDBPlain      = 18, /*!< Circled number bullet. */
    msoBulletCircleNumWDBlackPlain = 20, /*!< Circled number WD black bullet. */
    msoBulletCircleNumWDWhitePlain = 19, /*!< Circled number WD white bullet. */
    msoBulletHebrewAlphaDash       = 25, /*!< Hebrew alphabetical bullet with dash. */
    msoBulletHindiAlpha1Period     = 40, /*!< Hindi alphabetical bullet 1 with period. */
    msoBulletHindiAlphaPeriod      = 36, /*!< Hindi alphabetical bullet with period. */
    msoBulletHindiNumParenRight    = 39, /*!< Hindi numbered bullet with closing parenthesis. */
    msoBulletHindiNumPeriod        = 37, /*!< Hindi numbered bullet with period. */
    msoBulletKanjiKoreanPeriod     = 27, /*!< Korean Kanji bullet with period. */
    msoBulletKanjiKoreanPlain      = 26, /*!< Korean Kanji bullet. */
    msoBulletKanjiSimpChinDBPeriod = 38, /*!< Simplified Chinese Kanji bulllet with period. */
    msoBulletRomanLCParenBoth      =  4, /*!< Lowercase roman bullet with opening and closing parentheses. */
    msoBulletRomanLCParenRight     =  5, /*!< Lowercase roman bullet with closing parenthesis. */
    msoBulletRomanLCPeriod         =  6, /*!< Lowercase roman bullet with period. */
    msoBulletRomanUCParenBoth      = 14, /*!< Uppercase roman bullet with opening and closing parentheses. */
    msoBulletRomanUCParenRight     = 15, /*!< Uppercase roman bullet with closing parenthesis. */
    msoBulletRomanUCPeriod         =  7, /*!< Uppercase roman bullet with period. */
    msoBulletSimpChinPeriod        = 17, /*!< Simplified Chinese bulllet with period. */
    msoBulletSimpChinPlain         = 16, /*!< Simplified Chinese bullet. */
    msoBulletStyleMixed            = -2, /*!< Return value only; indicates a combination of the other states. */
    msoBulletThaiAlphaParenBoth    = 32, /*!< Thai alphabetical bullet with opening and closing parentheses. */
    msoBulletThaiAlphaParenRight   = 31, /*!< Thai alphabetical bullet with closing parenthesis. */
    msoBulletThaiAlphaPeriod       = 30, /*!< Thai alphabetical bullet with period. */
    msoBulletThaiNumParenBoth      = 35, /*!< Thai numerical bullet with opening and closing parentheses. */
    msoBulletThaiNumParenRight     = 34, /*!< Thai numerical bullet with closing parenthesis. */
    msoBulletThaiNumPeriod         = 33, /*!< Thai numerical bullet with period. */
    msoBulletTradChinPeriod        = 22, /*!< Traditional Chinese bulllet with period. */
    msoBulletTradChinPlain         = 21, /*!< Traditional Chinese bulllet. */
};

/**
Specifies the menu group that a command bar pop-up control belongs to when the menu groups of the OLE server are merged with the menu groups of an OLE client (that is, when an object of the container application type is embedded in another application).

[Official VBA documentation for MsoOLEMenuGroup](https://docs.microsoft.com/office/vba/api/office.msoolemenugroup)
*/
enum MsoOLEMenuGroup
{
    msoOLEMenuGroupContainer =  2, /*!< **Container** menu */
    msoOLEMenuGroupEdit      =  1, /*!< **Edit** menu */
    msoOLEMenuGroupFile      =  0, /*!< **File** menu */
    msoOLEMenuGroupHelp      =  5, /*!< **Help** menu */
    msoOLEMenuGroupNone      = -1, /*!< Pop-up control is not merged. */
    msoOLEMenuGroupObject    =  3, /*!< **Object** menu */
    msoOLEMenuGroupWindow    =  4, /*!< **Window** menu */
};

/**
Indicates how to format the child nodes in an organization chart.

[Official VBA documentation for MsoOrgChartLayoutType](https://docs.microsoft.com/office/vba/api/office.msoorgchartlayouttype)
*/
enum MsoOrgChartLayoutType
{
    msoOrgChartLayoutBothHanging  =  2, /*!< Places child nodes vertically below the parent node on both the left and the right side. */
    msoOrgChartLayoutLeftHanging  =  3, /*!< Places child nodes vertically below the parent node on the left side. */
    msoOrgChartLayoutMixed        = -2, /*!< Return value for a parent node that has children formatted using more than one **MsoOrgChartLayoutType**. */
    msoOrgChartLayoutRightHanging =  4, /*!< Places child nodes vertically below the parent node on the right side. */
    msoOrgChartLayoutStandard     =  1, /*!< Places child nodes horizontally below the parent node. */
};

/**
Specifies orientation of an organization chart.

[Official VBA documentation for MsoOrgChartOrientation](https://docs.microsoft.com/office/vba/api/office.msoorgchartorientation)
*/
enum MsoOrgChartOrientation
{
    msoOrgChartOrientationMixed    = -2, /*!< Mixed orientation */
    msoOrgChartOrientationVertical =  1, /*!< Vertical orientation */
};

/**
Specifies orientation of an object when it is displayed or printed.

[Official VBA documentation for MsoOrientation](https://docs.microsoft.com/office/vba/api/office.msoorientation)
*/
enum MsoOrientation
{
    msoOrientationHorizontal =  1, /*!< Horizontal (landscape) orientation */
    msoOrientationMixed      = -2, /*!< Mixed orientation */
    msoOrientationVertical   =  2, /*!< Vertical (portrait) orientation */
};

/**
Specifies paragraph alignment for a text block.

[Official VBA documentation for MsoParagraphAlignment](https://docs.microsoft.com/office/vba/api/office.msoparagraphalignment)
*/
enum MsoParagraphAlignment
{
    msoAlignCenter         =  1, /*!< Specifies that the center of each line of text is aligned to the midpoint of the right and left text box margins, and the left and right edges of each line are ragged. */
    msoAlignDistribute     =  4, /*!< Specifies that the first and last characters of each line (except the last) are aligned to the left and right margins, and lines are filled by adding or subtracting the same amount from each character. The last line of the paragraph is aligned to the left margin if text direction is left-to-right, or to the right margin if text direction is right-to-left. */
    msoAlignJustify        =  3, /*!< Specifies that the first and last characters of each line (except the last) are aligned to the left and right margins, and lines are filled by adding or subtracting space between and within words. The last line of the paragraph is aligned to the left margin if text direction is left-to-right, or to the right margin if text direction is right-to-left. */
    msoAlignJustifyLow     =  6, /*!< Specifies the alignment or adjustment of kashida length in Arabic text. Kashida are special characters used to extend the joiner between two Arabic characters. */
    msoAlignLeft           =  0, /*!< Specifies that the leftmost character of each line is aligned to the left margin, and the right edge of each line is ragged. This is the default alignment for paragraphs with left-to-right text direction. */
    msoAlignMixed          = -2, /*!< Uses a combination of alignment styles. */
    msoAlignRight          =  2, /*!< Specifies that the rightmost character of each line is aligned to the right margin, and the left edge of each line is ragged. This is the default alignment for paragraphs with right-to-left text direction. */
    msoAlignThaiDistribute =  5, /*!< Specifies that the first and last characters of each line (except the last) are aligned to the left and right margins, and lines are filled by adding or subtracting space between (but not within) words. The last line of the paragraph is aligned to the left margin. */
};

/**
Specifies the format of a file or folder path.

[Official VBA documentation for MsoPathFormat](https://docs.microsoft.com/office/vba/api/office.msopathformat)
*/
enum MsoPathFormat
{
    msoPathType1     =  1, /*!< Represents the Type1 format. */
    msoPathType2     =  2, /*!< Represents the Type2 format. */
    msoPathType3     =  3, /*!< Represents the Type3 format. */
    msoPathType4     =  4, /*!< Represents the Type4 format. */
    msoPathType5     =  5, /*!< Represents the Type5 format. */
    msoPathType6     =  6, /*!< Represents the Type6 format. */
    msoPathType7     =  7, /*!< Represents the Type7 format. */
    msoPathType8     =  8, /*!< Represents the Type8 format. */
    msoPathType9     =  9, /*!< Represents the Type9 format. */
    msoPathTypeMixed = -2, /*!< Represents a mixed format. */
    msoPathTypeNone  =  0, /*!< Represents no format. */
};

/**
Specifies the fill pattern used in a shape.

[Official VBA documentation for MsoPatternType](https://docs.microsoft.com/office/vba/api/office.msopatterntype)
*/
enum MsoPatternType
{
    msoPattern10Percent              =  2, /*!< 10% of the foreground color */
    msoPattern20Percent              =  3, /*!< 20% of the foreground color */
    msoPattern25Percent              =  4, /*!< 25% of the foreground color */
    msoPattern30Percent              =  5, /*!< 30% of the foreground color */
    msoPattern40Percent              =  6, /*!< 40% of the foreground color */
    msoPattern50Percent              =  7, /*!< 50% of the foreground color */
    msoPattern5Percent               =  1, /*!< 5% of the foreground color */
    msoPattern60Percent              =  8, /*!< 60% of the foreground color */
    msoPattern70Percent              =  9, /*!< 70% of the foreground color */
    msoPattern75Percent              = 10, /*!< 75% of the foreground color */
    msoPattern80Percent              = 11, /*!< 80% of the foreground color */
    msoPattern90Percent              = 12, /*!< 90% of the foreground color */
    msoPatternCross                  = 51, /*!< Cross */
    msoPatternDarkDownwardDiagonal   = 15, /*!< Dark Downward Diagonal */
    msoPatternDarkHorizontal         = 13, /*!< Dark Horizontal */
    msoPatternDarkUpwardDiagonal     = 16, /*!< Dark Upward Diagonal */
    msoPatternDarkVertical           = 14, /*!< Dark Vertical */
    msoPatternDashedDownwardDiagonal = 28, /*!< Dashed Downward Diagonal */
    msoPatternDashedHorizontal       = 32, /*!< Dashed Horizontal */
    msoPatternDashedUpwardDiagonal   = 27, /*!< Dashed Upward Diagonal */
    msoPatternDashedVertical         = 31, /*!< Dashed Vertical */
    msoPatternDiagonalBrick          = 40, /*!< Diagonal Brick */
    msoPatternDiagonalCross          = 54, /*!< Diagonal Cross */
    msoPatternDivot                  = 46, /*!< Pattern Divot */
    msoPatternDottedDiamond          = 24, /*!< Dotted Diamond */
    msoPatternDottedGrid             = 45, /*!< Dotted Grid */
    msoPatternDownwardDiagonal       = 52, /*!< Downward Diagonal */
    msoPatternHorizontal             = 49, /*!< Horizontal */
    msoPatternHorizontalBrick        = 35, /*!< Horizontal Brick */
    msoPatternLargeCheckerBoard      = 36, /*!< Large Checker Board */
    msoPatternLargeConfetti          = 33, /*!< Large Confetti */
    msoPatternLargeGrid              = 34, /*!< Large Grid */
    msoPatternLightDownwardDiagonal  = 21, /*!< Light Downward Diagonal */
    msoPatternLightHorizontal        = 19, /*!< Light Horizontal */
    msoPatternLightUpwardDiagonal    = 22, /*!< Light Upward Diagonal */
    msoPatternLightVertical          = 20, /*!< Light Vertical */
    msoPatternMixed                  = -2, /*!< Mixed pattern */
    msoPatternNarrowHorizontal       = 30, /*!< Narrow Horizontal */
    msoPatternNarrowVertical         = 29, /*!< Narrow Vertical */
    msoPatternOutlinedDiamond        = 41, /*!< Outlined Diamond */
    msoPatternPlaid                  = 42, /*!< Plaid */
    msoPatternShingle                = 47, /*!< Shingle */
    msoPatternSmallCheckerBoard      = 17, /*!< Small Checker Board */
    msoPatternSmallConfetti          = 37, /*!< Small Confetti */
    msoPatternSmallGrid              = 23, /*!< Small Grid */
    msoPatternSolidDiamond           = 39, /*!< Solid Diamond */
    msoPatternSphere                 = 43, /*!< Sphere */
    msoPatternTrellis                = 18, /*!< Trellis */
    msoPatternUpwardDiagonal         = 53, /*!< Upward Diagonal */
    msoPatternVertical               = 50, /*!< Vertical */
    msoPatternWave                   = 48, /*!< Wave */
    msoPatternWeave                  = 44, /*!< Weave */
    msoPatternWideDownwardDiagonal   = 25, /*!< Wide Downward Diagonal */
    msoPatternWideUpwardDiagonal     = 26, /*!< Wide Upward Diagonal */
    msoPatternZigZag                 = 38, /*!< Zig Zag */
};

/**
Specifies an Information Rights Management (IRM) permission type for a document.

[Official VBA documentation for MsoPermission](https://docs.microsoft.com/office/vba/api/office.msopermission)
*/
enum MsoPermission
{
    msoPermissionChange      = 15, /*!< Permission to change */
    msoPermissionEdit        =  2, /*!< Permission to edit */
    msoPermissionExtract     =  8, /*!< Permission to extract */
    msoPermissionFullControl = 64, /*!< Full control permissions */
    msoPermissionObjModel    = 32, /*!< Permission to access the object model programmatically */
    msoPermissionPrint       = 16, /*!< Permission to print */
    msoPermissionRead        =  1, /*!< Permission to read */
    msoPermissionSave        =  4, /*!< Permission to save */
    msoPermissionView        =  1, /*!< Permission to view */
};

/**
Specifies contact-picker field types.

[Official VBA documentation for MsoPickerField](https://docs.microsoft.com/office/vba/api/office.msopickerfield)
*/
enum MsoPickerField
{
    msoPickerFieldUnknown  = 0, /*!< An unknown type of field */
    msoPickerFieldDateTime = 1, /*!< A **DateTime** field */
    msoPickerFieldNumber   = 2, /*!< A number field */
    msoPickerFieldText     = 3, /*!< A text field */
    msoPickerFieldUser     = 4, /*!< A user or group field */
    msoPickerFieldMax      = 5, /*!< The sentinel value of the enumeration */
};

/**
Specifies the color transformation applied to a picture.

[Official VBA documentation for MsoPictureColorType](https://docs.microsoft.com/office/vba/api/office.msopicturecolortype)
*/
enum MsoPictureColorType
{
    msoPictureAutomatic     =  1, /*!< Default color transformation */
    msoPictureBlackAndWhite =  3, /*!< Black-and-white transformation */
    msoPictureGrayscale     =  2, /*!< Grayscale transformation */
    msoPictureMixed         = -2, /*!< Mixed transformation */
    msoPictureWatermark     =  4, /*!< Watermark transformation */
};

/**
Specifies whether a picture will be compressed or not when inserted into a file.

[Official VBA documentation for MsoPictureCompress](https://docs.microsoft.com/office/vba/api/office.msopicturecompress)
*/
enum MsoPictureCompress
{
    msoPictureCompressDocDefault = -1, /*!< The picture is compressed or not depending on the settings for the document. */
    msoPictureCompressFalse      =  0, /*!< The picture is not compressed. */
    msoPictureCompressTrue       =  1, /*!< The picture will be compressed. */
};

/**
Specifies constants that define the types of picture effects.

[Official VBA documentation for MsoPictureEffectType](https://docs.microsoft.com/office/vba/api/office.msopictureeffecttype)
*/
enum MsoPictureEffectType
{
    msoEffectBackgroundRemoval  =  1, /*!< Background removal effect */
    msoEffectBlur               =  2, /*!< Blur effect */
    msoEffectBrightnessContrast =  3, /*!< Brightness contrast effect */
    msoEffectCement             =  4, /*!< Cement effect */
    msoEffectChalkSketch        =  5, /*!< Chalk sketch effect */
    msoEffectColorTemperature   =  6, /*!< Color temperature effect */
    msoEffectCrisscrossEtching  =  7, /*!< Crisscross etching effect */
    msoEffectCutout             =  8, /*!< Cutout effect */
    msoEffectFilmGrain          =  9, /*!< Film grain effect */
    msoEffectGlass              = 10, /*!< Glass effect */
    msoEffectGlowDiffused       = 11, /*!< Diffused glow effect */
    msoEffectGlowEdges          = 12, /*!< Glow edges effect */
    msoEffectLightScreen        = 13, /*!< Light screen effect */
    msoEffectLineDrawing        = 14, /*!< Line drawing effect */
    msoEffectMarker             = 15, /*!< Marker effect */
    msoEffectMosaicBubbles      = 16, /*!< Mosaic bubbles */
    msoEffectNone               = 17, /*!< No effect */
    msoEffectPaintBrush         = 18, /*!< Paintbrush effect */
    msoEffectPaintStrokes       = 19, /*!< Paint strokes effect */
    msoEffectPastelsSmooth      = 20, /*!< Smooth pastel effect */
    msoEffectPencilGrayscale    = 21, /*!< Pencil greyscale effect */
    msoEffectPencilSketch       = 22, /*!< Pencil sketch effect */
    msoEffectPhotocopy          = 23, /*!< Photocopy effect */
    msoEffectPlasticWrap        = 24, /*!< Plastic wrap effect */
    msoEffectSaturation         = 25, /*!< Saturation effect */
    msoEffectSharpenSoften      = 26, /*!< Sharpen soften effect */
    msoEffectTexturizer         = 27, /*!< Texturizer effect */
    msoEffectWatercolorSponge   = 28, /*!< Watercolor sponge effect */
};

/**
Indicates the effects camera type used by the specified object.

[Official VBA documentation for MsoPresetCamera](https://docs.microsoft.com/office/vba/api/office.msopresetcamera)
*/
enum MsoPresetCamera
{
    msoCameraIsometricBottomDown                 = 23, /*!< Specifies Isometric Bottom Down. */
    msoCameraIsometricBottomUp                   = 22, /*!< Specifies Isometric Bottom Up. */
    msoCameraIsometricLeftDown                   = 25, /*!< Specifies Isometric Left Down. */
    msoCameraIsometricLeftUp                     = 24, /*!< Specifies Isometric Left Up. */
    msoCameraIsometricOffAxis1Left               = 28, /*!< Specifies Isometric OffAxis1 Left. */
    msoCameraIsometricOffAxis1Right              = 29, /*!< Specifies Isometric OffAxis1 Right. */
    msoCameraIsometricOffAxis1Top                = 30, /*!< Specifies Isometric OffAxis1 Top. */
    msoCameraIsometricOffAxis2Left               = 31, /*!< Specifies Isometric OffAxis2 Left. */
    msoCameraIsometricOffAxis2Right              = 32, /*!< Specifies Isometric OffAxis2 Right. */
    msoCameraIsometricOffAxis2Top                = 33, /*!< Specifies Isometric OffAxis2 Top. */
    msoCameraIsometricOffAxis3Bottom             = 36, /*!< Specifies Isometric OffAxis3 Bottom. */
    msoCameraIsometricOffAxis3Left               = 34, /*!< Specifies Isometric OffAxis3 Left. */
    msoCameraIsometricOffAxis3Right              = 35, /*!< Specifies Isometric OffAxis3 Right. */
    msoCameraIsometricOffAxis4Bottom             = 39, /*!< Specifies Isometric OffAxis4 Bottom. */
    msoCameraIsometricOffAxis4Left               = 37, /*!< Specifies Isometric OffAxis4 Left. */
    msoCameraIsometricOffAxis4Right              = 38, /*!< Specifies Isometric OffAxis4 Right. */
    msoCameraIsometricRightDown                  = 27, /*!< Specifies Isometric Right Down. */
    msoCameraIsometricRightUp                    = 26, /*!< Specifies Isometric Right Up. */
    msoCameraIsometricTopDown                    = 21, /*!< Specifies Isometric Top Down. */
    msoCameraIsometricTopUp                      = 20, /*!< Specifies Isometric Top Up. */
    msoCameraLegacyObliqueBottom                 =  8, /*!< Specifies Legacy Oblique Bottom. */
    msoCameraLegacyObliqueBottomLeft             =  7, /*!< Specifies Legacy Oblique Lower Left. */
    msoCameraLegacyObliqueBottomRight            =  9, /*!< Specifies Legacy Oblique Lower Right. */
    msoCameraLegacyObliqueFront                  =  5, /*!< Specifies Legacy Oblique Front. */
    msoCameraLegacyObliqueLeft                   =  4, /*!< Specifies Legacy Oblique Left. */
    msoCameraLegacyObliqueRight                  =  6, /*!< Specifies Legacy Oblique Right. */
    msoCameraLegacyObliqueTop                    =  2, /*!< Specifies Legacy Oblique Top. */
    msoCameraLegacyObliqueTopLeft                =  1, /*!< Specifies Legacy Oblique Upper Left. */
    msoCameraLegacyObliqueTopRight               =  3, /*!< Specifies Legacy Oblique Upper Right. */
    msoCameraLegacyPerspectiveBottom             = 17, /*!< Specifies Legacy Perspective Bottom. */
    msoCameraLegacyPerspectiveBottomLeft         = 16, /*!< Specifies Legacy Perspective Lower Left. */
    msoCameraLegacyPerspectiveBottomRight        = 18, /*!< Specifies Legacy Perspective Lower Right. */
    msoCameraLegacyPerspectiveFront              = 14, /*!< Specifies Legacy Perspective Front. */
    msoCameraLegacyPerspectiveLeft               = 13, /*!< Specifies Legacy Perspective Left. */
    msoCameraLegacyPerspectiveRight              = 15, /*!< Specifies Legacy Perspective Right. */
    msoCameraLegacyPerspectiveTop                = 11, /*!< Specifies Legacy Perspective Top. */
    msoCameraLegacyPerspectiveTopLeft            = 10, /*!< Specifies Legacy Perspective Upper Left. */
    msoCameraLegacyPerspectiveTopRight           = 12, /*!< Specifies Legacy Perspective Upper Right. */
    msoCameraObliqueBottom                       = 46, /*!< Specifies Oblique Bottom. */
    msoCameraObliqueBottomLeft                   = 45, /*!< Specifies Oblique Lower Left. */
    msoCameraObliqueBottomRight                  = 47, /*!< Specifies Oblique Lower Right. */
    msoCameraObliqueLeft                         = 43, /*!< Specifies Oblique Left. */
    msoCameraObliqueRight                        = 44, /*!< Specifies Oblique Right. */
    msoCameraObliqueTop                          = 41, /*!< Specifies Oblique Top. */
    msoCameraObliqueTopLeft                      = 40, /*!< Specifies Oblique Upper Left. */
    msoCameraObliqueTopRight                     = 42, /*!< Specifies Oblique Upper Right. */
    msoCameraOrthographicFront                   = 19, /*!< Specifies Orthographic Front. */
    msoCameraPerspectiveAbove                    = 51, /*!< Specifies Perspective Above. */
    msoCameraPerspectiveAboveLeftFacing          = 53, /*!< Specifies Perspective Above Left Facing. */
    msoCameraPerspectiveAboveRightFacing         = 54, /*!< Specifies Perspective Above Right Facing. */
    msoCameraPerspectiveBelow                    = 52, /*!< Specifies Perspective Below. */
    msoCameraPerspectiveContrastingLeftFacing    = 55, /*!< Specifies Perspective Contrasting Left Facing. */
    msoCameraPerspectiveContrastingRightFacing   = 56, /*!< Specifies Perspective Contrasting Right Facing. */
    msoCameraPerspectiveFront                    = 48, /*!< Specifies Perspective Front. */
    msoCameraPerspectiveHeroicExtremeLeftFacing  = 59, /*!< Specifies Perspective Heroic Extreme Left Facing. */
    msoCameraPerspectiveHeroicExtremeRightFacing = 60, /*!< Specifies Perspective Heroic Extreme Right Facing. */
    msoCameraPerspectiveHeroicLeftFacing         = 57, /*!< Specifies Perspective Heroic Left Facing. */
    msoCameraPerspectiveHeroicRightFacing        = 58, /*!< Specifies Perspective Heroic Right Facing. */
    msoCameraPerspectiveLeft                     = 49, /*!< Specifies Perspective Left. */
    msoCameraPerspectiveRelaxed                  = 61, /*!< Specifies Perspective Relaxed. */
    msoCameraPerspectiveRelaxedModerately        = 62, /*!< Specifies Perspective Relaxed Moderately. */
    msoCameraPerspectiveRight                    = 50, /*!< Specifies Perspective Right. */
    msoPresetCameraMixed                         = -2, /*!< Specifies a mixed effect. */
};

/**
Specifies the direction that the extrusion's sweep path takes away from the extruded shape (the front face of the extrusion). Used with the **PresetExtrusionDirection** property of the **ThreeDFormat** object.

[Official VBA documentation for MsoPresetExtrusionDirection](https://docs.microsoft.com/office/vba/api/office.msopresetextrusiondirection)
*/
enum MsoPresetExtrusionDirection
{
    msoExtrusionBottom               =  2, /*!< Lower part */
    msoExtrusionBottomLeft           =  3, /*!< Lower left */
    msoExtrusionBottomRight          =  1, /*!< Lower right */
    msoExtrusionLeft                 =  6, /*!< Left */
    msoExtrusionNone                 =  5, /*!< No extrusion */
    msoExtrusionRight                =  4, /*!< Right */
    msoExtrusionTop                  =  8, /*!< Upper part */
    msoExtrusionTopLeft              =  9, /*!< Upper left */
    msoExtrusionTopRight             =  7, /*!< Upper right */
    msoPresetExtrusionDirectionMixed = -2, /*!< Return value only; indicates a combination of the other states. */
};

/**
Specifies which predefined gradient to use to fill a shape.

[Official VBA documentation for MsoPresetGradientType](https://docs.microsoft.com/office/vba/api/office.msopresetgradienttype)
*/
enum MsoPresetGradientType
{
    msoGradientBrass       = 20, /*!< Brass gradient */
    msoGradientCalmWater   =  8, /*!< Calm Water gradient */
    msoGradientChrome      = 21, /*!< Chrome gradient */
    msoGradientChromeII    = 22, /*!< Chrome II gradient */
    msoGradientDaybreak    =  4, /*!< Daybreak gradient */
    msoGradientDesert      =  6, /*!< Desert gradient */
    msoGradientEarlySunset =  1, /*!< Early Sunset gradient */
    msoGradientFire        =  9, /*!< Fire gradient */
    msoGradientFog         = 10, /*!< Fog gradient */
    msoGradientGold        = 18, /*!< Gold gradient */
    msoGradientGoldII      = 19, /*!< Gold II gradient */
    msoGradientHorizon     =  5, /*!< Horizon gradient */
    msoGradientLateSunset  =  2, /*!< Late Sunset gradient */
    msoGradientMahogany    = 15, /*!< Mahogany gradient */
    msoGradientMoss        = 11, /*!< Moss gradient */
    msoGradientNightfall   =  3, /*!< Nightfall gradient */
    msoGradientOcean       =  7, /*!< Ocean gradient */
    msoGradientParchment   = 14, /*!< Parchment gradient */
    msoGradientPeacock     = 12, /*!< Peacock gradient */
    msoGradientRainbow     = 16, /*!< Rainbow gradient */
    msoGradientRainbowII   = 17, /*!< Rainbow II gradient */
    msoGradientSapphire    = 24, /*!< Sapphire gradient */
    msoGradientSilver      = 23, /*!< Silver gradient */
    msoGradientWheat       = 13, /*!< Wheat gradient */
    msoPresetGradientMixed = -2, /*!< Mixed gradient */
};

/**
Specifies the location of lighting on an extruded (three-dimensional) shape relative to the shape.

[Official VBA documentation for MsoPresetLightingDirection](https://docs.microsoft.com/office/vba/api/office.msopresetlightingdirection)
*/
enum MsoPresetLightingDirection
{
    msoLightingBottom               =  8, /*!< Lighting comes from the lower part. */
    msoLightingBottomLeft           =  7, /*!< Lighting comes from the lower-left. */
    msoLightingBottomRight          =  9, /*!< Lighting comes from the lower-right. */
    msoLightingLeft                 =  4, /*!< Lighting comes from the left. */
    msoLightingNone                 =  5, /*!< No lighting. */
    msoLightingRight                =  6, /*!< Lighting comes from the right. */
    msoLightingTop                  =  2, /*!< Lighting comes from the upper part. */
    msoLightingTopLeft              =  1, /*!< Lighting comes from the upper-left. */
    msoLightingTopRight             =  3, /*!< Lighting comes from the upper-right. */
    msoPresetLightingDirectionMixed = -2, /*!< Not supported. */
};

/**
Specifies the intensity of light used on a shape.

[Official VBA documentation for MsoPresetLightingSoftness](https://docs.microsoft.com/office/vba/api/office.msopresetlightingsoftness)
*/
enum MsoPresetLightingSoftness
{
    msoLightingBright              =  3, /*!< Bright light */
    msoLightingDim                 =  1, /*!< Dim light */
    msoLightingNormal              =  2, /*!< Normal light */
    msoPresetLightingSoftnessMixed = -2, /*!< Not supported */
};

/**
Specifies the extrusion surface material. Used with the **PresetMaterial** property of the **ThreeDFormat** object.

[Official VBA documentation for MsoPresetMaterial](https://docs.microsoft.com/office/vba/api/office.msopresetmaterial)
*/
enum MsoPresetMaterial
{
    msoMaterialClear             = 13, /*!< Clear */
    msoMaterialDarkEdge          = 11, /*!< DarkEdge */
    msoMaterialFlat              = 14, /*!< Flat */
    msoMaterialMatte             =  1, /*!< Matte */
    msoMaterialMatte2            =  5, /*!< Matte2 */
    msoMaterialMetal             =  3, /*!< Metal */
    msoMaterialMetal2            =  7, /*!< Metal2 */
    msoMaterialPlastic           =  2, /*!< Plastic */
    msoMaterialPlastic2          =  6, /*!< Plastic2 */
    msoMaterialPowder            = 10, /*!< Powder */
    msoMaterialSoftEdge          = 12, /*!< Soft Edge */
    msoMaterialSoftMetal         = 15, /*!< Soft Metal */
    msoMaterialTranslucentPowder =  9, /*!< Translucent Powder */
    msoMaterialWarmMatte         =  8, /*!< Warm Matte */
    msoMaterialWireFrame         =  4, /*!< Wireframe */
    msoPresetMaterialMixed       = -2, /*!< Mixed Material */
};

/**
Specifies what text effect to use on a **WordArt** object. Refer to the **WordArt Gallery** dialog box in any Microsoft Office product that supports WordArt to see a preview of each effect. The values of the **MsoPresetTextEffect** constants correspond to the formats listed in the **WordArt Gallery** dialog box (numbered from left to right and from top to bottom).

[Official VBA documentation for MsoPresetTextEffect](https://docs.microsoft.com/office/vba/api/office.msopresettexteffect)
*/
enum MsoPresetTextEffect
{
    msoTextEffect1     =  0, /*!< First text effect */
    msoTextEffect10    =  9, /*!< Tenth text effect */
    msoTextEffect11    = 10, /*!< Eleventh text effect */
    msoTextEffect12    = 11, /*!< Twelfth text effect */
    msoTextEffect13    = 12, /*!< Thirteenth text effect */
    msoTextEffect14    = 13, /*!< Fourteenth text effect */
    msoTextEffect15    = 14, /*!< Fifteenth text effect */
    msoTextEffect16    = 15, /*!< Sixteenth text effect */
    msoTextEffect17    = 16, /*!< Seventeenth text effect */
    msoTextEffect18    = 17, /*!< Eighteenth text effect */
    msoTextEffect19    = 18, /*!< Nineteenth text effect */
    msoTextEffect2     =  1, /*!< Second text effect */
    msoTextEffect20    = 19, /*!< Twentieth text effect */
    msoTextEffect21    = 20, /*!< Twenty-first text effect */
    msoTextEffect22    = 21, /*!< Twenty-second text effect */
    msoTextEffect23    = 22, /*!< Twenty-third text effect */
    msoTextEffect24    = 23, /*!< Twenty-fourth text effect */
    msoTextEffect25    = 24, /*!< Twenty-fifth text effect */
    msoTextEffect26    = 25, /*!< Twenty-sixth text effect */
    msoTextEffect27    = 26, /*!< Twenty-seventh text effect */
    msoTextEffect28    = 27, /*!< Twenty-eighth text effect */
    msoTextEffect29    = 28, /*!< Twenty-ninth text effect */
    msoTextEffect3     =  2, /*!< Third text effect */
    msoTextEffect30    = 29, /*!< Thirtieth text effect */
    msoTextEffect31    = 30, /*!< Thirty-first text effect */
    msoTextEffect32    = 31, /*!< Thirty-second text effect */
    msoTextEffect33    = 32, /*!< Thirty-third text effect */
    msoTextEffect34    = 33, /*!< Thirty-fourth text effect */
    msoTextEffect35    = 34, /*!< Thirty-fifth text effect */
    msoTextEffect36    = 35, /*!< Thirty-sixth text effect */
    msoTextEffect37    = 36, /*!< Thirty-seventh text effect */
    msoTextEffect38    = 37, /*!< Thirty-eighth text effect */
    msoTextEffect39    = 38, /*!< Thirty-ninth text effect */
    msoTextEffect4     =  3, /*!< Fourth text effect */
    msoTextEffect40    = 39, /*!< Fortieth text effect */
    msoTextEffect41    = 40, /*!< Forty-first text effect */
    msoTextEffect42    = 41, /*!< Forty-second text effect */
    msoTextEffect43    = 42, /*!< Forty-third text effect */
    msoTextEffect44    = 43, /*!< Forty-fourth text effect */
    msoTextEffect45    = 44, /*!< Forty-fifth text effect */
    msoTextEffect46    = 45, /*!< Forty-sixth text effect */
    msoTextEffect47    = 46, /*!< Forty-seventh text effect */
    msoTextEffect48    = 47, /*!< Forty-eighth text effect */
    msoTextEffect49    = 48, /*!< Forty-ninth text effect */
    msoTextEffect5     =  4, /*!< Fifth text effect */
    msoTextEffect50    = 49, /*!< Fiftieth text effect */
    msoTextEffect6     =  5, /*!< Sixth text effect */
    msoTextEffect7     =  6, /*!< Seventh text effect */
    msoTextEffect8     =  7, /*!< Eighth text effect */
    msoTextEffect9     =  8, /*!< Ninth text effect */
    msoTextEffectMixed = -2, /*!< Not used */
};

/**
Specifies shape of WordArt text. You can see a preview of each text effect shape by selecting **WordArt Shape** on the **WordArt** toolbar.

[Official VBA documentation for MsoPresetTextEffectShape](https://docs.microsoft.com/office/vba/api/office.msopresettexteffectshape)
*/
enum MsoPresetTextEffectShape
{
    msoTextEffectShapeArchDownCurve         = 10, /*!< Text is an arch that curves down. */
    msoTextEffectShapeArchDownPour          = 14, /*!< Text is a 3D arch that curves down. */
    msoTextEffectShapeArchUpCurve           =  9, /*!< Text is an arch that curves up. */
    msoTextEffectShapeArchUpPour            = 13, /*!< Text is a 3D arch that curves up. */
    msoTextEffectShapeButtonCurve           = 12, /*!< Text is curved around a center "button." */
    msoTextEffectShapeButtonPour            = 16, /*!< Text is seen in 3D, curved around a center "button." */
    msoTextEffectShapeCanDown               = 20, /*!< Text is stretched to fill the height of the shape, with only a slight curve down. */
    msoTextEffectShapeCanUp                 = 19, /*!< Text is stretched to fill the height of the shape, with only a slight curve up. */
    msoTextEffectShapeCascadeDown           = 40, /*!< Text slants up and to the right as font size decreases. */
    msoTextEffectShapeCascadeUp             = 39, /*!< Text slants down and to the right as font size increases. */
    msoTextEffectShapeChevronDown           =  6, /*!< Text slants up to its center point and then slants down. */
    msoTextEffectShapeChevronUp             =  5, /*!< Text slants down to its center point and then slants up. */
    msoTextEffectShapeCircleCurve           = 11, /*!< Text follows a circle, reading clockwise. */
    msoTextEffectShapeCirclePour            = 15, /*!< Text has a 3D effect and follows a circle, reading clockwise. */
    msoTextEffectShapeCurveDown             = 18, /*!< Text curves down and to the right as font size decreases. */
    msoTextEffectShapeCurveUp               = 17, /*!< Text curves down and to the right as font size increases. */
    msoTextEffectShapeDeflate               = 26, /*!< Font size decreases to the text's midpoint, then increases to the starting size. */
    msoTextEffectShapeDeflateBottom         = 28, /*!< Font size decreases to the text's midpoint, then increases to the starting size, while keeping the top of the text along the same curve. */
    msoTextEffectShapeDeflateInflate        = 31, /*!< Font size increases to the text's midpoint, then decreases to the starting size. */
    msoTextEffectShapeDeflateInflateDeflate = 32, /*!< Font size decreases, increases, and decreases again across the text. */
    msoTextEffectShapeDeflateTop            = 30, /*!< Font size decreases to the text's midpoint, then increases to the starting size, while keeping the bottom of the text along the same curve. */
    msoTextEffectShapeDoubleWave1           = 23, /*!< Text follows a line that curves up, then down, then up and down again. */
    msoTextEffectShapeDoubleWave2           = 24, /*!< Text follows a line that curves down, then up, then down and up again. */
    msoTextEffectShapeFadeDown              = 36, /*!< Top of the text appears to be closer to the viewer than bottom of the text. */
    msoTextEffectShapeFadeLeft              = 34, /*!< Left side of text appears to be closer to the viewer than right side. */
    msoTextEffectShapeFadeRight             = 33, /*!< Right side of text appears to be closer to the viewer than left side. */
    msoTextEffectShapeFadeUp                = 35, /*!< Bottom of text appears to be closer to the viewer than top. */
    msoTextEffectShapeInflate               = 25, /*!< Font size of text increases to its center point, then decreases. Center point of each letter is on the same straight line. */
    msoTextEffectShapeInflateBottom         = 27, /*!< Font size of text increases to its center point, then decreases. Center point of each letter follows an arch that curves downward. */
    msoTextEffectShapeInflateTop            = 29, /*!< Font size of text increases to its center point, then decreases. Center point of each letter follows an arch that curves upward. */
    msoTextEffectShapeMixed                 = -2, /*!< Not used. */
    msoTextEffectShapePlainText             =  1, /*!< No shape applied. */
    msoTextEffectShapeRingInside            =  7, /*!< Text appears to be written inside a 3D ring. */
    msoTextEffectShapeRingOutside           =  8, /*!< Text appears to be written outside a 3D ring. */
    msoTextEffectShapeSlantDown             = 38, /*!< Text slants down and to the right. */
    msoTextEffectShapeSlantUp               = 37, /*!< Text slants up and to the right. */
    msoTextEffectShapeStop                  =  2, /*!< Text follows the shape of a stop sign. */
    msoTextEffectShapeTriangleDown          =  4, /*!< Text slants up, then down. */
    msoTextEffectShapeTriangleUp            =  3, /*!< Text slants down, then up. */
    msoTextEffectShapeWave1                 = 21, /*!< Text follows a wave up, then down and up again. */
    msoTextEffectShapeWave2                 = 22, /*!< Text follows a wave down, then up and down again. */
};

/**
Specifies texture to be used to fill a shape.

[Official VBA documentation for MsoPresetTexture](https://docs.microsoft.com/office/vba/api/office.msopresettexture)
*/
enum MsoPresetTexture
{
    msoPresetTextureMixed     = -2, /*!< Not used */
    msoTextureBlueTissuePaper = 17, /*!< Blue tissue paper texture */
    msoTextureBouquet         = 20, /*!< Bouquet texture */
    msoTextureBrownMarble     = 11, /*!< Brown marble texture */
    msoTextureCanvas          =  2, /*!< Canvas texture */
    msoTextureCork            = 21, /*!< Cork texture */
    msoTextureDenim           =  3, /*!< Denim texture */
    msoTextureFishFossil      =  7, /*!< Fish fossil texture */
    msoTextureGranite         = 12, /*!< Granite texture */
    msoTextureGreenMarble     =  9, /*!< Green marble texture */
    msoTextureMediumWood      = 24, /*!< Medium wood texture */
    msoTextureNewsprint       = 13, /*!< Newsprint texture */
    msoTextureOak             = 23, /*!< Oak texture */
    msoTexturePaperBag        =  6, /*!< Paper bag texture */
    msoTexturePapyrus         =  1, /*!< Papyrus texture */
    msoTextureParchment       = 15, /*!< Parchment texture */
    msoTexturePinkTissuePaper = 18, /*!< Pink tissue paper texture */
    msoTexturePurpleMesh      = 19, /*!< Purple mesh texture */
    msoTextureRecycledPaper   = 14, /*!< Recycled paper texture */
    msoTextureSand            =  8, /*!< Sand texture */
    msoTextureStationery      = 16, /*!< Stationery texture */
    msoTextureWalnut          = 22, /*!< Walnut texture */
    msoTextureWaterDroplets   =  5, /*!< Water droplets texture */
    msoTextureWhiteMarble     = 10, /*!< White marble texture */
    msoTextureWovenMat        =  4, /*!< Woven mat texture */
};

/**
Specifies an extrusion (three-dimensional) format. The **MsoPresetThreeDFormat** constants are numbered according to the order (left to right, top to bottom) in which they are shown on the **3D Style** button on the **Formatting** toolbar. 

[Official VBA documentation for MsoPresetThreeDFormat](https://docs.microsoft.com/office/vba/api/office.msopresetthreedformat)
*/
enum MsoPresetThreeDFormat
{
    msoPresetThreeDFormatMixed = -2, /*!< Not used */
    msoThreeD1                 =  1, /*!< First 3D format */
    msoThreeD10                = 10, /*!< Tenth 3D format */
    msoThreeD11                = 11, /*!< Eleventh 3D format */
    msoThreeD12                = 12, /*!< Twelfth 3D format */
    msoThreeD13                = 13, /*!< Thirteenth 3D format */
    msoThreeD14                = 14, /*!< Fourteenth 3D format */
    msoThreeD15                = 15, /*!< Fifteenth 3D format */
    msoThreeD16                = 16, /*!< Sixteenth 3D format */
    msoThreeD17                = 17, /*!< Seventeenth 3D format */
    msoThreeD18                = 18, /*!< Eighteenth 3D format */
    msoThreeD19                = 19, /*!< Nineteenth 3D format */
    msoThreeD2                 =  2, /*!< Second 3D format */
    msoThreeD20                = 20, /*!< Twentieth 3D format */
    msoThreeD3                 =  3, /*!< Third 3D format */
    msoThreeD4                 =  4, /*!< Fourth 3D format */
    msoThreeD5                 =  5, /*!< Fifth 3D format */
    msoThreeD6                 =  6, /*!< Sixth 3D format */
    msoThreeD7                 =  7, /*!< Seventh 3D format */
    msoThreeD8                 =  8, /*!< Eighth 3D format */
    msoThreeD9                 =  9, /*!< Ninth 3D format */
};

/**
Indicates the type of recolor to use when changing a color scheme.

[Official VBA documentation for MsoRecolorType](https://docs.microsoft.com/office/vba/api/office.msorecolortype)
*/
enum MsoRecolorType
{
    msoRecolorType1     =  1, /*!< Specifies recolor Type1. */
    msoRecolorType10    = 10, /*!< Specifies recolor Type10. */
    msoRecolorType2     =  2, /*!< Specifies recolor Type2. */
    msoRecolorType3     =  3, /*!< Specifies recolor Type3. */
    msoRecolorType4     =  4, /*!< Specifies recolor Type4. */
    msoRecolorType5     =  5, /*!< Specifies recolor Type5. */
    msoRecolorType6     =  6, /*!< Specifies recolor Type6. */
    msoRecolorType7     =  7, /*!< Specifies recolor Type7. */
    msoRecolorType8     =  8, /*!< Specifies recolor Type8. */
    msoRecolorType9     =  9, /*!< Specifies recolor Type9. */
    msoRecolorTypeMixed = -2, /*!< Specifies a mixture of recolor types. */
    msoRecolorTypeNone  =  0, /*!< Specifies no recolor type. */
};

/**
Specifies the type of the **ReflectionFormat** object.

[Official VBA documentation for MsoReflectionType](https://docs.microsoft.com/office/vba/api/office.msoreflectiontype)
*/
enum MsoReflectionType
{
    msoReflectionType1     =  1, /*!< Type 1 */
    msoReflectionType2     =  2, /*!< Type 2 */
    msoReflectionType3     =  3, /*!< Type 3 */
    msoReflectionType4     =  4, /*!< Type 4 */
    msoReflectionType5     =  5, /*!< Type 5 */
    msoReflectionType6     =  6, /*!< Type 6 */
    msoReflectionType7     =  7, /*!< Type 7 */
    msoReflectionType8     =  8, /*!< Type 8 */
    msoReflectionType9     =  9, /*!< Type 9 */
    msoReflectionTypeMixed = -2, /*!< Return value only; indicates a combination of the other states. */
    msoReflectionTypeNone  =  0, /*!< No reflection type */
};

/**
Specifies where a node is added to a diagram relative to existing nodes.

[Official VBA documentation for MsoRelativeNodePosition](https://docs.microsoft.com/office/vba/api/office.msorelativenodeposition)
*/
enum MsoRelativeNodePosition
{
    msoAfterLastSibling   = 4, /*!< Node is added after last sibling. */
    msoAfterNode          = 2, /*!< Node is added after current node. */
    msoBeforeFirstSibling = 3, /*!< Node is added before first sibling. */
    msoBeforeNode         = 1, /*!< Node is added before current node. */
};

/**
Specifies which part of the shape retains its position when the shape is scaled.

[Official VBA documentation for MsoScaleFrom](https://docs.microsoft.com/office/vba/api/office.msoscalefrom)
*/
enum MsoScaleFrom
{
    msoScaleFromBottomRight = 2, /*!< Shape's lower right corner retains its position. */
    msoScaleFromMiddle      = 1, /*!< Shape's midpoint retains its position. */
    msoScaleFromTopLeft     = 0, /*!< Shape's upper left corner retains its position. */
};

/**
Specifies the ideal screen resolution to be used to view a document in a web browser.

[Official VBA documentation for MsoScreenSize](https://docs.microsoft.com/office/vba/api/office.msoscreensize)
*/
enum MsoScreenSize
{
    msoScreenSize1024x768  =  4, /*!< 1024x768 resolution */
    msoScreenSize1152x882  =  5, /*!< 1152x882 resolution */
    msoScreenSize1152x900  =  6, /*!< 1152x900 resolution */
    msoScreenSize1280x1024 =  7, /*!< 1280x1024 resolution */
    msoScreenSize1600x1200 =  8, /*!< 1600x1200 resolution */
    msoScreenSize1800x1440 =  9, /*!< 1800x1440 resolution */
    msoScreenSize1920x1200 = 10, /*!< 1920x1200 resolution */
    msoScreenSize544x376   =  0, /*!< 544x376 resolution */
    msoScreenSize640x480   =  1, /*!< 640x480 resolution */
    msoScreenSize720x512   =  2, /*!< 720x512 resolution */
    msoScreenSize800x600   =  3, /*!< 800x600 resolution */
};

/**
Specifies the type for a segment. Used with the **Insert** and **AddNodes** methods of the **FreeformBuilder** object.

[Official VBA documentation for MsoSegmentType](https://docs.microsoft.com/office/vba/api/office.msosegmenttype)
*/
enum MsoSegmentType
{
    msoSegmentCurve = 1, /*!< Curve */
    msoSegmentLine  = 0, /*!< Line */
};

/**
Specifies the type of shadowing effect.

[Official VBA documentation for MsoShadowStyle](https://docs.microsoft.com/office/vba/api/office.msoshadowstyle)
*/
enum MsoShadowStyle
{
    msoShadowStyleInnerShadow =  1, /*!< Specifies the inner shadow effect. */
    msoShadowStyleMixed       = -2, /*!< Specifies a combination of inner and outer shadow effects. */
    msoShadowStyleOuterShadow =  2, /*!< Specifies the outer shadow effect. */
};

/**
Specifies the type of shadow displayed with a shape.The **msoShadowType** constants are numbered according to their order (left to right, top to bottom) within the **Shadow Style** set shown in the **Drawing** toolbar.

[Official VBA documentation for MsoShadowType](https://docs.microsoft.com/office/vba/api/office.msoshadowtype)
*/
enum MsoShadowType
{
    msoShadow1     =  1, /*!< First shadow type */
    msoShadow10    = 10, /*!< Tenth shadow type */
    msoShadow11    = 11, /*!< Eleventh shadow type */
    msoShadow12    = 12, /*!< Twelfth shadow type */
    msoShadow13    = 13, /*!< Thirteenth shadow type */
    msoShadow14    = 14, /*!< Fourteenth shadow type */
    msoShadow15    = 15, /*!< Fifteenth shadow type */
    msoShadow16    = 16, /*!< Sixteenth shadow type */
    msoShadow17    = 17, /*!< Seventeenth shadow type */
    msoShadow18    = 18, /*!< Eighteenth shadow type */
    msoShadow19    = 19, /*!< Nineteenth shadow type */
    msoShadow2     =  2, /*!< Second shadow type */
    msoShadow20    = 20, /*!< Twentieth shadow type */
    msoShadow3     =  3, /*!< Third shadow type */
    msoShadow4     =  4, /*!< Fourth shadow type */
    msoShadow5     =  5, /*!< Fifth shadow type */
    msoShadow6     =  6, /*!< Sixth shadow type */
    msoShadow7     =  7, /*!< Seventh shadow type */
    msoShadow8     =  8, /*!< Eighth shadow type */
    msoShadow9     =  9, /*!< Ninth shadow type */
    msoShadowMixed = -2, /*!< Not supported */
};

/**
Indicates the line and shape style.

[Official VBA documentation for MsoShapeStyleIndex](https://docs.microsoft.com/office/vba/api/office.msoshapestyleindex)
*/
enum MsoShapeStyleIndex
{
    msoLineStylePreset1     = 10001, /*!< Line style 1 */
    msoLineStylePreset10    = 10010, /*!< Line style 10 */
    msoLineStylePreset11    = 10011, /*!< Line style 11 */
    msoLineStylePreset12    = 10012, /*!< Line style 12 */
    msoLineStylePreset13    = 10013, /*!< Line style 13 */
    msoLineStylePreset14    = 10014, /*!< Line style 14 */
    msoLineStylePreset15    = 10015, /*!< Line style 15 */
    msoLineStylePreset16    = 10016, /*!< Line style 16 */
    msoLineStylePreset17    = 10017, /*!< Line style 17 */
    msoLineStylePreset18    = 10018, /*!< Line style 18 */
    msoLineStylePreset19    = 10019, /*!< Line style 19 */
    msoLineStylePreset2     = 10002, /*!< Line style 2 */
    msoLineStylePreset20    = 10020, /*!< Line style 20 */
    msoLineStylePreset3     = 10003, /*!< Line style 3 */
    msoLineStylePreset4     = 10004, /*!< Line style 4 */
    msoLineStylePreset5     = 10005, /*!< Line style 5 */
    msoLineStylePreset6     = 10006, /*!< Line style 6 */
    msoLineStylePreset7     = 10007, /*!< Line style 7 */
    msoLineStylePreset8     = 10008, /*!< Line style 8 */
    msoLineStylePreset9     = 10009, /*!< Line style 9 */
    msoShapeStylePreset1    =     1, /*!< Shape style 1 */
    msoShapeStylePreset10   =    10, /*!< Shape style 10 */
    msoShapeStylePreset11   =    11, /*!< Shape style 11 */
    msoShapeStylePreset12   =    12, /*!< Shape style 12 */
    msoShapeStylePreset13   =    13, /*!< Shape style 13 */
    msoShapeStylePreset14   =    14, /*!< Shape style 14 */
    msoShapeStylePreset15   =    15, /*!< Shape style 15 */
    msoShapeStylePreset16   =    16, /*!< Shape style 16 */
    msoShapeStylePreset17   =    17, /*!< Shape style 17 */
    msoShapeStylePreset18   =    18, /*!< Shape style 18 */
    msoShapeStylePreset19   =    19, /*!< Shape style 19 */
    msoShapeStylePreset2    =     2, /*!< Shape style 2 */
    msoShapeStylePreset20   =    20, /*!< Shape style 20 */
    msoShapeStylePreset3    =     3, /*!< Shape style 3 */
    msoShapeStylePreset4    =     4, /*!< Shape style 4 */
    msoShapeStylePreset5    =     5, /*!< Shape style 5 */
    msoShapeStylePreset6    =     6, /*!< Shape style 6 */
    msoShapeStylePreset7    =     7, /*!< Shape style 7 */
    msoShapeStylePreset8    =     8, /*!< Shape style 8 */
    msoShapeStylePreset9    =     9, /*!< Shape style 9 */
    msoShapeStyleMixed      =    -2, /*!< A mix of shape styles */
    msoShapeStyleNotAPreset =     0, /*!< No shape style */
};

/**
Specifies the type of a shape or range of shapes.

[Official VBA documentation for MsoShapeType](https://docs.microsoft.com/office/vba/api/office.msoshapetype)
*/
enum MsoShapeType
{
    mso3DModel           = 30, /*!< 3D model */
    msoAutoShape         =  1, /*!< AutoShape */
    msoCallout           =  2, /*!< Callout */
    msoCanvas            = 20, /*!< Canvas */
    msoChart             =  3, /*!< Chart */
    msoComment           =  4, /*!< Comment */
    msoContentApp        = 27, /*!< Content Office Add-in */
    msoDiagram           = 21, /*!< Diagram */
    msoEmbeddedOLEObject =  7, /*!< Embedded OLE object */
    msoFormControl       =  8, /*!< Form control */
    msoFreeform          =  5, /*!< Freeform */
    msoGraphic           = 28, /*!< Graphic */
    msoGroup             =  6, /*!< Group */
    msoIgxGraphic        = 24, /*!< SmartArt graphic */
    msoInk               = 22, /*!< Ink */
    msoInkComment        = 23, /*!< Ink comment */
    msoLine              =  9, /*!< Line */
    msoLinked3DModel     = 31, /*!< Linked 3D model */
    msoLinkedGraphic     = 29, /*!< Linked graphic */
    msoLinkedOLEObject   = 10, /*!< Linked OLE object */
    msoLinkedPicture     = 11, /*!< Linked picture */
    msoMedia             = 16, /*!< Media */
    msoOLEControlObject  = 12, /*!< OLE control object */
    msoPicture           = 13, /*!< Picture */
    msoPlaceholder       = 14, /*!< Placeholder */
    msoScriptAnchor      = 18, /*!< Script anchor */
    msoShapeTypeMixed    = -2, /*!< Mixed shape type */
    msoSlicer            = 25, /*!< Slicer */
    msoTable             = 19, /*!< Table */
    msoTextBox           = 17, /*!< Text box */
    msoTextEffect        = 15, /*!< Text effect */
    msoWebVideo          = 26, /*!< Web video */
};

/**
Specifies the priority for a shared workspace task.

[Official VBA documentation for MsoSharedWorkspaceTaskPriority](https://docs.microsoft.com/office/vba/api/office.msosharedworkspacetaskpriority)
*/
enum MsoSharedWorkspaceTaskPriority
{
    msoSharedWorkspaceTaskPriorityHigh   = 1, /*!< High priority */
    msoSharedWorkspaceTaskPriorityLow    = 3, /*!< Low priority */
    msoSharedWorkspaceTaskPriorityNormal = 2, /*!< Normal priority */
};

/**
Specifies the status of a shared workspace task.

[Official VBA documentation for MsoSharedWorkspaceTaskStatus](https://docs.microsoft.com/office/vba/api/office.msosharedworkspacetaskstatus)
*/
enum MsoSharedWorkspaceTaskStatus
{
    msoSharedWorkspaceTaskStatusCompleted  = 3, /*!< Completed */
    msoSharedWorkspaceTaskStatusDeferred   = 4, /*!< Deferred */
    msoSharedWorkspaceTaskStatusInProgress = 2, /*!< In progress */
    msoSharedWorkspaceTaskStatusNotStarted = 1, /*!< Not started */
    msoSharedWorkspaceTaskStatusWaiting    = 5, /*!< Waiting */
};

/**
Specifies properties of the signature subset. These settings act as filters for signature sets.

[Official VBA documentation for MsoSignatureSubset](https://docs.microsoft.com/office/vba/api/office.msosignaturesubset)
*/
enum MsoSignatureSubset
{
    msoSignatureSubsetAll                    = 5, /*!< All non-visible signatures plus all signature lines. */
    msoSignatureSubsetSignatureLines         = 2, /*!< All signature lines. */
    msoSignatureSubsetSignatureLinesSigned   = 3, /*!< Signature lines that have been signed. */
    msoSignatureSubsetSignatureLinesUnsigned = 4, /*!< Signature lines that have not been signed. */
    msoSignatureSubsetSignaturesAllSigs      = 0, /*!< All non-visible signatures plus all signed signature lines. */
    msoSignatureSubsetSignaturesNonVisible   = 1, /*!< All non-visible signatures. */
};

/**
Specifies constants that define the different ways to add a new node to the data model in SmartArt.

[Official VBA documentation for MsoSmartArtNodePosition](https://docs.microsoft.com/office/vba/api/office.msosmartartnodeposition)
*/
enum MsoSmartArtNodePosition
{
    msoSmartArtNodeAbove   = 4, /*!< Insert a node above the SmartArt node in the data model. */
    msoSmartArtNodeAfter   = 2, /*!< Insert a node after the SmartArt node in the data model. Corresponds to clicking **Add Shape After** in the SmartArt ribbon. */
    msoSmartArtNodeBefore  = 3, /*!< Insert a node before the SmartArt node in the data model. */
    msoSmartArtNodeBelow   = 5, /*!< Insert a node below the SmartArt node in the data model. */
    msoSmartArtNodeDefault = 1, /*!< The default. Corresponds to clicking **Add Shape** in the SmartArt ribbon. */
};

/**
Specifies constants that define the types of SmartArt nodes.

[Official VBA documentation for MsoSmartArtNodeType](https://docs.microsoft.com/office/vba/api/office.msosmartartnodetype)
*/
enum MsoSmartArtNodeType
{
    msoSmartArtNodeTypeAssistant = 2, /*!< Defines an assistant node, which can be used within hierarchical diagrams. */
    msoSmartArtNodeTypeDefault   = 1, /*!< The default node type. */
};

/**
Represents the soft edge effect in Office graphics.

[Official VBA documentation for MsoSoftEdgeType](https://docs.microsoft.com/office/vba/api/office.msosoftedgetype)
*/
enum MsoSoftEdgeType
{
    msoSoftEdgeType1    =  1, /*!< Soft Edge Type 1 */
    msoSoftEdgeType2    =  2, /*!< Soft Edge Type 2 */
    msoSoftEdgeType3    =  3, /*!< Soft Edge Type 3 */
    msoSoftEdgeType4    =  4, /*!< Soft Edge Type 4 */
    msoSoftEdgeType5    =  5, /*!< Soft Edge Type 5 */
    msoSoftEdgeType6    =  6, /*!< Soft Edge Type 6 */
    msoSoftEdgeTypeNone =  0, /*!< No soft edge */
    SoftEdgeTypeMixed   = -2, /*!< A mix of soft edge types */
};

/**
Specifies how conflicts should be resolved when synchronizing a shared document. Used with the **ResolveConflict** method of the **Sync** object.

[Official VBA documentation for MsoSyncConflictResolutionType](https://docs.microsoft.com/office/vba/api/office.msosyncconflictresolutiontype)
*/
enum MsoSyncConflictResolutionType
{
    msoSyncConflictClientWins = 0, /*!< Replace the server copy with the local copy. */
    msoSyncConflictMerge      = 2, /*!< Merge changes made to the server copy into the local copy. To resolve the conflict with the merged changes winning, you must save the active document after merging changes, and then call the **ResolveConflict** method again with the **msoSyncConflictClientWins** option. */
    msoSyncConflictServerWins = 1, /*!< Replace the local copy with the server copy. */
};

/**
Specifies the type of tab stop.

[Official VBA documentation for MsoTabStopType](https://docs.microsoft.com/office/vba/api/office.msotabstoptype)
*/
enum MsoTabStopType
{
    msoTabStopCenter  =  2, /*!< Center tab stop */
    msoTabStopDecimal =  4, /*!< Decimal tab stop */
    msoTabStopLeft    =  1, /*!< Left tab stop */
    msoTabStopMixed   = -2, /*!< Return value only; indicates a combination of the other states. */
    msoTabStopRight   =  3, /*!< Right tab stop */
};

/**
Specifies target browser for documents viewed in a web browser.

[Official VBA documentation for MsoTargetBrowser](https://docs.microsoft.com/office/vba/api/office.msotargetbrowser)
*/
enum MsoTargetBrowser
{
    msoTargetBrowserIE4 = 2, /*!< Microsoft Internet Explorer 4.0 or later */
    msoTargetBrowserIE5 = 3, /*!< Microsoft Internet Explorer 5 or later */
    msoTargetBrowserIE6 = 4, /*!< Microsoft Internet Explorer 6 or later */
    msoTargetBrowserV3  = 0, /*!< Microsoft Internet Explorer 3.0, Netscape Navigator 3, or later */
    msoTargetBrowserV4  = 1, /*!< Microsoft Internet Explorer 4.0, Netscape Navigator 4, or later */
};

/**
Specifies the capitalization of the text.

[Official VBA documentation for MsoTextCaps](https://docs.microsoft.com/office/vba/api/office.msotextcaps)
*/
enum MsoTextCaps
{
    msoAllCaps   =  2, /*!< Display the text as all uppercase letters. */
    msoCapsMixed = -2, /*!< Display the text as mixed uppercase and lowercase letters. */
    msoNoCaps    =  0, /*!< Display the text with no uppercase letters. */
    msoSmallCaps =  1, /*!< Display the text with all lowercase letters. */
};

/**
Specifies the capitalization of text.

[Official VBA documentation for MsoTextChangeCase](https://docs.microsoft.com/office/vba/api/office.msotextchangecase)
*/
enum MsoTextChangeCase
{
    msoCaseLower    = 2, /*!< Display the text as lowercase characters. */
    msoCaseSentence = 1, /*!< Display the text as sentence case characters. Sentence case specifies that the first letter of the sentence is capitalized and that all others should be lowercase (with some exceptions such as proper nouns, and acronyms). */
    msoCaseTitle    = 4, /*!< Display the text as title case characters. Title case specifies that the first letter of each word is capitalized and that all others should be lowercase. In some cases short articles, prepositions, and conjunctions are not capitalized. */
    msoCaseToggle   = 5, /*!< Indicates that lowercase text should be converted to uppercase and that uppercase text should be converted to lowercase text. */
    msoCaseUpper    = 3, /*!< Display the text as uppercase characters. */
};

/**
Indicates the type of text wrap.

[Official VBA documentation for MsoTextCharWrap](https://docs.microsoft.com/office/vba/api/office.msotextcharwrap)
*/
enum MsoTextCharWrap
{
    msoCharWrapMixed    = -2, /*!< Specifies a mixed text wrap. */
    msoCustomCharWrap   =  3, /*!< Specifies a custom text wrap scheme. */
    msoNoCharWrap       =  0, /*!< Specifies no text wrapping. */
    msoStandardCharWrap =  1, /*!< Specifies wrapping text around the standard boundary of an object. */
    msoStrictCharWrap   =  2, /*!< Specifies text wrapping that adheres to restrictions imposed by some languages such as Chinese and Japanese alphabets. */
};

/**
Specifies the direction that text runs.

[Official VBA documentation for MsoTextDirection](https://docs.microsoft.com/office/vba/api/office.msotextdirection)
*/
enum MsoTextDirection
{
    msoTextDirectionLeftToRight =  1, /*!< Text runs left to right. */
    msoTextDirectionMixed       = -2, /*!< Return value only; indicates a combination of the other states. */
    msoTextDirectionRightToLeft =  2, /*!< Text runs right to left. */
};

/**
Specifies alignment for WordArt text.

[Official VBA documentation for MsoTextEffectAlignment](https://docs.microsoft.com/office/vba/api/office.msotexteffectalignment)
*/
enum MsoTextEffectAlignment
{
    msoTextEffectAlignmentCentered       =  2, /*!< Centered. */
    msoTextEffectAlignmentLeft           =  1, /*!< Left-aligned. */
    msoTextEffectAlignmentLetterJustify  =  4, /*!< Text is justified. Spacing between letters may be adjusted to justify text. */
    msoTextEffectAlignmentMixed          = -2, /*!< Not used. */
    msoTextEffectAlignmentRight          =  3, /*!< Right- aligned. */
    msoTextEffectAlignmentStretchJustify =  6, /*!< Text is justified. Letters may be stretched to justify text. */
    msoTextEffectAlignmentWordJustify    =  5, /*!< Text is justified. Spacing between words (but not letters) may be adjusted to justify text. */
};

/**
Indicates the text alignment scheme used for an object.

[Official VBA documentation for MsoTextFontAlign](https://docs.microsoft.com/office/vba/api/office.msotextfontalign)
*/
enum MsoTextFontAlign
{
    msoFontAlignAuto     =  0, /*!< Specifies that the text alignment will be determined by the Office application. */
    msoFontAlignBaseline =  3, /*!< Specifies that the font is aligned to the baseline of the object. */
    msoFontAlignBottom   =  4, /*!< Specifies that the font is aligned to the bottom of the object. */
    msoFontAlignCenter   =  2, /*!< Specifies that the font is aligned to the center of the object. */
    msoFontAlignMixed    = -2, /*!< Specifies that there is a mix of text alignments used with the object. */
    msoFontAlignTop      =  1, /*!< Specifies that the font is aligned to the top of the object. */
};

/**
Specifies orientation for text.

[Official VBA documentation for MsoTextOrientation](https://docs.microsoft.com/office/vba/api/office.msotextorientation)
*/
enum MsoTextOrientation
{
    msoTextOrientationDownward                 =  3, /*!< Downward */
    msoTextOrientationHorizontal               =  1, /*!< Horizontal */
    msoTextOrientationHorizontalRotatedFarEast =  6, /*!< Horizontal and rotated as required for Asian language support */
    msoTextOrientationMixed                    = -2, /*!< Not supported */
    msoTextOrientationUpward                   =  2, /*!< Upward */
    msoTextOrientationVertical                 =  5, /*!< Vertical */
    msoTextOrientationVerticalFarEast          =  4, /*!< Vertical as required for Asian language support */
};

/**
Indicates the number of times a character is printed to darken the image.

[Official VBA documentation for MsoTextStrike](https://docs.microsoft.com/office/vba/api/office.msotextstrike)
*/
enum MsoTextStrike
{
    msoDoubleStrike =  2, /*!< Specifies that the character is printed twice. */
    msoNoStrike     =  0, /*!< Specifies that the character is not printed. */
    msoSingleStrike =  1, /*!< Specifies that the character is printed once. */
    msoStrikeMixed  = -2, /*!< Specifies that the text can contain a combination of doublestrike and single strike characters. */
};

/**
Indicates the text alignment against tab stops or line breaks.

[Official VBA documentation for MsoTextTabAlign](https://docs.microsoft.com/office/vba/api/office.msotexttabalign)
*/
enum MsoTextTabAlign
{
    msoTabAlignCenter  =  1, /*!< Specifies that the following text up to next tab or line break is centered on the designated tab stop. */
    msoTabAlignDecimal =  3, /*!< Specifies that the following text is searched for the first occurrence of the character representing the decimal point. The text up to the next tab or line break is then aligned such that the decimal point starts at the designated tab stop. */
    msoTabAlignLeft    =  0, /*!< Specifies that the following text starts immediately after the designated tab stop (the default). */
    msoTabAlignMixed   = -2, /*!< Specifies that mixed text alignment against tab stops is used. */
    msoTabAlignRight   =  2, /*!< Specifies that the following text up to the next tab or line break is rendered flush right to the designated tab stop. */
};

/**
Indicates the type of underline for text.

[Official VBA documentation for MsoTextUnderlineType](https://docs.microsoft.com/office/vba/api/office.msotextunderlinetype)
*/
enum MsoTextUnderlineType
{
    msoNoUnderline                  =  0, /*!< Specifies no underline. */
    msoUnderlineDashHeavyLine       =  8, /*!< Specifies a dash underline. */
    msoUnderlineDashLine            =  7, /*!< Specifies a dash line underline. */
    msoUnderlineDashLongHeavyLine   = 10, /*!< Specifies a long heavy line underline. */
    msoUnderlineDashLongLine        =  9, /*!< Specifies a dashed long line underline. */
    msoUnderlineDotDashHeavyLine    = 12, /*!< Specifies a dot dash heavy line underline. */
    msoUnderlineDotDashLine         = 11, /*!< Specifies a dot dash line underline. */
    msoUnderlineDotDotDashHeavyLine = 14, /*!< Specifies a dot dot dash heavy line underline. */
    msoUnderlineDotDotDashLine      = 13, /*!< Specifies a dot dot dash line underline. */
    msoUnderlineDottedHeavyLine     =  6, /*!< Specifies a dotted heavy line underline. */
    msoUnderlineDottedLine          =  5, /*!< Specifies a dotted line underline. */
    msoUnderlineDoubleLine          =  3, /*!< Specifies a double line underline. */
    msoUnderlineHeavyLine           =  4, /*!< Specifies a heavy line underline. */
    msoUnderlineMixed               = -2, /*!< Specifies a mixed of underline types. */
    msoUnderlineSingleLine          =  2, /*!< Specifies a single line underline. */
    msoUnderlineWavyDoubleLine      = 17, /*!< Specifies a wavy double line underline. */
    msoUnderlineWavyHeavyLine       = 16, /*!< Specifies a wavy heavy line underline. */
    msoUnderlineWavyLine            = 15, /*!< Specifies a wavy line underline. */
    msoUnderlineWords               =  1, /*!< Specifies underlining words. */
};

/**
Specifies the alignment (the origin of the coordinate grid) for the tiling of the texture fill.

[Official VBA documentation for MsoTextureAlignment](https://docs.microsoft.com/office/vba/api/office.msotexturealignment)
*/
enum MsoTextureAlignment
{
    msoTextureAlignmentMixed = -2, /*!< Return value only; indicates a combination of the other states. */
    msoTextureBottom         =  7, /*!< Bottom alignment */
    msoTextureBottomLeft     =  6, /*!< Bottom-left alignment */
    msoTextureBottomRight    =  8, /*!< Bottom-right alignment */
    msoTextureCenter         =  4, /*!< Center alignment */
    msoTextureLeft           =  3, /*!< Left alignment */
    msoTextureRight          =  5, /*!< Right alignment */
    msoTextureTop            =  1, /*!< Top alignment */
    msoTextureTopLeft        =  0, /*!< Top-left alignment */
    msoTextureTopRight       =  2, /*!< Top-right alignment */
};

/**
Specifies the texture type for the selected fill.

[Official VBA documentation for MsoTextureType](https://docs.microsoft.com/office/vba/api/office.msotexturetype)
*/
enum MsoTextureType
{
    msoTexturePreset      =  1, /*!< Preset texture type */
    msoTextureTypeMixed   = -2, /*!< Return value only; indicates a combination of the other states. */
    msoTextureUserDefined =  2, /*!< User-defined texture type */
};

/**
Indicates the Office theme color.

[Official VBA documentation for MsoThemeColorIndex](https://docs.microsoft.com/office/vba/api/office.msothemecolorindex)
*/
enum MsoThemeColorIndex
{
    msoNotThemeColor               =  0, /*!< Specifies no theme color. */
    msoThemeColorAccent1           =  5, /*!< Specifies the Accent 1 theme color. */
    msoThemeColorAccent2           =  6, /*!< Specifies the Accent 2 theme color. */
    msoThemeColorAccent3           =  7, /*!< Specifies the Accent 3 theme color. */
    msoThemeColorAccent4           =  8, /*!< Specifies the Accent 4 theme color. */
    msoThemeColorAccent5           =  9, /*!< Specifies the Accent 5 theme color. */
    msoThemeColorAccent6           = 10, /*!< Specifies the Accent 6 theme color. */
    msoThemeColorBackground1       = 14, /*!< Specifies the Background 1 theme color. */
    msoThemeColorBackground2       = 16, /*!< Specifies the Background 2 theme color. */
    msoThemeColorDark1             =  1, /*!< Specifies the Dark 1 theme color. */
    msoThemeColorDark2             =  3, /*!< Specifies the Dark 2 theme color. */
    msoThemeColorFollowedHyperlink = 12, /*!< Specifies the theme color for a clicked hyperlink. */
    msoThemeColorHyperlink         = 11, /*!< Specifies the theme color for a hyperlink. */
    msoThemeColorLight1            =  2, /*!< Specifies the Light 1 theme color. */
    msoThemeColorLight2            =  4, /*!< Specifies the Light 2 theme color. */
    msoThemeColorMixed             = -2, /*!< Specifies a mixed color theme. */
    msoThemeColorText1             = 13, /*!< Specifies the Text 1 theme color. */
    msoThemeColorText2             = 15, /*!< Specifies the Text 2 theme color. */
};

/**
Indicates the color scheme for an Office theme.

[Official VBA documentation for MsoThemeColorSchemeIndex](https://docs.microsoft.com/office/vba/api/office.msothemecolorschemeindex)
*/
enum MsoThemeColorSchemeIndex
{
    msoThemeAccent1           =  5, /*!< Specifies color scheme Accent 1. */
    msoThemeAccent2           =  6, /*!< Specifies color scheme Accent 2. */
    msoThemeAccent3           =  7, /*!< Specifies color scheme Accent 3. */
    msoThemeAccent4           =  8, /*!< Specifies color scheme Accent 4. */
    msoThemeAccent5           =  9, /*!< Specifies color scheme Accent 5. */
    msoThemeAccent6           = 10, /*!< Specifies color scheme Accent 6. */
    msoThemeDark1             =  1, /*!< Specifies color scheme Dark 1. */
    msoThemeDark2             =  3, /*!< Specifies color scheme Dark 2. */
    msoThemeFollowedHyperlink = 12, /*!< Specifies a color scheme for a clicked hyperlink. */
    msoThemeHyperlink         = 11, /*!< Specifies a color scheme for a hyperlink. */
    msoThemeLight1            =  2, /*!< Specifies color scheme Light 1. */
    msoThemeLight2            =  4, /*!< Specifies color scheme Light 2. */
};

/**
Specifies a tri-state value.

[Official VBA documentation for MsoTriState](https://docs.microsoft.com/office/vba/api/office.msotristate)
*/
enum MsoTriState
{
    msoCTrue          =  1, /*!< Not supported */
    msoFalse          =  0, /*!< False */
    msoTriStateMixed  = -2, /*!< Not supported */
    msoTriStateToggle = -3, /*!< Not supported */
    msoTrue           = -1, /*!< True */
};

/**
Specifies the vertical alignment of text in a text frame. Used with the **VerticalAnchor** property of the **TextFrame** object.

[Official VBA documentation for MsoVerticalAnchor](https://docs.microsoft.com/office/vba/api/office.msoverticalanchor)
*/
enum MsoVerticalAnchor
{
    msoAnchorBottom         =  4, /*!< Aligns text to bottom of text frame. */
    msoAnchorBottomBaseLine =  5, /*!< Anchors bottom of text string to current position, regardless of text resizing. When you resize text without baseline anchoring, text centers itself on previous position. */
    msoAnchorMiddle         =  3, /*!< Centers text vertically. */
    msoAnchorTop            =  1, /*!< Aligns text to top of text frame. */
    msoAnchorTopBaseline    =  2, /*!< Anchors bottom of text string to current position, regardless of text resizing. When you resize text without baseline anchoring, text centers itself on previous position. */
    msoVerticalAnchorMixed  = -2, /*!< Return value only; indicates a combination of the other states. */
};

/**
Indicates various image warping formats.

[Official VBA documentation for MsoWarpFormat](https://docs.microsoft.com/office/vba/api/office.msowarpformat)
*/
enum MsoWarpFormat
{
    msoWarpFormat1     =  0, /*!< Specifies Warp Format 1. */
    msoWarpFormat10    =  9, /*!< Specifies Warp Format 10. */
    msoWarpFormat11    = 10, /*!< Specifies Warp Format 11. */
    msoWarpFormat12    = 11, /*!< Specifies Warp Format 12. */
    msoWarpFormat13    = 12, /*!< Specifies Warp Format 13. */
    msoWarpFormat14    = 13, /*!< Specifies Warp Format 14. */
    msoWarpFormat15    = 14, /*!< Specifies Warp Format 15. */
    msoWarpFormat16    = 15, /*!< Specifies Warp Format 16. */
    msoWarpFormat17    = 16, /*!< Specifies Warp Format 17. */
    msoWarpFormat18    = 17, /*!< Specifies Warp Format 18. */
    msoWarpFormat19    = 18, /*!< Specifies Warp Format 19. */
    msoWarpFormat2     =  1, /*!< Specifies Warp Format 2. */
    msoWarpFormat20    = 19, /*!< Specifies Warp Format 20. */
    msoWarpFormat21    = 20, /*!< Specifies Warp Format 21. */
    msoWarpFormat22    = 21, /*!< Specifies Warp Format 22. */
    msoWarpFormat23    = 22, /*!< Specifies Warp Format 23. */
    msoWarpFormat24    = 23, /*!< Specifies Warp Format 24. */
    msoWarpFormat25    = 24, /*!< Specifies Warp Format 25. */
    msoWarpFormat26    = 25, /*!< Specifies Warp Format 26. */
    msoWarpFormat27    = 26, /*!< Specifies Warp Format 27. */
    msoWarpFormat28    = 27, /*!< Specifies Warp Format 28. */
    msoWarpFormat29    = 28, /*!< Specifies Warp Format 29. */
    msoWarpFormat3     =  2, /*!< Specifies Warp Format 3. */
    msoWarpFormat30    = 29, /*!< Specifies Warp Format 30. */
    msoWarpFormat31    = 30, /*!< Specifies Warp Format 31. */
    msoWarpFormat32    = 31, /*!< Specifies Warp Format 32. */
    msoWarpFormat33    = 32, /*!< Specifies Warp Format 33. */
    msoWarpFormat34    = 33, /*!< Specifies Warp Format 34. */
    msoWarpFormat35    = 34, /*!< Specifies Warp Format 35. */
    msoWarpFormat36    = 35, /*!< Specifies Warp Format 36. */
    msoWarpFormat37    = 36, /*!< Specifies Warp Format 37. */
    msoWarpFormat4     =  3, /*!< Specifies Warp Format 4. */
    msoWarpFormat5     =  4, /*!< Specifies Warp Format 5. */
    msoWarpFormat6     =  5, /*!< Specifies Warp Format 6. */
    msoWarpFormat7     =  6, /*!< Specifies Warp Format 7. */
    msoWarpFormat8     =  7, /*!< Specifies Warp Format 8. */
    msoWarpFormat9     =  8, /*!< Specifies Warp Format 9. */
    msoWarpFormatMixed = -2, /*!< Specifies a mix of warp formats. */
};

/**
Specifies the context under which a wizard's callback procedure is called. Used as an argument in a callback procedure designed for use with a custom wizard.

[Official VBA documentation for MsoWizardMsgType](https://docs.microsoft.com/office/vba/api/office.msowizardmsgtype)
*/
enum MsoWizardMsgType
{
    msoWizardMsgLocalStateOff = 2, /*!< User clicked the right button in the decision or branch balloon. */
    msoWizardMsgLocalStateOn  = 1, /*!< Not supported. */
    msoWizardMsgResuming      = 5, /*!< Passed to the **ActivateWizard** method if **msoWizardActResume** is specified for the Act argument. */
    msoWizardMsgShowHelp      = 3, /*!< User clicked the left button in the decision or branch balloon. */
    msoWizardMsgSuspending    = 4, /*!< Passed to the **ActivateWizard** method if **msoWizardActSuspend** is specified for the Act argument. */
};

/**
Specifies where in the z-order a shape should be moved relative to other shapes.

[Official VBA documentation for MsoZOrderCmd](https://docs.microsoft.com/office/vba/api/office.msozordercmd)
*/
enum MsoZOrderCmd
{
    msoBringForward       = 2, /*!< Bring shape forward. */
    msoBringInFrontOfText = 4, /*!< Bring shape in front of text. Used only in Microsoft Word. */
    msoBringToFront       = 0, /*!< Bring shape to the front. */
    msoSendBackward       = 3, /*!< Send shape backward. */
    msoSendBehindText     = 5, /*!< Send shape behind text. Used only in Microsoft Word. */
    msoSendToBack         = 1, /*!< Send shape to the back. */
};

/**
Specifies constants that define the styles of the slabs on the **File** tab.

[Official VBA documentation for OutSpaceSlabStyle](https://docs.microsoft.com/office/vba/api/office.outspaceslabstyle)
*/
enum OutSpaceSlabStyle
{
    OutSpaceSlabStyleError   = 2, /*!< Error style */
    OutSpaceSlabStyleNormal  = 0, /*!< Normal style */
    OutSpaceSlabStyleWarning = 1, /*!< Warning style */
};

/**
Specifies constants that define the size of the controls on the ribbon.

[Official VBA documentation for RibbonControlSize](https://docs.microsoft.com/office/vba/api/office.ribboncontrolsize)
*/
enum RibbonControlSize
{
    RibbonControlSizeLarge   = 1, /*!< Large controls */
    RibbonControlSizeRegular = 0, /*!< Small controls */
};

/**
Indicates additional information about a signature.

[Official VBA documentation for SignatureDetail](https://docs.microsoft.com/office/vba/api/office.signaturedetail)
*/
enum SignatureDetail
{
    sigdetApplicationName       =  1, /*!< Specifies the application name. */
    sigdetApplicationVersion    =  2, /*!< Specifies the application version. */
    sigdetColorDepth            =  8, /*!< Specifies the color depth. */
    sigdetDelSuggSigner         = 16, /*!< Specifies the suggested signer delegate. */
    sigdetDelSuggSignerEmail    = 20, /*!< Specifies the suggested signer's delegate's email. */
    sigdetDelSuggSignerEmailSet = 21, /*!< Indicates whether an email for a suggested signer delegate has been specified. */
    sigdetDelSuggSignerLine2    = 18, /*!< Specifies the suggested signer's delegate's signature line. */
    sigdetDelSuggSignerLine2Set = 19, /*!< Specifies the set of suggested signer's delegate's signature lines. */
    sigdetDelSuggSignerSet      = 17, /*!< Specifies the set of suggested signer's delegates. */
    sigdetDocPreviewImg         = 10, /*!< Specifies the document preview image. */
    sigdetHashAlgorithm         = 14, /*!< Specifies the hash algorithm. */
    sigdetHorizResolution       =  6, /*!< Specifies the horizontal resolution. */
    sigdetIPCurrentView         = 12, /*!< Specifies the IP current view. */
    sigdetIPFormHash            = 11, /*!< Specifies the IP form hash. */
    sigdetLocalSigningTime      =  0, /*!< Specifies the local signing time. */
    sigdetNumberOfMonitors      =  5, /*!< Specifies the number of monitors. */
    sigdetOfficeVersion         =  3, /*!< Specifies the Office version. */
    sigdetShouldShowViewWarning = 15, /*!< Specifies the Should Show View Warning setting. */
    sigdetSignatureType         = 13, /*!< Specifies the signature type. */
    sigdetSignedData            =  9, /*!< Specifies the signed data. */
    sigdetVertResolution        =  7, /*!< Specifies the vertical resolution. */
    sigdetWindowsVersion        =  4, /*!< Specifies the Windows version. */
};

/**
Indicates the signature line image.

[Official VBA documentation for SignatureLineImage](https://docs.microsoft.com/office/vba/api/office.signaturelineimage)
*/
enum SignatureLineImage
{
    siglnimgSignedInvalid    = 3, /*!< The SignedInvalid image */
    siglnimgSignedValid      = 2, /*!< The SignedValid image */
    siglnimgSoftwareRequired = 0, /*!< The SoftwareRequired image */
    siglnimgUnsigned         = 1, /*!< The Unsigned image */
};

/**
Specifies properties of a signature provider.

[Official VBA documentation for SignatureProviderDetail](https://docs.microsoft.com/office/vba/api/office.signatureproviderdetail)
*/
enum SignatureProviderDetail
{
    sigprovdetHashAlgorithm = 1, /*!< Hash algorithm used to hash the data in the file. */
    sigprovdetUIOnly        = 2, /*!< Indicates that the signature provider only uses a custom user interface. */
    sigprovdetUrl           = 0, /*!< The URL of the signature provider. */
};

/**
Specifies properties of a signature.

[Official VBA documentation for SignatureType](https://docs.microsoft.com/office/vba/api/office.signaturetype)
*/
enum SignatureType
{
    sigtypeMax           = 3, /*!< Specifies the maximum number of the signature types available in the current version of Office. */
    sigtypeNonVisible    = 1, /*!< A signature that is not visible in the content of the document. */
    sigtypeSignatureLine = 2, /*!< A signature that is visible in the content of the document. */
    sigtypeUnknown       = 0, /*!< A signature not generated by Office. */
};

} // namespace wxAutoExcel


#endif //_WXAUTOEXCEL_ENUMS_H
