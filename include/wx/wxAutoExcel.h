/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// License:     MIT license
/////////////////////////////////////////////////////////////////////////////


#ifndef _WXAUTOEXCEL_H
#define _WXAUTOEXCEL_H

#include <wx/variant.h>
#include <wx/colour.h>

#include "wx/wxAutoExcel_defs.h"

#include "wx/wxAutoExcel_tribool.h"

#include "wx/wxAutoExcelApplication.h"
#include "wx/wxAutoExcelWorkbooks.h"
#include "wx/wxAutoExcelWorkbook.h"
#include "wx/wxAutoExcelSheets.h"
#include "wx/wxAutoExcelSheet.h"
#include "wx/wxAutoExcelWorksheets.h"
#include "wx/wxAutoExcelWorksheet.h"
#include "wx/wxAutoExcelWindows.h"
#include "wx/wxAutoExcelPanes.h"
#include "wx/wxAutoExcelTab.h"
#include "wx/wxAutoExcelRange.h"
#include "wx/wxAutoExcelHeadersFooters.h"
#include "wx/wxAutoExcelPages.h"
#include "wx/wxAutoExcelPageSetup.h"
#include "wx/wxAutoExcelComments.h"
#include "wx/wxAutoExcelAreas.h"
#include "wx/wxAutoExcelFont.h"
#include "wx/wxAutoExcelCharacters.h"
#include "wx/wxAutoExcelBorders.h"
#include "wx/wxAutoExcelGradient.h"
#include "wx/wxAutoExcelInterior.h"
#include "wx/wxAutoExcelStyles.h"

#include "wx/wxAutoExcelAutoFilter.h"
#include "wx/wxAutoExcelFilters.h"
#include "wx/wxAutoExcelSort.h"
#include "wx/wxAutoExcelSortFields.h"

#include "wx/wxAutoExcelHyperlinks.h"

#include "wx/wxAutoExcelSheetViews.h"
#include "wx/wxAutoExcelWorksheetView.h"

#include "wx/wxAutoExcelValidation.h"

#include "wx/wxAutoExcelErrors.h"
#include "wx/wxAutoExcelErrorCheckingOptions.h"

#include "wx/wxAutoExcelAboveAverage.h"
#include "wx/wxAutoExcelDatabar.h"
#include "wx/wxAutoExcelConditionValue.h"
#include "wx/wxAutoExcelColorScale.h"
#include "wx/wxAutoExcelColorScaleCriteria.h"
#include "wx/wxAutoExcelDataBarBorder.h"
#include "wx/wxAutoExcelFormatColor.h"
#include "wx/wxAutoExcelFormatConditions.h"
#include "wx/wxAutoExcelTop10.h"
#include "wx/wxAutoExcelUniqueValues.h"
#include "wx/wxAutoExcelIconCriteria.h"
#include "wx/wxAutoExcelIconSets.h"
#include "wx/wxAutoExcelIconSetCondition.h"
#include "wx/wxAutoExcelNegativeBarFormat.h"

#include "wx/wxAutoExcelShape.h"
#include "wx/wxAutoExcelShapes.h"
#include "wx/wxAutoExcelShapeRange.h"
#include "wx/wxAutoExcelGroupShapes.h"
#include "wx/wxAutoExcelShapeNodes.h"
#include "wx/wxAutoExcelFreeformBuilder.h"
#include "wx/wxAutoExcelAdjustments.h"

#include "wx/wxAutoExcelCalloutFormat.h"
#include "wx/wxAutoExcelConnectorFormat.h"
#include "wx/wxAutoExcelControlFormat.h"
#include "wx/wxAutoExcelFillFormat.h"
#include "wx/wxAutoExcelGlowFormat.h"
#include "wx/wxAutoExcelLineFormat.h"
#include "wx/wxAutoExcelLinkFormat.h"
#include "wx/wxAutoExcelOLEFormat.h"
#include "wx/wxAutoExcelPictureFormat.h"
#include "wx/wxAutoExcelReflectionFormat.h"
#include "wx/wxAutoExcelShadowFormat.h"
#include "wx/wxAutoExcelSoftEdgeFormat.h"
#include "wx/wxAutoExcelColorFormat.h"
#include "wx/wxAutoExcelTextEffectFormat.h"
#include "wx/wxAutoExcelThreeDFormat.h"

#include "wx/wxAutoExcelModel3DFormat.h"

#include "wx/wxAutoExcelTextColumn2.h"
#include "wx/wxAutoExcelTextFrame.h"
#include "wx/wxAutoExcelTextFrame2.h"
#include "wx/wxAutoExcelTextRange2.h"
#include "wx/wxAutoExcelParagraphFormat2.h"
#include "wx/wxAutoExcelBulletFormat2.h"
#include "wx/wxAutoExcelFont2.h"
#include "wx/wxAutoExcelTabStops2.h"

#include "wx/wxAutoExcelIcon.h"
#include "wx/wxAutoExcelGraphic.h"

#include "wx/wxAutoExcelAxisTitle.h"
#include "wx/wxAutoExcelAxes.h"
#include "wx/wxAutoExcelChart.h"
#include "wx/wxAutoExcelCharts.h"
#include "wx/wxAutoExcelChartArea.h"
#include "wx/wxAutoExcelChartCategory.h"
#include "wx/wxAutoExcelChartColorFormat.h"
#include "wx/wxAutoExcelChartFillFormat.h"
#include "wx/wxAutoExcelChartFormat.h"
#include "wx/wxAutoExcelChartGroups.h"
#include "wx/wxAutoExcelChartObjects.h"
#include "wx/wxAutoExcelChartView.h"
#include "wx/wxAutoExcelChartTitle.h"
#include "wx/wxAutoExcelDataLabels.h"
#include "wx/wxAutoExcelDataTable.h"
#include "wx/wxAutoExcelDisplayUnitLabel.h"
#include "wx/wxAutoExcelDownBars.h"
#include "wx/wxAutoExcelDropLines.h"
#include "wx/wxAutoExcelErrorBars.h"
#include "wx/wxAutoExcelFloor.h"
#include "wx/wxAutoExcelGridlines.h"
#include "wx/wxAutoExcelHiLoLines.h"
#include "wx/wxAutoExcelLeaderLines.h"
#include "wx/wxAutoExcelLegend.h"
#include "wx/wxAutoExcelLegendEntries.h"
#include "wx/wxAutoExcelLegendKey.h"
#include "wx/wxAutoExcelPlotArea.h"
#include "wx/wxAutoExcelPoints.h"
#include "wx/wxAutoExcelSeries.h"
#include "wx/wxAutoExcelSeriesCollection.h"
#include "wx/wxAutoExcelSeriesLines.h"
#include "wx/wxAutoExcelTickLabels.h"
#include "wx/wxAutoExcelTrendLines.h"
#include "wx/wxAutoExcelUpBars.h"
#include "wx/wxAutoExcelWalls.h"

#include "wx/wxAutoExcelSparkAxes.h"
#include "wx/wxAutoExcelSparkColor.h"
#include "wx/wxAutoExcelSparkline.h"
#include "wx/wxAutoExcelSparklineGroups.h"
#include "wx/wxAutoExcelSparkPoints.h"

#include "wx/wxAutoExcelNames.h"

#include "wx/wxAutoExcelPageBreaks.h"

#include "wx/wxAutoExcelOLEObjects.h"

#include "wx/wxAutoExcelDocumentProperties.h"
#include "wx/wxAutoExcelRecentFiles.h"
#include "wx/wxAutoExcelLanguageSettings.h"
#include "wx/wxAutoExcelAddIns.h"

#include "wx/wxAutoExcelListRow.h"
#include "wx/wxAutoExcelListColumn.h"
#include "wx/wxAutoExcelListDataFormat.h"
#include "wx/wxAutoExcelListObject.h"
#include "wx/wxAutoExcelTableObject.h"
#include "wx/wxAutoExcelTableStyleElement.h"
#include "wx/wxAutoExcelTableStyle.h"

#include "wx/wxAutoExcelAllowEditRanges.h"
#include "wx/wxAutoExcelAuthor.h"
#include "wx/wxAutoExcelCommentsThreaded.h"
#include "wx/wxAutoExcelCustomProperties.h"
#include "wx/wxAutoExcelDisplayFormat.h"
#include "wx/wxAutoExcelFileExportConverters.h"
#include "wx/wxAutoExcelMultiThreadedCalculation.h"
#include "wx/wxAutoExcelOutline.h"
#include "wx/wxAutoExcelProtection.h"
#include "wx/wxAutoExcelUserAccess.h"

#include "wx/wxAutoExcel_version.h"

#endif //_WXAUTOEXCEL_H
