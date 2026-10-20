#!/usr/bin/env python3
"""
Generates tests/getter_lists.h, used by tests/wxExcel_getters.cpp.

For each wxAutoExcel class in CLASSES, the header defines a macro listing
calls of all public getters (GetXxx() methods) which can be called without
arguments, including the inherited ones, and a macro listing the properties
with a setter taking the type the getter returns (WXAUTOEXCEL_SETTERS_...):

    #define WXAUTOEXCEL_GETTERS_wxExcelLegend(CHECK, object) \
        CHECK(object, GetAutoScaleFont) \
        ...

Getters listed in EXCLUDED and setters listed in SETTERS_EXCLUDED are written
as comments with the reason instead.

Run it from any folder after adding getters or classes:
    python tests/tools/generate_getter_lists.py
"""

import os
import re

ROOT = os.path.normpath(os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", ".."))
INCLUDE_DIR = os.path.join(ROOT, "include", "wx")
OUTPUT = os.path.join(ROOT, "tests", "getter_lists.h")

# The classes to generate the getter lists for.
CLASSES = [
    # workbooks and application
    "wxExcelApplication",
    "wxExcelWorkbooks",
    "wxExcelWorkbook",
    "wxExcelDocumentProperties",
    "wxExcelDocumentProperty",
    "wxExcelSheets",
    "wxExcelWorksheet",
    "wxExcelTab",
    "wxExcelOutline",
    "wxExcelNames",
    "wxExcelName",
    "wxExcelStyles",
    "wxExcelStyle",
    # charts
    "wxExcelChartObjects",
    "wxExcelChartObject",
    "wxExcelChart",
    "wxExcelChartArea",
    "wxExcelChartTitle",
    "wxExcelChartGroups",
    "wxExcelChartGroup",
    "wxExcelPlotArea",
    "wxExcelLegend",
    "wxExcelAxes",
    "wxExcelAxis",
    "wxExcelAxisTitle",
    "wxExcelSeriesCollection",
    "wxExcelSeries",
    "wxExcelPoints",
    "wxExcelPoint",
    "wxExcelDataLabels",
    "wxExcelDataLabel",
    # shapes
    "wxExcelShapes",
    "wxExcelShape",
    "wxExcelShapeRange",
    "wxExcelGroupShapes",
    "wxExcelTextFrame2",
    "wxExcelTextRange2",
    "wxExcelFont2",
    "wxExcelFillFormat",
    "wxExcelLineFormat",
    "wxExcelColorFormat",
    "wxExcelThreeDFormat",
    # tables and sorting
    "wxExcelListObjects",
    "wxExcelListObject",
    "wxExcelListColumns",
    "wxExcelListColumn",
    "wxExcelListRows",
    "wxExcelListRow",
    "wxExcelTableStyle",
    "wxExcelSort",
    "wxExcelSortFields",
    "wxExcelSortField",
    "wxExcelAutoFilter",
    "wxExcelFilters",
    "wxExcelFilter",
    # conditional formats
    "wxExcelFormatConditions",
    "wxExcelFormatCondition",
    "wxExcelAboveAverage",
    "wxExcelTop10",
    "wxExcelUniqueValues",
    "wxExcelDatabar",
    "wxExcelConditionValue",
    "wxExcelColorScale",
    "wxExcelColorScaleCriteria",
    "wxExcelColorScaleCriterion",
    "wxExcelIconSetCondition",
    "wxExcelIconCriteria",
    "wxExcelIconCriterion",
    # ranges, also with cells having different values
    "wxExcelRange",
    "wxExcelFont",
    "wxExcelInterior",
    "wxExcelBorders",
    "wxExcelBorder",
    # other range related objects
    "wxExcelValidation",
    "wxExcelComments",
    "wxExcelComment",
    "wxExcelHyperlinks",
    "wxExcelHyperlink",
    "wxExcelCharacters",
    "wxExcelAreas",
    # page setup
    "wxExcelPageSetup",
    "wxExcelPages",
    "wxExcelPage",
    "wxExcelHeaderFooter",
    "wxExcelPageBreaks",
    "wxExcelPageBreak",
]

# Getters which cannot be checked in the state the tests create, as
# {class: {getter: reason}}. Excel fails when getting these properties in
# that state, which was verified to happen also in VBA.
SCOPE_PIVOT = "only for conditional formats of PivotTables"
FORMULA_ONLY = "fails unless the condition is based on a formula"
CHART_SHEET = "only for chart sheets, the test uses an embedded chart"
HISTOGRAM = "only for histogram charts"
RADAR = "only for radar charts"
CATEGORY_AXIS = "only for category axes, the test uses the value axis"
DATE_AXIS = "only for date axes"
PICTURE_FILL = "only for picture fills"
MARKERS = "only for chart types with markers"
TEXTURE = "only for texture fills"
ONE_SHAPE = "only for a range of one shape, the test uses two"

EXCLUDED = {
    "wxExcelApplication": {
        "GetHinstance": "fails in 64-bit Excel, where GetHinstancePtr() must be used",
        "GetOpenFilename": "a method showing a file dialog",
        "GetSaveAsFilename": "a method showing a file dialog",
    },
    "wxExcelWorkbook": {
        "GetAutoUpdateSaveChanges": "only for shared workbooks",
    },
    "wxExcelDocumentProperty": {
        "GetLinkSource": "only for properties linked to content",
    },
    "wxExcelName": {
        "GetCategory": "only for names of macros and functions",
        "GetCategoryLocal": "only for names of macros and functions",
        "GetShortcutKey": "only for names of macros",
    },
    "wxExcelStyle": {
        "GetMergeCells": "Excel fails getting it for a style",
    },
    "wxExcelFormatCondition": {
        "GetFormula2": "only for the between and not between operators",
        "GetScopeType": SCOPE_PIVOT,
        "GetText": "only for text conditions",
        "GetTextOperator": "only for text conditions",
    },
    "wxExcelAboveAverage": {"GetScopeType": SCOPE_PIVOT},
    "wxExcelTop10": {"GetScopeType": SCOPE_PIVOT},
    "wxExcelUniqueValues": {"GetScopeType": SCOPE_PIVOT},
    "wxExcelDatabar": {"GetFormula": FORMULA_ONLY, "GetScopeType": SCOPE_PIVOT},
    "wxExcelColorScale": {"GetFormula": FORMULA_ONLY, "GetScopeType": SCOPE_PIVOT},
    "wxExcelIconSetCondition": {"GetFormula": FORMULA_ONLY, "GetScopeType": SCOPE_PIVOT},
    "wxExcelShape": {
        "GetChart": "only for chart shapes",
        "GetFormControlType": "only for form controls",
        "GetGroupItems": "only for groups",
        "GetHyperlink": "fails when the shape has no hyperlink",
        "GetParentGroup": "only for shapes in a group",
    },
    "wxExcelFillFormat": {
        "GetGradientDegree": "only for gradient fills",
        "GetTextureAlignment": TEXTURE,
        "GetTextureHorizontalScale": TEXTURE,
        "GetTextureName": TEXTURE,
        "GetTextureOffsetX": TEXTURE,
        "GetTextureOffsetY": TEXTURE,
        "GetTextureTile": TEXTURE,
        "GetTextureVerticalScale": TEXTURE,
    },
    "wxExcelShapeRange": {
        "GetChart": "only for a range of one chart shape",
        "GetGroupItems": "only for a range of one group",
        "GetID": ONE_SHAPE,
        "GetName": ONE_SHAPE,
        "GetParentGroup": "only for a range of one shape in a group",
        "GetVertices": "only for a range of one freeform",
        "GetZOrderPosition": ONE_SHAPE,
    },
    "wxExcelListObject": {
        "GetSharePointURL": "only for lists linked to SharePoint",
        "GetTableObject": "only for tables created by a query",
    },
    "wxExcelSortField": {
        "GetCustomOrder": "Excel returns no usable value unless a custom order is set",
        "GetSortOnValue": "only when sorting by a color or an icon",
    },
    "wxExcelFilter": {
        "GetCriteria2": "only for filters with two criteria",
    },
    "wxExcelChart": {
        "GetCodeName": CHART_SHEET,
        "GetDataTable": "only for charts with a data table",
        "GetIndex": CHART_SHEET,
        "GetNext": CHART_SHEET,
        "GetPrevious": CHART_SHEET,
        "GetProtectContents": CHART_SHEET,
        "GetProtectDrawingObjects": CHART_SHEET,
        "GetProtectionMode": CHART_SHEET,
        "GetTab": CHART_SHEET,
        "GetVisible": CHART_SHEET,
    },
    "wxExcelChartGroup": {
        "GetBinsCountValue": HISTOGRAM,
        "GetBinsOverflowEnabled": HISTOGRAM,
        "GetBinsOverflowValue": HISTOGRAM,
        "GetBinsType": HISTOGRAM,
        "GetBinsUnderflowEnabled": HISTOGRAM,
        "GetBinsUnderflowValue": HISTOGRAM,
        "GetBinWidthValue": HISTOGRAM,
        "GetHasRadarAxisLabels": RADAR,
        "GetRadarAxisLabels": RADAR,
    },
    "wxExcelAxis": {
        "GetAxisBetweenCategories": CATEGORY_AXIS,
        "GetBaseUnit": DATE_AXIS,
        "GetBaseUnitIsAuto": DATE_AXIS,
        "GetCategoryNames": CATEGORY_AXIS,
        "GetCategoryType": CATEGORY_AXIS,
        "GetDisplayUnitLabel": "only when a display unit is set",
        "GetLogBase": "only for logarithmic scales",
        "GetMajorUnitScale": DATE_AXIS,
        "GetMinorUnitScale": DATE_AXIS,
        "GetTickLabelSpacing": CATEGORY_AXIS,
        "GetTickLabelSpacingIsAuto": CATEGORY_AXIS,
        "GetTickMarkSpacing": CATEGORY_AXIS,
    },
    "wxExcelSeries": {
        "GetApplyPictToEnd": PICTURE_FILL,
        "GetApplyPictToFront": PICTURE_FILL,
        "GetApplyPictToSides": PICTURE_FILL,
        "GetBubbleSizes": "only for bubble charts",
        "GetErrorBars": "only for series with error bars",
        "GetGeoProjectionType": "only for map charts",
        "GetLeaderLines": "only for chart types with leader lines",
        "GetMarkerBackgroundColor": MARKERS,
        "GetMarkerBackgroundColorIndex": MARKERS,
        "GetMarkerForegroundColor": MARKERS,
        "GetMarkerForegroundColorIndex": MARKERS,
        "GetParentDataLabelOption": "only for treemap and sunburst charts",
        "GetQuartileCalculationInclusiveMedian": "only for box and whisker charts",
        "GetRegionLabelOptions": "only for map charts",
        "GetSmooth": "only for line and scatter charts",
    },
    "wxExcelPoint": {
        "GetApplyPictToEnd": PICTURE_FILL,
        "GetApplyPictToFront": PICTURE_FILL,
        "GetApplyPictToSides": PICTURE_FILL,
        "GetIsTotal": "only for waterfall charts",
        "GetSecondaryPlot": "only for pie of pie and bar of pie charts",
    },
    "wxExcelHyperlink": {
        "GetEmailSubject": "only for mailto hyperlinks",
        "GetShape": "only for hyperlinks of shapes",
    },
    "wxExcelRange": {
        "GetCurrentArray": "only for cells of an array formula",
        "GetHidden": "only for entire rows or columns",
        "GetLocationInTable": "only for cells of a PivotTable",
        "GetName": "only for named ranges",
        "GetOutlineLevel": "only for entire rows or columns",
        "GetPageBreak": "only for entire rows or columns",
        "GetPrevious": "fails for the ranges used in the test, also in VBA",
        "GetShowDetail": "only for outlines and PivotTables",
        "GetSummary": "only for outline summary rows or columns",
    },
    "wxExcelInterior": {
        "GetInvertIfNegative": "only for chart elements",
    },
    "wxExcelBorder": {
        "GetThemeColor": "fails after the setters of Borders, checked before, make the color explicit",
    },
}


# Setters which cannot be checked by setting the value returned by the
# getter, as {class: {setter: reason}}. Excel fails setting these properties
# also in VBA, when they are set in the same state and order as in the tests.
SET_FAILS = "Excel fails setting the value it returns"
SHARED_WORKBOOK = "only for shared workbooks"
THREE_D_CHART = "only for 3-D charts"
CUSTOM_POSITION = "fails after setting Left makes the position custom"
EXPLICIT_COLOR = "fails after setting Color makes the color explicit"
BUILTIN_STYLE = "read-only for built-in styles"
ARROWHEADS = "only for lines, the test uses a rectangle"

SETTERS_EXCLUDED = {
    "wxExcelApplication": {
        "SetConstrainNumeric": SET_FAILS,
        "SetHeight": "fails while the Excel window is hidden",
        "SetLeft": "fails while the Excel window is hidden",
        "SetTop": "fails while the Excel window is hidden",
        "SetWidth": "fails while the Excel window is hidden",
    },
    "wxExcelWorkbook": {
        "SetAutoUpdateFrequency": SHARED_WORKBOOK,
        "SetChangeHistoryDuration": SHARED_WORKBOOK,
        "SetHighlightChangesOnScreen": SHARED_WORKBOOK,
        "SetListChangesOnNewSheet": SHARED_WORKBOOK,
        "SetShowConflictHistory": SHARED_WORKBOOK,
    },
    "wxExcelDocumentProperty": {
        "SetLinkToContent": "read-only for built-in properties",
        "SetType": "read-only for built-in properties",
    },
    "wxExcelTab": {"SetThemeColor": "fails for a tab without a theme color"},
    "wxExcelChart": {
        "SetElevation": THREE_D_CHART,
        "SetHeightPercent": THREE_D_CHART,
        "SetName": "fails for embedded charts",
        "SetPerspective": THREE_D_CHART,
        "SetRightAngleAxes": THREE_D_CHART,
    },
    "wxExcelLegend": {"SetPosition": CUSTOM_POSITION},
    "wxExcelDataLabel": {"SetPosition": CUSTOM_POSITION},
    "wxExcelSeries": {
        "SetExplosion": "only for pie charts",
        "SetPictureType": PICTURE_FILL,
        "SetPictureUnit": PICTURE_FILL,
        "SetPictureUnit2": PICTURE_FILL,
    },
    "wxExcelPoint": {
        "SetPictureType": PICTURE_FILL,
        "SetPictureUnit": PICTURE_FILL,
        "SetPictureUnit2": PICTURE_FILL,
    },
    "wxExcelTableStyle": {
        "SetShowAsAvailablePivotTableStyle": BUILTIN_STYLE,
        "SetShowAsAvailableSlicerStyle": BUILTIN_STYLE,
        "SetShowAsAvailableTableStyle": BUILTIN_STYLE,
        "SetShowAsAvailableTimelineStyle": BUILTIN_STYLE,
    },
    "wxExcelInterior": {
        "SetPatternThemeColor": EXPLICIT_COLOR,
        "SetThemeColor": EXPLICIT_COLOR,
    },
    "wxExcelShape": {"SetGraphicStyle": SET_FAILS},
    "wxExcelTextFrame2": {"SetWordArtformat": SET_FAILS},
    "wxExcelFont2": {"SetWordArtformat": SET_FAILS},
    "wxExcelFillFormat": {"SetRotateWithObject": SET_FAILS},
    "wxExcelLineFormat": {
        "SetBeginArrowheadLength": ARROWHEADS,
        "SetBeginArrowheadStyle": ARROWHEADS,
        "SetBeginArrowheadWidth": ARROWHEADS,
        "SetEndArrowheadLength": ARROWHEADS,
        "SetEndArrowheadStyle": ARROWHEADS,
        "SetEndArrowheadWidth": ARROWHEADS,
        "SetPattern": SET_FAILS,
    },
    "wxExcelThreeDFormat": {
        "SetPresetCamera": SET_FAILS,
        "SetPresetLightingSoftness": SET_FAILS,
    },
    "wxExcelShapeRange": {
        "SetAutoShapeType": ONE_SHAPE,
        "SetBackgroundStyle": ONE_SHAPE,
        "SetGraphicStyle": ONE_SHAPE,
        "SetLeft": ONE_SHAPE,
        "SetShapeStyle": "fails for shapes with different styles, for which Excel returns msoShapeStyleMixed",
    },
    "wxExcelAboveAverage": {
        "SetCalcFor": "only for PivotTables",
        "SetNumStDev": "only for rules based on the standard deviation",
    },
    "wxExcelTop10": {"SetCalcFor": "only for PivotTables"},
    "wxExcelColorScaleCriterion": {
        "SetValue": "only for criteria of the number, percent, or percentile type",
    },
}


def strip_comments(text):
    """Removes the comments and the preprocessor lines."""
    text = re.sub(r"/\*.*?\*/", "", text, flags=re.S)
    text = re.sub(r"//[^\n]*", "", text)
    return re.sub(r"^\s*#[^\n]*", "", text, flags=re.M)


def normalize_type(declaration):
    """Returns the type without const, & and spaces, e.g. "wxColour" for "const wxColour&"."""
    return re.sub(r"\bconst\b|&|\s", "", declaration)


def parse_headers():
    """Returns {class name: (base class names, {getter name: result type},
    {setter name: parameter type})}."""
    classes = {}
    for name in sorted(os.listdir(INCLUDE_DIR)):
        if not (name.startswith("wxAutoExcel") and name.endswith(".h")):
            continue
        with open(os.path.join(INCLUDE_DIR, name), encoding="latin-1") as f:
            text = strip_comments(f.read())
        for match in re.finditer(r"class\s+WXDLLIMPEXP_WXAUTOEXCEL\s+(\w+)\s*(?::\s*([^{]+))?\{", text):
            end = match.end()
            depth = 1
            while depth and end < len(text):
                depth += {"{": 1, "}": -1}.get(text[end], 0)
                end += 1
            bases = re.findall(r"(?:public|protected|private)?\s*(\w+)", match.group(2) or "")
            access = "private"
            getters = {}
            setters = {}
            for declaration in text[match.end():end - 1].split(";"):
                for specifier in re.findall(r"\b(public|protected|private)\s*:", declaration):
                    access = specifier
                declaration = re.sub(r".*\b(?:public|protected|private)\s*:", "", declaration, flags=re.S)
                # static functions such as wxExcelApplication::GetInstance() are not getters
                if access != "public" or re.match(r"\s*static\b", declaration):
                    continue
                function = re.match(r"\s*(?:virtual\s+)?([\w:<>\s\*&,]+?)\s+((?:Get|Set)\w+)\s*\((.*)\)\s*(?:const)?\s*$",
                                    declaration, re.S)
                if not function or function.group(2).endswith("_"):
                    continue
                result, name = function.group(1), function.group(2)
                params = [p for p in function.group(3).split(",") if p.strip() and p.strip() != "void"]
                if name.startswith("Get"):
                    if all("=" in p for p in params) and name not in getters:
                        getters[name] = normalize_type(result)
                # setters with one required parameter, e.g. "const wxColour& color"
                elif normalize_type(result) == "void" and params and "=" not in params[0] \
                        and all("=" in p for p in params[1:]):
                    parameter = re.match(r"(.*?)\b\w+\s*$", params[0].strip(), re.S)
                    if parameter and name not in setters:
                        setters[name] = normalize_type(parameter.group(1))
            classes[match.group(1)] = (bases, getters, setters)
    return classes


def all_members(classes, name, index):
    """Returns {name: type} of the getters (index 1) or the setters (index 2)
    of the class including the inherited ones."""
    result = {}
    for base in classes[name][0]:
        if base in classes:
            for member, member_type in all_members(classes, base, index).items():
                result.setdefault(member, member_type)
    for member, member_type in classes[name][index].items():
        result.setdefault(member, member_type)
    return result


def main():
    classes = parse_headers()
    lines = [
        "// This file is generated by tests/tools/generate_getter_lists.py, do not edit it.",
        "",
        "#ifndef WXAUTOEXCEL_TESTS_GETTER_LISTS_H",
        "#define WXAUTOEXCEL_TESTS_GETTER_LISTS_H",
    ]
    for name in CLASSES:
        excluded = EXCLUDED.get(name, {})
        getters = all_members(classes, name, 1)
        setters = all_members(classes, name, 2)
        unknown = set(excluded) - set(getters)
        if unknown:
            raise SystemExit(f"Excluded getters not found in {name}: {', '.join(sorted(unknown))}")
        excluded_setters = SETTERS_EXCLUDED.get(name, {})
        unknown = set(excluded_setters) - set(setters)
        if unknown:
            raise SystemExit(f"Excluded setters not found in {name}: {', '.join(sorted(unknown))}")

        lines += ["", f"#define WXAUTOEXCEL_GETTERS_{name}(CHECK, object) \\"]
        for getter in getters:
            if getter in excluded:
                lines.append(f"    /* {getter}: {excluded[getter]} */ \\")
            else:
                lines.append(f"    CHECK(object, {getter}) \\")
        lines.append(f"    /* {len(getters)} getters */")

        # the setters whose value can be got with the getter of the same type,
        # or of type wxString for a setter taking wxVariant, e.g. SetFormula()
        properties = [setter[3:] for setter in setters
                      if getters.get("Get" + setter[3:]) == setters[setter]
                      or (setters[setter] == "wxVariant" and getters.get("Get" + setter[3:]) == "wxString")]
        lines += ["", f"#define WXAUTOEXCEL_SETTERS_{name}(CHECK, object) \\"]
        for prop in properties:
            if "Set" + prop in excluded_setters:
                lines.append(f"    /* Set{prop}: {excluded_setters['Set' + prop]} */ \\")
            elif "Get" + prop in excluded:
                lines.append(f"    /* Set{prop}: the getter is excluded */ \\")
            else:
                lines.append(f"    CHECK(object, {prop}) \\")
        lines.append(f"    /* {len(properties)} setters */")
    lines += ["", "#endif // WXAUTOEXCEL_TESTS_GETTER_LISTS_H", ""]
    with open(OUTPUT, "w", newline="\n") as f:
        f.write("\n".join(lines))
    print(f"Written {OUTPUT}")


if __name__ == "__main__":
    main()
