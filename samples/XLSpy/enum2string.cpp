#include "enum2string.h"


wxString XlMeasurementUnits_ToStr(XlMeasurementUnits val)
{
    switch ( val )
    {
        case  xlCentimeters: return "xlCentimeters";
        case  xlInches: return "xlInches";
        case  xlMillimeters: return "xlMillimeters";

        default: return wxString::Format("Unknown (%ld)", (long)val);
    }
}

wxString XlFileFormat_ToStr(XlFileFormat val)
{
    switch ( val )
    {
        // case xlAddIn : return "xlAddIn";
        case xlAddIn8 : return "xlAddIn/xlAddIn8";
        case xlCSV : return "xlCSV";
        case xlCSVMac : return "xlCSVMac";
        case xlCSVMSDOS : return "xlCSVMSDOS";
        case xlCSVWindows : return "xlCSVWindows";
        case xlCurrentPlatformText : return "xlCurrentPlatformText";
        case xlDBF2 : return "xlDBF2";
        case xlDBF3 : return "xlDBF3";
        case xlDBF4 : return "xlDBF4";
        case xlDIF : return "xlDIF";
        case xlExcel12 : return "xlExcel12";
        case xlExcel2 : return "xlExcel2";
        case xlExcel2FarEast : return "xlExcel2FarEast";
        case xlExcel3 : return "xlExcel3";
        case xlExcel4 : return "xlExcel4";
        case xlExcel4Workbook : return "xlExcel4Workbook";
        // case xlExcel5 : return "xlExcel5";
        case xlExcel7 : return "xlExcel5/xlExcel7";
        case xlExcel8 : return "xlExcel8";
        case xlExcel9795 : return "xlExcel9795";
        case xlHtml : return "xlHtml";
        case xlIntlAddIn : return "xlIntlAddIn";
        case xlIntlMacro : return "xlIntlMacro";
        case xlOpenDocumentSpreadsheet : return "xlOpenDocumentSpreadsheet";
        case xlOpenXMLAddIn : return "xlOpenXMLAddIn";
        case xlOpenXMLTemplate : return "xlOpenXMLTemplate";
        case xlOpenXMLTemplateMacroEnabled : return "xlOpenXMLTemplateMacroEnabled";
        // case xlOpenXMLWorkbook : return "xlOpenXMLWorkbook";
        case xlOpenXMLWorkbookMacroEnabled : return "xlOpenXMLWorkbookMacroEnabled";
        case xlSYLK : return "xlSYLK";
        // case xlTemplate : return "xlTemplate";
        case xlTemplate8 : return "xlTemplate/xlTemplate8";
        case xlTextMac : return "xlTextMac";
        case xlTextMSDOS : return "xlTextMSDOS";
        case xlTextPrinter : return "xlTextPrinter";
        case xlTextWindows : return "xlTextWindows";
        case xlUnicodeText : return "xlUnicodeText";
        case xlWebArchive : return "xlWebArchive";
        case xlWJ2WD1 : return "xlWJ2WD1";
        case xlWJ3 : return "xlWJ3";
        case xlWJ3FJ3 : return "xlWJ3FJ3";
        case xlWK1 : return "xlWK1";
        case xlWK1ALL : return "xlWK1ALL";
        case xlWK1FMT : return "xlWK1FMT";
        case xlWK3 : return "xlWK3";
        case xlWK3FM3 : return "xlWK3FM3";
        case xlWK4 : return "xlWK4";
        case xlWKS : return "xlWKS";
        case xlWorkbookDefault : return "xlWorkbookDefault/xlOpenXMLWorkbook";
        case xlWorkbookNormal : return "xlWorkbookNormal";
        case xlWorks2FarEast : return "xlWorks2FarEast";
        case xlWQ1 : return "xlWQ1";
        case xlXMLSpreadsheet : return "xlXMLSpreadsheet";

        default: return wxString::Format("Unknown (%ld)", (long)val);
    }
}

wxString MsoAutomationSecurity_ToStr(MsoAutomationSecurity val)
{
    switch ( val )
    {
        case msoAutomationSecurityByUI : return "msoAutomationSecurityByUI";
        case msoAutomationSecurityForceDisable : return "msoAutomationSecurityForceDisable";
        case msoAutomationSecurityLow : return "msoAutomationSecurityLow";

        default: return wxString::Format("Unknown (%ld)", (long)val);
    }
}

wxString MsoShapeType_ToStr(MsoShapeType val)
{
    switch ( val )
    {
        case msoAutoShape: return "AutoShape";
        case msoCallout: return "Callout";
        case msoCanvas: return "Canvas";
        case msoChart: return "Chart";
        case msoComment: return "Comment";
        case msoDiagram: return "Diagram";
        case msoEmbeddedOLEObject: return "Embedded OLE object";
        case msoFormControl: return "Form control";
        case msoFreeform: return "Freeform";
        case msoGroup: return "Group";
        case msoIgxGraphic: return "SmartArt graphic";
        case msoInk: return "Ink";
        case msoInkComment: return "Ink comment";
        case msoLine: return "Line";
        case msoLinkedOLEObject: return "Linked OLE object";
        case msoLinkedPicture: return "Linked picture";
        case msoMedia: return "Media";
        case msoOLEControlObject: return "OLE control object";
        case msoPicture: return "Picture";
        case msoPlaceholder: return "Placeholder";
        case msoScriptAnchor: return "Script anchor";
        case msoShapeTypeMixed: return "Mixed shape type";
        case msoTable: return "Table";
        case msoTextBox: return "Text box";
        case msoTextEffect: return "Text effect";

        default: return wxString::Format("Unknown (%ld)", (long)val);
    }
}


wxString XlChartType_ToStr(XlChartType val)
{
    switch ( val )
    {
        case xl3DArea : return "xl3DArea";
        case xl3DAreaStacked : return "xl3DAreaStacked";
        case xl3DAreaStacked100 : return "xl3DAreaStacked100";
        case xl3DBarClustered : return "xl3DBarClustered";
        case xl3DBarStacked : return "xl3DBarStacked";
        case xl3DBarStacked100 : return "xl3DBarStacked100";
        case xl3DColumn : return "xl3DColumn";
        case xl3DColumnClustered : return "xl3DColumnClustered";
        case xl3DColumnStacked : return "xl3DColumnStacked";
        case xl3DColumnStacked100 : return "xl3DColumnStacked100";
        case xl3DLine : return "xl3DLine";
        case xl3DPie : return "xl3DPie";
        case xl3DPieExploded : return "xl3DPieExploded";
        case xlArea : return "xlArea";
        case xlAreaStacked : return "xlAreaStacked";
        case xlAreaStacked100 : return "xlAreaStacked100";
        case xlBarClustered : return "xlBarClustered";
        case xlBarOfPie : return "xlBarOfPie";
        case xlBarStacked : return "xlBarStacked";
        case xlBarStacked100 : return "xlBarStacked100";
        case xlBubble : return "xlBubble";
        case xlBubble3DEffect : return "xlBubble3DEffect";
        case xlColumnClustered : return "xlColumnClustered";
        case xlColumnStacked : return "xlColumnStacked";
        case xlColumnStacked100 : return "xlColumnStacked100";
        case xlConeBarClustered : return "xlConeBarClustered";
        case xlConeBarStacked : return "xlConeBarStacked";
        case xlConeBarStacked100 : return "xlConeBarStacked100";
        case xlConeCol : return "xlConeCol";
        case xlConeColClustered : return "xlConeColClustered";
        case xlConeColStacked : return "xlConeColStacked";
        case xlConeColStacked100 : return "xlConeColStacked100";
        case xlCylinderBarClustered : return "xlCylinderBarClustered";
        case xlCylinderBarStacked : return "xlCylinderBarStacked";
        case xlCylinderBarStacked100 : return "xlCylinderBarStacked100";
        case xlCylinderCol : return "xlCylinderCol";
        case xlCylinderColClustered : return "xlCylinderColClustered";
        case xlCylinderColStacked : return "xlCylinderColStacked";
        case xlCylinderColStacked100 : return "xlCylinderColStacked100";
        case xlDoughnut : return "xlDoughnut";
        case xlDoughnutExploded : return "xlDoughnutExploded";
        case xlLine : return "xlLine";
        case xlLineMarkers : return "xlLineMarkers";
        case xlLineMarkersStacked : return "xlLineMarkersStacked";
        case xlLineMarkersStacked100 : return "xlLineMarkersStacked100";
        case xlLineStacked : return "xlLineStacked";
        case xlLineStacked100 : return "xlLineStacked100";
        case xlPie : return "xlPie";
        case xlPieExploded : return "xlPieExploded";
        case xlPieOfPie : return "xlPieOfPie";
        case xlPyramidBarClustered : return "xlPyramidBarClustered";
        case xlPyramidBarStacked : return "xlPyramidBarStacked";
        case xlPyramidBarStacked100 : return "xlPyramidBarStacked100";
        case xlPyramidCol : return "xlPyramidCol";
        case xlPyramidColClustered : return "xlPyramidColClustered";
        case xlPyramidColStacked : return "xlPyramidColStacked";
        case xlPyramidColStacked100 : return "xlPyramidColStacked100";
        case xlRadar : return "xlRadar";
        case xlRadarFilled : return "xlRadarFilled";
        case xlRadarMarkers : return "xlRadarMarkers";
        case xlStockHLC : return "xlStockHLC";
        case xlStockOHLC : return "xlStockOHLC";
        case xlStockVHLC : return "xlStockVHLC";
        case xlStockVOHLC : return "xlStockVOHLC";
        case xlSurface : return "xlSurface";
        case xlSurfaceTopView : return "xlSurfaceTopView";
        case xlSurfaceTopViewWireframe : return "xlSurfaceTopViewWireframe";
        case xlSurfaceWireframe : return "xlSurfaceWireframe";
        case xlXYScatter : return "xlXYScatter";
        case xlXYScatterLines : return "xlXYScatterLines";
        case xlXYScatterLinesNoMarkers : return "xlXYScatterLinesNoMarkers";
        case xlXYScatterSmooth : return "xlXYScatterSmooth";
        case xlXYScatterSmoothNoMarkers : return "xlXYScatterSmoothNoMarkers";

        default: return wxString::Format("Unknown (%ld)", (long)val);
    }
}

wxString XlSheetVisibility_ToStr(XlSheetVisibility val)
{
    switch ( val )
    {
        case xlSheetHidden: return "xlSheetHidden";
        case xlSheetVeryHidden: return "xlSheetVeryHidden";
        case xlSheetVisible: return "xlSheetVisible";

        default: return wxString::Format("Unknown (%ld)", (long)val);
    }
}

wxString XlPaperSize_ToStr(XlPaperSize val)
{
    switch ( val )
    {
        case xlPaper10x14 : return "xlPaper10x14";
        case xlPaper11x17 : return "xlPaper11x17";
        case xlPaperA3 : return "xlPaperA3";
        case xlPaperA4 : return "xlPaperA4";
        case xlPaperA4Small : return "xlPaperA4Small";
        case xlPaperA5 : return "xlPaperA5";
        case xlPaperB4 : return "xlPaperB4";
        case xlPaperB5 : return "xlPaperB5";
        case xlPaperCsheet : return "xlPaperCsheet";
        case xlPaperDsheet : return "xlPaperDsheet";
        case xlPaperEnvelope10 : return "xlPaperEnvelope10";
        case xlPaperEnvelope11 : return "xlPaperEnvelope11";
        case xlPaperEnvelope12 : return "xlPaperEnvelope12";
        case xlPaperEnvelope14 : return "xlPaperEnvelope14";
        case xlPaperEnvelope9 : return "xlPaperEnvelope9";
        case xlPaperEnvelopeB4 : return "xlPaperEnvelopeB4";
        case xlPaperEnvelopeB5 : return "xlPaperEnvelopeB5";
        case xlPaperEnvelopeB6 : return "xlPaperEnvelopeB6";
        case xlPaperEnvelopeC3 : return "xlPaperEnvelopeC3";
        case xlPaperEnvelopeC4 : return "xlPaperEnvelopeC4";
        case xlPaperEnvelopeC5 : return "xlPaperEnvelopeC5";
        case xlPaperEnvelopeC6 : return "xlPaperEnvelopeC6";
        case xlPaperEnvelopeC65 : return "xlPaperEnvelopeC65";
        case xlPaperEnvelopeDL : return "xlPaperEnvelopeDL";
        case xlPaperEnvelopeItaly : return "xlPaperEnvelopeItaly";
        case xlPaperEnvelopeMonarch : return "xlPaperEnvelopeMonarch";
        case xlPaperEnvelopePersonal : return "xlPaperEnvelopePersonal";
        case xlPaperEsheet : return "xlPaperEsheet";
        case xlPaperExecutive : return "xlPaperExecutive";
        case xlPaperFanfoldLegalGerman : return "xlPaperFanfoldLegalGerman";
        case xlPaperFanfoldStdGerman : return "xlPaperFanfoldStdGerman";
        case xlPaperFanfoldUS : return "xlPaperFanfoldUS";
        case xlPaperFolio : return "xlPaperFolio";
        case xlPaperLedger : return "xlPaperLedger";
        case xlPaperLegal : return "xlPaperLegal";
        case xlPaperLetter : return "xlPaperLetter";
        case xlPaperLetterSmall : return "xlPaperLetterSmall";
        case xlPaperNote : return "xlPaperNote";
        case xlPaperQuarto : return "xlPaperQuarto";
        case xlPaperStatement : return "xlPaperStatement";
        case xlPaperTabloid : return "xlPaperTabloid";
        case xlPaperUser : return "xlPaperUser";

        default: return wxString::Format("Unknown (%ld)", (long)val);
    }
}


wxString XlListObjectSourceType_ToStr(XlListObjectSourceType val)
{
    switch ( val )
    {
        case xlSrcExternal: return "xlSrcExternal";
        case xlSrcQuery:    return "xlSrcQuery";
        case xlSrcRange:    return "xlSrcRange";
        case xlSrcXml:      return "xlSrcXml";
        case xlSrcModel:    return "xlSrcModel";
    }

    return wxString::Format("Unknown (%ld)", (long)val);
}

/*

wxString _ToStr(int val)
{
    switch ( val )
    {
        x
    }

    return wxString::Format("Unknown (%ld)", (long)val);
}

*/