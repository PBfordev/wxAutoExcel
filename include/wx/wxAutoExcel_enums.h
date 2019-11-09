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

    /*! @brief Commmon MS Excel constants
    */
    enum XlConstants {
        xlAll  = -4104,
        xlAutomatic  = -4105,
        xlBoth  = 1,
        xlCenter  = -4108,
        xlChecker  = 9,
        xlCircle  = 8,
        xlCorner  = 2,
        xlCrissCross  = 16,
        xlCross  = 4,
        xlDiamond  = 2,
        xlDistributed  = -4117,
        xlDoubleAccounting  = 5,
        xlFixedValue  = 1,
        xlFormats  = -4122,
        xlGray16  = 17,
        xlGray8  = 18,
        xlGrid  = 15,
        xlHigh  = -4127,
        xlInside  = 2,
        xlJustify  = -4130,
        xlLightDown  = 13,
        xlLightHorizontal  = 11,
        xlLightUp  = 14,
        xlLightVertical  = 12,
        xlLow  = -4134,
        xlManual  = -4135,
        xlMinusValues  = 3,
        xlModule  = -4141,
        xlNextToAxis  = 4,
        xlNone  = -4142,
        xlNotes  = -4144,
        xlOff  = -4146,
        xlOn  = 1,
        xlPercent  = 2,
        xlPlus  = 9,
        xlPlusValues  = 2,
        xlSemiGray75  = 10,
        xlShowLabel  = 4,
        xlShowLabelAndPercent  = 5,
        xlShowPercent  = 3,
        xlShowValue  = 2,
        xlSimple  = -4154,
        xlSingle  = 2,
        xlSingleAccounting  = 4,
        xlSolid  = 1,
        xlSquare  = 1,
        xlStar  = 5,
        xlStError  = 4,
        xlToolbarButton  = 2,
        xlTriangle  = 3,
        xlGray25  = -4124,
        xlGray50  = -4125,
        xlGray75  = -4126,
        xlBottom  = -4107,
        xlLeft  = -4131,
        xlRight  = -4152,
        xlTop  = -4160,
        xl3DBar  = -4099,
        xl3DSurface  = -4103,
        xlBar  = 2,
        xlColumn  = 3,
        xlCombination  = -4111,
        xlCustom  = -4114,
        xlDefaultAutoFormat  = -1,
        xlMaximum  = 2,
        xlMinimum  = 4,
        xlOpaque  = 3,
        xlTransparent  = 2,
        xlBidi  = -5000,
        xlLatin  = -5001,
        xlContext  = -5002,
        xlLTR  = -5003,
        xlRTL  = -5004,
        xlFullScript  = 1,
        xlPartialScript  = 2,
        xlMixedScript  = 3,
        xlMixedAuthorizedScript  = 4,
        xlVisualCursor  = 2,
        xlLogicalCursor  = 1,
        xlSystem  = 1,
        xlPartial  = 3,
        xlHindiNumerals  = 3,
        xlBidiCalendar  = 3,
        xlGregorian  = 2,
        xlComplete  = 4,
        xlScale  = 3,
        xlClosed  = 3,
        xlColor1  = 7,
        xlColor2  = 8,
        xlColor3  = 9,
        xlConstants  = 2,
        xlContents  = 2,
        xlBelow  = 1,
        xlCascade  = 7,
        xlCenterAcrossSelection  = 7,
        xlChart4  = 2,
        xlChartSeries  = 17,
        xlChartShort  = 6,
        xlChartTitles  = 18,
        xlClassic1  = 1,
        xlClassic2  = 2,
        xlClassic3  = 3,
        xl3DEffects1  = 13,
        xl3DEffects2  = 14,
        xlAbove  = 0,
        xlAccounting1  = 4,
        xlAccounting2  = 5,
        xlAccounting3  = 6,
        xlAccounting4  = 17,
        xlAdd  = 2,
        xlDebugCodePane  = 13,
        xlDesktop  = 9,
        xlDirect  = 1,
        xlDivide  = 5,
        xlDoubleClosed  = 5,
        xlDoubleOpen  = 4,
        xlDoubleQuote  = 1,
        xlEntireChart  = 20,
        xlExcelMenus  = 1,
        xlExtended  = 3,
        xlFill  = 5,
        xlFirst  = 0,
        xlFloating  = 5,
        xlFormula  = 5,
        xlGeneral  = 1,
        xlGridline  = 22,
        xlIcons  = 1,
        xlImmediatePane  = 12,
        xlInteger  = 2,
        xlLast  = 1,
        xlLastCell  = 11,
        xlList1  = 10,
        xlList2  = 11,
        xlList3  = 12,
        xlLocalFormat1  = 15,
        xlLocalFormat2  = 16,
        xlLong  = 3,
        xlLotusHelp  = 2,
        xlMacrosheetCell  = 7,
        xlMixed  = 2,
        xlMultiply  = 4,
        xlNarrow  = 1,
        xlNoDocuments  = 3,
        xlOpen  = 2,
        xlOutside  = 3,
        xlReference  = 4,
        xlSemiautomatic  = 2,
        xlShort  = 1,
        xlSingleQuote  = 2,
        xlStrict  = 2,
        xlSubtract  = 3,
        xlTextBox  = 16,
        xlTiled  = 1,
        xlTitleBar  = 8,
        xlToolbar  = 1,
        xlVisible  = 12,
        xlWatchPane  = 11,
        xlWide  = 3,
        xlWorkbookTab  = 6,
        xlWorksheet4  = 1,
        xlWorksheetCell  = 3,
        xlWorksheetShort  = 5,
        xlAllExceptBorders  = 7,
        xlLeftToRight  = 2,
        xlTopToBottom  = 1,
        xlVeryHidden  = 2,
        xlDrawingObject  = 14,
    };

    /*!Specifies a language setting in a Microsoft Office application. The msoAppLanguageID enumeration is used with the LanguageSettings member of the Application object to determine the language used for the install language, the user interface language, or the Help language.

    [MSDN documentation for MsoAppLanguageID](http://msdn.microsoft.com/en-us/library/aa432459%28v=office.12%29.aspx)
    */
    enum MsoAppLanguageID {
        msoLanguageIDExeMode = 4, /*!< Execution mode language. */
        msoLanguageIDHelp = 3, /*!< Help language. */
        msoLanguageIDInstall = 1, /*!< Install language. */
        msoLanguageIDUI = 2, /*!< User interface language. */
        msoLanguageIDUIPrevious = 5, /*!< User interface language used prior to the current user interface language. */
    };

    /*!Specifies the length of the arrowhead at the end of a line.

    [MSDN documentation for MsoArrowheadLength](http://msdn.microsoft.com/en-us/library/office/aa432460.aspx).
    */
    enum MsoArrowheadLength {
        msoArrowheadLengthMedium = 2 , /*!< Medium. */
        msoArrowheadLengthMixed = -2 , /*!< Return value only; indicates a combination of the other states in the specified shape range. */
        msoArrowheadLong = 3 , /*!< Long. */
        msoArrowheadShort = 1 , /*!< Short. */
    };

    /*!Defines how to align specified objects relative to one another.

    [MSDN documentation for MsoAlignCmd](http://msdn.microsoft.com/en-us/library/office/aa432457.aspx).
    */
    enum MsoAlignCmd {
        msoAlignBottoms = 5 , /*!< Align bottoms of specified objects. */
        msoAlignCenters = 1 , /*!< Align centers of specified objects. */
        msoAlignLefts = 0 , /*!< Align left sides of specified objects. */
        msoAlignMiddles = 4 , /*!< Align middles of specified objects. */
        msoAlignRights = 2 , /*!< Align right sides of specified objects. */
        msoAlignTops = 3 , /*!< Align tops of specified objects. */
    };

    /*!Specifies the style of the arrowhead at the end of a line.

    [MSDN documentation for MsoArrowheadStyle](http://msdn.microsoft.com/en-us/library/office/aa432462.aspx).
    */
    enum MsoArrowheadStyle {
        msoArrowheadDiamond = 5 , /*!< Diamond-shaped. */
        msoArrowheadNone = 1 , /*!< No arrowhead. */
        msoArrowheadOpen = 3 , /*!< Open. */
        msoArrowheadOval = 6 , /*!< Oval-shaped. */
        msoArrowheadStealth = 4 , /*!< Stealth-shaped. */
        msoArrowheadStyleMixed = -2 , /*!< Return value only; indicates a combination of the other states.  */
        msoArrowheadTriangle = 2 , /*!< Triangular. */
    };
    /*!Specifies the width of the arrowhead at the end of a line.

    [MSDN documentation for MsoArrowheadWidth](http://msdn.microsoft.com/en-us/library/office/aa432467.aspx).
    */
    enum MsoArrowheadWidth {
        msoArrowheadNarrow = 1 , /*!< Narrow. */
        msoArrowheadWide = 3 , /*!< Wide. */
        msoArrowheadWidthMedium = 2 , /*!< Medium. */
        msoArrowheadWidthMixed = -2 , /*!< Return value only; indicates a combination of the other states.  */
    };

    /*!  @brief Specifies the security mode an application uses when programmatically opening files.

    [MSDN documentation for MsoAutomationSecurity](http://msdn.microsoft.com/en-us/library/aa432468.aspx).
    */
    enum MsoAutomationSecurity {
        msoAutomationSecurityByUI = 2, /*!< Uses the security setting specified in the Security dialog box.*/
        msoAutomationSecurityForceDisable =	3, /*!< Disables all macros in all files opened programmatically, without showing any security alerts.*/
        msoAutomationSecurityLow =	1, /*!< Enables all macros. This is the default value when the application is started.*/
    };

    /*!Specifies the shape type.

    [MSDN documentation for MsoAutoShapeType](http://msdn.microsoft.com/en-us/library/office/aa432469.aspx).
    */
    enum MsoAutoShapeType {
        msoShape16pointStar = 94 , /*!< 16-point star. */
        msoShape24pointStar = 95 , /*!< 24-point star. */
        msoShape32pointStar = 96 , /*!< 32-point star. */
        msoShape4pointStar = 91 , /*!< 4-point star. */
        msoShape5pointStar = 92 , /*!< 5-point star. */
        msoShape8pointStar = 93 , /*!< 8-point star. */
        msoShapeActionButtonBackorPrevious = 129 , /*!< Back or Previous button. Supports mouse-click and mouse-over actions. */
        msoShapeActionButtonBeginning = 131 , /*!< Beginning button. Supports mouse-click and mouse-over actions. */
        msoShapeActionButtonCustom = 125 , /*!< Button with no default picture or text. Supports mouse-click and mouse-over actions. */
        msoShapeActionButtonDocument = 134 , /*!< Document button. Supports mouse-click and mouse-over actions. */
        msoShapeActionButtonEnd = 132 , /*!< End button. Supports mouse-click and mouse-over actions. */
        msoShapeActionButtonForwardorNext = 130 , /*!< Forward or Next button. Supports mouse-click and mouse-over actions. */
        msoShapeActionButtonHelp = 127 , /*!< Help button. Supports mouse-click and mouse-over actions. */
        msoShapeActionButtonHome = 126 , /*!< Home button. Supports mouse-click and mouse-over actions. */
        msoShapeActionButtonInformation = 128 , /*!< Information button. Supports mouse-click and mouse-over actions. */
        msoShapeActionButtonMovie = 136 , /*!< Movie button. Supports mouse-click and mouse-over actions. */
        msoShapeActionButtonReturn = 133 , /*!< Return button. Supports mouse-click and mouse-over actions. */
        msoShapeActionButtonSound = 135 , /*!< Sound button. Supports mouse-click and mouse-over actions. */
        msoShapeArc = 25 , /*!< Arc. */
        msoShapeBalloon = 137 , /*!< Balloon. */
        msoShapeBentArrow = 41 , /*!< Block arrow that follows a curved 90-degree angle.  */
        msoShapeBentUpArrow = 44 , /*!< Block arrow that follows a sharp 90-degree angle. Points up by default. */
        msoShapeBevel = 15 , /*!< Bevel. */
        msoShapeBlockArc = 20 , /*!< Block arc. */
        msoShapeCan = 13 , /*!< Can. */
        msoShapeChevron = 52 , /*!< Chevron. */
        msoShapeCircularArrow = 60 , /*!< Block arrow that follows a curved 180-degree angle. */
        msoShapeCloudCallout = 108 , /*!< Cloud callout. */
        msoShapeCross = 11 , /*!< Cross. */
        msoShapeCube = 14 , /*!< Cube. */
        msoShapeCurvedDownArrow = 48 , /*!< Block arrow that curves down. */
        msoShapeCurvedDownRibbon = 100 , /*!< Ribbon banner that curves down. */
        msoShapeCurvedLeftArrow = 46 , /*!< Block arrow that curves left. */
        msoShapeCurvedRightArrow = 45 , /*!< Block arrow that curves right. */
        msoShapeCurvedUpArrow = 47 , /*!< Block arrow that curves up. */
        msoShapeCurvedUpRibbon = 99 , /*!< Ribbon banner that curves up. */
        msoShapeDiamond = 4 , /*!< Diamond. */
        msoShapeDonut = 18 , /*!< Donut. */
        msoShapeDoubleBrace = 27 , /*!< Double brace. */
        msoShapeDoubleBracket = 26 , /*!< Double bracket. */
        msoShapeDoubleWave = 104 , /*!< Double wave. */
        msoShapeDownArrow = 36 , /*!< Block arrow that points down. */
        msoShapeDownArrowCallout = 56 , /*!< Callout with arrow that points down. */
        msoShapeDownRibbon = 98 , /*!< Ribbon banner with center area below ribbon ends. */
        msoShapeExplosion1 = 89 , /*!< Explosion. */
        msoShapeExplosion2 = 90 , /*!< Explosion. */
        msoShapeFlowchartAlternateProcess = 62 , /*!< Alternate process flowchart symbol. */
        msoShapeFlowchartCard = 75 , /*!< Card flowchart symbol. */
        msoShapeFlowchartCollate = 79 , /*!< Collate flowchart symbol. */
        msoShapeFlowchartConnector = 73 , /*!< Connector flowchart symbol. */
        msoShapeFlowchartData = 64 , /*!< Data flowchart symbol. */
        msoShapeFlowchartDecision = 63 , /*!< Decision flowchart symbol. */
        msoShapeFlowchartDelay = 84 , /*!< Delay flowchart symbol. */
        msoShapeFlowchartDirectAccessStorage = 87 , /*!< Direct access storage flowchart symbol. */
        msoShapeFlowchartDisplay = 88 , /*!< Display flowchart symbol. */
        msoShapeFlowchartDocument = 67 , /*!< Document flowchart symbol. */
        msoShapeFlowchartExtract = 81 , /*!< Extract flowchart symbol. */
        msoShapeFlowchartInternalStorage = 66 , /*!< Internal storage flowchart symbol. */
        msoShapeFlowchartMagneticDisk = 86 , /*!< Magnetic disk flowchart symbol. */
        msoShapeFlowchartManualInput = 71 , /*!< Manual input flowchart symbol. */
        msoShapeFlowchartManualOperation = 72 , /*!< Manual operation flowchart symbol. */
        msoShapeFlowchartMerge = 82 , /*!< Merge flowchart symbol. */
        msoShapeFlowchartMultidocument = 68 , /*!< Multi-document flowchart symbol. */
        msoShapeFlowchartOffpageConnector = 74 , /*!< Off-page connector flowchart symbol. */
        msoShapeFlowchartOr = 78 , /*!< "Or" flowchart symbol. */
        msoShapeFlowchartPredefinedProcess = 65 , /*!< Predefined process flowchart symbol. */
        msoShapeFlowchartPreparation = 70 , /*!< Preparation flowchart symbol. */
        msoShapeFlowchartProcess = 61 , /*!< Process flowchart symbol. */
        msoShapeFlowchartPunchedTape = 76 , /*!< Punched tape flowchart symbol. */
        msoShapeFlowchartSequentialAccessStorage = 85 , /*!< Sequential access storage flowchart symbol. */
        msoShapeFlowchartSort = 80 , /*!< Sort flowchart symbol. */
        msoShapeFlowchartStoredData = 83 , /*!< Stored data flowchart symbol. */
        msoShapeFlowchartSummingJunction = 77 , /*!< Summing junction flowchart symbol. */
        msoShapeFlowchartTerminator = 69 , /*!< Terminator flowchart symbol. */
        msoShapeFoldedCorner = 16 , /*!< Folded corner. */
        msoShapeHeart = 21 , /*!< Heart. */
        msoShapeHexagon = 10 , /*!< Hexagon. */
        msoShapeHorizontalScroll = 102 , /*!< Horizontal scroll. */
        msoShapeIsoscelesTriangle = 7 , /*!< Isosceles triangle. */
        msoShapeLeftArrow = 34 , /*!< Block arrow that points left. */
        msoShapeLeftArrowCallout = 54 , /*!< Callout with arrow that points left. */
        msoShapeLeftBrace = 31 , /*!< Left brace. */
        msoShapeLeftBracket = 29 , /*!< Left bracket. */
        msoShapeLeftRightArrow = 37 , /*!< Block arrow with arrowheads that point both left and right. */
        msoShapeLeftRightArrowCallout = 57 , /*!< Callout with arrowheads that point both left and right. */
        msoShapeLeftRightUpArrow = 40 , /*!< Block arrow with arrowheads that point left, right, and up. */
        msoShapeLeftUpArrow = 43 , /*!< Block arrow with arrowheads that point left and up. */
        msoShapeLightningBolt = 22 , /*!< Lightning bolt. */
        msoShapeLineCallout1 = 109 , /*!< Callout with border and horizontal callout line. */
        msoShapeLineCallout1AccentBar = 113 , /*!< Callout with horizontal accent bar. */
        msoShapeLineCallout1BorderandAccentBar = 121 , /*!< Callout with border and horizontal accent bar. */
        msoShapeLineCallout1NoBorder = 117 , /*!< Callout with horizontal line. */
        msoShapeLineCallout2 = 110 , /*!< Callout with diagonal straight line. */
        msoShapeLineCallout2AccentBar = 114 , /*!< Callout with diagonal callout line and accent bar. */
        msoShapeLineCallout2BorderandAccentBar = 122 , /*!< Callout with border, diagonal straight line, and accent bar. */
        msoShapeLineCallout2NoBorder = 118 , /*!< Callout with no border and diagonal callout line. */
        msoShapeLineCallout3 = 111 , /*!< Callout with angled line. */
        msoShapeLineCallout3AccentBar = 115 , /*!< Callout with angled callout line and accent bar. */
        msoShapeLineCallout3BorderandAccentBar = 123 , /*!< Callout with border, angled callout line, and accent bar. */
        msoShapeLineCallout3NoBorder = 119 , /*!< Callout with no border and angled callout line. */
        msoShapeLineCallout4 = 112 , /*!< Callout with callout line segments forming a U-shape. */
        msoShapeLineCallout4AccentBar = 116 , /*!< Callout with accent bar and callout line segments forming a U-shape. */
        msoShapeLineCallout4BorderandAccentBar = 124 , /*!< Callout with border, accent bar, and callout line segments forming a U-shape. */
        msoShapeLineCallout4NoBorder = 120 , /*!< Callout with no border and callout line segments forming a U-shape. */
        msoShapeMixed = -2 , /*!< Return value only; indicates a combination of the other states.  */
        msoShapeMoon = 24 , /*!< Moon. */
        msoShapeNoSymbol = 19 , /*!< "No" symbol. */
        msoShapeNotchedRightArrow = 50 , /*!< Notched block arrow that points right. */
        msoShapeNotPrimitive = 138 , /*!< Not supported. */
        msoShapeOctagon = 6 , /*!< Octagon. */
        msoShapeOval = 9 , /*!< Oval. */
        msoShapeOvalCallout = 107 , /*!< Oval-shaped callout. */
        msoShapeParallelogram = 2 , /*!< Parallelogram. */
        msoShapePentagon = 51 , /*!< Pentagon. */
        msoShapePlaque = 28 , /*!< Plaque. */
        msoShapeQuadArrow = 39 , /*!< Block arrows that point up, down, left, and right. */
        msoShapeQuadArrowCallout = 59 , /*!< Callout with arrows that point up, down, left, and right. */
        msoShapeRectangle = 1 , /*!< Rectangle. */
        msoShapeRectangularCallout = 105 , /*!< Rectangular callout. */
        msoShapeRegularPentagon = 12 , /*!< Pentagon. */
        msoShapeRightArrow = 33 , /*!< Block arrow that points right. */
        msoShapeRightArrowCallout = 53 , /*!< Callout with arrow that points right. */
        msoShapeRightBrace = 32 , /*!< Right brace. */
        msoShapeRightBracket = 30 , /*!< Right bracket. */
        msoShapeRightTriangle = 8 , /*!< Right triangle. */
        msoShapeRoundedRectangle = 5 , /*!< Rounded rectangle. */
        msoShapeRoundedRectangularCallout = 106 , /*!< Rounded rectangle-shaped callout. */
        msoShapeSmileyFace = 17 , /*!< Smiley face. */
        msoShapeStripedRightArrow = 49 , /*!< Block arrow that points right with stripes at the tail. */
        msoShapeSun = 23 , /*!< Sun. */
        msoShapeTrapezoid = 3 , /*!< Trapezoid. */
        msoShapeUpArrow = 35 , /*!< Block arrow that points up. */
        msoShapeUpArrowCallout = 55 , /*!< Callout with arrow that points up. */
        msoShapeUpDownArrow = 38 , /*!< Block arrow that points up and down. */
        msoShapeUpDownArrowCallout = 58 , /*!< Callout with arrows that point up and down. */
        msoShapeUpRibbon = 97 , /*!< Ribbon banner with center area above ribbon ends. */
        msoShapeUTurnArrow = 42 , /*!< Block arrow forming a U shape. */
        msoShapeVerticalScroll = 101 , /*!< Vertical scroll. */
        msoShapeWave = 103 , /*!< Wave. */
    };

    /*!Determines the type of automatic sizing allowed.

    [MSDN documentation for MsoAutoSize](http://msdn.microsoft.com/en-us/library/office/aa432470.aspx).
    */
    enum MsoAutoSize {
        msoAutoSizeMixed = -2 , /*!< A combination of automatic sizing schemes are used. */
        msoAutoSizeNone = 0 , /*!< No autosizing. */
        msoAutoSizeShapeToFitText = 1 , /*!< The shape is adjusted to fit the text. */
        msoAutoSizeTextToFitShape = 2 , /*!< The text is adjusted to fit the shape. */
    };

    /*!Indicates the background style for an object.

    [MSDN documentation for MsoBackgroundStyleIndex](http://msdn.microsoft.com/en-us/library/office/aa432471.aspx).
    */
    enum MsoBackgroundStyleIndex {
        msoBackgroundStyle1 = 1 , /*!< Specifies Style1. */
        msoBackgroundStyle10 = 10 , /*!< Specifies Style10. */
        msoBackgroundStyle11 = 11 , /*!< Specifies Style11. */
        msoBackgroundStyle12 = 12 , /*!< Specifies Style12. */
        msoBackgroundStyle2 = 2 , /*!< Specifies Style2. */
        msoBackgroundStyle3 = 3 , /*!< Specifies Style3. */
        msoBackgroundStyle4 = 4 , /*!< Specifies Style4. */
        msoBackgroundStyle5 = 5 , /*!< Specifies Style5. */
        msoBackgroundStyle6 = 6 , /*!< Specifies Style6. */
        msoBackgroundStyle7 = 7 , /*!< Specifies Style7. */
        msoBackgroundStyle8 = 8 , /*!< Specifies Style8. */
        msoBackgroundStyle9 = 9 , /*!< Specifies Style9. */
        msoBackgroundStyleMixed = -2 , /*!< Specifies a combination of styles. */
        msoBackgroundStyleNone = 0 , /*!< Specifies no styles. */
    };

    /*!Specifies baseline text alignment.

    [MSDN documentation for MsoBaselineAlignment](http://msdn.microsoft.com/en-us/library/office/ff862048.aspx).
    */
    enum MsoBaselineAlignment {
        msoBaselineAlignAuto = 5, /*!< Automatic alignment. */
        msoBaselineAlignBaseline = 1, /*!< Baseline alignment. */
        msoBaselineAlignCenter = 3, /*!< Center alignment. */
        msoBaselineAlignFarEast50 = 4, /*!< East Asia 50 alignment. */
        msoBaselineAlignMixed = -2, /*!< Return value only; indicates a combination of the other states. */
        msoBaselineAlignTop = 2, /*!< Top alignment. */
    };

    /*!Indicates the bevel type of a ThreeDFormat object.

    [MSDN documentation for MsoBevelType](http://msdn.microsoft.com/en-us/library/office/aa432479.aspx).
    */
    enum MsoBevelType {
        msoBevelAngle = 6 , /*!< Specifies an Angle bevel. */
        msoBevelArtDeco = 13 , /*!< Specifies an ArtDeco bevel. */
        msoBevelCircle = 3 , /*!< Specifies a Circle bevel. */
        msoBevelConvex = 8 , /*!< Specifies a Convex bevel. */
        msoBevelCoolSlant = 9 , /*!< Specifies a CoolSlant bevel. */
        msoBevelCross = 5 , /*!< Specifies a Cross bevel. */
        msoBevelDivot = 10 , /*!< Specifies a Divot bevel. */
        msoBevelHardEdge = 12 , /*!< Specifies a HardEdge bevel. */
        msoBevelNone = 1 , /*!< Specifies no bevel. */
        msoBevelRelaxedInset = 2 , /*!< Specifies a RelaxedInset bevel. */
        msoBevelRiblet = 11 , /*!< Specifies a Riblet bevel. */
        msoBevelSlope = 4 , /*!< Specifies a Slope bevel. */
        msoBevelSoftRound = 7 , /*!< Specifies a SoftRound bevel. */
        msoBevelTypeMixed = -2 , /*!< Specifies a mixed type bevel. */
    };

    /*!Specifies the size of the angle between the callout line and the side of the callout text box.

    [MSDN documentation for MsoCalloutAngleType](http://msdn.microsoft.com/en-us/library/office/aa432485.aspx).
    */
    enum MsoCalloutAngleType {
        msoCalloutAngle30 = 2 , /*!< 30Â° angle. */
        msoCalloutAngle45 = 3 , /*!< 45Â° angle. */
        msoCalloutAngle60 = 4 , /*!< 60Â° angle. */
        msoCalloutAngle90 = 5 , /*!< 90Â° angle. */
        msoCalloutAngleAutomatic = 1 , /*!< Default angle. Angle can be changed as you drag the object. */
        msoCalloutAngleMixed = -2 , /*!< Return value only; indicates a combination of the other states.  */
    };

    /*!Specifies how a shape appears when viewed in black-and-white mode.

    [MSDN documentation for MsoBlackWhiteMode](http://msdn.microsoft.com/en-us/library/office/aa432480.aspx).
    */
    enum MsoBlackWhiteMode {
        msoBlackWhiteAutomatic = 1 , /*!< Default behavior. */
        msoBlackWhiteBlack = 8 , /*!< Black. */
        msoBlackWhiteBlackTextAndLine = 6 , /*!< White with grayscale fill. */
        msoBlackWhiteDontShow = 10 , /*!< Not shown. */
        msoBlackWhiteGrayOutline = 5 , /*!< Gray with white fill. */
        msoBlackWhiteGrayScale = 2 , /*!< Grayscale. */
        msoBlackWhiteHighContrast = 7 , /*!< Black with white fill. */
        msoBlackWhiteInverseGrayScale = 4 , /*!< Inverse grayscale. */
        msoBlackWhiteLightGrayScale = 3 , /*!< Light grayscale. */
        msoBlackWhiteMixed = -2 , /*!< Not supported. */
        msoBlackWhiteWhite = 9 , /*!< White. */
    };

    /*!

    [MSDN documentation for MsoBulletType](http://msdn.microsoft.com/en-us/library/office/ff861105.aspx).
    */
    enum MsoBulletType {
        msoBulletMixed = -2, /*!< Return value only; indicates a combination of the other states. */
        msoBulletNone = 0, /*!< No bullets. */
        msoBulletNumbered = 2, /*!< Numbered bullets. */
        msoBulletPicture = 3, /*!< Picture bullets. */
        msoBulletUnnumbered = 1, /*!< Unnumbered bullets. */

    };

    /*!Specifies starting position of the callout line relative to the text bounding box.

    [MSDN documentation for MsoCalloutDropType](http://msdn.microsoft.com/en-us/library/office/aa432486.aspx).
    */
    enum MsoCalloutDropType {
        msoCalloutDropBottom = 4 , /*!< Bottom. */
        msoCalloutDropCenter = 3 , /*!< Center. */
        msoCalloutDropCustom = 1 , /*!< Custom. If this value is used as the value for the PresetDrop property, the Drop and AutoAttach properties of the CalloutFormat object are used to determine where the callout line attaches to the text box.  */
        msoCalloutDropMixed = -2 , /*!< Return value only; indicates a combination of the other states.  */
        msoCalloutDropTop = 2 , /*!< Top. */
    };

    /*!Specifies the type of callout line.

    [MSDN documentation for MsoCalloutType](http://msdn.microsoft.com/en-us/library/office/aa432487.aspx).
    */
    enum MsoCalloutType {
        msoCalloutFour = 4 , /*!< Callout line made up of two line segments. Callout line is attached on right side of text bounding box. */
        msoCalloutMixed = -2 , /*!< Return value only; indicates a combination of the other states.  */
        msoCalloutOne = 1 , /*!< Single, horizontal callout line. */
        msoCalloutThree = 3 , /*!< Callout line made up of two line segments. Callout line is attached on left side of text bounding box. */
        msoCalloutTwo = 2 , /*!< Single, angled callout line. */
    };

    /*!Specifies whether and how to display chart elements.

    [MSDN documentation for MsoChartElementType](http://msdn.microsoft.com/en-us/library/office/ff864118.aspx).
    */
    enum MsoChartElementType {
        msoElementChartFloorNone = 1200, /*!< Do not display chart floor. */
        msoElementChartFloorShow = 1201, /*!< Display chart floor. */
        msoElementChartTitleAboveChart = 2, /*!< Display title above chart. */
        msoElementChartTitleCenteredOverlay = 1, /*!< Display title as centered overlay. */
        msoElementChartTitleNone = 0, /*!< Do not display chart title. */
        msoElementChartWallNone = 1100, /*!< Do not display chart wall. */
        msoElementChartWallShow = 1101, /*!< Dispaly chart wall. */
        msoElementDataLabelBestFit = 210, /*!< Use best fit for data label. */
        msoElementDataLabelBottom = 209, /*!< Display data label at bottom. */
        msoElementDataLabelCenter = 202, /*!< Display data label in center. */
        msoElementDataLabelInsideBase = 204, /*!< Display data label inside at the base. */
        msoElementDataLabelInsideEnd = 203, /*!< Display data label inside at the end. */
        msoElementDataLabelLeft = 206, /*!< Display data label to the left. */
        msoElementDataLabelNone = 200, /*!< Do not display data label. */
        msoElementDataLabelOutSideEnd = 205, /*!< Display data label outside at the end. */
        msoElementDataLabelRight = 207, /*!< Display data label to the right. */
        msoElementDataLabelShow = 201, /*!< Display data label. */
        msoElementDataLabelTop = 208, /*!< Display data label at the top. */
        msoElementDataTableNone = 500, /*!< Do not display data table. */
        msoElementDataTableShow = 501, /*!< Display data table. */
        msoElementDataTableWithLegendKeys = 502, /*!< Display data table with legend keys. */
        msoElementErrorBarNone = 700, /*!< Do not display error bar. */
        msoElementErrorBarPercentage = 702, /*!< Display percentage error bar. */
        msoElementErrorBarStandardDeviation = 703, /*!< Display standard deviation error bar. */
        msoElementErrorBarStandardError = 701, /*!< Display standard error bar. */
        msoElementLegendBottom = 104, /*!< Display legend at the bottom. */
        msoElementLegendLeft = 103, /*!< Display legend at the left. */
        msoElementLegendLeftOverlay = 106, /*!< Overlay legend at the left. */
        msoElementLegendNone = 100, /*!< Do not display legend. */
        msoElementLegendRight = 101, /*!< Display legend at the right. */
        msoElementLegendRightOverlay = 105, /*!< Overlay legend at the right. */
        msoElementLegendTop = 102, /*!< Display legend at the top. */
        msoElementLineDropHiLoLine = 804, /*!< Display drop high/low line. */
        msoElementLineDropLine = 801, /*!< Display drop line. */
        msoElementLineHiLoLine = 802, /*!< Display high/low line. */
        msoElementLineNone = 800, /*!< Do not display line. */
        msoElementLineSeriesLine = 803, /*!< Display series line. */
        msoElementPlotAreaNone = 1000, /*!< Do not display plot area. */
        msoElementPlotAreaShow = 1001, /*!< Display plot area. */
        msoElementPrimaryCategoryAxisBillions = 374, /*!< Use billions for primary category axis units. */
        msoElementPrimaryCategoryAxisLogScale = 375, /*!< Use log scale for primary category axis. */
        msoElementPrimaryCategoryAxisMillions = 373, /*!< Use millions for primary category axis units. */
        msoElementPrimaryCategoryAxisNone = 348, /*!< Do not display primary category axis. */
        msoElementPrimaryCategoryAxisReverse = 351, /*!< Reverse primary category axis. */
        msoElementPrimaryCategoryAxisShow = 349, /*!< Show primary category axis. */
        msoElementPrimaryCategoryAxisThousands = 372, /*!< Use thousands for primary category axis units. */
        msoElementPrimaryCategoryAxisTitleAdjacentToAxis = 301, /*!< Display primary category axis title adjacent to the axis. */
        msoElementPrimaryCategoryAxisTitleBelowAxis = 302, /*!< Display primary category axis title below the axis. */
        msoElementPrimaryCategoryAxisTitleHorizontal = 305, /*!< Display primary category axis title horizontally. */
        msoElementPrimaryCategoryAxisTitleNone = 300, /*!< Do not display primary category axis title. */
        msoElementPrimaryCategoryAxisTitleRotated = 303, /*!< Rotate primary category axis title. */
        msoElementPrimaryCategoryAxisTitleVertical = 304, /*!< Display primary category axis title vertically. */
        msoElementPrimaryCategoryAxisWithoutLabels = 350, /*!< Display primary category axis without labels. */
        msoElementPrimaryCategoryGridLinesMajor = 334, /*!< Display major gridlines along primary category axis. */
        msoElementPrimaryCategoryGridLinesMinor = 333, /*!< Display minor gridlines along primary category axis. */
        msoElementPrimaryCategoryGridLinesMinorMajor = 335, /*!< Display both major and minor gridlines along primary category axis. */
        msoElementPrimaryCategoryGridLinesNone = 332, /*!< Do not display grid lines along primary category axis. */
        msoElementPrimaryValueAxisBillions = 356, /*!< Use billions for primary value axis units. */
        msoElementPrimaryValueAxisLogScale = 357, /*!< Use log scale for primary value axis. */
        msoElementPrimaryValueAxisMillions = 355, /*!< Use millions for primary value axis units. */
        msoElementPrimaryValueAxisNone = 352, /*!< Do not display primary value axis. */
        msoElementPrimaryValueAxisShow = 353, /*!< Show primary value axis */
        msoElementPrimaryValueAxisThousands = 354, /*!< Use thousands for primary value axis units. */
        msoElementPrimaryValueAxisTitleAdjacentToAxis = 306, /*!< Place primary value axis title adjacent to the axis. */
        msoElementPrimaryValueAxisTitleBelowAxis = 308, /*!< Place primary value axis title below the axis. */
        msoElementPrimaryValueAxisTitleHorizontal = 311, /*!< Display primary value axis title horizontally. */
        msoElementPrimaryValueAxisTitleNone = 306, /*!< Do not display primary value axis title. */
        msoElementPrimaryValueAxisTitleRotated = 309, /*!< Rotate primary value axis title. */
        msoElementPrimaryValueAxisTitleVertical = 310, /*!< Display primary value axis title vertically. */
        msoElementPrimaryValueGridLinesMajor = 330, /*!< Display major gridlines along primary value axis. */
        msoElementPrimaryValueGridLinesMinor = 329, /*!< Display minor gridlines along primary value axis. */
        msoElementPrimaryValueGridLinesMinorMajor = 331, /*!< Display both major and minor gridlines along primary value axis. */
        msoElementPrimaryValueGridLinesNone = 328, /*!< Do not display grid lines along primary value axis. */
        msoElementSecondaryCategoryAxisBillions = 378, /*!< Use billions for secondary category axis units. */
        msoElementSecondaryCategoryAxisLogScale = 379, /*!< Use log scale for secondary category axis. */
        msoElementSecondaryCategoryAxisMillions = 377, /*!< Use millions for secondary category axis units. */
        msoElementSecondaryCategoryAxisNone = 358, /*!< Do not display secondary category axis. */
        msoElementSecondaryCategoryAxisReverse = 361, /*!< Reverse secondary category axis. */
        msoElementSecondaryCategoryAxisShow = 359, /*!< Display secondary category axis. */
        msoElementSecondaryCategoryAxisThousands = 376, /*!< Use thousands for secondary category axis units. */
        msoElementSecondaryCategoryAxisTitleAdjacentToAxis = 313, /*!< Dispaly secondary category axis title adjacent to axis. */
        msoElementSecondaryCategoryAxisTitleBelowAxis = 314, /*!< Dispaly secondary category axis title below axis. */
        msoElementSecondaryCategoryAxisTitleHorizontal = 317, /*!< Display secondary category axis title horizontally. */
        msoElementSecondaryCategoryAxisTitleNone = 312, /*!< Do not display secondary category axis title. */
        msoElementSecondaryCategoryAxisTitleRotated = 315, /*!< Rotate secondary category axis title. */
        msoElementSecondaryCategoryAxisTitleVertical = 316, /*!< Display secondary category axis title vertically. */
        msoElementSecondaryCategoryAxisWithoutLabels = 360, /*!< Display secondary category axis without labels. */
        msoElementSecondaryCategoryGridLinesMajor = 342, /*!< Display major gridlines along secondary category axis. */
        msoElementSecondaryCategoryGridLinesMinor = 341, /*!< Display minor gridlines along secondary category axis. */
        msoElementSecondaryCategoryGridLinesMinorMajor = 343, /*!< Display both major and minor gridlines along secondary category axis. */
        msoElementSecondaryCategoryGridLinesNone = 340, /*!< Do not display grid lines along secondary category axis. */
        msoElementSecondaryValueAxisBillions = 366, /*!< Use billions for secondary value axis units. */
        msoElementSecondaryValueAxisLogScale = 367, /*!< Use log scale for secondary value axis. */
        msoElementSecondaryValueAxisMillions = 365, /*!< Use millions for secondary value axis units. */
        msoElementSecondaryValueAxisNone = 362, /*!< Do not display secondary value axis. */
        msoElementSecondaryValueAxisShow = 363, /*!< Display secondary value axis. */
        msoElementSecondaryValueAxisThousands = 364, /*!< Use thousands for secondary value axis units. */
        msoElementSecondaryValueAxisTitleAdjacentToAxis = 319, /*!< Display secondary value axis title adjacent to axis. */
        msoElementSecondaryValueAxisTitleBelowAxis = 320, /*!< Display secondary value axis title below axis. */
        msoElementSecondaryValueAxisTitleHorizontal = 323, /*!< Display secondary value axis title horizontally. */
        msoElementSecondaryValueAxisTitleNone = 318, /*!< Do not display secondary value axis title. */
        msoElementSecondaryValueAxisTitleRotated = 321, /*!< Rotate secondary value axis title. */
        msoElementSecondaryValueAxisTitleVertical = 322, /*!< Display secondary value axis title vertically. */
        msoElementSecondaryValueGridLinesMajor = 338, /*!< Display major gridlines along secondary value axis. */
        msoElementSecondaryValueGridLinesMinor = 337, /*!< Display minor gridlines along secondary value axis. */
        msoElementSecondaryValueGridLinesMinorMajor = 339, /*!< Display both major and minor gridlines along secondary value axis. */
        msoElementSecondaryValueGridLinesNone = 336, /*!< Do not display gridlines along secondary value axis. */
        msoElementSeriesAxisGridLinesMajor = 346, /*!< Display major gridlines along series axis. */
        msoElementSeriesAxisGridLinesMinor = 345, /*!< Display minor gridlines along series axis. */
        msoElementSeriesAxisGridLinesMinorMajor = 347, /*!< Display both major and minor gridlines along series axis. */
        msoElementSeriesAxisGridLinesNone = 344, /*!< Do not display gridlines along series axis. */
        msoElementSeriesAxisNone = 368, /*!< Do not display series axis. */
        msoElementSeriesAxisReverse = 371, /*!< Reverse series axis. */
        msoElementSeriesAxisShow = 369, /*!< Display series axis. */
        msoElementSeriesAxisTitleHorizontal = 327, /*!< Display series axis title horizontally. */
        msoElementSeriesAxisTitleNone = 324, /*!< Do not display series axis title. */
        msoElementSeriesAxisTitleRotated = 325, /*!< Rotate series axis title. */
        msoElementSeriesAxisTitleVertical = 326, /*!< Display series axis title vertically. */
        msoElementSeriesAxisWithoutLabeling = 370, /*!< Display series axis title without labeling. */
        msoElementTrendlineAddExponential = 602, /*!< Add an exponential trendline. */
        msoElementTrendlineAddLinear = 601, /*!< Add a linear trendline. */
        msoElementTrendlineAddLinearForecast = 603, /*!< Add a linear forecast. */
        msoElementTrendlineAddTwoPeriodMovingAverage = 604, /*!< Add a two-period moving average. */
        msoElementTrendlineNone = 600, /*!< Do not display trendline. */
        msoElementUpDownBarsNone = 900, /*!< Do not display up/down bars. */
        msoElementUpDownBarsShow = 901, /*!< Display up/down bars. */
    };


    /*! Specifies Clipboard formats. Since Excel 2007.

    [MSDN documentation for MsoClipboardFormat](http://msdn.microsoft.com/en-us/library/office/ff860826.aspx).
    */
    enum MsoClipboardFormat {
        msoClipboardFormatHTML = 2, /*!< HTML format. */
        msoClipboardFormatMixed = -2, /*!< Return value only; indicates a combination of the other states. */
        msoClipboardFormatNative = 1, /*!< Native format. */
        msoClipboardFormatPlainText = 4, /*!< Plain text format. */
        msoClipboardFormatRTF = 3, /*!< RTF format. */

    };


    /*!Specifies the color type.

    [MSDN documentation for MsoColorType](http://msdn.microsoft.com/en-us/library/office/aa432491.aspx).
    */
    enum MsoColorType {
        msoColorTypeCMS = 4 , /*!< Color Management System color type. */
        msoColorTypeCMYK = 3 , /*!< Color is determined by values of cyan, magenta, yellow, and black. */
        msoColorTypeInk = 5 , /*!< Not supported. */
        msoColorTypeMixed = -2 , /*!< Not supported. */
        msoColorTypeRGB = 1 , /*!< Color is determined by values of red, green, and blue. */
        msoColorTypeScheme = 2 , /*!< Color is defined by an application-specific scheme. */
    };

    /*!Specifies the connector type.

    [MSDN documentation for MsoConnectorType](http://msdn.microsoft.com/en-us/library/office/aa432496.aspx).
    */

    enum MsoConnectorType {
        msoConnectorCurve = 3,
        msoConnectorElbow = 2,
        msoConnectorStraight = 1,
        msoConnectorTypeMixed = -2,
    };

    /*!Specifies how to evenly distribute a collection of shapes. 

    [MSDN documentation for MsoDistributeCmd](http://msdn.microsoft.com/en-us/library/office/aa432507.aspx).
    */
    enum MsoDistributeCmd {
        msoDistributeHorizontally = 0 , /*!< Distribute horizontally. */
        msoDistributeVertically = 1 , /*!< Distribute vertically. */
    };

    
    /*!Specifies the data type for a document property.. 

    [MSDN documentation for MsoDocProperties](http://msdn.microsoft.com/en-us/library/office/ff864634%28v=office.14%29.aspx).
    */
    enum MsoDocProperties {
        msoPropertyTypeBoolean = 2, /*!< Boolean value. */
        msoPropertyTypeDate = 3, /*!< Date value. */
        msoPropertyTypeFloat = 5, /*!< Floating point value. */
        msoPropertyTypeNumber = 1, /*!< Integer value. */
        msoPropertyTypeString = 4, /*!< String value. */

    };

    /*!Specifies the editing type of a node.

    [MSDN documentation for MsoEditingType](http://msdn.microsoft.com/en-us/library/office/aa432510.aspx).
    */
    enum MsoEditingType {
        msoEditingAuto = 0 , /*!< Editing type is appropriate to the segments being connected. */
        msoEditingCorner = 1 , /*!< Corner node. */
        msoEditingSmooth = 2 , /*!< Smooth node. */
        msoEditingSymmetric = 3 , /*!< Symmetric node. */
    };


    /*!  @brief Specifies the document encoding (code page or character set) for the Web browser to use when a user views a saved document.

    [MSDN documentation for MsoEncoding](http://msdn.microsoft.com/en-us/library/aa432511.aspx).
    */
    enum MsoEncoding {
        msoEncodingArabic = 1256 , /*!< Arabic. */
        msoEncodingArabicASMO = 708 , /*!< Arabic ASMO. */
        msoEncodingArabicAutoDetect = 51256 , /*!< Web browser auto-detects type of Arabic encoding to use. */
        msoEncodingArabicTransparentASMO = 720 , /*!< Transparent Arabic. */
        msoEncodingAutoDetect = 50001 , /*!< Web browser auto-detects type of encoding to use. */
        msoEncodingBaltic = 1257 , /*!< Baltic. */
        msoEncodingCentralEuropean = 1250 , /*!< Central European. */
        msoEncodingCyrillic = 1251 , /*!< Cyrillic. */
        msoEncodingCyrillicAutoDetect = 51251 , /*!< Web browser auto-detects type of Cyrillic encoding to use. */
        msoEncodingEBCDICArabic = 20420 , /*!< Extended Binary Coded Decimal Interchange Code (EBCDIC) Arabic. */
        msoEncodingEBCDICDenmarkNorway = 20277 , /*!< EBCDIC as used in Denmark and Norway. */
        msoEncodingEBCDICFinlandSweden = 20278 , /*!< EBCDIC as used in Finland and Sweden. */
        msoEncodingEBCDICFrance = 20297 , /*!< EBCDIC as used in France. */
        msoEncodingEBCDICGermany = 20273 , /*!< EBCDIC as used in Germany. */
        msoEncodingEBCDICGreek = 20423 , /*!< EBCDIC as used in the Greek language. */
        msoEncodingEBCDICGreekModern = 875 , /*!< EBCDIC as used in the Modern Greek language. */
        msoEncodingEBCDICHebrew = 20424 , /*!< EBCDIC as used in the Hebrew language. */
        msoEncodingEBCDICIcelandic = 20871 , /*!< EBCDIC as used in Iceland. */
        msoEncodingEBCDICInternational = 500 , /*!< International EBCDIC. */
        msoEncodingEBCDICItaly = 20280 , /*!< EBCDIC as used in Italy. */
        msoEncodingEBCDICJapaneseKatakanaExtended = 20290 , /*!< EBCDIC as used with Japanese Katakana (extended). */
        msoEncodingEBCDICJapaneseKatakanaExtendedAndJapanese = 50930 , /*!< EBCDIC as used with Japanese Katakana (extended) and Japanese. */
        msoEncodingEBCDICJapaneseLatinExtendedAndJapanese = 50939 , /*!< EBCDIC as used with Japanese Latin (extended) and Japanese. */
        msoEncodingEBCDICKoreanExtended = 20833 , /*!< EBCDIC as used with Korean (extended). */
        msoEncodingEBCDICKoreanExtendedAndKorean = 50933 , /*!< EBCDIC as used with Korean (extended) and Korean. */
        msoEncodingEBCDICLatinAmericaSpain = 20284 , /*!< EBCDIC as used in Latin America and Spain. */
        msoEncodingEBCDICMultilingualROECELatin2 = 870 , /*!< EBCDIC Multilingual ROECE (Latin 2). */
        msoEncodingEBCDICRussian = 20880 , /*!< EBCDIC as used with Russian. */
        msoEncodingEBCDICSerbianBulgarian = 21025 , /*!< EBCDIC as used with Serbian and Bulgarian. */
        msoEncodingEBCDICSimplifiedChineseExtendedAndSimplifiedChinese = 50935 , /*!< EBCDIC as used with Simplified Chinese (extended) and Simplified Chinese. */
        msoEncodingEBCDICThai = 20838 , /*!< EBCDIC as used with Thai. */
        msoEncodingEBCDICTurkish = 20905 , /*!< EBCDIC as used with Turkish. */
        msoEncodingEBCDICTurkishLatin5 = 1026 , /*!< EBCDIC as used with Turkish (Latin 5). */
        msoEncodingEBCDICUnitedKingdom = 20285 , /*!< EBCDIC as used in the United Kingdom. */
        msoEncodingEBCDICUSCanada = 37 , /*!< EBCDIC as used in the United States and Canada. */
        msoEncodingEBCDICUSCanadaAndJapanese = 50931 , /*!< EBCDIC as used in the United States and Canada, and with Japanese. */
        msoEncodingEBCDICUSCanadaAndTraditionalChinese = 50937 , /*!< EBCDIC as used in the United States and Canada, and with Traditional Chinese. */
        msoEncodingEUCChineseSimplifiedChinese = 51936 , /*!< Extended Unix Code (EUC) as used with Chinese and Simplified Chinese. */
        msoEncodingEUCJapanese = 51932 , /*!< EUC as used with Japanese. */
        msoEncodingEUCKorean = 51949 , /*!< EUC as used with Korean. */
        msoEncodingEUCTaiwaneseTraditionalChinese = 51950 , /*!< EUC as used with Taiwanese and Traditional Chinese. */
        msoEncodingEuropa3 = 29001 , /*!< Europa. */
        msoEncodingExtAlphaLowercase = 21027 , /*!< Extended Alpha lowercase. */
        msoEncodingGreek = 1253 , /*!< Greek. */
        msoEncodingGreekAutoDetect = 51253 , /*!< Web browser auto-detects type of Greek encoding to use. */
        msoEncodingHebrew = 1255 , /*!< Hebrew. */
        msoEncodingHZGBSimplifiedChinese = 52936 , /*!< Simplified Chinese (HZGB). */
        msoEncodingIA5German = 20106 , /*!< German (International Alphabet No. 5, or IA5). */
        msoEncodingIA5IRV = 20105 , /*!< IA5, International Reference Version (IRV). */
        msoEncodingIA5Norwegian = 20108 , /*!< IA5 as used with Norwegian. */
        msoEncodingIA5Swedish = 20107 , /*!< IA5 as used with Swedish. */
        msoEncodingISCIIAssamese = 57006 , /*!< Indian Script Code for Information Interchange (ISCII) as used with Assamese. */
        msoEncodingISCIIBengali = 57003 , /*!< ISCII as used with Bengali. */
        msoEncodingISCIIDevanagari = 57002 , /*!< ISCII as used with Devanagari. */
        msoEncodingISCIIGujarati = 57010 , /*!< ISCII as used with Gujarati. */
        msoEncodingISCIIKannada = 57008 , /*!< ISCII as used with Kannada. */
        msoEncodingISCIIMalayalam = 57009 , /*!< ISCII as used with Malayalam. */
        msoEncodingISCIIOriya = 57007 , /*!< ISCII as used with Oriya. */
        msoEncodingISCIIPunjabi = 57011 , /*!< ISCII as used with Punjabi. */
        msoEncodingISCIITamil = 57004 , /*!< ISCII as used with Tamil. */
        msoEncodingISCIITelugu = 57005 , /*!< ISCII as used with Telugu. */
        msoEncodingISO2022CNSimplifiedChinese = 50229 , /*!< ISO 2022-CN encoding as used with Simplified Chinese. */
        msoEncodingISO2022CNTraditionalChinese = 50227 , /*!< ISO 2022-CN encoding as used with Traditional Chinese. */
        msoEncodingISO2022JPJISX02011989 = 50222 , /*!< ISO 2022-JP */
        msoEncodingISO2022JPJISX02021984 = 50221 , /*!< ISO 2022-JP */
        msoEncodingISO2022JPNoHalfwidthKatakana = 50220 , /*!< ISO 2022-JP with no half-width Katakana. */
        msoEncodingISO2022KR = 50225 , /*!< ISO 2022-KR. */
        msoEncodingISO6937NonSpacingAccent = 20269 , /*!< ISO 6937 Non-Spacing Accent. */
        msoEncodingISO885915Latin9 = 28605 , /*!< ISO 8859-15 with Latin 9. */
        msoEncodingISO88591Latin1 = 28591 , /*!< ISO 8859-1 Latin 1. */
        msoEncodingISO88592CentralEurope = 28592 , /*!< ISO 8859-2 Central Europe. */
        msoEncodingISO88593Latin3 = 28593 , /*!< ISO 8859-3 Latin 3. */
        msoEncodingISO88594Baltic = 28594 , /*!< ISO 8859-4 Baltic. */
        msoEncodingISO88595Cyrillic = 28595 , /*!< ISO 8859-5 Cyrillic. */
        msoEncodingISO88596Arabic = 28596 , /*!< ISA 8859-6 Arabic. */
        msoEncodingISO88597Greek = 28597 , /*!< ISO 8859-7 Greek. */
        msoEncodingISO88598Hebrew = 28598 , /*!< ISO 8859-8 Hebrew. */
        msoEncodingISO88598HebrewLogical = 38598 , /*!< ISO 8859-8 Hebrew (Logical). */
        msoEncodingISO88599Turkish = 28599 , /*!< ISO 8859-9 Turkish. */
        msoEncodingJapaneseAutoDetect = 50932 , /*!< Web browser auto-detects type of Japanese encoding to use. */
        msoEncodingJapaneseShiftJIS = 932 , /*!< Japanese (Shift-JIS). */
        msoEncodingKOI8R = 20866 , /*!< KOI8-R. */
        msoEncodingKOI8U = 21866 , /*!< K0I8-U. */
        msoEncodingKorean = 949 , /*!< Korean. */
        msoEncodingKoreanAutoDetect = 50949 , /*!< Web browser auto-detects type of Korean encoding to use. */
        msoEncodingKoreanJohab = 1361 , /*!< Korean (Johab). */
        msoEncodingMacArabic = 10004 , /*!< Macintosh Arabic. */
        msoEncodingMacCroatia = 10082 , /*!< Macintosh Croatian. */
        msoEncodingMacCyrillic = 10007 , /*!< Macintosh Cyrillic. */
        msoEncodingMacGreek1 = 10006 , /*!< Macintosh Greek. */
        msoEncodingMacHebrew = 10005 , /*!< Macintosh Hebrew. */
        msoEncodingMacIcelandic = 10079 , /*!< Macintosh Icelandic. */
        msoEncodingMacJapanese = 10001 , /*!< Macintosh Japanese. */
        msoEncodingMacKorean = 10003 , /*!< Macintosh Korean. */
        msoEncodingMacLatin2 = 10029 , /*!< Macintosh Latin 2. */
        msoEncodingMacRoman = 10000 , /*!< Macintosh Roman. */
        msoEncodingMacRomania = 10010 , /*!< Macintosh Romanian. */
        msoEncodingMacSimplifiedChineseGB2312 = 10008 , /*!< Macintosh Simplified Chinese (GB 2312). */
        msoEncodingMacTraditionalChineseBig5 = 10002 , /*!< Macintosh Traditional Chinese (Big 5). */
        msoEncodingMacTurkish = 10081 , /*!< Macintosh Turkish. */
        msoEncodingMacUkraine = 10017 , /*!< Macintosh Ukrainian. */
        msoEncodingOEMArabic = 864 , /*!< OEM as used with Arabic. */
        msoEncodingOEMBaltic = 775 , /*!< OEM as used with Baltic. */
        msoEncodingOEMCanadianFrench = 863 , /*!< OEM as used with Canadian French. */
        msoEncodingOEMCyrillic = 855 , /*!< OEM as used with Cyrillic. */
        msoEncodingOEMCyrillicII = 866 , /*!< OEM as used with Cyrillic II. */
        msoEncodingOEMGreek437G = 737 , /*!< OEM as used with Greek 437G. */
        msoEncodingOEMHebrew = 862 , /*!< OEM as used with Hebrew. */
        msoEncodingOEMIcelandic = 861 , /*!< OEM as used with Icelandic. */
        msoEncodingOEMModernGreek = 869 , /*!< OEM as used with Modern Greek. */
        msoEncodingOEMMultilingualLatinI = 850 , /*!< OEM as used with multi-lingual Latin I. */
        msoEncodingOEMMultilingualLatinII = 852 , /*!< OEM as used with multi-lingual Latin II. */
        msoEncodingOEMNordic = 865 , /*!< OEM as used with Nordic languages. */
        msoEncodingOEMPortuguese = 860 , /*!< OEM as used with Portuguese. */
        msoEncodingOEMTurkish = 857 , /*!< OEM as used with Turkish. */
        msoEncodingOEMUnitedStates = 437 , /*!< OEM as used in the United States. */
        msoEncodingSimplifiedChineseAutoDetect = 50936 , /*!< Web browser auto-detects type of Simplified Chinese encoding to use. */
        msoEncodingSimplifiedChineseGB18030 = 54936 , /*!< Simplified Chinese GB 18030. */
        msoEncodingSimplifiedChineseGBK = 936 , /*!< Simplified Chinese GBK. */
        msoEncodingT61 = 20261 , /*!< T61. */
        msoEncodingTaiwanCNS = 20000 , /*!< Taiwan CNS. */
        msoEncodingTaiwanEten = 20002 , /*!< Taiwan Eten. */
        msoEncodingTaiwanIBM5550 = 20003 , /*!< Taiwan IBM 5550. */
        msoEncodingTaiwanTCA = 20001 , /*!< Taiwan TCA. */
        msoEncodingTaiwanTeleText = 20004 , /*!< Taiwan Teletext. */
        msoEncodingTaiwanWang = 20005 , /*!< Taiwan Wang. */
        msoEncodingThai = 874 , /*!< Thai. */
        msoEncodingTraditionalChineseAutoDetect = 50950 , /*!< Web browser auto-detects type of Traditional Chinese encoding to use. */
        msoEncodingTraditionalChineseBig5 = 950 , /*!< Traditional Chinese Big 5. */
        msoEncodingTurkish = 1254 , /*!< Turkish. */
        msoEncodingUnicodeBigEndian = 1201 , /*!< Unicode big endian. */
        msoEncodingUnicodeLittleEndian = 1200 , /*!< Unicode little endian. */
        msoEncodingUSASCII = 20127 , /*!< United States ASCII. */
        msoEncodingUTF7 = 65000 , /*!< UTF-7 encoding. */
        msoEncodingUTF8 = 65001 , /*!< UTF-8 encoding. */
        msoEncodingVietnamese = 1258 , /*!< Vietnamese. */
        msoEncodingWestern = 1252 , /*!< Western. */
    };


    /*!Specifies how to use the value specified in the ExtraInfo property of the FollowHyperlink method.

    [MSDN documentation for MsoExtraInfoMethod](http://msdn.microsoft.com/en-us/library/office/aa432516.aspx).
    */
    enum MsoExtraInfoMethod {
        msoMethodGet = 0 , /*!< The value specified in the ExtraInfo property is a string that is appended to the address. */
        msoMethodPost = 1 , /*!< The value specified in the ExtraInfo property is posted as a string or byte array. */
    };

    /*!Specifies whether the extrusion color is based on the extruded shape's fill (the front face of the extrusion) and automatically changes when the shape's fill changes, or whether the extrusion color is independent of the shape's fill. Used with the 

    [MSDN documentation for MsoExtrusionColorType](http://msdn.microsoft.com/en-us/library/office/aa432517.aspx).
    */
    enum MsoExtrusionColorType {
        msoExtrusionColorAutomatic = 1 , /*!< Extrusion color is based on shape fill. */
        msoExtrusionColorCustom = 2 , /*!< Extrusion color is independent of shape fill. */
        msoExtrusionColorTypeMixed = -2 , /*!< Return value only; indicates a combination of the other states.  */
    };

    /*!  @brief Specifies how the application handles calls to methods and properties that require features not yet installed.

    [MSDN documentation for MsoFeatureInstall](http://msdn.microsoft.com/en-us/library/aa432519.aspx).
    */
    enum MsoFeatureInstall {
        msoFeatureInstallNone = 0, /*!< Generates a generic automation error at run time when uninstalled features are called. */
        msoFeatureInstallOnDemand = 1, /*!< Prompts the user to install new features. */
        msoFeatureInstallOnDemandWithUI = 2, /*!< Displays a progress meter during installation; does not prompt the user to install new features. */
    };


    /*!  @brief Specifies the file validation mode. Since Excel 2010

    [MSDN documentation for MsoFileValidationMode](http://msdn.microsoft.com/en-us/library/office/ff860813%28v=office.14%29.aspx).
    */
    enum MsoFileValidationMode {
        msoFileValidationDefault = 0, /*!< Validate the file (default). */
        msoFileValidationSkip = 1, /*!< Do not validate the file. */
    };

    /*!Specifies a shape's fill type.

    [MSDN documentation for MsoFillType](http://msdn.microsoft.com/en-us/library/office/aa432529.aspx).
    */
    enum MsoFillType {
        msoFillBackground = 5 , /*!< Fill is the same as the background. */
        msoFillGradient = 3 , /*!< Gradient fill. */
        msoFillMixed = -2 , /*!< Mixed fill. */
        msoFillPatterned = 2 , /*!< Patterned fill. */
        msoFillPicture = 6 , /*!< Picture fill. */
        msoFillSolid = 1 , /*!< Solid fill. */
        msoFillTextured = 4 , /*!< Textured fill. */
    };

    /*!Specifies whether a shape should be flipped horizontally or vertically.

    [MSDN documentation for MsoFlipCmd](http://msdn.microsoft.com/en-us/library/office/aa432532.aspx).
    */
    enum MsoFlipCmd {
        msoFlipHorizontal = 0 , /*!< Flip horizontally. */
        msoFlipVertical = 1 , /*!< Flip vertically. */
    };

    /*!Specifies the type of gradient used in a shape's fill.

    [MSDN documentation for MsoGradientColorType](http://msdn.microsoft.com/en-us/library/office/aa432534.aspx).
    */
    enum MsoGradientColorType {
        msoGradientColorMixed = -2 , /*!< Mixed gradient. */
        msoGradientOneColor = 1 , /*!< One-color gradient. */
        msoGradientPresetColors = 3 , /*!< Gradient colors set according to a built-in gradient of the set defined by the msoPresetGradientType constant. */
        msoGradientTwoColors = 2 , /*!< Two-color gradient. */
    };

    /*!  @brief Specifies the style for a gradient fill.

    [MSDN documentation for MsoFeatureInstall](http://msdn.microsoft.com/en-us/library/aa432535.aspx).
    */
    enum MsoGradientStyle {
        msoGradientDiagonalDown = 4, /*!<  Diagonal gradient moving from a top corner down to the opposite corner. */
        msoGradientDiagonalUp = 3, /*!<  Diagonal gradient moving from a bottom corner up to the opposite corner. */
        msoGradientFromCenter = 7, /*!<  Gradient running from the center out to the corners. */
        msoGradientFromCorner = 5, /*!<  Gradient running from a corner to the other three corners. */
        msoGradientFromTitle = 6, /*!<  Gradient running from the title outward. */
        msoGradientHorizontal = 1, /*!<  Gradient running horizontally across the shape. */
        msoGradientMixed = -2, /*!<  Gradient is mixed. */
        msoGradientVertical = 2, /*!<  Gradient running vertically down the shape. */
    };

    /*!Specifies the horizontal alignment of text in a text frame. Used with the 

    [MSDN documentation for MsoHorizontalAnchor](http://msdn.microsoft.com/en-us/library/office/aa432537.aspx).
    */
    enum MsoHorizontalAnchor {
        msoAnchorCenter = 2 , /*!< Text is centered horizontally. */
        msoAnchorNone = 1 , /*!< No alignment. */
        msoHorizontalAnchorMixed = -2 , /*!< Return value only; indicates a combination of the other states. */
    };

    /*!Specifies the type of hyperlink.

    [MSDN documentation for MsoHyperlinkType](http://msdn.microsoft.com/en-us/library/office/aa432538.aspx).
    */
    enum MsoHyperlinkType {
        msoHyperlinkInlineShape = 2 , /*!< Hyperlink applies to an inline shape. Used only with Microsoft Word. */
        msoHyperlinkRange = 0 , /*!< Hyperlink applies to a Range object. */
        msoHyperlinkShape = 1 , /*!< Hyperlink applies to a Shape object. */
    };    

    /*!  @brief Specifies the language identifier.

    [MSDN documentation for MsoLanguageID](http://msdn.microsoft.com/en-us/library/aa432635.aspx).
    */

    enum MsoLanguageID
    {
        msoLanguageIDAfrikaans = 1078, /*!< The Afrikaans language. */
        msoLanguageIDAlbanian = 1052, /*!< The Albanian language. */
        msoLanguageIDAmharic = 1118, /*!< The Amharic language. */
        msoLanguageIDArabic = 1025, /*!< The Arabic language. */
        msoLanguageIDArabicAlgeria = 5121, /*!< The Arabic Algeria language. */
        msoLanguageIDArabicBahrain = 15361, /*!< The Arabic Bahrain language. */
        msoLanguageIDArabicEgypt = 3073, /*!< The Arabic Egypt language. */
        msoLanguageIDArabicIraq = 2049, /*!< The Arabic Iraq language. */
        msoLanguageIDArabicJordan = 11265, /*!< The Arabic Jordan language. */
        msoLanguageIDArabicKuwait = 13313, /*!< The Arabic Kuwait language. */
        msoLanguageIDArabicLebanon = 12289, /*!< The Arabic Lebanon language. */
        msoLanguageIDArabicLibya = 4097, /*!< The Arabic Libya language. */
        msoLanguageIDArabicMorocco = 6145, /*!< The Arabic Morocco language. */
        msoLanguageIDArabicOman = 8193, /*!< The Arabic Oman language. */
        msoLanguageIDArabicQatar = 16385, /*!< The Arabic Qatar language. */
        msoLanguageIDArabicSyria = 10241, /*!< The Arabic Syria language. */
        msoLanguageIDArabicTunisia = 7169, /*!< The Arabic Tunisia language. */
        msoLanguageIDArabicUAE = 14337, /*!< The Arabic UAE language. */
        msoLanguageIDArabicYemen = 9217, /*!< The Arabic Yemen language. */
        msoLanguageIDArmenian = 1067, /*!< The Armenian language. */
        msoLanguageIDAssamese = 1101, /*!< The Assamese language. */
        msoLanguageIDAzeriCyrillic = 2092, /*!< The Azeri Cyrillic language. */
        msoLanguageIDAzeriLatin = 1068, /*!< The Azeri Latin language. */
        msoLanguageIDBasque = 1069, /*!< The Basque language. */
        msoLanguageIDBelgianDutch = 2067, /*!< The Belgian Dutch language. */
        msoLanguageIDBelgianFrench = 2060, /*!< The Belgian French language. */
        msoLanguageIDBengali = 1093, /*!< The Bengali language. */
        msoLanguageIDBosnian = 4122, /*!< The Bosnian language. */
        msoLanguageIDBosnianBosniaHerzegovinaCyrillic = 8218, /*!< The Bosnian Bosnia Herzegovina Cyrillic language. */
        msoLanguageIDBosnianBosniaHerzegovinaLatin = 5146, /*!< The Bosnian Bosnia Herzegovina Latin language. */
        msoLanguageIDBrazilianPortuguese = 1046, /*!< The Brazilian Portuguese language. */
        msoLanguageIDBulgarian = 1026, /*!< The Bulgarian language. */
        msoLanguageIDBurmese = 1109, /*!< The Burmese language. */
        msoLanguageIDByelorussian = 1059, /*!< The Byelorussian language. */
        msoLanguageIDCatalan = 1027, /*!< The Catalan language. */
        msoLanguageIDCherokee = 1116, /*!< The Cherokee language. */
        msoLanguageIDChineseHongKongSAR = 3076, /*!< The Chinese Hong Kong SAR language. */
        msoLanguageIDChineseMacaoSAR = 5124, /*!< The Chinese Macao SAR language. */
        msoLanguageIDChineseSingapore = 4100, /*!< The Chinese Singapore language. */
        msoLanguageIDCroatian = 1050, /*!< The Croatian language. */
        msoLanguageIDCzech = 1029, /*!< The Czech language. */
        msoLanguageIDDanish = 1030, /*!< The Danish language. */
        msoLanguageIDDivehi = 1125, /*!< The Divehi language. */
        msoLanguageIDDutch = 1043, /*!< The Dutch language. */
        msoLanguageIDEdo = 1126, /*!< The Edo language. */
        msoLanguageIDEnglishAUS = 3081, /*!< The English AUS language. */
        msoLanguageIDEnglishBelize = 10249, /*!< The English Belize language. */
        msoLanguageIDEnglishCanadian = 4105, /*!< The English Canadian language. */
        msoLanguageIDEnglishCaribbean = 9225, /*!< The English Caribbean language. */
        msoLanguageIDEnglishIndonesia = 14345, /*!< The English Indonesia language. */
        msoLanguageIDEnglishIreland = 6153, /*!< The English Ireland language. */
        msoLanguageIDEnglishJamaica = 8201, /*!< The English Jamaica language. */
        msoLanguageIDEnglishNewZealand = 5129, /*!< The English NewZealand language. */
        msoLanguageIDEnglishPhilippines = 13321, /*!< The English Philippines language. */
        msoLanguageIDEnglishSouthAfrica = 7177, /*!< The English South Africa language. */
        msoLanguageIDEnglishTrinidadTobago = 11273, /*!< The English Trinidad Tobago language. */
        msoLanguageIDEnglishUK = 2057, /*!< The English UK language. */
        msoLanguageIDEnglishUS = 1033, /*!< The English US language. */
        msoLanguageIDEnglishZimbabwe = 12297, /*!< The English Zimbabwe language. */
        msoLanguageIDEstonian = 1061, /*!< The Estonian language. */
        msoLanguageIDFaeroese = 1080, /*!< The Faeroese language. */
        msoLanguageIDFarsi = 1065, /*!< The Farsi language. */
        msoLanguageIDFilipino = 1124, /*!< The Filipino language. */
        msoLanguageIDFinnish = 1035, /*!< The Finnish language. */
        msoLanguageIDFrench = 1036, /*!< The French language. */
        msoLanguageIDFrenchCameroon = 11276, /*!< The French Cameroon language. */
        msoLanguageIDFrenchCanadian = 3084, /*!< The French Canadian language. */
        msoLanguageIDFrenchCotedIvoire = 12300, /*!< The French Coted Ivoire language. */
        msoLanguageIDFrenchHaiti = 15372, /*!< The French Haiti language. */
        msoLanguageIDFrenchLuxembourg = 5132, /*!< The French Luxembourg language. */
        msoLanguageIDFrenchMali = 13324, /*!< The French Mali language. */
        msoLanguageIDFrenchMonaco = 6156, /*!< The French Monaco language. */
        msoLanguageIDFrenchMorocco = 14348, /*!< The French Morocco language. */
        msoLanguageIDFrenchReunion = 8204, /*!< The French Reunion language. */
        msoLanguageIDFrenchSenegal = 10252, /*!< The French Senegal language. */
        msoLanguageIDFrenchWestIndies = 7180, /*!< The French West Indies language. */
        msoLanguageIDFranchCongoDRC = 9228, /*!< The French Congo DRC language. */
        msoLanguageIDFrisianNetherlands = 1122, /*!< The Frisian Netherlands language. */
        msoLanguageIDFulfulde = 1127, /*!< The Fulfulde language. */
        msoLanguageIDGaelicIreland = 2108, /*!< The Gaelic Ireland language. */
        msoLanguageIDGaelicScotland = 1084, /*!< The Gaelic Scotland language. */
        msoLanguageIDGalician = 1110, /*!< The Galician language. */
        msoLanguageIDGeorgian = 1079, /*!< The Georgian language. */
        msoLanguageIDGerman = 1031, /*!< The German language. */
        msoLanguageIDGermanAustria = 3079, /*!< The German Austria language. */
        msoLanguageIDGermanLiechtenstein = 5127, /*!< The German Liechtenstein language. */
        msoLanguageIDGermanLuxembourg = 4103, /*!< The German Luxembourg language. */
        msoLanguageIDGreek = 1032, /*!< The Greek language. */
        msoLanguageIDGuarani = 1140, /*!< The Guarani language. */
        msoLanguageIDGujarati = 1095, /*!< The Gujarati language. */
        msoLanguageIDHausa = 1128, /*!< The Hausa language. */
        msoLanguageIDHawaiian = 1141, /*!< The Hawaiian language. */
        msoLanguageIDHebrew = 1037, /*!< The Hebrew language. */
        msoLanguageIDHindi = 1081, /*!< The Hindi language. */
        msoLanguageIDHungarian = 1038, /*!< The Hungarian language. */
        msoLanguageIDIbibio = 1129, /*!< The Ibibio language. */
        msoLanguageIDIcelandic = 1039, /*!< The Icelandic language. */
        msoLanguageIDIgbo = 1136, /*!< The Igbo language. */
        msoLanguageIDIndonesian = 1057, /*!< The Indonesian language. */
        msoLanguageIDInuktitut = 1117, /*!< The Inuktitut language. */
        msoLanguageIDItalian = 1040, /*!< The Italian language. */
        msoLanguageIDJapanese = 1041, /*!< The Japanese language. */
        msoLanguageIDKannada = 1099, /*!< The Kannada language. */
        msoLanguageIDKanuri = 1137, /*!< The Kanuri language. */
        msoLanguageIDKashmiri = 1120, /*!< The Kashmiri language. */
        msoLanguageIDKashmiriDevanagari = 2144, /*!< The Kashmiri Devanagari language. */
        msoLanguageIDKazakh = 1087, /*!< The Kazakh language. */
        msoLanguageIDKhmer = 1107, /*!< The Khmer language. */
        msoLanguageIDKirghiz = 1088, /*!< The Kirghiz language. */
        msoLanguageIDKonkani = 1111, /*!< The Konkani language. */
        msoLanguageIDKorean = 1042, /*!< The Korean language. */
        msoLanguageIDKyrgyz = 1088, /*!< The Kyrgyz language. */
        msoLanguageIDLao = 1108, /*!< The Lao language. */
        msoLanguageIDLatin = 1142, /*!< The Latin language. */
        msoLanguageIDLatvian = 1062, /*!< The Latvian language. */
        msoLanguageIDLithuanian = 1063, /*!< The Lithuanian language. */
        msoLanguageIDMacedoninanFYROM = 1071, /*!< The Macedonian FYROM language. */
        msoLanguageIDMalayalam = 1100, /*!< The Malayalam language. */
        msoLanguageIDMalayBruneiDarussalam = 2110, /*!< The Malay Brunei Darussalam language. */
        msoLanguageIDMalaysian = 1086, /*!< The Malaysian language. */
        msoLanguageIDMaltese = 1082, /*!< The Maltese language. */
        msoLanguageIDManipuri = 1112, /*!< The Manipuri language. */
        msoLanguageIDMaori = 1153, /*!< The Maori language. */
        msoLanguageIDMarathi = 1102, /*!< The Marathi language. */
        msoLanguageIDMexicanSpanish = 2058, /*!< The Mexican Spanish language. */
        msoLanguageIDMixed = -2, /*!< The Mixed language. */
        msoLanguageIDMongolian = 1104, /*!< The Mongolian language. */
        msoLanguageIDNepali = 1121, /*!< The Nepali language. */
        msoLanguageIDNone = 0, /*!< No language specified. */
        msoLanguageIDNoProofing = 1024, /*!< No proofing. */
        msoLanguageIDNorwegianBokmol = 1044, /*!< The Norwegian Bokmol language. */
        msoLanguageIDNorwegianNynorsk = 2068, /*!< The Norwegian Nynorsk language. */
        msoLanguageIDOriya = 1096, /*!< The Oriya language. */
        msoLanguageIDOromo = 1138, /*!< The Oromo language. */
        msoLanguageIDPashto = 1123, /*!< The Pashto language. */
        msoLanguageIDPolish = 1045, /*!< The Polish language. */
        msoLanguageIDPortuguese = 2070, /*!< The Portuguese language. */
        msoLanguageIDPunjabi = 1094, /*!< The Punjabi language. */
        msoLanguageIDQuechuaBolivia = 1131, /*!< The Quechua Bolivia language. */
        msoLanguageIDQuechuaEcuador = 2155, /*!< The Quechua Ecuador language. */
        msoLanguageIDQuechuaPeru = 3179, /*!< The Quechua Peru language. */
        msoLanguageIDRhaetoRomanic = 1047, /*!< The Rhaeto Romanic language. */
        msoLanguageIDRomanian = 1048, /*!< The Romanian language. */
        msoLanguageIDRomanianMoldova = 2072, /*!< The Romanian Moldova language. */
        msoLanguageIDRussian = 1049, /*!< The Russian language. */
        msoLanguageIDRussianMoldova = 2073, /*!< The Russian Moldova language. */
        msoLanguageIDSamiLappish = 1083, /*!< The Sami Lappish language. */
        msoLanguageIDSanskrit = 1103, /*!< The Sanskrit language. */
        msoLanguageIDSepedi = 1132, /*!< The Sepedi language. */
        msoLanguageIDSerbianBosniaHerzegovinaCyrillic = 7194, /*!< The Serbian Bosnia Herzegovina Cyrillic language. */
        msoLanguageIDSerbianBosniaHerzegovinaLatin = 6170, /*!< The Serbian Bosnia Herzegovina Latin language. */
        msoLanguageIDSerbianCyrillic = 3098, /*!< The Serbian Cyrillic language. */
        msoLanguageIDSerbianLatin = 2074, /*!< The Serbian Latin language. */
        msoLanguageIDSesotho = 1072, /*!< The Sesotho language. */
        msoLanguageIDSimplifiedChinese = 2052, /*!< The Simplified Chinese language. */
        msoLanguageIDSindhi = 1113, /*!< The Sindhi language. */
        msoLanguageIDSindhiPakistan = 2137, /*!< The Sindhi Pakistan language. */
        msoLanguageIDSinhalese = 1115, /*!< The Sinhalese language. */
        msoLanguageIDSlovak = 1051, /*!< The Slovak language. */
        msoLanguageIDSlovenian = 1060, /*!< The Slovenian language. */
        msoLanguageIDSomali = 1143, /*!< The Somali language. */
        msoLanguageIDSorbian = 1070, /*!< The Sorbian language. */
        msoLanguageIDSpanish = 1034, /*!< The Spanish language. */
        msoLanguageIDSpanishArgentina = 11274, /*!< The Spanish Argentina language. */
        msoLanguageIDSpanishBolivia = 16394, /*!< The Spanish Bolivia language. */
        msoLanguageIDSpanishChile = 13322, /*!< The Spanish Chile language. */
        msoLanguageIDSpanishColombia = 9226, /*!< The Spanish Colombia language. */
        msoLanguageIDSpanishCostaRica = 5130, /*!< The Spanish Costa Rica language. */
        msoLanguageIDSpanishDominicanRepublic = 7178, /*!< The Spanish Dominican Republic language. */
        msoLanguageIDSpanishEcuador = 12298, /*!< The Spanish Ecuador language. */
        msoLanguageIDSpanishElSalvador = 17418, /*!< The Spanish El Salvador language. */
        msoLanguageIDSpanishGuatemala = 4106, /*!< The Spanish Guatemala language. */
        msoLanguageIDSpanishHonduras = 18442, /*!< The Spanish Honduras language. */
        msoLanguageIDSpanishModernSort = 3082, /*!< The Spanish Modern Sort language. */
        msoLanguageIDSpanishNicaragua = 19466, /*!< The Spanish Nicaragua language. */
        msoLanguageIDSpanishPanama = 6154, /*!< The Spanish Panama language. */
        msoLanguageIDSpanishParaguay = 15370, /*!< The Spanish Paraguay language. */
        msoLanguageIDSpanishPeru = 10250, /*!< The Spanish Peru language. */
        msoLanguageIDSpanishPuertoRico = 20490, /*!< The Spanish Puerto Rico language. */
        msoLanguageIDSpanishUruguay = 14346, /*!< The Spanish Uruguay language. */
        msoLanguageIDSpanishVenezuela = 8202, /*!< The Spanish Venezuela language. */
        msoLanguageIDSutu = 1072, /*!< The Sutu language. */
        msoLanguageIDSwahili = 1089, /*!< The Swahili language. */
        msoLanguageIDSwedish = 1053, /*!< The Swedish language. */
        msoLanguageIDSwedishFinland = 2077, /*!< The Swedish Finland language. */
        msoLanguageIDSwissFrench = 4108, /*!< The Swiss French language. */
        msoLanguageIDSwissGerman = 2055, /*!< The Swiss German language. */
        msoLanguageIDSwissItalian = 2064, /*!< The Swiss Italian language. */
        msoLanguageIDSyriac = 1114, /*!< The Syriac language. */
        msoLanguageIDTajik = 1064, /*!< The Tajik language. */
        msoLanguageIDTamazight = 1119, /*!< The Tamazight language. */
        msoLanguageIDTamazightLatin = 2143, /*!< The Tamazight Latin language. */
        msoLanguageIDTamil = 1097, /*!< The Tamil language. */
        msoLanguageIDTatar = 1092, /*!< The Tatar language. */
        msoLanguageIDTelugu = 1098, /*!< The Telugu language. */
        msoLanguageIDThai = 1054, /*!< The Thai language. */
        msoLanguageIDTibetan = 1105, /*!< The Tibetan language. */
        msoLanguageIDTigrignaEritrea = 2163, /*!< The Tigrigna Eritrea language. */
        msoLanguageIDTigrignaEthiopic = 1139, /*!< The Tigrigna Ethiopic language. */
        msoLanguageIDTraditionalChinese = 1028, /*!< The Traditional Chinese language. */
        msoLanguageIDTsonga = 1073, /*!< The Tsonga language. */
        msoLanguageIDTswana = 1074, /*!< The Tswana language. */
        msoLanguageIDTurkish = 1055, /*!< The Turkish language. */
        msoLanguageIDTurkmen = 1090, /*!< The Turkmen language. */
        msoLanguageIDUkrainian = 1058, /*!< The Ukrainian language. */
        msoLanguageIDUrdu = 1056, /*!< The Urdu language. */
        msoLanguageIDUzbekCyrillic = 2115, /*!< The Uzbek Cyrillic language. */
        msoLanguageIDUzbekLatin = 1091, /*!< The Uzbek Latin language. */
        msoLanguageIDVenda = 1075, /*!< The Venda language. */
        msoLanguageIDVietnamese = 1066, /*!< The Vietnamese language. */
        msoLanguageIDWelsh = 1106, /*!< The Welsh language. */
        msoLanguageIDXhosa = 1076, /*!< The Xhosa language. */
        msoLanguageIDYi = 1144, /*!< The Yi language. */
        msoLanguageIDYiddish = 1085, /*!< The Yiddish language. */
        msoLanguageIDYoruba = 1130, /*!< The Yoruba language. */
        msoLanguageIDZulu = 1077, /*!< The Zulu language. */
    };


    /*!Indicates the effects lighting for an object.

    [MSDN documentation for MsoLightRigType](http://msdn.microsoft.com/en-us/library/office/aa432638.aspx).
    */
    enum MsoLightRigType {
        msoLightRigBalanced = 14 , /*!< Specifies the Balanced effect. */
        msoLightRigBrightRoom = 27 , /*!< Specifies the BrightRoom effect. */
        msoLightRigChilly = 22 , /*!< Specifies the Chilly effect. */
        msoLightRigContrasting = 18 , /*!< Specifies the Contrasting effect. */
        msoLightRigFlat = 24 , /*!< Specifies the Flat effect. */
        msoLightRigFlood = 17 , /*!< Specifies the Flood effect. */
        msoLightRigFreezing = 23 , /*!< Specifies the Freezing effect. */
        msoLightRigGlow = 26 , /*!< Specifies the Glow effect. */
        msoLightRigHarsh = 16 , /*!< Specifies the Harsh effect. */
        msoLightRigLegacyFlat1 = 1 , /*!< Specifies the LegacyFlat1 effect. */
        msoLightRigLegacyFlat2 = 2 , /*!< Specifies the LegacyFlat2 effect. */
        msoLightRigLegacyFlat3 = 3 , /*!< Specifies the LegacyFlat3 effect. */
        msoLightRigLegacyFlat4 = 4 , /*!< Specifies the LegacyFlat4 effect. */
        msoLightRigLegacyHarsh1 = 9 , /*!< Specifies the LegacyHarsh1 effect. */
        msoLightRigLegacyHarsh2 = 10 , /*!< Specifies the LegacyHarsh2 effect. */
        msoLightRigLegacyHarsh3 = 11 , /*!< Specifies the LegacyHarsh3 effect. */
        msoLightRigLegacyHarsh4 = 12 , /*!< Specifies the LegacyHarsh4 effect. */
        msoLightRigLegacyNormal1 = 5 , /*!< Specifies the LegacyNormal1 effect. */
        msoLightRigLegacyNormal2 = 6 , /*!< Specifies the LegacyNormal2 effect. */
        msoLightRigLegacyNormal3 = 7 , /*!< Specifies the LegacyNormal3 effect. */
        msoLightRigLegacyNormal4 = 8 , /*!< Specifies the LegacyNormal4 effect. */
        msoLightRigMixed = -2 , /*!< Specifies the Mixed effect. */
        msoLightRigMorning = 19 , /*!< Specifies the Morning effect. */
        msoLightRigSoft = 15 , /*!< Specifies the Soft effect. */
        msoLightRigSunrise = 20 , /*!< Specifies the Sunrise effect. */
        msoLightRigSunset = 21 , /*!< Specifies the Sunset effect. */
        msoLightRigThreePoint = 13 , /*!< Specifies the ThreePoint effect. */
        msoLightRigTwoPoint = 25 , /*!< Specifies the TwoPoint effect. */
    };

    /*!Specifies the dash style for a line. 

    [MSDN documentation for MsoLineDashStyle](http://msdn.microsoft.com/en-us/library/office/aa432639.aspx).
    */
    enum MsoLineDashStyle {
        msoLineDash = 4 , /*!< Line consists of dashes only. */
        msoLineDashDot = 5 , /*!< Line is a dash-dot pattern. */
        msoLineDashDotDot = 6 , /*!< Line is a dash-dot-dot pattern. */
        msoLineDashStyleMixed = -2 , /*!< Not supported. */
        msoLineLongDash = 7 , /*!< Line consists of long dashes. */
        msoLineLongDashDot = 8 , /*!< Line is a long dash-dot pattern. */
        msoLineRoundDot = 3 , /*!< Line is made up of round dots. */
        msoLineSolid = 1 , /*!< Line is solid. */
        msoLineSquareDot = 2 , /*!< Line is made up of square dots. */
    };
    /*!Specifies the style for a line.

    [MSDN documentation for MsoLineStyle](http://msdn.microsoft.com/en-us/library/office/aa432640.aspx).
    */
    enum MsoLineStyle {
        msoLineSingle = 1 , /*!< Single line. */
        msoLineStyleMixed = -2 , /*!< Not supported. */
        msoLineThickBetweenThin = 5 , /*!< Thick line with a thin line on each side. */
        msoLineThickThin = 4 , /*!< Thick line next to thin line. For horizontal lines, thick line is above thin line. For vertical lines, thick line is to the left of the thin line. */
        msoLineThinThick = 3 , /*!< Thick line next to thin line. For horizontal lines, thick line is below thin line. For vertical lines, thick line is to the right of the thin line. */
        msoLineThinThin = 2 , /*!< Two thin lines. */
    };

    /*!Specifies numbered bullet styles.

    [MSDN documentation for MsoNumberedBulletStyle](http://msdn.microsoft.com/en-us/library/office/ff860574.aspx).
    */
    enum MsoNumberedBulletStyle {
        msoBulletAlphaLCParenBoth = 8, /*!< Lowercase alphabetical bullet with opening and closing parentheses. */
        msoBulletAlphaLCParenRight = 9, /*!< Lowercase alphabetical bullet with closing parenthesis. */
        msoBulletAlphaLCPeriod = 0, /*!< Lowercase alphabetical bullet with period. */
        msoBulletAlphaUCParenBoth = 10, /*!< Uppercase alphabetical bullet with opening and closing parentheses. */
        msoBulletAlphaUCParenRight = 11, /*!< Uppercase alphabetical bullet with closing parenthesis. */
        msoBulletAlphaUCPeriod = 1, /*!< Uppercase alphabetical bullet with period. */
        msoBulletArabicAbjadDash = 24, /*!< Arabic Abjad bullet with a dash. */
        msoBulletArabicAlphaDash = 23, /*!< Arabic alphabetical bullet with a dash. */
        msoBulletArabicDBPeriod = 29, /*!< Arabic DB bullet with period. */
        msoBulletArabicDBPlain = 28, /*!< Plain Arabic DB bullet. */
        msoBulletArabicParenBoth = 12, /*!< Arabic bullet with opening and closing parentheses. */
        msoBulletArabicParenRight = 2, /*!< Arabic bullet with closing parenthesis. */
        msoBulletArabicPeriod = 3, /*!< Arabic bullet with period. */
        msoBulletArabicPlain = 13, /*!< Plain Arabic bullet. */
        msoBulletCircleNumDBPlain = 18, /*!< Circled number bullet. */
        msoBulletCircleNumWDBlackPlain = 20, /*!< Circled number WD black bullet. */
        msoBulletCircleNumWDWhitePlain = 19, /*!< Circled number WD white bullet. */
        msoBulletHebrewAlphaDash = 25, /*!< Hebrew alphabetical bullet with dash. */
        msoBulletHindiAlpha1Period = 40, /*!< Hindi alphabetical bullet 1 with period. */
        msoBulletHindiAlphaPeriod = 36, /*!< Hindi alphabetical bullet with period. */
        msoBulletHindiNumParenRight = 39, /*!< Hindi numbered bullet with closing parenthesis. */
        msoBulletHindiNumPeriod = 37, /*!< Hindi numbered bullet with period. */
        msoBulletKanjiKoreanPeriod = 27, /*!< Korean Kanji bullet with period. */
        msoBulletKanjiKoreanPlain = 26, /*!< Korean Kanji bullet. */
        msoBulletKanjiSimpChinDBPeriod = 38, /*!< Simplified Chinese Kanji bulllet with period. */
        msoBulletRomanLCParenBoth = 4, /*!< Lowercase roman bullet with opening and closing parentheses. */
        msoBulletRomanLCParenRight = 5, /*!< Lowercase roman bullet with closing parenthesis. */
        msoBulletRomanLCPeriod = 6, /*!< Lowercase roman bullet with period. */
        msoBulletRomanUCParenBoth = 14, /*!< Uppercase roman bullet with opening and closing parentheses. */
        msoBulletRomanUCParenRight = 15, /*!< Uppercase roman bullet with closing parenthesis. */
        msoBulletRomanUCPeriod = 7, /*!< Uppercase roman bullet with period. */
        msoBulletSimpChinPeriod = 17, /*!< Simplified Chinese bulllet with period. */
        msoBulletSimpChinPlain = 16, /*!< Simplified Chinese bullet. */
        msoBulletStyleMixed = -2, /*!< Return value only; indicates a combination of the other states. */
        msoBulletThaiAlphaParenBoth = 32, /*!< Thai alphabetical bullet with opening and closing parentheses. */
        msoBulletThaiAlphaParenRight = 31, /*!< Thai alphabetical bullet with closing parenthesis. */
        msoBulletThaiAlphaPeriod = 30, /*!< Thai alphabetical bullet with period. */
        msoBulletThaiNumParenBoth = 35, /*!< Thai numerical bullet with opening and closing parentheses. */
        msoBulletThaiNumParenRight = 34, /*!< Thai numerical bullet with closing parenthesis. */
        msoBulletThaiNumPeriod = 33, /*!< Thai numerical bullet with period. */
        msoBulletTradChinPeriod = 22, /*!< Traditional Chinese bulllet with period. */
        msoBulletTradChinPlain = 21, /*!< Traditional Chinese bulllet. */      
    };

    /*!Specifies paragraph alignment for a text block.

    [MSDN documentation for MsoParagraphAlignment](http://msdn.microsoft.com/en-us/library/office/ff862771.aspx).
    */
    enum MsoParagraphAlignment {
        msoAlignCenter = 1, /*!< Specifies that the center of each line of text is aligned to the midpoint of the right and left text box margins, and the left and right edges of each line are ragged. */
        msoAlignDistribute = 4, /*!< Specifies that the first and last characters of each line (except the last) are aligned to the left and right margins, and lines are filled by adding or subtracting the same amount from each character. The last line of the paragraph is aligned to the left margin if text direction is left-to-right, or to the right margin if text direction is right-to-left. */
        msoAlignJustify = 3, /*!< Specifies that the first and last characters of each line (except the last) are aligned to the left and right margins, and lines are filled by adding or subtracting space between and within words. The last line of the paragraph is aligned to the left margin if text direction is left-to-right, or to the right margin if text direction is right-to-left. */
        msoAlignJustifyLow = 6, /*!< Specifies the alignment or adjustment of kashida length in Arabic text. Kashida are special characters used to extend the joiner between two Arabic characters. */
        msoAlignLeft = 0, /*!< Specifies that the leftmost character of each line is aligned to the left margin, and the right edge of each line is ragged. This is the default alignment for paragraphs with left-to-right text direction. */
        msoAlignMixed = -2, /*!< Use a combination of alignment styles. */
        msoAlignRight = 2, /*!< Specifies that the rightmost character of each line is aligned to the right margin, and the left edge of each line is ragged. This is the default alignment for paragraphs with right-to-left text direction. */
        msoAlignThaiDistribute = 5, /*!< Specifies that the first and last characters of each line (except the last) are aligned to the left and right margins, and lines are filled by adding or subtracting space between (but not within) words. The last line of the paragraph is aligned to the left margin. */
    };


    /*!Specifies the format of a file or folder path.

    [MSDN documentation for MsoPathFormat](http://msdn.microsoft.com/en-us/library/office/aa432652.aspx).
    */
    enum MsoPathFormat {
        msoPathType1 = 1 , /*!< Represents the Type1 format. */
        msoPathType2 = 2 , /*!< Represents the Type2 format. */
        msoPathType3 = 3 , /*!< Represents the Type3 format. */
        msoPathType4 = 4 , /*!< Represents the Type4 format. */
        msoPathType5 = 5 , /*!< Represents the Type5 format. */
        msoPathType6 = 6 , /*!< Represents the Type6 format. */
        msoPathType7 = 7 , /*!< Represents the Type7 format. */
        msoPathType8 = 8 , /*!< Represents the Type8 format. */
        msoPathType9 = 9 , /*!< Represents the Type9 format. */
        msoPathTypeMixed = -2 , /*!< Represents a mixed format. */
        msoPathTypeNone = 0 , /*!< Represents no format. */
    };

    /*!Specifies the fill pattern used in a shape.

    [MSDN documentation for MsoPatternType](http://msdn.microsoft.com/en-us/library/office/aa432653.aspx).
    */
    enum MsoPatternType {
        msoPattern10Percent = 2 , /*!< 10% of the foreground color. */
        msoPattern20Percent = 3 , /*!< 20% of the foreground color. */
        msoPattern25Percent = 4 , /*!< 25% of the foreground color. */
        msoPattern30Percent = 5 , /*!< 30% of the foreground color. */
        msoPattern40Percent = 6 , /*!< 40% of the foreground color. */
        msoPattern50Percent = 7 , /*!< 50% of the foreground color. */
        msoPattern5Percent = 1 , /*!< 5% of the foreground color. */
        msoPattern60Percent = 8 , /*!< 60% of the foreground color. */
        msoPattern70Percent = 9 , /*!< 70% of the foreground color. */
        msoPattern75Percent = 10 , /*!< 75% of the foreground color. */
        msoPattern80Percent = 11 , /*!< 80% of the foreground color. */
        msoPattern90Percent = 12 , /*!< 90% of the foreground color. */
        msoPatternCross = 51 , /*!< Cross */
        msoPatternDarkDownwardDiagonal = 15 , /*!< Dark Downward Diagonal */
        msoPatternDarkHorizontal = 13 , /*!< Dark Horizontal */
        msoPatternDarkUpwardDiagonal = 16 , /*!< Dark Upward Diagonal */
        msoPatternDarkVertical = 14 , /*!< Dark Vertical */
        msoPatternDashedDownwardDiagonal = 28 , /*!< Dashed Downward Diagonal */
        msoPatternDashedHorizontal = 32 , /*!< Dashed Horizontal */
        msoPatternDashedUpwardDiagonal = 27 , /*!< Dashed Upward Diagonal */
        msoPatternDashedVertical = 31 , /*!< Dashed Vertical */
        msoPatternDiagonalBrick = 40 , /*!< Diagonal Brick */
        msoPatternDiagonalCross = 54 , /*!< Diagonal Cross */
        msoPatternDivot = 46 , /*!< Pattern Divot */
        msoPatternDottedDiamond = 24 , /*!< Dotted Diamond */
        msoPatternDottedGrid = 45 , /*!< Dotted Grid */
        msoPatternDownwardDiagonal = 52 , /*!< Downward Diagonal */
        msoPatternHorizontal = 49 , /*!< Horizontal */
        msoPatternHorizontalBrick = 35 , /*!< Horizontal Brick */
        msoPatternLargeCheckerBoard = 36 , /*!< Large Checker Board */
        msoPatternLargeConfetti = 33 , /*!< Large Confetti */
        msoPatternLargeGrid = 34 , /*!< Large Grid */
        msoPatternLightDownwardDiagonal = 21 , /*!< Light Downward Diagonal */
        msoPatternLightHorizontal = 19 , /*!< Light Horizontal */
        msoPatternLightUpwardDiagonal = 22 , /*!< Light Upward Diagonal */
        msoPatternLightVertical = 20 , /*!< Light Vertical */
        msoPatternMixed = -2 , /*!< Mixed pattern */
        msoPatternNarrowHorizontal = 30 , /*!< Narrow Horizontal */
        msoPatternNarrowVertical = 29 , /*!< Narrow Vertical */
        msoPatternOutlinedDiamond = 41 , /*!< Outlined Diamond */
        msoPatternPlaid = 42 , /*!< Plaid */
        msoPatternShingle = 47 , /*!< Shingle */
        msoPatternSmallCheckerBoard = 17 , /*!< Small Checker Board */
        msoPatternSmallConfetti = 37 , /*!< Small Confetti */
        msoPatternSmallGrid = 23 , /*!< Small Grid */
        msoPatternSolidDiamond = 39 , /*!< Solid Diamond */
        msoPatternSphere = 43 , /*!< Sphere */
        msoPatternTrellis = 18 , /*!< Trellis */
        msoPatternUpwardDiagonal = 53 , /*!< Upward Diagonal */
        msoPatternVertical = 50 , /*!< Vertical */
        msoPatternWave = 48 , /*!< Wave */
        msoPatternWeave = 44 , /*!< Weave */
        msoPatternWideDownwardDiagonal = 25 , /*!< Wide Downward Diagonal */
        msoPatternWideUpwardDiagonal = 26 , /*!< Wide Upward Diagonal */
        msoPatternZigZag = 38 , /*!< Zig Zag */
    };

    /*!Specifies the color transformation applied to a picture.

    [MSDN documentation for MsoPictureColorType](http://msdn.microsoft.com/en-us/library/office/aa432655.aspx).
    */
    enum MsoPictureColorType {
        msoPictureAutomatic = 1 , /*!< Default color transformation. */
        msoPictureBlackAndWhite = 3 , /*!< Black-and-white transformation. */
        msoPictureGrayscale = 2 , /*!< Grayscale transformation. */
        msoPictureMixed = -2 , /*!< Mixed transformation. */
        msoPictureWatermark = 4 , /*!< Watermark transformation. */
    };


    /*!Indicates the effects camera type used by the specified object.

    [MSDN documentation for MsoPresetCamera](http://msdn.microsoft.com/en-us/library/office/aa432656.aspx).
    */
    enum MsoPresetCamera {
        msoCameraIsometricBottomDown = 23 , /*!< Specifies Isometric Bottom Down. */
        msoCameraIsometricBottomUp = 22 , /*!< Specifies Isometric Bottom Up. */
        msoCameraIsometricLeftDown = 25 , /*!< Specifies Isometric Left Down. */
        msoCameraIsometricLeftUp = 24 , /*!< Specifies Isometric Left Up. */
        msoCameraIsometricOffAxis1Left = 28 , /*!< Specifies Isometric OffAxis1 Left. */
        msoCameraIsometricOffAxis1Right = 29 , /*!< Specifies Isometric OffAxis1 Right. */
        msoCameraIsometricOffAxis1Top = 30 , /*!< Specifies Isometric OffAxis1 Top. */
        msoCameraIsometricOffAxis2Left = 31 , /*!< Specifies Isometric OffAxis2 Left. */
        msoCameraIsometricOffAxis2Right = 32 , /*!< Specifies Isometric OffAxis2 Right. */
        msoCameraIsometricOffAxis2Top = 33 , /*!< Specifies Isometric OffAxis2 Top. */
        msoCameraIsometricOffAxis3Bottom = 36 , /*!< Specifies Isometric OffAxis3 Bottom. */
        msoCameraIsometricOffAxis3Left = 34 , /*!< Specifies Isometric OffAxis3 Left. */
        msoCameraIsometricOffAxis3Right = 35 , /*!< Specifies Isometric OffAxis3 Right. */
        msoCameraIsometricOffAxis4Bottom = 39 , /*!< Specifies Isometric OffAxis4 Bottom. */
        msoCameraIsometricOffAxis4Left = 37 , /*!< Specifies Isometric OffAxis4 Left. */
        msoCameraIsometricOffAxis4Right = 38 , /*!< Specifies Isometric OffAxis4 Right. */
        msoCameraIsometricRightDown = 27 , /*!< Specifies Isometric Right Down. */
        msoCameraIsometricRightUp = 26 , /*!< Specifies Isometric Right Up. */
        msoCameraIsometricTopDown = 21 , /*!< Specifies Isometric Top Down. */
        msoCameraIsometricTopUp = 20 , /*!< Specifies Isometric Top Up. */
        msoCameraLegacyObliqueBottom = 8 , /*!< Specifies Legacy Oblique Bottom. */
        msoCameraLegacyObliqueBottomLeft = 7 , /*!< Specifies Legacy Oblique Lower Left. */
        msoCameraLegacyObliqueBottomRight = 9 , /*!< Specifies Legacy Oblique Lower Right. */
        msoCameraLegacyObliqueFront = 5 , /*!< Specifies Legacy Oblique Front. */
        msoCameraLegacyObliqueLeft = 4 , /*!< Specifies Legacy Oblique Left. */
        msoCameraLegacyObliqueRight = 6 , /*!< Specifies Legacy Oblique Right. */
        msoCameraLegacyObliqueTop = 2 , /*!< Specifies Legacy Oblique Top. */
        msoCameraLegacyObliqueTopLeft = 1 , /*!< Specifies Legacy Oblique Upper Left. */
        msoCameraLegacyObliqueTopRight = 3 , /*!< Specifies Legacy Oblique Upper Right. */
        msoCameraLegacyPerspectiveBottom = 17 , /*!< Specifies Legacy Perspective Bottom. */
        msoCameraLegacyPerspectiveBottomLeft = 16 , /*!< Specifies Legacy Perspective Lower Left. */
        msoCameraLegacyPerspectiveBottomRight = 18 , /*!< Specifies Legacy Perspective Lower Right. */
        msoCameraLegacyPerspectiveFront = 14 , /*!< Specifies Legacy Perspective Front. */
        msoCameraLegacyPerspectiveLeft = 13 , /*!< Specifies Legacy Perspective Left. */
        msoCameraLegacyPerspectiveRight = 15 , /*!< Specifies Legacy Perspective Right. */
        msoCameraLegacyPerspectiveTop = 11 , /*!< Specifies Legacy Perspective Top. */
        msoCameraLegacyPerspectiveTopLeft = 10 , /*!< Specifies Legacy Perspective Upper Left. */
        msoCameraLegacyPerspectiveTopRight = 12 , /*!< Specifies Legacy Perspective Upper Right. */
        msoCameraObliqueBottom = 46 , /*!< Specifies Oblique Bottom. */
        msoCameraObliqueBottomLeft = 45 , /*!< Specifies Oblique Lower Left. */
        msoCameraObliqueBottomRight = 47 , /*!< Specifies Oblique Lower Right. */
        msoCameraObliqueLeft = 43 , /*!< Specifies Oblique Left. */
        msoCameraObliqueRight = 44 , /*!< Specifies Oblique Right. */
        msoCameraObliqueTop = 41 , /*!< Specifies Oblique Top. */
        msoCameraObliqueTopLeft = 40 , /*!< Specifies Oblique Upper Left. */
        msoCameraObliqueTopRight = 42 , /*!< Specifies Oblique Upper Right. */
        msoCameraOrthographicFront = 19 , /*!< Specifies Orthographic Front. */
        msoCameraPerspectiveAbove = 51 , /*!< Specifies Perspective Above. */
        msoCameraPerspectiveAboveLeftFacing = 53 , /*!< Specifies Perspective Above Left Facing. */
        msoCameraPerspectiveAboveRightFacing = 54 , /*!< Specifies Perspective Above Right Facing. */
        msoCameraPerspectiveBelow = 52 , /*!< Specifies Perspective Below. */
        msoCameraPerspectiveContrastingLeftFacing = 55 , /*!< Specifies Perspective Contrasting Left Facing. */
        msoCameraPerspectiveContrastingRightFacing = 56 , /*!< Specifies Perspective Contrasting Right Facing. */
        msoCameraPerspectiveFront = 48 , /*!< Specifies Perspective Front. */
        msoCameraPerspectiveHeroicExtremeLeftFacing = 59 , /*!< Specifies Perspective Heroic Extreme Left Facing. */
        msoCameraPerspectiveHeroicExtremeRightFacing = 60 , /*!< Specifies Perspective Heroic Extreme Right Facing. */
        msoCameraPerspectiveHeroicLeftFacing = 57 , /*!< Specifies Perspective Heroic Left Facing. */
        msoCameraPerspectiveHeroicRightFacing = 58 , /*!< Specifies Perspective Heroic Right Facing. */
        msoCameraPerspectiveLeft = 49 , /*!< Specifies Perspective Left. */
        msoCameraPerspectiveRelaxed = 61 , /*!< Specifies Perspective Relaxed. */
        msoCameraPerspectiveRelaxedModerately = 62 , /*!< Specifies Perspective Relaxed Moderately. */
        msoCameraPerspectiveRight = 50 , /*!< Specifies Perspective Right. */
        msoPresetCameraMixed = -2 , /*!< Specifies a mixed effect. */
    };

    /*!Specifies the direction that the extrusion's sweep path takes away from the extruded shape (the front face of the extrusion).

    [MSDN documentation for MsoPresetExtrusionDirection](http://msdn.microsoft.com/en-us/library/office/aa432657.aspx).
    */
    enum MsoPresetExtrusionDirection {
        msoExtrusionBottom = 2 , /*!< Lower part. */
        msoExtrusionBottomLeft = 3 , /*!< Lower left. */
        msoExtrusionBottomRight = 1 , /*!< Lower right. */
        msoExtrusionLeft = 6 , /*!< Left. */
        msoExtrusionNone = 5 , /*!< No extrusion. */
        msoExtrusionRight = 4 , /*!< Right. */
        msoExtrusionTop = 8 , /*!< Upper part. */
        msoExtrusionTopLeft = 9 , /*!< Upper left. */
        msoExtrusionTopRight = 7 , /*!< Upper right. */
        msoPresetExtrusionDirectionMixed = -2 , /*!< Return value only; indicates a combination of the other states.  */
    };

    /*!Specifies which predefined gradient to use to fill a shape.

    [MSDN documentation for MsoPresetGradientType](http://msdn.microsoft.com/en-us/library/office/aa432658.aspx).
    */
    enum MsoPresetGradientType {
        msoGradientBrass = 20 , /*!< Brass gradient. */
        msoGradientCalmWater = 8 , /*!< Calm Water gradient. */
        msoGradientChrome = 21 , /*!< Chrome gradient. */
        msoGradientChromeII = 22 , /*!< Chrome II gradient. */
        msoGradientDaybreak = 4 , /*!< Daybreak gradient. */
        msoGradientDesert = 6 , /*!< Desert gradient. */
        msoGradientEarlySunset = 1 , /*!< Early Sunset gradient. */
        msoGradientFire = 9 , /*!< Fire gradient. */
        msoGradientFog = 10 , /*!< Fog gradient. */
        msoGradientGold = 18 , /*!< Gold gradient. */
        msoGradientGoldII = 19 , /*!< Gold II gradient. */
        msoGradientHorizon = 5 , /*!< Horizon gradient. */
        msoGradientLateSunset = 2 , /*!< Late Sunset gradient. */
        msoGradientMahogany = 15 , /*!< Mahogany gradient. */
        msoGradientMoss = 11 , /*!< Moss gradient. */
        msoGradientNightfall = 3 , /*!< Nightfall gradient. */
        msoGradientOcean = 7 , /*!< Ocean gradient. */
        msoGradientParchment = 14 , /*!< Parchment gradient. */
        msoGradientPeacock = 12 , /*!< Peacock gradient. */
        msoGradientRainbow = 16 , /*!< Rainbow gradient. */
        msoGradientRainbowII = 17 , /*!< Rainbow II gradient. */
        msoGradientSapphire = 24 , /*!< Sapphire gradient. */
        msoGradientSilver = 23 , /*!< Silver gradient. */
        msoGradientWheat = 13 , /*!< Wheat gradient. */
        msoPresetGradientMixed = -2 , /*!< Mixed gradient. */
    };

    /*!Specifies the location of lighting on an extruded (three-dimensional) shape relative to the shape.

    [MSDN documentation for MsoPresetLightingDirection](http://msdn.microsoft.com/en-us/library/office/aa432659.aspx).
    */
    enum MsoPresetLightingDirection {
        msoLightingBottom = 8 , /*!< Lighting comes from the lower part. */
        msoLightingBottomLeft = 7 , /*!< Lighting comes from the lower left. */
        msoLightingBottomRight = 9 , /*!< Lighting comes from the lower right. */
        msoLightingLeft = 4 , /*!< Lighting comes from the left. */
        msoLightingNone = 5 , /*!< No lighting. */
        msoLightingRight = 6 , /*!< Lighting comes from the right. */
        msoLightingTop = 2 , /*!< Lighting comes from the upper part. */
        msoLightingTopLeft = 1 , /*!< Lighting comes from the upper left. */
        msoLightingTopRight = 3 , /*!< Lighting comes from the upper right. */
        msoPresetLightingDirectionMixed = -2 , /*!< Not supported. */
    };

    /*!Specifies the intensity of light used on a shape.

    [MSDN documentation for MsoPresetLightingSoftness](http://msdn.microsoft.com/en-us/library/office/aa432660.aspx).
    */
    enum MsoPresetLightingSoftness {
        msoLightingBright = 3 , /*!< Bright light. */
        msoLightingDim = 1 , /*!< Dim light. */
        msoLightingNormal = 2 , /*!< Normal light. */
        msoPresetLightingSoftnessMixed = -2 , /*!< Not supported. */
    };

    /*!Specifies the extrusion surface material.

    [MSDN documentation for MsoPresetMaterial](http://msdn.microsoft.com/en-us/library/office/aa432661.aspx).
    */
    enum MsoPresetMaterial {
        msoMaterialClear = 13 , /*!< Clear */
        msoMaterialDarkEdge = 11 , /*!< DarkEdge */
        msoMaterialFlat = 14 , /*!< Flat */
        msoMaterialMatte = 1 , /*!< Matte */
        msoMaterialMatte2 = 5 , /*!< Matte2 */
        msoMaterialMetal = 3 , /*!< Metal */
        msoMaterialMetal2 = 7 , /*!< Metal2 */
        msoMaterialPlastic = 2 , /*!< Plastic */
        msoMaterialPlastic2 = 6 , /*!< Plastic2 */
        msoMaterialPowder = 10 , /*!< Powder */
        msoMaterialSoftEdge = 12 , /*!< Soft Edge */
        msoMaterialSoftMetal = 15 , /*!< Soft Metal */
        msoMaterialTranslucentPowder = 9 , /*!< Translucent Powder */
        msoMaterialWarmMatte = 8 , /*!< Warm Matte */
        msoMaterialWireFrame = 4 , /*!< Wireframe */
        msoPresetMaterialMixed = -2 , /*!< Mixed Material */
    };

    /*! Specifies what text effect to use on a WordArt object. 

    [MSDN documentation for MsoPresetTextEffect](http://msdn.microsoft.com/en-us/library/office/aa432662.aspx).
    */
    enum MsoPresetTextEffect {
        msoTextEffect1 = 0 , /*!< First text effect. */
        msoTextEffect10 = 9 , /*!< Tenth text effect. */
        msoTextEffect11 = 10 , /*!< Eleventh text effect. */
        msoTextEffect12 = 11 , /*!< Twelfth text effect. */
        msoTextEffect13 = 12 , /*!< Thirteenth text effect. */
        msoTextEffect14 = 13 , /*!< Fourteenth text effect. */
        msoTextEffect15 = 14 , /*!< Fifteenth text effect. */
        msoTextEffect16 = 15 , /*!< Sixteenth text effect. */
        msoTextEffect17 = 16 , /*!< Seventeenth text effect. */
        msoTextEffect18 = 17 , /*!< Eighteenth text effect. */
        msoTextEffect19 = 18 , /*!< Nineteenth text effect. */
        msoTextEffect2 = 1 , /*!< Second text effect. */
        msoTextEffect20 = 19 , /*!< Twentieth text effect. */
        msoTextEffect21 = 20 , /*!< Twenty-first text effect. */
        msoTextEffect22 = 21 , /*!< Twenty-second text effect. */
        msoTextEffect23 = 22 , /*!< Twenty-third text effect. */
        msoTextEffect24 = 23 , /*!< Twenty-fourth text effect. */
        msoTextEffect25 = 24 , /*!< Twenty-fifth text effect. */
        msoTextEffect26 = 25 , /*!< Twenty-sixth text effect. */
        msoTextEffect27 = 26 , /*!< Twenty-seventh text effect. */
        msoTextEffect28 = 27 , /*!< Twenty-eighth text effect. */
        msoTextEffect29 = 28 , /*!< Twenty-ninth text effect. */
        msoTextEffect3 = 2 , /*!< Third text effect. */
        msoTextEffect30 = 29 , /*!< Thirtieth text effect. */
        msoTextEffect4 = 3 , /*!< Fourth text effect. */
        msoTextEffect5 = 4 , /*!< Fifth text effect. */
        msoTextEffect6 = 5 , /*!< Sixth text effect. */
        msoTextEffect7 = 6 , /*!< Seventh text effect. */
        msoTextEffect8 = 7 , /*!< Eighth text effect. */
        msoTextEffect9 = 8 , /*!< Ninth text effect. */
        msoTextEffectMixed = -2 , /*!< Not used. */
    };

    /*!Specifies shape of WordArt text. 

    [MSDN documentation for MsoPresetTextEffectShape](http://msdn.microsoft.com/en-us/library/office/aa432663.aspx).
    */
    enum MsoPresetTextEffectShape {
        msoTextEffectShapeArchDownCurve = 10 , /*!< Text is an arch that curves down. */
        msoTextEffectShapeArchDownPour = 14 , /*!< Text is a 3-D arch that curves down. */
        msoTextEffectShapeArchUpCurve = 9 , /*!< Text is an arch that curves up. */
        msoTextEffectShapeArchUpPour = 13 , /*!< Text is a 3-D arch that curves up. */
        msoTextEffectShapeButtonCurve = 12 , /*!< Text is curved around a center "button." */
        msoTextEffectShapeButtonPour = 16 , /*!< Text is seen in 3-D, curved around a center "button." */
        msoTextEffectShapeCanDown = 20 , /*!< Text is stretched to fill the height of the shape, with only a slight curve down. */
        msoTextEffectShapeCanUp = 19 , /*!< Text is stretched to fill the height of the shape, with only a slight curve up. */
        msoTextEffectShapeCascadeDown = 40 , /*!< Text slants up and to the right as font size decreases. */
        msoTextEffectShapeCascadeUp = 39 , /*!< Text slants down and to the right as font size increases. */
        msoTextEffectShapeChevronDown = 6 , /*!< Text slants up to its center point and then slants down. */
        msoTextEffectShapeChevronUp = 5 , /*!< Text slants down to its center point and then slants up. */
        msoTextEffectShapeCircleCurve = 11 , /*!< Text follows a circle, reading clockwise. */
        msoTextEffectShapeCirclePour = 15 , /*!< Text has a 3-D effect and follows a circle, reading clockwise. */
        msoTextEffectShapeCurveDown = 18 , /*!< Text curves down and to the right as font size decreases. */
        msoTextEffectShapeCurveUp = 17 , /*!< Text curves down and to the right as font size increases. */
        msoTextEffectShapeDeflate = 26 , /*!< Font size decreases to the text's midpoint, then increases to the starting size. */
        msoTextEffectShapeDeflateBottom = 28 , /*!< Font size decreases to the text's midpoint, then increases to the starting size, while keeping the top of the text along the same curve. */
        msoTextEffectShapeDeflateInflate = 31 , /*!< Font size increases to the text's midpoint, then decreases to the starting size. */
        msoTextEffectShapeDeflateInflateDeflate = 32 , /*!< Font size decreases, increases, and decreases again across the text. */
        msoTextEffectShapeDeflateTop = 30 , /*!< Font size decreases to the text's midpoint, then increases to the starting size, while keeping the bottom of the text along the same curve. */
        msoTextEffectShapeDoubleWave1 = 23 , /*!< Text follows a line that curves up, then down, then up and down again. */
        msoTextEffectShapeDoubleWave2 = 24 , /*!< Text follows a line that curves down, then up, then down and up again. */
        msoTextEffectShapeFadeDown = 36 , /*!< Top of the text appears to be closer to the viewer than bottom of the text. */
        msoTextEffectShapeFadeLeft = 34 , /*!< Left side of text appears to be closer to the viewer than right side. */
        msoTextEffectShapeFadeRight = 33 , /*!< Right side of text appears to be closer to the viewer than left side. */
        msoTextEffectShapeFadeUp = 35 , /*!< Bottom of text appears to be closer to the viewer than top. */
        msoTextEffectShapeInflate = 25 , /*!< Font size of text increases to its center point, then decreases. Center point of each letter is on the same straight line.  */
        msoTextEffectShapeInflateBottom = 27 , /*!< Font size of text increases to its center point, then decreases. Center point of each letter follows an arch that curves downward. */
        msoTextEffectShapeInflateTop = 29 , /*!< Font size of text increases to its center point, then decreases. Center point of each letter follows an arch that curves upward. */
        msoTextEffectShapeMixed = -2 , /*!< Not used. */
        msoTextEffectShapePlainText = 1 , /*!< No shape applied. */
        msoTextEffectShapeRingInside = 7 , /*!< Text appears to be written inside a 3-D ring. */
        msoTextEffectShapeRingOutside = 8 , /*!< Text appears to be written outside a 3-D ring. */
        msoTextEffectShapeSlantDown = 38 , /*!< Text slants down and to the right. */
        msoTextEffectShapeSlantUp = 37 , /*!< Text slants up and to the right. */
        msoTextEffectShapeStop = 2 , /*!< Text follows the shape of a stop sign. */
        msoTextEffectShapeTriangleDown = 4 , /*!< Text slants up, then down. */
        msoTextEffectShapeTriangleUp = 3 , /*!< Text slants down, then up. */
        msoTextEffectShapeWave1 = 21 , /*!< Text follows a wave up, then down and up again. */
        msoTextEffectShapeWave2 = 22 , /*!< Text follows a wave down, then up and down again. */
    };

    /*!Specifies the type of tab stop.

    [MSDN documentation for MsoTabStopType](http://msdn.microsoft.com/en-us/library/office/ff863302.aspx).
    */
    enum MsoTabStopType {
        msoTabStopCenter = 2, /*!< Center tab stop. */
        msoTabStopDecimal = 4, /*!< Decimal tab stop. */
        msoTabStopLeft = 1, /*!< Left tab stop. */
        msoTabStopMixed = -2, /*!< Return value only; indicates a combination of the other states. */
        msoTabStopRight = 3, /*!< Right tab stop. */

    };

    /*!Specifies the capitalization of the text.

    [MSDN documentation for MsoTextCaps](http://msdn.microsoft.com/en-us/library/office/ff865244.aspx).
    */
    enum MsoTextCaps {
        msoAllCaps = 2, /*!< Display the text as all uppercase letters. */
        msoCapsMixed = -2, /*!< Display the text as mixed uppercase and lowercase letters. */
        msoNoCaps = 0, /*!< Display the text with no uppercase letters. */
        msoSmallCaps = 1, /*!< Display the text as with all lowercasee letters. */

    };

    /*!Specifies the capitalization of text.

    [MSDN documentation for MsoTextChangeCase](http://msdn.microsoft.com/en-us/library/office/ff865194.aspx).
    */
    enum MsoTextChangeCase {
        msoCaseLower = 2, /*!< Display the text as lowercase characters. */
        msoCaseSentence = 1, /*!< Display the text as sentence case characters. Sentence case specifies that the first letter of the sentence is capitalized and that all others should be lowercase (with some exceptions such as proper nouns, and acronyms). */
        msoCaseTitle = 4, /*!< Display the text as title case characters. Title case specifies that the first letter of each word is capitalized and that all others should be lowercase. In some cases short articles, prepositions, and conjunctions are not capitalized. */
        msoCaseToggle = 5, /*!< Indicates that lowercase text should be converted to uppercase and that uppercase text should be converted to lowercase text. */
        msoCaseUpper = 3, /*!< Display the text as uppercase characters. */
    };

    /*!Indicates the number of times a character is printed to darken the image.
    [MSDN documentation for MsoTextStrike](http://msdn.microsoft.com/en-us/library/office/ff861180.aspx).
    */
    enum MsoTextStrike {
        msoDoubleStrike = 2, /*!< Specifies that the character is printed twice. */
        msoNoStrike = 0, /*!< Specifies that the character is not printed. */
        msoSingleStrike = 1, /*!< Specifies that the character is printed once. */
        msoStrikeMixed = -2, /*!< Specifies that the text can contain a combination of doublestrike and single strike characters. */
    };

    /*!Specifies texture to be used to fill a shape.

    [MSDN documentation for MsoPresetTexture](http://msdn.microsoft.com/en-us/library/office/aa432664.aspx).
    */
    enum MsoPresetTexture {
        msoPresetTextureMixed = -2 , /*!< Not used. */
        msoTextureBlueTissuePaper = 17 , /*!< Blue tissue paper texture. */
        msoTextureBouquet = 20 , /*!< Bouquet texture. */
        msoTextureBrownMarble = 11 , /*!< Brown marble texture. */
        msoTextureCanvas = 2 , /*!< Canvas texture. */
        msoTextureCork = 21 , /*!< Cork texture. */
        msoTextureDenim = 3 , /*!< Denim texture. */
        msoTextureFishFossil = 7 , /*!< Fish fossil texture. */
        msoTextureGranite = 12 , /*!< Granite texture. */
        msoTextureGreenMarble = 9 , /*!< Green marble texture. */
        msoTextureMediumWood = 24 , /*!< Medium wood texture. */
        msoTextureNewsprint = 13 , /*!< Newsprint texture. */
        msoTextureOak = 23 , /*!< Oak texture. */
        msoTexturePaperBag = 6 , /*!< Paper bag texture. */
        msoTexturePapyrus = 1 , /*!< Papyrus texture. */
        msoTextureParchment = 15 , /*!< Parchment texture. */
        msoTexturePinkTissuePaper = 18 , /*!< Pink tissue paper texture. */
        msoTexturePurpleMesh = 19 , /*!< Purple mesh texture. */
        msoTextureRecycledPaper = 14 , /*!< Recycled paper texture. */
        msoTextureSand = 8 , /*!< Sand texture. */
        msoTextureStationery = 16 , /*!< Stationery texture. */
        msoTextureWalnut = 22 , /*!< Walnut texture. */
        msoTextureWaterDroplets = 5 , /*!< Water droplets texture. */
        msoTextureWhiteMarble = 10 , /*!< White marble texture. */
        msoTextureWovenMat = 4 , /*!< Woven mat texture. */
    };

    /*!Specifies an extrusion (three-dimensional) format.The 

    [MSDN documentation for MsoPresetThreeDFormat](http://msdn.microsoft.com/en-us/library/office/aa432665.aspx).
    */
    enum MsoPresetThreeDFormat {
        msoPresetThreeDFormatMixed = -2 , /*!< Not used. */
        msoThreeD1 = 1 , /*!< First 3-D format. */
        msoThreeD10 = 10 , /*!< Tenth 3-D format. */
        msoThreeD11 = 11 , /*!< Eleventh 3-D format. */
        msoThreeD12 = 12 , /*!< Twelfth 3-D format. */
        msoThreeD13 = 13 , /*!< Thirteenth 3-D format. */
        msoThreeD14 = 14 , /*!< Fourteenth 3-D format. */
        msoThreeD15 = 15 , /*!< Fifteenth 3-D format. */
        msoThreeD16 = 16 , /*!< Sixteenth 3-D format. */
        msoThreeD17 = 17 , /*!< Seventeenth 3-D format. */
        msoThreeD18 = 18 , /*!< Eighteenth 3-D format. */
        msoThreeD19 = 19 , /*!< Nineteenth 3-D format. */
        msoThreeD2 = 2 , /*!< Second 3-D format. */
        msoThreeD20 = 20 , /*!< Twentieth 3-D format. */
        msoThreeD3 = 3 , /*!< Third 3-D format. */
        msoThreeD4 = 4 , /*!< Fourth 3-D format. */
        msoThreeD5 = 5 , /*!< Fifth 3-D format. */
        msoThreeD6 = 6 , /*!< Sixth 3-D format. */
        msoThreeD7 = 7 , /*!< Seventh 3-D format. */
        msoThreeD8 = 8 , /*!< Eighth 3-D format. */
        msoThreeD9 = 9 , /*!< Ninth 3-D format. */
    };

    /*!Specifies the type of the ReflectionFormat.

    [MSDN documentation for MsoTextureType](http://msdn.microsoft.com/en-us/library/office/ff862544.aspx).
    */
    enum MsoReflectionType {
        msoReflectionType1 = 1, /*!< Type 1. */
        msoReflectionType2 = 2, /*!< Type 2. */
        msoReflectionType3 = 3, /*!< Type 3. */
        msoReflectionType4 = 4, /*!< Type 4. */
        msoReflectionType5 = 5, /*!< Type 5. */
        msoReflectionType6 = 6, /*!< Type 6. */
        msoReflectionType7 = 7, /*!< Type 7. */
        msoReflectionType8 = 8, /*!< Type 8. */
        msoReflectionType9 = 9, /*!< Type 9. */
        msoReflectionTypeMixed = -2, /*!< Return value only; indicates a combination of the other states. */
        msoReflectionTypeNone = 0, /*!< No reflection type. */
    };



    /*!Specifies which part of the shape retains its position when the shape is scaled.

    [MSDN documentation for MsoScaleFrom](http://msdn.microsoft.com/en-us/library/office/aa432670.aspx).
    */
    enum MsoScaleFrom {
        msoScaleFromBottomRight = 2 , /*!< Shape's lower right corner retains its position. */
        msoScaleFromMiddle = 1 , /*!< Shape's midpoint retains its position. */
        msoScaleFromTopLeft = 0 , /*!< Shape's upper left corner retains its position. */
    };

    /*!Specifies the type of shadowing effect.

    [MSDN documentation for MsoShadowStyle](http://msdn.microsoft.com/en-us/library/office/aa432675.aspx).
    */
    enum MsoShadowStyle {
        msoShadowStyleInnerShadow = 1 , /*!< Specifies the inner shadow effect. */
        msoShadowStyleMixed = -2 , /*!< Specifies a combination of  inner and outer  shadow effects. */
        msoShadowStyleOuterShadow = 2 , /*!< Specifies the outer shadow effect. */
    };

    /*!Indicates the line and shape style.

    [MSDN documentation for MsoShapeStyleIndex](http://msdn.microsoft.com/en-us/library/office/aa432677.aspx).
    */
    enum MsoShapeStyleIndex {
        msoLineStyle1 = 10001 , /*!< Line Style 1 */
        msoLineStyle10 = 10010 , /*!< Line Style 10 */
        msoLineStyle11 = 10011 , /*!< Line Style 11 */
        msoLineStyle12 = 10012 , /*!< Line Style 12 */
        msoLineStyle13 = 10013 , /*!< Line Style 13 */
        msoLineStyle14 = 10014 , /*!< Line Style 14 */
        msoLineStyle15 = 10015 , /*!< Line Style 15 */
        msoLineStyle16 = 10016 , /*!< Line Style 16 */
        msoLineStyle17 = 10017 , /*!< Line Style 17 */
        msoLineStyle18 = 10018 , /*!< Line Style 18 */
        msoLineStyle19 = 10019 , /*!< Line Style 19 */
        msoLineStyle2 = 10002 , /*!< Line Style 2 */
        msoLineStyle20 = 10020 , /*!< Line Style 20 */
        msoLineStyle3 = 10003 , /*!< Line Style 3 */
        msoLineStyle4 = 10004 , /*!< Line Style 4 */
        msoLineStyle5 = 10005 , /*!< Line Style 5 */
        msoLineStyle6 = 10006 , /*!< Line Style 6 */
        msoLineStyle7 = 10007 , /*!< Line Style 7 */
        msoLineStyle8 = 10008 , /*!< Line Style 8 */
        msoLineStyle9 = 10009 , /*!< Line Style 9 */
        msoShapeStyle1 = 1 , /*!< Shape Style 1 */
        msoShapeStyle10 = 10 , /*!< Shape Style 10 */
        msoShapeStyle11 = 11 , /*!< Shape Style 11 */
        msoShapeStyle12 = 12 , /*!< Shape Style 12 */
        msoShapeStyle13 = 13 , /*!< Shape Style 13 */
        msoShapeStyle14 = 14 , /*!< Shape Style 14 */
        msoShapeStyle15 = 15 , /*!< Shape Style 15 */
        msoShapeStyle16 = 16 , /*!< Shape Style 16 */
        msoShapeStyle17 = 17 , /*!< Shape Style 17 */
        msoShapeStyle18 = 18 , /*!< Shape Style 18 */
        msoShapeStyle19 = 19 , /*!< Shape Style 19 */
        msoShapeStyle2 = 2 , /*!< Shape Style 2 */
        msoShapeStyle20 = 20 , /*!< Shape Style 20 */
        msoShapeStyle3 = 3 , /*!< Shape Style 3 */
        msoShapeStyle4 = 4 , /*!< Shape Style 4 */
        msoShapeStyle5 = 5 , /*!< Shape Style 5 */
        msoShapeStyle6 = 6 , /*!< Shape Style 6 */
        msoShapeStyle7 = 7 , /*!< Shape Style 7 */
        msoShapeStyle8 = 8 , /*!< Shape Style 8 */
        msoShapeStyle9 = 9 , /*!< Shape Style 9 */
        msoShapeStyleMixed = -2 , /*!< A mix of shape styles. */
        msoShapeStyleNone = 0 , /*!< No shape style. */
    };
    /*!Specifies the type of a shape or range of shapes.

    [MSDN documentation for MsoShapeType](http://msdn.microsoft.com/en-us/library/office/aa432678.aspx).
    */
    enum MsoShapeType {
        msoAutoShape = 1 , /*!< AutoShape. */
        msoCallout = 2 , /*!< Callout. */
        msoCanvas = 20 , /*!< Canvas. */
        msoChart = 3 , /*!< Chart. */
        msoComment = 4 , /*!< Comment. */
        msoDiagram = 21 , /*!< Diagram. */
        msoEmbeddedOLEObject = 7 , /*!< Embedded OLE object. */
        msoFormControl = 8 , /*!< Form control. */
        msoFreeform = 5 , /*!< Freeform. */
        msoGroup = 6 , /*!< Group. */
        msoIgxGraphic = 24 , /*!< SmartArt graphic */
        msoInk = 22 , /*!< Ink */
        msoInkComment = 23 , /*!< Ink comment */
        msoLine = 9 , /*!< Line */
        msoLinkedOLEObject = 10 , /*!< Linked OLE object */
        msoLinkedPicture = 11 , /*!< Linked picture */
        msoMedia = 16 , /*!< Media */
        msoOLEControlObject = 12 , /*!< OLE control object */
        msoPicture = 13 , /*!< Picture */
        msoPlaceholder = 14 , /*!< Placeholder */
        msoScriptAnchor = 18 , /*!< Script anchor */
        msoShapeTypeMixed = -2 , /*!< Mixed shape type */
        msoTable = 19 , /*!< Table */
        msoTextBox = 17 , /*!< Text box */
        msoTextEffect = 15 , /*!< Text effect */
    };

    /*!Specifies the type for a segment. 

    [MSDN documentation for MsoSegmentType](http://msdn.microsoft.com/en-us/library/office/aa432674.aspx).
    */
    enum MsoSegmentType {
        msoSegmentCurve = 1 , /*!< Curve. */
        msoSegmentLine = 0 , /*!< Line. */
    };

    /*!Specifies the type of shadow displayed with a shape.The 

    [MSDN documentation for MsoShadowType](http://msdn.microsoft.com/en-us/library/office/aa432676.aspx).
    */
    enum MsoShadowType {
        msoShadow1 = 1 , /*!< First shadow type. */
        msoShadow10 = 10 , /*!< Tenth shadow type. */
        msoShadow11 = 11 , /*!< Eleventh shadow type. */
        msoShadow12 = 12 , /*!< Twelfth shadow type. */
        msoShadow13 = 13 , /*!< Thirteenth shadow type. */
        msoShadow14 = 14 , /*!< Fourteenth shadow type. */
        msoShadow15 = 15 , /*!< Fifteenth shadow type. */
        msoShadow16 = 16 , /*!< Sixteenth shadow type. */
        msoShadow17 = 17 , /*!< Seventeenth shadow type. */
        msoShadow18 = 18 , /*!< Eighteenth shadow type. */
        msoShadow19 = 19 , /*!< Nineteenth shadow type. */
        msoShadow2 = 2 , /*!< Second shadow type. */
        msoShadow20 = 20 , /*!< Twentieth shadow type. */
        msoShadow3 = 3 , /*!< Third shadow type. */
        msoShadow4 = 4 , /*!< Fourth shadow type. */
        msoShadow5 = 5 , /*!< Fifth shadow type. */
        msoShadow6 = 6 , /*!< Sixth shadow type. */
        msoShadow7 = 7 , /*!< Seventh shadow type. */
        msoShadow8 = 8 , /*!< Eighth shadow type. */
        msoShadow9 = 9 , /*!< Ninth shadow type. */
        msoShadowMixed = -2 , /*!< Not supported. */
    };

    /*!


    [MSDN documentation for MsoSoftEdgeType](http://msdn.microsoft.com/en-us/library/office/aa432682.aspx).
    */
    enum MsoSoftEdgeType {
        msoSoftEdgeType1 = 1, /*!< Soft Edge Type 1 */
        msoSoftEdgeType2 = 2, /*!< Soft Edge Type 2 */
        msoSoftEdgeType3 = 3, /*!< Soft Edge Type 3 */
        msoSoftEdgeType4 = 4, /*!< Soft Edge Type 4 */
        msoSoftEdgeType5 = 5, /*!< Soft Edge Type 5 */
        msoSoftEdgeType6 = 6, /*!< Soft Edge Type 6 */
        msoSoftEdgeTypeNone = 0, /*!< No soft edge. */
        SoftEdgeTypeMixed = -2, /*!< A mix of soft edge types. */
    };

    /*!Specifies alignment for WordArt text.

    [MSDN documentation for MsoTextEffectAlignment](http://msdn.microsoft.com/en-us/library/office/aa432694.aspx).
    */
    enum MsoTextEffectAlignment {
        msoTextEffectAlignmentCentered = 2 , /*!< Centered. */
        msoTextEffectAlignmentLeft = 1 , /*!< Left-aligned. */
        msoTextEffectAlignmentLetterJustify = 4 , /*!< Text is justified. Spacing between letters may be adjusted to justify text. */
        msoTextEffectAlignmentMixed = -2 , /*!< Not used. */
        msoTextEffectAlignmentRight = 3 , /*!< Right- aligned. */
        msoTextEffectAlignmentStretchJustify = 6 , /*!< Text is justified. Letters may be stretched to justify text. */
        msoTextEffectAlignmentWordJustify = 5 , /*!< Text is justified. Spacing between words (but not letters) may be adjusted to justify text. */
    };


    /*!Specifies the direction that text runs.    
    */
    enum MsoTextDirection {
        msoTextDirectionLeftToRight = 1, /*!< Text runs left to right. */
        msoTextDirectionMixed = -2, /*!< Return value only; indicates a combination of the other states. */
        msoTextDirectionRightToLeft = 2, /*!< Text runs right to left. */
    };

    /*!Specifies orientation for text.

    [MSDN documentation for MsoTextOrientation](http://msdn.microsoft.com/en-us/library/office/aa432696.aspx).
    */
    enum MsoTextOrientation {
        msoTextOrientationDownward = 3 , /*!< Downward. */
        msoTextOrientationHorizontal = 1 , /*!< Horizontal. */
        msoTextOrientationHorizontalRotatedFarEast = 6 , /*!< Horizontal and rotated as required for Asian language support. */
        msoTextOrientationMixed = -2 , /*!< Not supported. */
        msoTextOrientationUpward = 2 , /*!< Upward. */
        msoTextOrientationVertical = 5 , /*!< Vertical. */
        msoTextOrientationVerticalFarEast = 4 , /*!< Vertical as required for Asian language support. */
    };
    /*!Indicates the type of underline for text.

    [MSDN documentation for MsoTextUnderlineType](http://msdn.microsoft.com/en-us/library/office/ff864159.aspx).
    */
    enum MsoTextUnderlineType {
        msoNoUnderline = 0, /*!< Specifies no underline. */
        msoUnderlineDashHeavyLine = 8, /*!< Specifies a dash underline. */
        msoUnderlineDashLine = 7, /*!< Specifies a dash line underline. */
        msoUnderlineDashLongHeavyLine = 10, /*!< Specifies a long heavy line underline. */
        msoUnderlineDashLongLine = 9, /*!< Specifies a dashed long line underline. */
        msoUnderlineDotDashHeavyLine = 12, /*!< Specifies a dot dash heavy line underline. */
        msoUnderlineDotDashLine = 11, /*!< Specifies a dot dash line underline. */
        msoUnderlineDotDotDashHeavyLine = 14, /*!< Specifies a dot dot dash heavy line underline. */
        msoUnderlineDotDotDashLine = 13, /*!< Specifies a dot dot dash line underline. */
        msoUnderlineDottedHeavyLine = 6, /*!< Specifies a dotted heavy line underline. */
        msoUnderlineDottedLine = 5, /*!< Specifies a dotted line underline. */
        msoUnderlineDoubleLine = 3, /*!< Specifies a double line underline. */
        msoUnderlineHeavyLine = 4, /*!< Specifies a heavy line underline. */
        msoUnderlineMixed = -2, /*!< Specifies a mixed of underline types. */
        msoUnderlineSingleLine = 2, /*!< Specifies a single line underline. */
        msoUnderlineWavyDoubleLine = 17, /*!< Specifies a wavy double line underline. */
        msoUnderlineWavyHeavyLine = 16, /*!< Specifies a wavy heavy line underline. */
        msoUnderlineWavyLine = 15, /*!< Specifies a wavy line underline. */
        msoUnderlineWords = 1, /*!< Specifies underlining words. */

    };


    /*!Specifies the alignment (the origin of the coordinate grid) for the tiling of the texture fill.

    [MSDN documentation for MsoTextureType](http://msdn.microsoft.com/en-us/library/office/aa432700.aspx).
    */
    enum MsoTextureAlignment {
        msoTextureAlignmentMixed = -2,  /*!<  Return value only; indicates a combination of the other states. */
        msoTextureBottom = 7,  /*!<  Bottom alignment. */
        msoTextureBottomLeft = 6,  /*!<  Bottom-left alignment. */
        msoTextureBottomRight = 8,  /*!<  Bottom-right alignment. */
        msoTextureCenter = 4,  /*!<  Center alignment. */
        msoTextureLeft = 3,  /*!<  Left alignment. */
        msoTextureRight = 5,  /*!<  Right alignment. */
        msoTextureTop = 1,  /*!<  Top alignment. */
        msoTextureTopLeft = 0,  /*!<  Top-left alignment. */
        msoTextureTopRight = 2,  /*!<  Top-right alignment. */
    };

    /*!Specifies the texture type for the selected fill.

    [MSDN documentation for MsoTextureType](http://msdn.microsoft.com/en-us/library/office/aa432700.aspx).
    */
    enum MsoTextureType {
        msoTexturePreset = 1 , /*!< Preset texture type. */
        msoTextureTypeMixed = -2 , /*!< Return value only; indicates a combination of the other states.  */
        msoTextureUserDefined = 2 , /*!< User-defined texture type. */
    };

    /*!Indicates the Office theme color.

    [MSDN documentation for MsoThemeColorIndex](http://msdn.microsoft.com/en-us/library/office/aa432702.aspx).
    */
    enum MsoThemeColorIndex {
        msoNotThemeColor = 0 , /*!< Specifies no theme color. */
        msoThemeColorAccent1 = 5 , /*!< Specifies the Accent 1 theme color. */
        msoThemeColorAccent2 = 6 , /*!< Specifies the Accent 2 theme color. */
        msoThemeColorAccent3 = 7 , /*!< Specifies the Accent 3 theme color. */
        msoThemeColorAccent4 = 8 , /*!< Specifies the Accent 4 theme color. */
        msoThemeColorAccent5 = 9 , /*!< Specifies the Accent 5 theme color. */
        msoThemeColorAccent6 = 10 , /*!< Specifies the Accent 6 theme color. */
        msoThemeColorBackground1 = 14 , /*!< Specifies the Background 1 theme color. */
        msoThemeColorBackground2 = 16 , /*!< Specifies the Background 2 theme color. */
        msoThemeColorDark1 = 1 , /*!< Specifies the Dark 1 theme color. */
        msoThemeColorDark2 = 3 , /*!< Specifies the Dark 2 theme color. */
        msoThemeColorFollowedHyperlink = 12 , /*!< Specifies the theme color for a clicked hyperlink. */
        msoThemeColorHyperlink = 11 , /*!< Specifies the theme color for a hyperlink. */
        msoThemeColorLight1 = 2 , /*!< Specifies the Light 1 theme color. */
        msoThemeColorLight2 = 4 , /*!< Specifies the Light 2  theme color. */
        msoThemeColorMixed = -2 , /*!< Specifies a mixed color theme. */
        msoThemeColorText1 = 13 , /*!< Specifies the Text 1 theme color. */
        msoThemeColorText2 = 15 , /*!< Specifies the Text 2 theme color. */
    };

    /*!Specifies a tri-state variable.

    [MSDN documentation for MsoTriState](http://msdn.microsoft.com/en-us/library/office/aa432714.aspx).
    */
    enum MsoTriState {
        msoCTrue = 1 , /*!< Not supported. */
        msoFalse = 0 , /*!< False. */
        msoTriStateMixed = -2 , /*!< Not supported. */
        msoTriStateToggle = -3 , /*!< Not supported. */
        msoTrue = -1 , /*!< True. */
    };

    /*!Specifies the vertical alignment of text in a text frame. Used with the 

    [MSDN documentation for MsoVerticalAnchor](http://msdn.microsoft.com/en-us/library/office/aa432720.aspx).
    */
    enum MsoVerticalAnchor {
        msoAnchorBottom = 4 , /*!< Aligns text to bottom of text frame. */
        msoAnchorBottomBaseLine = 5 , /*!< Anchors bottom of text string to current position, regardless of text resizing. When you resize text without baseline anchoring, text centers itself on previous position. */
        msoAnchorMiddle = 3 , /*!< Centers text vertically. */
        msoAnchorTop = 1 , /*!< Aligns text to top of text frame. */
        msoAnchorTopBaseline = 2 , /*!< Anchors bottom of text string to current position, regardless of text resizing. When you resize text without baseline anchoring, text centers itself on previous position. */
        msoVerticalAnchorMixed = -2 , /*!< Return value only; indicates a combination of the other states.  */
    };

    /*!Indicates various image warping formats.

    [MSDN documentation for MsoWarpFormat](http://msdn.microsoft.com/en-us/library/office/aa432721.aspx).
    */
    enum MsoWarpFormat {
        msoWarpFormat1 = 0 , /*!< Specifies Warp Format 1. */
        msoWarpFormat10 = 9 , /*!< Specifies Warp Format 10. */
        msoWarpFormat11 = 10 , /*!< Specifies Warp Format 11. */
        msoWarpFormat12 = 11 , /*!< Specifies Warp Format 12. */
        msoWarpFormat13 = 12 , /*!< Specifies Warp Format 13. */
        msoWarpFormat14 = 13 , /*!< Specifies Warp Format 14. */
        msoWarpFormat15 = 14 , /*!< Specifies Warp Format 15. */
        msoWarpFormat16 = 15 , /*!< Specifies Warp Format 16. */
        msoWarpFormat17 = 16 , /*!< Specifies Warp Format 17. */
        msoWarpFormat18 = 17 , /*!< Specifies Warp Format 18. */
        msoWarpFormat19 = 18 , /*!< Specifies Warp Format 19. */
        msoWarpFormat2 = 1 , /*!< Specifies Warp Format 2. */
        msoWarpFormat20 = 19 , /*!< Specifies Warp Format 20. */
        msoWarpFormat21 = 20 , /*!< Specifies Warp Format 21. */
        msoWarpFormat22 = 21 , /*!< Specifies Warp Format 22. */
        msoWarpFormat23 = 22 , /*!< Specifies Warp Format 23. */
        msoWarpFormat24 = 23 , /*!< Specifies Warp Format 24. */
        msoWarpFormat25 = 24 , /*!< Specifies Warp Format 25. */
        msoWarpFormat26 = 25 , /*!< Specifies Warp Format 26. */
        msoWarpFormat27 = 26 , /*!< Specifies Warp Format 27. */
        msoWarpFormat28 = 27 , /*!< Specifies Warp Format 28. */
        msoWarpFormat29 = 28 , /*!< Specifies Warp Format 29. */
        msoWarpFormat3 = 2 , /*!< Specifies Warp Format 3. */
        msoWarpFormat30 = 29 , /*!< Specifies Warp Format 30. */
        msoWarpFormat4 = 3 , /*!< Specifies Warp Format 4. */
        msoWarpFormat5 = 4 , /*!< Specifies Warp Format 5. */
        msoWarpFormat6 = 5 , /*!< Specifies Warp Format 6. */
        msoWarpFormat7 = 6 , /*!< Specifies Warp Format 7. */
        msoWarpFormat8 = 7 , /*!< Specifies Warp Format 8. */
        msoWarpFormat9 = 8 , /*!< Specifies Warp Format 9. */
        msoWarpFormatMixed = -2 , /*!< Specifies a mix of warp formats. */
    };

    /*!Specifies where in the z-order a shape should be moved relative to other shapes.

    [MSDN documentation for MsoZOrderCmd](http://msdn.microsoft.com/en-us/library/office/aa432726.aspx).
    */
    enum MsoZOrderCmd {
        msoBringForward = 2 , /*!< Bring shape forward. */
        msoBringInFrontOfText = 4 , /*!< Bring shape in front of text. */
        msoBringToFront = 0 , /*!< Bring shape to the front. */
        msoSendBackward = 3 , /*!< Send shape backward. */
        msoSendBehindText = 5 , /*!< Send shape behind text. */
        msoSendToBack = 1 , /*!< Send shape to the back. */
    };


    /*!  @brief Specifies if the values are above or below average. Since Excel 2007.

    [MSDN documentation for XlAboveBelow](http://msdn.microsoft.com/en-us/library/bb240918.aspx).
    */
    enum XlAboveBelow {
        XlAboveAverage = 0 , /*!< Above average. */
        XlAboveStdDev = 4 , /*!< Above standard deviation. */
        XlBelowAverage = 1 , /*!< Below average. */
        XlBelowStdDev = 5 , /*!< Below standard deviation. */
        XlEqualAboveAverage = 2 , /*!< Equal above average. */
        XlEqualBelowAverage = 3 , /*!< Equal below average. */
    };
    /*!  @brief Specifies the action that should be performed. Since Excel 2007.

    [MSDN documentation for XlActionType](http://msdn.microsoft.com/en-us/library/bb240924.aspx).
    */
    enum XlActionType {
        xlActionTypeDrillthrough = 256 , /*!< Drill through. */
        xlActionTypeReport = 128 , /*!< Report. */
        xlActionTypeRowset = 16 , /*!< Rowset. */
        xlActionTypeUrl = 1 , /*!< URL. */
    };
    /*!  @brief Specifies country/region and international settings. Since Excel 2007.

    [MSDN documentation for XlApplicationInternational](http://msdn.microsoft.com/en-us/library/bb240927.aspx).
    */
    enum XlApplicationInternational {
        xl24HourClock = 33 , /*!< True if you're using 24-hour time; False if you're using 12-hour time. */
        xl4DigitYears = 43 , /*!< True if you're using four-digit years; False if you're using two-digit years. */
        xlAlternateArraySeparator = 16 , /*!< Alternate array item separator to be used if the current array separator is the same as the decimal separator. */
        xlColumnSeparator = 14 , /*!< Character used to separate columns in array literals. */
        xlCountryCode = 1 , /*!< Country/Region version of Microsoft Excel. */
        xlCountrySetting = 2 , /*!< Current country/region setting in the Windows Control Panel. */
        xlCurrencyBefore = 37 , /*!< True if the currency symbol precedes the currency values; False if it follows them. */
        xlCurrencyCode = 25 , /*!< Currency symbol. */
        xlCurrencyDigits = 27 , /*!< Number of decimal digits to be used in currency formats. */
        xlCurrencyLeadingZeros = 40 , /*!< True if leading zeros are displayed for zero currency values. */
        xlCurrencyMinusSign = 38 , /*!< True if you're using a minus sign for negative numbers; False if you're using parentheses. */
        xlCurrencyNegative = 28 , /*!< Currency format for negative currency values:0 = (symbolx) or (xsymbol)1 = -symbolx or -xsymbol2 = symbol-x or x-symbol3 = symbolx- or xsymbol-where symbol is the currency symbol of the country or region. Note that the position of the currency symbol is determined by xlCurrencyBefore. */
        xlCurrencySpaceBefore = 36 , /*!< True if a space is added before the currency symbol. */
        xlCurrencyTrailingZeros = 39 , /*!< True if trailing zeros are displayed for zero currency values. */
        xlDateOrder = 32 , /*!< Order of date elements:0 = month-day-year1 = day-month-year2 = year-month-day */
        xlDateSeparator = 17 , /*!< Date separator (/). */
        xlDayCode = 21 , /*!< Day symbol (d). */
        xlDayLeadingZero = 42 , /*!< True if a leading zero is displayed in days. */
        xlDecimalSeparator = 3 , /*!< Decimal separator. */
        xlGeneralFormatName = 26 , /*!< Name of the General number format. */
        xlHourCode = 22 , /*!< Hour symbol (h). */
        xlLeftBrace = 12 , /*!< Character used instead of the left brace ({) in array literals. */
        xlLeftBracket = 10 , /*!< Character used instead of the left bracket ([) in R1C1-style relative references. */
        xlListSeparator = 5 , /*!< List separator. */
        xlLowerCaseColumnLetter = 9 , /*!< Lowercase column letter. */
        xlLowerCaseRowLetter = 8 , /*!< Lowercase row letter. */
        xlMDY = 44 , /*!< True if the date order is month-day-year for dates displayed in the long form; False if the date order is day-month-year. */
        xlMetric = 35 , /*!< True if you're using the metric system; False if you're using the English measurement system. */
        xlMinuteCode = 23 , /*!< Minute symbol (m). */
        xlMonthCode = 20 , /*!< Month symbol (m). */
        xlMonthLeadingZero = 41 , /*!< True if a leading zero is displayed in months (when months are displayed as numbers). */
        xlMonthNameChars = 30 , /*!< Always returns three characters for backward compatibility. Abbreviated month names are read from Microsoft Windows and can be any length. */
        xlNoncurrencyDigits = 29 , /*!< Number of decimal digits to be used in noncurrency formats. */
        xlNonEnglishFunctions = 34 , /*!< True if you're not displaying functions in English. */
        xlRightBrace = 13 , /*!< Character used instead of the right brace (}) in array literals. */
        xlRightBracket = 11 , /*!< Character used instead of the right bracket (]) in R1C1-style references. */
        xlRowSeparator = 15 , /*!< Character used to separate rows in array literals. */
        xlSecondCode = 24 , /*!< Second symbol (s). */
        xlThousandsSeparator = 4 , /*!< Zero or thousands separator. */
        xlTimeLeadingZero = 45 , /*!< True if a leading zero is displayed in times. */
        xlTimeSeparator = 18 , /*!< Time separator (:). */
        xlUpperCaseColumnLetter = 7 , /*!< Uppercase column letter. */
        xlUpperCaseRowLetter = 6 , /*!< Uppercase row letter (for R1C1-style references). */
        xlWeekdayNameChars = 31 , /*!< Always returns three characters for backward compatibility. Abbreviated weekday names are read from Microsoft Windows and can be any length. */
        xlYearCode = 19 , /*!< Year symbol in number formats (y). */
    };
    /*!  @brief Specifies which range name is listed first when a cell reference is replaced by a row-oriented and column-oriented range name. Since Excel 2007.

    [MSDN documentation for XlApplyNamesOrder](http://msdn.microsoft.com/en-us/library/bb240931.aspx).
    */
    enum XlApplyNamesOrder {
        xlColumnThenRow = 2 , /*!< Columns listed before rows. */
        xlRowThenColumn = 1 , /*!< Rows listed before columns. */
    };
    /*!  @brief Specifies spelling rules for the Arabic spelling checker. Since Excel 2007.

    [MSDN documentation for XlArabicModes](http://msdn.microsoft.com/en-us/library/bb240934.aspx).
    */
    enum XlArabicModes {
        xlArabicBothStrict = 3 , /*!< The spelling checker uses spelling rules regarding both Arabic words ending with the letter yaa and Arabic words beginning with an alef hamza. */
        xlArabicNone = 0 , /*!< The spelling checker ignores spelling rules regarding either Arabic words ending with the letter yaa or Arabic words beginning with an alef hamza. */
        xlArabicStrictAlefHamza = 1 , /*!< The spelling checker uses spelling rules regarding Arabic words beginning with an alef hamza. */
        xlArabicStrictFinalYaa = 2 , /*!< The spelling checker uses spelling rules regarding Arabic words ending with the letter yaa. */
    };
    /*!  @brief Specifies how windows are arranged on the screen. Since Excel 2007.

    [MSDN documentation for XlArrangeStyle](http://msdn.microsoft.com/en-us/library/bb240941.aspx).
    */
    enum XlArrangeStyle {
        xlArrangeStyleCascade = 7 , /*!< Windows are cascaded. */
        xlArrangeStyleHorizontal = -4128 , /*!< Windows are arranged horizontally. */
        xlArrangeStyleTiled = 1 , /*!< Default. Windows are tiled. */
        xlArrangeStyleVertical = -4166 , /*!< Windows are arranged vertically. */
    };
    /*!  @brief Specifies the length of the arrowhead at the end of a line. Since Excel 2007.

    [MSDN documentation for XlArrowHeadLength](http://msdn.microsoft.com/en-us/library/bb240944.aspx).
    */
    enum XlArrowHeadLength {
        xlArrowHeadLengthLong = 3 , /*!< Longest arrowhead. */
        xlArrowHeadLengthMedium = -4138 , /*!< Medium-length arrowhead. */
        xlArrowHeadLengthShort = 1 , /*!< Shortest arrowhead. */
    };
    /*!  @brief Specifies the type of arrowhead to apply at the end of a line. Since Excel 2007.

    [MSDN documentation for XlArrowHeadStyle](http://msdn.microsoft.com/en-us/library/bb240945.aspx).
    */
    enum XlArrowHeadStyle {
        xlArrowHeadStyleClosed = 3 , /*!< Small arrowhead with curved edge at connection to line. */
        xlArrowHeadStyleDoubleClosed = 5 , /*!< Large diamond-shaped arrowhead. */
        xlArrowHeadStyleDoubleOpen = 4 , /*!< Large arrowhead with curved edge at connection to line. */
        xlArrowHeadStyleNone = -4142 , /*!< No arrowhead. */
        xlArrowHeadStyleOpen = 2 , /*!< Large triangular arrowhead. */
    };
    /*!  @brief Specifies the width of the arrowhead at the end of a line. Since Excel 2007.

    [MSDN documentation for XlArrowHeadWidth](http://msdn.microsoft.com/en-us/library/bb240948.aspx).
    */
    enum XlArrowHeadWidth {
        xlArrowHeadWidthMedium = -4138 , /*!< Medium-width arrowhead. */
        xlArrowHeadWidthNarrow = 1 , /*!< Narrowest arrowhead. */
        xlArrowHeadWidthWide = 3 , /*!< Widest arrowhead. */
    };
    /*!  @brief Specifies how the target range is to be filled, based on the contents of the source range. Since Excel 2007.

    [MSDN documentation for XlAutoFillType](http://msdn.microsoft.com/en-us/library/bb240952.aspx).
    */
    enum XlAutoFillType {
        xlFillCopy = 1 , /*!< Copy the values and formats from the source range to the target range, repeating if necessary. */
        xlFillDays = 5 , /*!< Extend the names of the days of the week in the source range into the target range. Formats are copied from the source range to the target range, repeating if necessary. */
        xlFillDefault = 0 , /*!< Excel determines the values and formats used to fill the target range. */
        xlFillFormats = 3 , /*!< Copy only the formats from the source range to the target range, repeating if necessary. */
        xlFillMonths = 7 , /*!< Extend the names of the months in the source range into the target range. Formats are copied from the source range to the target range, repeating if necessary. */
        xlFillSeries = 2 , /*!< Extend the values in the source range into the target range as a series (for example, '1, 2' is extended as '3, 4, 5'). Formats are copied from the source range to the target range, repeating if necessary. */
        xlFillValues = 4 , /*!< Copy only the values from the source range to the target range, repeating if necessary. */
        xlFillWeekdays = 6 , /*!< Extend the names of the days of the workweek in the source range into the target range. Formats are copied from the source range to the target range, repeating if necessary. */
        xlFillYears = 8 , /*!< Extend the years in the source range into the target range. Formats are copied from the source range to the target range, repeating if necessary. */
        xlGrowthTrend = 10 , /*!< Extend the numeric values from the source range into the target range, assuming that the relationships between the numbers in the source range are multiplicative (for example, '1, 2,' is extended as '4, 8, 16', assuming that each number is a result of multiplying the previous number by some value). Formats are copied from the source range to the target range, repeating if necessary. */
        xlLinearTrend = 9 , /*!< Extend the numeric values from the source range into the target range, assuming that the relationships between the numbers is additive (for example, '1, 2,' is extended as '3, 4, 5', assuming that each number is a result of adding some value to the previous number). Formats are copied from the source range to the target range, repeating if necessary. */
    };
    /*!  @brief Specifies the operator to use to associate two criteria applied by a filter. Since Excel 2007.

    [MSDN documentation for XlAutoFilterOperator](http://msdn.microsoft.com/en-us/library/bb240957.aspx).
    */
    enum XlAutoFilterOperator {
        xlAnd = 1 , /*!< Logical AND of Criteria1 and Criteria2. */
        xlBottom10Items = 4 , /*!< Lowest-valued items displayed (number of items specified in Criteria1). */
        xlBottom10Percent = 6 , /*!< Lowest-valued items displayed (percentage specified in Criteria1). */
        xlFilterCellColor = 8 , /*!< Color of the cell */
        xlFilterDynamic = 11 , /*!< Dynamic filter */
        xlFilterFontColor = 9 , /*!< Color of the font */
        xlFilterIcon = 10 , /*!< Filter icon */
        xlFilterValues = 7 , /*!< Filter values */
        xlOr = 2 , /*!< Logical OR of Criteria1 or Criteria2. */
        xlTop10Items = 3 , /*!< Highest-valued items displayed (number of items specified in Criteria1). */
        xlTop10Percent = 5 , /*!< Highest-valued items displayed (percentage specified in Criteria1). */
    };
    /*!  @brief Specifies the point on the specified axis where the other axis crosses.

    [MSDN documentation for XlAxisCrosses](http://msdn.microsoft.com/en-us/library/bb240960.aspx).
    */
    enum XlAxisCrosses {
        xlAxisCrossesAutomatic = -4105 , /*!< Microsoft Excel sets the axis crossing point. */
        xlAxisCrossesCustom = -4114 , /*!< The CrossesAt property specifies the axis crossing point. */
        xlAxisCrossesMaximum = 2 , /*!< The axis crosses at the maximum value. */
        xlAxisCrossesMinimum = 4 , /*!< The axis crosses at the minimum value. */
    };
    /*!  @brief Specifies the type of axis group.

    [MSDN documentation for XlAxisGroup](http://msdn.microsoft.com/en-us/library/bb240962.aspx).
    */
    enum XlAxisGroup {
        xlPrimary = 1 , /*!< Primary axis group. */
        xlSecondary = 2 , /*!< Secondary axis group. */
    };
    /*!  @brief Specifies the axis type.

    [MSDN documentation for XlAxisType](http://msdn.microsoft.com/en-us/library/bb240966.aspx).
    */
    enum XlAxisType {
        xlCategory = 1 , /*!< Axis displays categories. */
        xlSeriesAxis = 3 , /*!< Axis displays data series. */
        xlValue = 2 , /*!< Axis displays values. */
    };
    /*!  @brief Specifies the background type for text in charts. Since Excel 2007.

    [MSDN documentation for XlBackground](http://msdn.microsoft.com/en-us/library/bb240967.aspx).
    */
    enum XlBackground {
        xlBackgroundAutomatic = -4105 , /*!< Excel controls the background. */
        xlBackgroundOpaque = 3 , /*!< Opaque background. */
        xlBackgroundTransparent = 2 , /*!< Transparent background. */
    };
    /*!  @brief Specifies the shape used with the 3-D bar or column chart.

    [MSDN documentation for XlBarShape](http://msdn.microsoft.com/en-us/library/bb240970.aspx).
    */
    enum XlBarShape {
        xlBox = 0 , /*!< Box. */
        xlConeToMax = 5 , /*!< Cone, truncated at value. */
        xlConeToPoint = 4 , /*!< Cone, coming to point at value. */
        xlCylinder = 3 , /*!< Cylinder. */
        xlPyramidToMax = 2 , /*!< Pyramid, truncated at value. */
        xlPyramidToPoint = 1 , /*!< Pyramid, coming to point at value. */
    };
    

    /*!  @brief Constants passed to and returned by the ChartGroup.BinsType property. Since Excel 2016.

    [MSDN documentation for XlBinsType](https://msdn.microsoft.com/VBA/Excel-VBA/articles/xlbinstype-enumeration-excel).
    */
    enum XlBinsType {
        xlBinsTypeAutomatic = 0,    /*!< Sets bins type automatically. */
        xlBinsTypeCategorical = 1,  /*!< Sets bins type by category. */
        xlBinsTypeManual = 2,       /*!< Sets bins type manually. */
        xlBinsTypeBinSize = 3,      /*!< Sets bins type by size. */
        xlBinsTypeBinCount = 4,     /*!< Sets bins type by count. */
    } ;

    /*!  @brief Specifies the border to be retrieved.

    [MSDN documentation for XlBordersIndex](http://msdn.microsoft.com/en-us/library/bb240971.aspx).
    */
    enum XlBordersIndex {
        xlDiagonalDown = 5 , /*!< Border running from the upper left-hand corner to the lower right of each cell in the range. */
        xlDiagonalUp = 6 , /*!< Border running from the lower left-hand corner to the upper right of each cell in the range. */
        xlEdgeBottom = 9 , /*!< Border at the bottom of the range. */
        xlEdgeLeft = 7 , /*!< Border at the left-hand edge of the range. */
        xlEdgeRight = 10 , /*!< Border at the right-hand edge of the range. */
        xlEdgeTop = 8 , /*!< Border at the top of the range. */
        xlInsideHorizontal = 12 , /*!< Horizontal borders for all cells in the range except borders on the outside of the range. */
        xlInsideVertical = 11 , /*!< Vertical borders for all the cells in the range except borders on the outside of the range. */
    };
    /*!  @brief Specifies the weight of the border around a range.

    [MSDN documentation for XlBorderWeight](http://msdn.microsoft.com/en-us/library/bb240972.aspx).
    */
    enum XlBorderWeight {
        xlHairline = 1 , /*!< Hairline (thinnest border). */
        xlMedium = -4138 , /*!< Medium. */
        xlThick = 4 , /*!< Thick (widest border). */
        xlThin = 2 , /*!< Thin. */
    };
    /*!  @brief Specifies which dialog box to display. Since Excel 2007.

    [MSDN documentation for XlBuiltInDialog](http://msdn.microsoft.com/en-us/library/office/ff194519%28v=office.14%29.aspx).
    */
    enum XlBuiltInDialog {
        xlDialogActivate = 103, /*!< Activate dialog box . */
        xlDialogActiveCellFont = 476, /*!< Active Cell Font dialog box . */
        xlDialogAddChartAutoformat = 390, /*!< Add Chart Autoformat dialog box . */
        xlDialogAddinManager = 321, /*!< Addin Manager dialog box . */
        xlDialogAlignment = 43, /*!< Alignment dialog box . */
        xlDialogApplyNames = 133, /*!< Apply Names dialog box . */
        xlDialogApplyStyle = 212, /*!< Apply Style dialog box . */
        xlDialogAppMove = 170, /*!< AppMove dialog box . */
        xlDialogAppSize = 171, /*!< AppSize dialog box . */
        xlDialogArrangeAll = 12, /*!< Arrange All dialog box . */
        xlDialogAssignToObject = 213, /*!< Assign To Object dialog box . */
        xlDialogAssignToTool = 293, /*!< Assign To Tool dialog box . */
        xlDialogAttachText = 80, /*!< Attach Text dialog box . */
        xlDialogAttachToolbars = 323, /*!< Attach Toolbars dialog box . */
        xlDialogAutoCorrect = 485, /*!< Auto Correct dialog box . */
        xlDialogAxes = 78, /*!< Axes dialog box . */
        xlDialogBorder = 45, /*!< Border dialog box . */
        xlDialogCalculation = 32, /*!< Calculation dialog box . */
        xlDialogCellProtection = 46, /*!< Cell Protection dialog box . */
        xlDialogChangeLink = 166, /*!< Change Link dialog box . */
        xlDialogChartAddData = 392, /*!< Chart Add Data dialog box . */
        xlDialogChartLocation = 527, /*!< Chart Location dialog box . */
        xlDialogChartOptionsDataLabelMultiple = 724, /*!< Chart Options DataLabel Multiple dialog box . */
        xlDialogChartOptionsDataLabels = 505, /*!< Chart Options DataLabels dialog box . */
        xlDialogChartOptionsDataTable = 506, /*!< Chart Options DataTable dialog box . */
        xlDialogChartSourceData = 540, /*!< Chart SourceData dialog box . */
        xlDialogChartTrend = 350, /*!< Chart Trend dialog box . */
        xlDialogChartType = 526, /*!< Chart Type dialog box . */
        xlDialogChartWizard = 288, /*!< ChartWizard dialog box . */
        xlDialogCheckboxProperties = 435, /*!< Checkbox Properties dialog box . */
        xlDialogClear = 52, /*!< Clear dialog box . */
        xlDialogColorPalette = 161, /*!< Color Palette dialog box . */
        xlDialogColumnWidth = 47, /*!< Column Width dialog box . */
        xlDialogCombination = 73, /*!< Combination dialog box . */
        xlDialogConditionalFormatting = 583, /*!< Conditional Formatting dialog box . */
        xlDialogConsolidate = 191, /*!< Consolidate dialog box . */
        xlDialogCopyChart = 147, /*!< Copy Chart dialog box . */
        xlDialogCopyPicture = 108, /*!< Copy Picture dialog box . */
        xlDialogCreateList = 796, /*!< Create List dialog box . */
        xlDialogCreateNames = 62, /*!< Create Names dialog box . */
        xlDialogCreatePublisher = 217, /*!< Create Publisher dialog box . */
        xlDialogCustomizeToolbar = 276, /*!< Customize Toolbar dialog box . */
        xlDialogCustomViews = 493, /*!< Custom Views dialog box . */
        xlDialogDataDelete = 36, /*!< Data Delete dialog box . */
        xlDialogDataLabel = 379, /*!< Data Label dialog box . */
        xlDialogDataLabelMultiple = 723, /*!< Data Label Multiple dialog box . */
        xlDialogDataSeries = 40, /*!< Data Series dialog box . */
        xlDialogDataValidation = 525, /*!< Data Validation dialog box . */
        xlDialogDefineName = 61, /*!< Define Name dialog box . */
        xlDialogDefineStyle = 229, /*!< Define Style dialog box . */
        xlDialogDeleteFormat = 111, /*!< Delete Format dialog box . */
        xlDialogDeleteName = 110, /*!< Delete Name dialog box . */
        xlDialogDemote = 203, /*!< Demote dialog box . */
        xlDialogDisplay = 27, /*!< Display dialog box . */
        xlDialogDocumentInspector = 862, /*!< Document Inspector dialog box . */
        xlDialogEditboxProperties = 438, /*!< Editbox Properties dialog box . */
        xlDialogEditColor = 223, /*!< Edit Color dialog box . */
        xlDialogEditDelete = 54, /*!< Edit Delete dialog box . */
        xlDialogEditionOptions = 251, /*!< Edition Options dialog box . */
        xlDialogEditSeries = 228, /*!< Edit Series dialog box . */
        xlDialogErrorbarX = 463, /*!< Errorbar X dialog box . */
        xlDialogErrorbarY = 464, /*!< Errorbar Y dialog box . */
        xlDialogErrorChecking = 732, /*!< Error Checking dialog box . */
        xlDialogEvaluateFormula = 709, /*!< Evaluate Formula dialog box . */
        xlDialogExternalDataProperties = 530, /*!< External Data Properties dialog box . */
        xlDialogExtract = 35, /*!< Extract dialog box . */
        xlDialogFileDelete = 6, /*!< File Delete dialog box . */
        xlDialogFileSharing = 481, /*!< File Sharing dialog box . */
        xlDialogFillGroup = 200, /*!< Fill Group dialog box . */
        xlDialogFillWorkgroup = 301, /*!< Fill Workgroup dialog box . */
        xlDialogFilter = 447, /*!< Dialog Filter dialog box . */
        xlDialogFilterAdvanced = 370, /*!< Filter Advanced dialog box . */
        xlDialogFindFile = 475, /*!< Find File dialog box . */
        xlDialogFont = 26, /*!< Font dialog box . */
        xlDialogFontProperties = 381, /*!< Font Properties dialog box . */
        xlDialogFormatAuto = 269, /*!< Format Auto dialog box . */
        xlDialogFormatChart = 465, /*!< Format Chart dialog box . */
        xlDialogFormatCharttype = 423, /*!< Format Charttype dialog box . */
        xlDialogFormatFont = 150, /*!< Format Font dialog box . */
        xlDialogFormatLegend = 88, /*!< Format Legend dialog box . */
        xlDialogFormatMain = 225, /*!< Format Main dialog box . */
        xlDialogFormatMove = 128, /*!< Format Move dialog box . */
        xlDialogFormatNumber = 42, /*!< Format Number dialog box . */
        xlDialogFormatOverlay = 226, /*!< Format Overlay dialog box . */
        xlDialogFormatSize = 129, /*!< Format Size dialog box . */
        xlDialogFormatText = 89, /*!< Format Text dialog box . */
        xlDialogFormulaFind = 64, /*!< Formula Find dialog box . */
        xlDialogFormulaGoto = 63, /*!< Formula Goto dialog box . */
        xlDialogFormulaReplace = 130, /*!< Formula Replace dialog box . */
        xlDialogFunctionWizard = 450, /*!< Function Wizard dialog box . */
        xlDialogGallery3dArea = 193, /*!< Gallery 3D Area dialog box . */
        xlDialogGallery3dBar = 272, /*!< Gallery 3D Bar dialog box . */
        xlDialogGallery3dColumn = 194, /*!< Gallery 3D Column dialog box . */
        xlDialogGallery3dLine = 195, /*!< Gallery 3D Line dialog box . */
        xlDialogGallery3dPie = 196, /*!< Gallery 3D Pie dialog box . */
        xlDialogGallery3dSurface = 273, /*!< Gallery 3D Surface dialog box . */
        xlDialogGalleryArea = 67, /*!< Gallery Area dialog box . */
        xlDialogGalleryBar = 68, /*!< Gallery Bar dialog box . */
        xlDialogGalleryColumn = 69, /*!< Gallery Column dialog box . */
        xlDialogGalleryCustom = 388, /*!< Gallery Custom dialog box . */
        xlDialogGalleryDoughnut = 344, /*!< Gallery Doughnut dialog box . */
        xlDialogGalleryLine = 70, /*!< Gallery Line dialog box . */
        xlDialogGalleryPie = 71, /*!< Gallery Pie dialog box . */
        xlDialogGalleryRadar = 249, /*!< Gallery Radar dialog box . */
        xlDialogGalleryScatter = 72, /*!< Gallery Scatter dialog box . */
        xlDialogGoalSeek = 198, /*!< Goal Seek dialog box . */
        xlDialogGridlines = 76, /*!< Gridlines dialog box . */
        xlDialogImportTextFile = 666, /*!< Import Text File dialog box . */
        xlDialogInsert = 55, /*!< Insert dialog box . */
        xlDialogInsertHyperlink = 596, /*!< Insert Hyperlink dialog box . */
        xlDialogInsertObject = 259, /*!< Insert Object dialog box . */
        xlDialogInsertPicture = 342, /*!< Insert Picture dialog box . */
        xlDialogInsertTitle = 380, /*!< Insert Title dialog box . */
        xlDialogLabelProperties = 436, /*!< Label Properties dialog box . */
        xlDialogListboxProperties = 437, /*!< Listbox Properties dialog box . */
        xlDialogMacroOptions = 382, /*!< Macro Options dialog box . */
        xlDialogMailEditMailer = 470, /*!< Mail Edit Mailer dialog box . */
        xlDialogMailLogon = 339, /*!< Mail Logon dialog box . */
        xlDialogMailNextLetter = 378, /*!< Mail Next Letter dialog box . */
        xlDialogMainChart = 85, /*!< Main Chart dialog box . */
        xlDialogMainChartType = 185, /*!< Main Chart Type dialog box . */
        xlDialogMenuEditor = 322, /*!< Menu Editor dialog box . */
        xlDialogMove = 262, /*!< Move dialog box . */
        xlDialogMyPermission = 834, /*!< My Permission dialog box . */
        xlDialogNameManager = 977, /*!< NameManager dialog box . */
        xlDialogNew = 119, /*!< New dialog box . */
        xlDialogNewName = 978, /*!< NewName dialog box . */
        xlDialogNewWebQuery = 667, /*!< New Web Query dialog box . */
        xlDialogNote = 154, /*!< Note dialog box . */
        xlDialogObjectProperties = 207, /*!< Object Properties dialog box . */
        xlDialogObjectProtection = 214, /*!< Object Protection dialog box . */
        xlDialogOpen = 1, /*!< Open dialog box . */
        xlDialogOpenLinks = 2, /*!< Open Links dialog box . */
        xlDialogOpenMail = 188, /*!< Open Mail dialog box . */
        xlDialogOpenText = 441, /*!< Open Text dialog box . */
        xlDialogOptionsCalculation = 318, /*!< Options Calculation dialog box . */
        xlDialogOptionsChart = 325, /*!< Options Chart dialog box . */
        xlDialogOptionsEdit = 319, /*!< Options Edit dialog box . */
        xlDialogOptionsGeneral = 356, /*!< Options General dialog box . */
        xlDialogOptionsListsAdd = 458, /*!< Options Lists Add dialog box . */
        xlDialogOptionsME = 647, /*!< OptionsME dialog box . */
        xlDialogOptionsTransition = 355, /*!< Options Transition dialog box . */
        xlDialogOptionsView = 320, /*!< Options View dialog box . */
        xlDialogOutline = 142, /*!< Outline dialog box . */
        xlDialogOverlay = 86, /*!< Overlay dialog box . */
        xlDialogOverlayChartType = 186, /*!< Overlay ChartType dialog box . */
        xlDialogPageSetup = 7, /*!< Page Setup dialog box . */
        xlDialogParse = 91, /*!< Parse dialog box . */
        xlDialogPasteNames = 58, /*!< Paste Names dialog box . */
        xlDialogPasteSpecial = 53, /*!< Paste Special dialog box . */
        xlDialogPatterns = 84, /*!< Patterns dialog box . */
        xlDialogPermission = 832, /*!< Permission dialog box . */
        xlDialogPhonetic = 656, /*!< Phonetic dialog box . */
        xlDialogPivotCalculatedField = 570, /*!< Pivot Calculated Field dialog box . */
        xlDialogPivotCalculatedItem = 572, /*!< Pivot Calculated Item dialog box . */
        xlDialogPivotClientServerSet = 689, /*!< Pivot Client Server Set dialog box . */
        xlDialogPivotFieldGroup = 433, /*!< Pivot Field Group dialog box . */
        xlDialogPivotFieldProperties = 313, /*!< Pivot Field Properties dialog box . */
        xlDialogPivotFieldUngroup = 434, /*!< Pivot Field Ungroup dialog box . */
        xlDialogPivotShowPages = 421, /*!< Pivot Show Pages dialog box . */
        xlDialogPivotSolveOrder = 568, /*!< Pivot Solve Order dialog box . */
        xlDialogPivotTableOptions = 567, /*!< Pivot Table Options dialog box . */
        xlDialogPivotTableSlicerConnections = 1183, /*!< Pivot Table Slicer Connections dialog box . */
        xlDialogPivotTableWhatIfAnalysisSettings = 1153, /*!< Pivot Table What If Analysis Settings dialog box . */
        xlDialogPivotTableWizard = 312, /*!< Pivot Table Wizard dialog box . */
        xlDialogPlacement = 300, /*!< Placement dialog box . */
        xlDialogPrint = 8, /*!< Print dialog box . */
        xlDialogPrinterSetup = 9, /*!< Printer Setup dialog box . */
        xlDialogPrintPreview = 222, /*!< Print Preview dialog box . */
        xlDialogPromote = 202, /*!< Promote dialog box . */
        xlDialogProperties = 474, /*!< Properties dialog box . */
        xlDialogPropertyFields = 754, /*!< Property Fields dialog box . */
        xlDialogProtectDocument = 28, /*!< Protect Document dialog box . */
        xlDialogProtectSharing = 620, /*!< Protect Sharing dialog box . */
        xlDialogPublishAsWebPage = 653, /*!< Publish As WebPage dialog box . */
        xlDialogPushbuttonProperties = 445, /*!< Pushbutton Properties dialog box . */
        xlDialogReplaceFont = 134, /*!< Replace Font dialog box . */
        xlDialogRoutingSlip = 336, /*!< This object or member has been deprecated, but it remains part of the object model for backward compatibility. You should not use it in new applications. . */
        xlDialogRowHeight = 127, /*!< Row Height dialog box . */
        xlDialogRun = 17, /*!< Run dialog box . */
        xlDialogSaveAs = 5, /*!< SaveAs dialog box . */
        xlDialogSaveCopyAs = 456, /*!< SaveCopyAs dialog box . */
        xlDialogSaveNewObject = 208, /*!< Save New Object dialog box . */
        xlDialogSaveWorkbook = 145, /*!< Save Workbook dialog box . */
        xlDialogSaveWorkspace = 285, /*!< Save Workspace dialog box . */
        xlDialogScale = 87, /*!< Scale dialog box . */
        xlDialogScenarioAdd = 307, /*!< Scenario Add dialog box . */
        xlDialogScenarioCells = 305, /*!< Scenario Cells dialog box . */
        xlDialogScenarioEdit = 308, /*!< Scenario Edit dialog box . */
        xlDialogScenarioMerge = 473, /*!< Scenario Merge dialog box . */
        xlDialogScenarioSummary = 311, /*!< Scenario Summary dialog box . */
        xlDialogScrollbarProperties = 420, /*!< Scrollbar Properties dialog box . */
        xlDialogSearch = 731, /*!< Search dialog box . */
        xlDialogSelectSpecial = 132, /*!< Select Special dialog box . */
        xlDialogSendMail = 189, /*!< Send Mail dialog box . */
        xlDialogSeriesAxes = 460, /*!< Series Axes dialog box . */
        xlDialogSeriesOptions = 557, /*!< Series Options dialog box . */
        xlDialogSeriesOrder = 466, /*!< Series Order dialog box . */
        xlDialogSeriesShape = 504, /*!< Series Shape dialog box . */
        xlDialogSeriesX = 461, /*!< Series X dialog box . */
        xlDialogSeriesY = 462, /*!< Series Y dialog box . */
        xlDialogSetBackgroundPicture = 509, /*!< Set Background Picture dialog box . */
        xlDialogSetManager = 1109, /*!< Set Manager dialog box . */
        xlDialogSetMDXEditor = 1208, /*!< Set MDX Editor dialog box . */
        xlDialogSetPrintTitles = 23, /*!< Set Print Titles dialog box . */
        xlDialogSetTupleEditorOnColumns = 1108, /*!< Set Tuple Editor On Columns dialog box . */
        xlDialogSetTupleEditorOnRows = 1107, /*!< Set Tuple Editor On Rows dialog box . */
        xlDialogSetUpdateStatus = 159, /*!< Set Update Status dialog box . */
        xlDialogShowDetail = 204, /*!< Show Detail dialog box . */
        xlDialogShowToolbar = 220, /*!< Show Toolbar dialog box . */
        xlDialogSize = 261, /*!< Size dialog box . */
        xlDialogSlicerCreation = 1182, /*!< Slicer Creation dialog box . */
        xlDialogSlicerPivotTableConnections = 1184, /*!< Slicer Pivot Table Connections dialog box . */
        xlDialogSlicerSettings = 1179, /*!< Slicer Settings dialog box . */
        xlDialogSort = 39, /*!< Sort dialog box . */
        xlDialogSortSpecial = 192, /*!< Sort Special dialog box . */
        xlDialogSparklineInsertColumn = 1134, /*!< Sparkline Insert Column dialog box . */
        xlDialogSparklineInsertLine = 1133, /*!< Sparkline Insert Line dialog box . */
        xlDialogSparklineInsertWinLoss = 1135, /*!< Sparkline Insert Win Loss dialog box . */
        xlDialogSplit = 137, /*!< Split dialog box . */
        xlDialogStandardFont = 190, /*!< Standard Font dialog box . */
        xlDialogStandardWidth = 472, /*!< Standard Width dialog box . */
        xlDialogStyle = 44, /*!< Style dialog box . */
        xlDialogSubscribeTo = 218, /*!< Subscribe To dialog box . */
        xlDialogSubtotalCreate = 398, /*!< Subtotal Create dialog box . */
        xlDialogSummaryInfo = 474, /*!< Summary Info dialog box . */
        xlDialogTable = 41, /*!< Table dialog box . */
        xlDialogTabOrder = 394, /*!< Tab Order dialog box . */
        xlDialogTextToColumns = 422, /*!< Text To Columns dialog box . */
        xlDialogUnhide = 94, /*!< Unhide dialog box . */
        xlDialogUpdateLink = 201, /*!< Update Link dialog box . */
        xlDialogVbaInsertFile = 328, /*!< VBA Insert File dialog box . */
        xlDialogVbaMakeAddin = 478, /*!< VBA Make Addin dialog box . */
        xlDialogVbaProcedureDefinition = 330, /*!< VBA Procedure Definition dialog box . */
        xlDialogView3d = 197, /*!< View 3D dialog box . */
        xlDialogWebOptionsBrowsers = 773, /*!< Web Options Browsers dialog box . */
        xlDialogWebOptionsEncoding = 686, /*!< Web Options Encoding dialog box . */
        xlDialogWebOptionsFiles = 684, /*!< Web Options Files dialog box . */
        xlDialogWebOptionsFonts = 687, /*!< Web Options Fonts dialog box . */
        xlDialogWebOptionsGeneral = 683, /*!< Web Options General dialog box . */
        xlDialogWebOptionsPictures = 685, /*!< Web Options Pictures dialog box . */
        xlDialogWindowMove = 14, /*!< Window Move dialog box . */
        xlDialogWindowSize = 13, /*!< Window Size dialog box . */
        xlDialogWorkbookAdd = 281, /*!< Workbook Add dialog box . */
        xlDialogWorkbookCopy = 283, /*!< Workbook Copy dialog box . */
        xlDialogWorkbookInsert = 354, /*!< Workbook Insert dialog box . */
        xlDialogWorkbookMove = 282, /*!< Workbook Move dialog box . */
        xlDialogWorkbookName = 386, /*!< Workbook Name dialog box . */
        xlDialogWorkbookNew = 302, /*!< Workbook New dialog box . */
        xlDialogWorkbookOptions = 284, /*!< Workbook Options dialog box . */
        xlDialogWorkbookProtect = 417, /*!< Workbook Protect dialog box . */
        xlDialogWorkbookTabSplit = 415, /*!< Workbook Tab Split dialog box . */
        xlDialogWorkbookUnhide = 384, /*!< Workbook Unhide dialog box . */
        xlDialogWorkgroup = 199, /*!< Workgroup dialog box . */
        xlDialogWorkspace = 95, /*!< Workspace dialog box . */
        xlDialogZoom = 256, /*!< Zoom dialog box . */
    };
    /*!  @brief Specifies what should be calculated. Since Excel 2007.

    [MSDN documentation for XlCalcFor](http://msdn.microsoft.com/en-us/library/bb240974.aspx).
    */
    enum XlCalcFor {
        xlAllValues = 0 , /*!< All values. */
        xlColGroups = 2 , /*!< Column groups. */
        xlRowGroups = 1 , /*!< Row groups. */
    };
    /*!  @brief Specifies the type of a calculated member in a PivotTable. Since Excel 2007.

    [MSDN documentation for XlCalculatedMemberType](http://msdn.microsoft.com/en-us/library/bb240975.aspx).
    */
    enum XlCalculatedMemberType {
        xlCalculatedMember = 0 , /*!< The member uses a Multidimensional Expression (MDX) formula. */
        xlCalculatedSet = 1 , /*!< The member contains an MDX formula for a set in a cube field. */
    };
    /*!  @brief Specifies the calculation mode. Since Excel 2007.

    [MSDN documentation for XlCalculation](http://msdn.microsoft.com/en-us/library/bb240978.aspx).
    */
    enum XlCalculation {
        xlCalculationAutomatic = -4105 , /*!< Excel controls recalculation. */
        xlCalculationManual = -4135 , /*!< Calculation is done when the user requests it. */
        xlCalculationSemiautomatic = 2 , /*!< Excel controls recalculation but ignores changes in tables. */
    };
    /*!  @brief Specifies which key interrupts recalculation. Since Excel 2007.

    [MSDN documentation for XlCalculationInterruptKey](http://msdn.microsoft.com/en-us/library/bb240981.aspx).
    */
    enum XlCalculationInterruptKey {
        xlAnyKey = 2 , /*!< Pressing any key interrupts recalculation. */
        xlEscKey = 1 , /*!< Pressing the ESC key interrupts recalculation. */
        xlNoKey = 0 , /*!< No key press can interrupt recalculation. */
    };
    /*!  @brief Specifies the calculation state of the application. Since Excel 2007.

    [MSDN documentation for XlCalculationState](http://msdn.microsoft.com/en-us/library/bb240982.aspx).
    */
    enum XlCalculationState {
        xlCalculating = 1 , /*!< Calculations in process. */
        xlDone = 0 , /*!< Calculations complete. */
        xlPending = 2 , /*!< Changes that trigger calculation have been made, but a recalculation has not yet been performed. */
    };
    /*!  @brief Specifies the type of the category axis. Since Excel 2007.

    [MSDN documentation for XlCategoryType](http://msdn.microsoft.com/en-us/library/bb240985.aspx).
    */
    enum XlCategoryType {
        xlAutomaticScale = -4105 , /*!< Excel controls the axis type. */
        xlCategoryScale = 2 , /*!< Axis groups data by an arbitrary set of categories. */
        xlTimeScale = 3 , /*!< Axis groups data on a time scale. */
    };
    /*!  @brief Specifies the way that rows on the specified worksheet are added or deleted to accommodate the number of rows in a recordset returned by a query. Since Excel 2007.

    [MSDN documentation for XlCellInsertionMode](http://msdn.microsoft.com/en-us/library/bb240989.aspx).
    */
    enum XlCellInsertionMode {
        xlInsertDeleteCells = 1 , /*!< Partial rows are inserted or deleted to match the exact number of rows required for the new recordset. */
        xlInsertEntireRows = 2 , /*!< Entire rows are inserted, if necessary, to accommodate any overflow. No cells or rows are deleted from the worksheet. */
        xlOverwriteCells = 0 , /*!< No new cells or rows are added to the worksheet. Data in surrounding cells is overwritten to accommodate any overflow. */
    };
    /*!  @brief Specifies the type of cells.

    [MSDN documentation for XlCellType](http://msdn.microsoft.com/en-us/library/bb240990.aspx).
    */
    enum XlCellType {
        xlCellTypeAllFormatConditions = -4172 , /*!< Cells of any format. */
        xlCellTypeAllValidation = -4174 , /*!< Cells having validation criteria. */
        xlCellTypeBlanks = 4 , /*!< Empty cells. */
        xlCellTypeComments = -4144 , /*!< Cells containing notes. */
        xlCellTypeConstants = 2 , /*!< Cells containing constants. */
        xlCellTypeFormulas = -4123 , /*!< Cells containing formulas. */
        xlCellTypeLastCell = 11 , /*!< The last cell in the used range. */
        xlCellTypeSameFormatConditions = -4173 , /*!< Cells having the same format. */
        xlCellTypeSameValidation = -4175 , /*!< Cells having the same validation criteria. */
        xlCellTypeVisible = 12 , /*!< All visible cells. */
    };
    
    /*!  @brief Specifies the position of the chart element. Since Excel 2007.

    [MSDN documentation for XlChartElementPosition](http://msdn.microsoft.com/en-us/library/bb240993.aspx).
    */
    enum XlChartElementPosition {
        xlChartElementPositionAutomatic = -4105 , /*!< Automatically sets the position of the chart element. */
        xlChartElementPositionCustom = -4114 , /*!< Specifies a specific position for the chart element. */
    };
    /*!  @brief Specifies a chart gallery.

    [MSDN documentation for XlChartGallery](http://msdn.microsoft.com/en-us/library/bb240995.aspx).
    */
    enum XlChartGallery {
        xlAnyGallery = 23 , /*!< Either of the galleries. */
        xlBuiltIn = 21 , /*!< The built-in gallery. */
        xlUserDefined = 22 , /*!< The user-defined gallery. */
    };
    /*!  @brief Specifies the type of the chart item.

    [MSDN documentation for XlChartItem](http://msdn.microsoft.com/en-us/library/bb240997.aspx).
    */
    enum XlChartItem {
        xlAxis = 21 , /*!< Axis. */
        xlAxisTitle = 17 , /*!< Axis title. */
        xlChartArea = 2 , /*!< Chart area. */
        xlChartTitle = 4 , /*!< Chart title. */
        xlCorners = 6 , /*!< Corners. */
        xlDataLabel = 0 , /*!< Data label. */
        xlDataTable = 7 , /*!< Data table. */
        xlDisplayUnitLabel = 30 , /*!< Display unit label. */
        xlDownBars = 20 , /*!< Down bars. */
        xlDropLines = 26 , /*!< Drop lines. */
        xlErrorBars = 9 , /*!< Error bars. */
        xlFloor = 23 , /*!< Floor. */
        xlHiLoLines = 25 , /*!< HiLo lines. */
        xlLeaderLines = 29 , /*!< Leader lines. */
        xlLegend = 24 , /*!< Legend. */
        xlLegendEntry = 12 , /*!< Legend entry. */
        xlLegendKey = 13 , /*!< Legend key. */
        xlMajorGridlines = 15 , /*!< Major gridlines. */
        xlMinorGridlines = 16 , /*!< Minor gridlines. */
        xlNothing = 28 , /*!< Nothing. */
        xlPivotChartCollapseEntireFieldButton = 34, 
        xlPivotChartDropZone = 32 , /*!< PivotChart drop zone. */
        xlPivotChartExpandEntireFieldButton = 33,
        xlPivotChartFieldButton = 31 , /*!< PivotChart field button. */
        xlPlotArea = 19 , /*!< Plot area. */
        xlRadarAxisLabels = 27 , /*!< Radar axis labels. */
        xlSeries = 3 , /*!< Series. */
        xlSeriesLines = 22 , /*!< Series lines. */
        xlShape = 14 , /*!< Shape. */
        xlTrendline = 8 , /*!< Trend line. */
        xlUpBars = 18 , /*!< Up bars. */
        xlWalls = 5 , /*!< Walls. */
        xlXErrorBars = 10 , /*!< X error bars. */
        xlYErrorBars = 11 , /*!< Y error bars. */
    };
    /*!  @brief Specifies where to relocate a chart.

    [MSDN documentation for XlChartLocation](http://msdn.microsoft.com/en-us/library/bb240998.aspx).
    */
    enum XlChartLocation {
        xlLocationAsNewSheet = 1 , /*!< Chart is moved to a new sheet. */
        xlLocationAsObject = 2 , /*!< Chart is to be embedded in an existing sheet. */
        xlLocationAutomatic = 3 , /*!< Excel controls chart location. */
    };
    /*!  @brief Specifies the placement of a user-selected picture on a bar in a 3-D bar or column.

    [MSDN documentation for XlChartPicturePlacement](http://msdn.microsoft.com/en-us/library/bb241002.aspx).
    */
    enum XlChartPicturePlacement {
        xlAllFaces = 7 , /*!< Display on all faces. */
        xlEnd = 2 , /*!< Display on end. */
        xlEndSides = 3 , /*!< Display on end and sides. */
        xlFront = 4 , /*!< Display on front. */
        xlFrontEnd = 6 , /*!< Display on front and end. */
        xlFrontSides = 5 , /*!< Display on front and sides. */
        xlSides = 1 , /*!< Display on sides. */
    };
    /*!  @brief Specifies how pictures are displayed on a column, bar picture chart, or legend key.

    [MSDN documentation for XlChartPictureType](http://msdn.microsoft.com/en-us/library/bb241004.aspx).
    */
    enum XlChartPictureType {
        xlStack = 2 , /*!< Picture is sized to repeat a maximum of 15 times in the longest stacked bar. */
        xlStackScale = 3 , /*!< Picture is sized to a specified number of units and repeated the length of the bar. */
        xlStretch = 1 , /*!< Picture is stretched the full length of the stacked bar. */
    };
    /*!  @brief Specifies the values displayed in the second chart in a pie chart or a bar of pie chart.

    [MSDN documentation for XlChartSplitType](http://msdn.microsoft.com/en-us/library/bb241007.aspx).
    */
    enum XlChartSplitType {
        xlSplitByCustomSplit = 4 , /*!< Arbitrary slides are displayed in the second chart. */
        xlSplitByPercentValue = 3 , /*!< Second chart displays values less than some percentage of the total value. The percentage is specified by the SplitValue property. */
        xlSplitByPosition = 1 , /*!< Second chart displays the smallest values in the data series. The number of values to display is specified by the SplitValue property. */
        xlSplitByValue = 2 , /*!< Second chart displays values less than the value specified by the SplitValue property. */
    };
    /*!  @brief Specifies the chart type.

    [MSDN documentation for XlChartType](http://msdn.microsoft.com/en-us/library/bb241008.aspx).
    */
    enum XlChartType {
        xl3DArea = -4098 , /*!< 3D Area. */
        xl3DAreaStacked = 78 , /*!< 3D Stacked Area. */
        xl3DAreaStacked100 = 79 , /*!< 100% Stacked Area. */
        xl3DBarClustered = 60 , /*!< 3D Clustered Bar. */
        xl3DBarStacked = 61 , /*!< 3D Stacked Bar. */
        xl3DBarStacked100 = 62 , /*!< 3D 100% Stacked Bar. */
        xl3DColumn = -4100 , /*!< 3D Column. */
        xl3DColumnClustered = 54 , /*!< 3D Clustered Column. */
        xl3DColumnStacked = 55 , /*!< 3D Stacked Column.  */
        xl3DColumnStacked100 = 56 , /*!< 3D 100% Stacked Column. */
        xl3DLine = -4101 , /*!< 3D Line. */
        xl3DPie = -4102 , /*!< 3D Pie. */
        xl3DPieExploded = 70 , /*!< Exploded 3D Pie. */
        xlArea = 1 , /*!< Area */
        xlAreaStacked = 76 , /*!< Stacked Area. */
        xlAreaStacked100 = 77 , /*!< 100% Stacked Area. */
        xlBarClustered = 57 , /*!< Clustered Bar. */
        xlBarOfPie = 71 , /*!< Bar of Pie. */
        xlBarStacked = 58 , /*!< Stacked Bar. */
        xlBarStacked100 = 59 , /*!< 100% Stacked Bar. */
        xlBoxwhisker = 121,
        xlBubble = 15 , /*!< Bubble. */
        xlBubble3DEffect = 87 , /*!< Bubble with 3D effects. */
        xlColumnClustered = 51 , /*!< Clustered Column. */
        xlColumnStacked = 52 , /*!< Stacked Column. */
        xlColumnStacked100 = 53 , /*!< 100% Stacked Column. */
        xlConeBarClustered = 102 , /*!< Clustered Cone Bar. */
        xlConeBarStacked = 103 , /*!< Stacked Cone Bar. */
        xlConeBarStacked100 = 104 , /*!< 100% Stacked Cone Bar. */
        xlConeCol = 105 , /*!< 3D Cone Column. */
        xlConeColClustered = 99 , /*!< Clustered Cone Column. */
        xlConeColStacked = 100 , /*!< Stacked Cone Column. */
        xlConeColStacked100 = 101 , /*!< 100% Stacked Cone Column. */
        xlCylinderBarClustered = 95 , /*!< Clustered Cylinder Bar. */
        xlCylinderBarStacked = 96 , /*!< Stacked Cylinder Bar. */
        xlCylinderBarStacked100 = 97 , /*!< 100% Stacked Cylinder Bar. */
        xlCylinderCol = 98 , /*!< 3D Cylinder Column. */
        xlCylinderColClustered = 92 , /*!< Clustered Cone Column. */
        xlCylinderColStacked = 93 , /*!< Stacked Cone Column. */
        xlCylinderColStacked100 = 94 , /*!< 100% Stacked Cylinder Column. */
        xlDoughnut = -4120 , /*!< Doughnut. */
        xlDoughnutExploded = 80 , /*!< Exploded Doughnut. */
        xlFunnel = 123,
        xlHistogram = 118,
        xlLine = 4 , /*!< Line. */
        xlLineMarkers = 65 , /*!< Line with Markers. */
        xlLineMarkersStacked = 66 , /*!< Stacked Line with Markers. */
        xlLineMarkersStacked100 = 67 , /*!< 100% Stacked Line with Markers. */
        xlLineStacked = 63 , /*!< Stacked Line. */
        xlLineStacked100 = 64 , /*!< 100% Stacked Line. */
        xlPareto = 122,
        xlPie = 5 , /*!< Pie. */
        xlPieExploded = 69 , /*!< Exploded Pie. */
        xlPieOfPie = 68 , /*!< Pie of Pie. */
        xlPyramidBarClustered = 109 , /*!< Clustered Pyramid Bar. */
        xlPyramidBarStacked = 110 , /*!< Stacked Pyramid Bar. */
        xlPyramidBarStacked100 = 111 , /*!< 100% Stacked Pyramid Bar. */
        xlPyramidCol = 112 , /*!< 3D Pyramid Column. */
        xlPyramidColClustered = 106 , /*!< Clustered Pyramid Column. */
        xlPyramidColStacked = 107 , /*!< Stacked Pyramid Column. */
        xlPyramidColStacked100 = 108 , /*!< 100% Stacked Pyramid Column. */
        xlRadar = -4151 , /*!< Radar. */
        xlRadarFilled = 82 , /*!< Filled Radar. */
        xlRadarMarkers = 81 , /*!< Radar with Data Markers. */
        xlRegionMap = 140,
        xlStockHLC = 88 , /*!< High-Low-Close. */
        xlStockOHLC = 89 , /*!< Open-High-Low-Close. */
        xlStockVHLC = 90 , /*!< Volume-High-Low-Close. */
        xlStockVOHLC = 91 , /*!< Volume-Open-High-Low-Close. */
        xlSunburst = 120,
        xlSurface = 83 , /*!< 3D Surface. */
        xlSurfaceTopView = 85 , /*!< Surface (Top View). */
        xlSurfaceTopViewWireframe = 86 , /*!< Surface (Top View wireframe). */
        xlTreemap = 117,
        xlWaterfall = 119,
        xlSurfaceWireframe = 84 , /*!< 3D Surface (wireframe). */
        xlXYScatter = -4169 , /*!< Scatter. */
        xlXYScatterLines = 74 , /*!< Scatter with Lines. */
        xlXYScatterLinesNoMarkers = 75 , /*!< Scatter with Lines and No Data Markers. */
        xlXYScatterSmooth = 72 , /*!< Scatter with Smoothed Lines. */
        xlXYScatterSmoothNoMarkers = 73 , /*!< Scatter with Smoothed Lines and No Data Markers. */
    };

    /*!  @brief Specifies the type of version for the document checked in when using the  Since Excel 2007.

    [MSDN documentation for XlCheckInVersionType](http://msdn.microsoft.com/en-us/library/bb241009.aspx).
    */
    enum XlCheckInVersionType {
        xlCheckInMajorVersion = 1 , /*!< Check in the major version. */
        xlCheckInMinorVersion = 0 , /*!< Check in the minor version. */
        xlCheckInOverwriteVersion = 2 , /*!< Overwrite current version on the server. */
    };
    /*!  @brief Specifies the format of an item on the Microsoft Windows clipboard. Since Excel 2007.

    [MSDN documentation for XlClipboardFormat](http://msdn.microsoft.com/en-us/library/bb241010.aspx).
    */
    enum XlClipboardFormat {
        xlClipboardFormatBIFF = 8 , /*!< Binary Interchange file format for Excel version 2.x */
        xlClipboardFormatBIFF12 = 63 , /*!< Binary Interchange file format 12 */
        xlClipboardFormatBIFF2 = 18 , /*!< Binary Interchange file format 2 */
        xlClipboardFormatBIFF3 = 20 , /*!< Binary Interchange file format 3 */
        xlClipboardFormatBIFF4 = 30 , /*!< Binary Interchange file format 4 */
        xlClipboardFormatBinary = 15 , /*!< Binary format */
        xlClipboardFormatBitmap = 9 , /*!< Bitmap format */
        xlClipboardFormatCGM = 13 , /*!< CGM format */
        xlClipboardFormatCSV = 5 , /*!< CSV format */
        xlClipboardFormatDIF = 4 , /*!< DIF format */
        xlClipboardFormatDspText = 12 , /*!< Dsp Text format */
        xlClipboardFormatEmbeddedObject = 21 , /*!< Embedded Object */
        xlClipboardFormatEmbedSource = 22 , /*!< Embedded Source */
        xlClipboardFormatLink = 11 , /*!< Link */
        xlClipboardFormatLinkSource = 23 , /*!< Link to the source file */
        xlClipboardFormatLinkSourceDesc = 32 , /*!< Link to the source description */
        xlClipboardFormatMovie = 24 , /*!< Movie */
        xlClipboardFormatNative = 14 , /*!< Native */
        xlClipboardFormatObjectDesc = 31 , /*!< Object description */
        xlClipboardFormatObjectLink = 19 , /*!< Object link */
        xlClipboardFormatOwnerLink = 17 , /*!< Link to the owner */
        xlClipboardFormatPICT = 2 , /*!< Picture */
        xlClipboardFormatPrintPICT = 3 , /*!< Print picture */
        xlClipboardFormatRTF = 7 , /*!< RTF format */
        xlClipboardFormatScreenPICT = 29 , /*!< Screen Picture */
        xlClipboardFormatStandardFont = 28 , /*!< Standard Font */
        xlClipboardFormatStandardScale = 27 , /*!< Standard Scale */
        xlClipboardFormatSYLK = 6 , /*!< SYLK */
        xlClipboardFormatTable = 16 , /*!< Table */
        xlClipboardFormatText = 0 , /*!< Text */
        xlClipboardFormatToolFace = 25 , /*!< Tool Face */
        xlClipboardFormatToolFacePICT = 26 , /*!< Tool Face Picture */
        xlClipboardFormatVALU = 1 , /*!< Value */
        xlClipboardFormatWK1 = 10 , /*!< Workbook */
    };
    /*!  @brief Specifies the value of the

    [MSDN documentation for XlCmdType](http://msdn.microsoft.com/en-us/library/bb241014.aspx).
    */
    enum XlCmdType {
        xlCmdCube = 1 , /*!< Contains a cube name for an OLAP data source. */
        xlCmdDefault = 4 , /*!< Contains command text that the OLE DB provider understands. */
        xlCmdList = 5 , /*!< Contains a pointer to list data. */
        xlCmdSql = 2 , /*!< Contains an SQL statement. */
        xlCmdTable = 3 , /*!< Contains a table name for accessing OLE DB data sources. */
    };
    /*!  @brief Specifies the color of a selected feature, such as a border, font, or fill. Since Excel 2007.

    [MSDN documentation for XlColorIndex](http://msdn.microsoft.com/en-us/library/bb241016.aspx).
    */
    enum XlColorIndex {
        xlColorIndexAutomatic = -4105 , /*!< Automatic color. */
        xlColorIndexNone = -4142 , /*!< No color. */
    };
    /*!  @brief Specifies how a column is to be parsed. Since Excel 2007.

    [MSDN documentation for XlColumnDataType](http://msdn.microsoft.com/en-us/library/bb241018.aspx).
    */
    enum XlColumnDataType {
        xlDMYFormat = 4 , /*!< DMY date format. */
        xlDYMFormat = 7 , /*!< DYM date format. */
        xlEMDFormat = 10 , /*!< EMD date format. */
        xlGeneralFormat = 1 , /*!< General. */
        xlMDYFormat = 3 , /*!< MDY date format. */
        xlMYDFormat = 6 , /*!< MYD date format. */
        xlSkipColumn = 9 , /*!< Column is not parsed. */
        xlTextFormat = 2 , /*!< Text. */
        xlYDMFormat = 8 , /*!< YDM date format. */
        xlYMDFormat = 5 , /*!< YMD date format. */
    };
    /*!  @brief Specifies the state of the command underlines in Microsoft Office Excel for the Macintosh. Since Excel 2007.

    [MSDN documentation for XlCommandUnderlines](http://msdn.microsoft.com/en-us/library/bb241022.aspx).
    */
    enum XlCommandUnderlines {
        xlCommandUnderlinesAutomatic = -4105 , /*!< Excel controls the display of command underlines. */
        xlCommandUnderlinesOff = -4146 , /*!< Command underlines are not displayed. */
        xlCommandUnderlinesOn = 1 , /*!< Command underlines are displayed. */
    };
    /*!  @brief Specifies the way that cells display comments and comment indicators. Since Excel 2007.

    [MSDN documentation for XlCommentDisplayMode](http://msdn.microsoft.com/en-us/library/bb241024.aspx).
    */
    enum XlCommentDisplayMode {
        xlCommentAndIndicator = 1 , /*!< Display comment and indicator at all times. */
        xlCommentIndicatorOnly = -1 , /*!< Display comment indicator only. Display comment when mouse pointer is moved over cell. */
        xlNoIndicator = 0 , /*!< Display neither the comment nor the comment indicator at any time. */
    };
    /*!  @brief Specifies the types of condition values that can be used. Since Excel 2007.

    [MSDN documentation for XlConditionValueTypes](http://msdn.microsoft.com/en-us/library/office/ff837624%28v=office.14%29.aspx).
    */
    enum XlConditionValueTypes {
        xlConditionValueAutomaticMax = 7, /*!< The longest data bar is proportional to the maximum value in the range. . */
        xlConditionValueAutomaticMin  = 6, /*!< The shortest data bar is proportional to the minimum value in the range. . */
        xlConditionValueFormula = 4, /*!< Formula is used. . */
        xlConditionValueHighestValue = 2, /*!< Highest value from the list of values. . */
        xlConditionValueLowestValue = 1, /*!< Lowest value from the list of values. . */
        xlConditionValueNone = -1, /*!< No conditional value. . */
        xlConditionValueNumber = 0, /*!< Number is used. . */
        xlConditionValuePercent = 3, /*!< Percentage is used. . */
        xlConditionValuePercentile = 5, /*!< Percentile is used. . */
    };
    /*!  @brief Specifies the type of database connection. Since Excel 2007.

    [MSDN documentation for XlConnectionType](http://msdn.microsoft.com/en-us/library/bb241031.aspx).
    */
    enum XlConnectionType {
        xlConnectionTypeODBC = 2 , /*!< ODBC */
        xlConnectionTypeOLEDB = 1 , /*!< OLEDB */
        xlConnectionTypeTEXT = 4 , /*!< Text */
        xlConnectionTypeWEB = 5 , /*!< Web */
        xlConnectionTypeXMLMAP = 3 , /*!< XML MAP */
    };
    /*!  @brief Specifies the subtotal function. Since Excel 2007.

    [MSDN documentation for XlConsolidationFunction](http://msdn.microsoft.com/en-us/library/bb241034.aspx).
    */
    enum XlConsolidationFunction {
        xlAverage = -4106 , /*!< Average. */
        xlCount = -4112 , /*!< Count. */
        xlCountNums = -4113 , /*!< Count numerical values only. */
        xlMax = -4136 , /*!< Maximum. */
        xlMin = -4139 , /*!< Minimum. */
        xlProduct = -4149 , /*!< Multiply. */
        xlStDev = -4155 , /*!< Standard deviation, based on a sample. */
        xlStDevP = -4156 , /*!< Standard deviation, based on the whole population. */
        xlSum = -4157 , /*!< Sum. */
        xlUnknown = 1000 , /*!< No subtotal function specified. */
        xlVar = -4164 , /*!< Variation, based on a sample. */
        xlVarP = -4165 , /*!< Variation, based on the whole population. */
    };
    /*!  @brief Specifies the operator used in a function. Since Excel 2007.

    [MSDN documentation for XlContainsOperator](http://msdn.microsoft.com/en-us/library/bb241040.aspx).
    */
    enum XlContainsOperator {
        xlBeginsWith = 2 , /*!< Begins with a specified value. */
        xlContains = 0 , /*!< Contains a specified value. */
        xlDoesNotContain = 1 , /*!< Does not contain the specified value. */
        xlEndsWith = 3 , /*!< Endswith the specified value */
    };
    /*!  @brief Specifies the format of the picture being copied. Since Excel 2007.

    [MSDN documentation for XlCopyPictureFormat](http://msdn.microsoft.com/en-us/library/bb241043.aspx).
    */
    enum XlCopyPictureFormat {
        xlBitmap = 2 , /*!< Bitmap (.bmp, .jpg, .gif). */
        xlPicture = -4147 , /*!< Drawn picture (.png, .wmf, .mix). */
    };
    /*!  @brief Specifies the processing for a file when it is opened. Since Excel 2007.

    [MSDN documentation for XlCorruptLoad](http://msdn.microsoft.com/en-us/library/bb241045.aspx).
    */
    enum XlCorruptLoad {
        xlExtractData = 2 , /*!< Workbook is opened in extract data mode. */
        xlNormalLoad = 0 , /*!< Workbook is opened normally. */
        xlRepairFile = 1 , /*!< Workbook is opened in repair mode. */
    };
    /*!  @brief Specifies the 32-bit creator code for Excel for Macintosh (decimal 1480803660, Hex 5843454C, string XCEL). Since Excel 2007.

    [MSDN documentation for XlCreator](http://msdn.microsoft.com/en-us/library/bb241047.aspx).
    */
    enum XlCreator {
        xlCreatorCode = 1480803660 , /*!< The Excel for Macintosh creator code. */
    };
    /*!  @brief Specifies the type of credentials method used. Since Excel 2007.

    [MSDN documentation for XlCredentialsMethod](http://msdn.microsoft.com/en-us/library/bb241049.aspx).
    */
    enum XlCredentialsMethod {
        CredentialsMethodIntegrated = 0 , /*!< Integrated */
        CredentialsMethodNone = 1 , /*!< No credentials used */
        CredentialsMethodStored = 2 , /*!< Use stored credentials */
    };
    /*!  @brief Specifies the subtype of the CubeField. Since Excel 2007.

    [MSDN documentation for XlCubeFieldSubType](http://msdn.microsoft.com/en-us/library/bb241055.aspx).
    */
    enum XlCubeFieldSubType {
        xlCubeAttribute = 4 , /*!< Attribute */
        xlCubeCalculatedMeasure = 5 , /*!< Calculated Measure */
        xlCubeHierarchy = 1 , /*!< Hierarchy */
        xlCubeKPIGoal = 7 , /*!< KPI Goal */
        xlCubeKPIStatus = 8 , /*!< KPI Status */
        xlCubeKPITrend = 9 , /*!< KPI Trend */
        xlCubeKPIValue = 6 , /*!< KPI Value */
        xlCubeKPIWeight = 10 , /*!< KPI Weight */
        xlCubeMeasure = 2 , /*!< Measure */
        xlCubeSet = 3 , /*!< Set */
    };
    /*!  @brief Specifies whether the OLAP field is a hierarchy, set, or measure field. Since Excel 2007.

    [MSDN documentation for XlCubeFieldType](http://msdn.microsoft.com/en-us/library/bb241056.aspx).
    */
    enum XlCubeFieldType {
        xlHierarchy = 1 , /*!< OLAP field is a hierarchy. */
        xlMeasure = 2 , /*!< OLAP field is a measure. */
        xlSet = 3 , /*!< OLAP field is a set. */
    };
    /*!  @brief Specifies whether status is Copy mode or Cut mode. Since Excel 2007.

    [MSDN documentation for XlCutCopyMode](http://msdn.microsoft.com/en-us/library/bb241059.aspx).
    */
    enum XlCutCopyMode {
        xlCopy = 1 , /*!< In Copy mode */
        xlCut = 2 , /*!< In Cut mode */
    };
    
    /*!  @brief Specifies the cell error number and value. Since Excel 2007.

    [MSDN documentation for XlCVError](http://msdn.microsoft.com/en-us/library/bb241062.aspx).
    */
    enum XlCVError {
        xlErrDiv0 = 2007 , /*!< cell displays \#DIV/0! */
        xlErrNA = 2042 , /*!< cell displays \#N/A */
        xlErrName = 2029 , /*!< cell displays \#NAME? */
        xlErrNull = 2000 , /*!< cell displays \#NULL! */
        xlErrNum = 2036 , /*!< cell displays \#NUM! */
        xlErrRef = 2023 , /*!< cell displays \#REF! */
        xlErrValue = 2015 , /*!< cell displays \#VALUE!*/
    };

    /*!  @brief Specifies the axis position for a range of cells with conditional formatting as data bars. Since Excel 2010.

    [MSDN documentation for XlDataDataBarAxisPosition](http://msdn.microsoft.com/en-us/library/office/ff821511%28v=office.14%29.aspx).
    */
    enum XlDataDataBarAxisPosition {        
        xlDataBarAxisAutomatic = 0, /*!< Display the axis at a variable position based on the ratio of the minimum negative value to the maximum positive value in the range. Positive values are displayed in a left-to-right direction. Negative values are displayed in a right-to-left direction. When all values are positive or all values are negative, no axis is displayed. . */
        xlDataBarAxisMidpoint = 1, /*!< Display the axis at the midpoint of the cell regardless of the set of values in the range. Positive values are displayed in a left-to-right direction. Negative values are displayed in a right-to-left direction. . */
        xlDataBarAxisNone = 2, /*!< No axis is displayed, and both positive and negative values are displayed in the left-to-right direction. . */
    };

    /*!  @brief Specifies the border of a data bar. Since Excel 2010.

    [MSDN documentation for XlDataDataBarBorder](http://msdn.microsoft.com/en-us/library/office/ff195307%28v=office.14%29.aspx).
    */
    enum XlDataBarBorder {        
        xlDataBarBorderNone = 0, /*!< The data bar has no border. . */
        xlDataBarBorderSolid = 1, /*!< The data bar has a solid border. . */
    };
    
    /*!  @brief Specifies whether to use the same border and fill color as postive data bars. Since Excel 2010.

    [MSDN documentation for XlDataDataBarNegativeColorType](http://msdn.microsoft.com/en-us/library/office/ff835606%28v=office.14%29.aspx).
    */
    enum XlDataBarNegativeColorType {            
        xlDataBarColor = 0, /*!< Use the color specified in the Negative Value and Axis Setting dialog box or by using the ColorType and BorderColorType properties of the NegativeBarFormat object. . */
        xlDataBarSameAsPositive = 1, /*!< Use the same color as positive data bars. . */
    };

    /*!  @brief Specifies how a data bar is filled with color. Since Excel 2010.

    [MSDN documentation for XlDataDataBarFillType](http://msdn.microsoft.com/en-us/library/office/ff196124%28v=office.14%29.aspx).
    */
    enum XlDataBarFillType {        
        xlDataBarFillGradient = 1, /*!< The data bar is filled with a color gradient. . */
        xlDataBarFillSolid = 0, /*!< The data bar is filled with solid color. . */
    };

    /*!  @brief Specifies where the data label is positioned.

    [MSDN documentation for XlDataLabelPosition](http://msdn.microsoft.com/en-us/library/bb241064.aspx).
    */
    enum XlDataLabelPosition {
        xlLabelPositionAbove = 0 , /*!< Data label is positioned above the data point. */
        xlLabelPositionBelow = 1 , /*!< Data label is positioned below the data point. */
        xlLabelPositionBestFit = 5 , /*!< Microsoft Office Excel 2007 sets the  position of the data label. */
        xlLabelPositionCenter = -4108 , /*!< Data label is centered on the data point or is inside a bar or pie chart. */
        xlLabelPositionCustom = 7 , /*!< Data label is in a custom position. */
        xlLabelPositionInsideBase = 4 , /*!< Data label is positioned inside the data point at the bottom edge. */
        xlLabelPositionInsideEnd = 3 , /*!< Data label is positioned inside the data point at the top edge. */
        xlLabelPositionLeft = -4131 , /*!< Data label is positioned to the left of the data point. */
        xlLabelPositionMixed = 6 , /*!< Data labels are in multiple positions. */
        xlLabelPositionOutsideEnd = 2 , /*!< Data label is positioned outside the data point at the top edge. */
        xlLabelPositionRight = -4152 , /*!< Data label is positioned to the right of the data point. */
    };
    /*!  @brief Specifies the separator used with data labels.

    [MSDN documentation for XlDataLabelSeparator](http://msdn.microsoft.com/en-us/library/bb241067.aspx).
    */
    enum XlDataLabelSeparator {
        xlDataLabelSeparatorDefault = 1 , /*!< Excel selects the separator. */
    };
    /*!  @brief Specifies the type of data label to apply.

    [MSDN documentation for XlDataLabelsType](http://msdn.microsoft.com/en-us/library/bb241068.aspx).
    */
    enum XlDataLabelsType {
        xlDataLabelsShowBubbleSizes = 6 , /*!< Show the size of the bubble in reference to the absolute value. */
        xlDataLabelsShowLabel = 4 , /*!< Category for the point. */
        xlDataLabelsShowLabelAndPercent = 5 , /*!< Percentage of the total, and category for the point. Available only for pie charts and doughnut charts. */
        xlDataLabelsShowNone = -4142 , /*!< No data labels. */
        xlDataLabelsShowPercent = 3 , /*!< Percentage of the total. Available only for pie charts and doughnut charts. */
        xlDataLabelsShowValue = 2 , /*!< Default value for the point (assumed if this argument is not specified). */
    };
    /*!  @brief Specifies the type of date to apply to a data series.

    [MSDN documentation for XlDataSeriesDate](http://msdn.microsoft.com/en-us/library/bb241205.aspx).
    */
    enum XlDataSeriesDate {
        xlDay = 1 , /*!< Day */
        xlMonth = 3 , /*!< Month */
        xlWeekday = 2 , /*!< Weekdays */
        xlYear = 4 , /*!< Year */
    };
    /*!  @brief Specifies the data series to create.

    [MSDN documentation for XlDataSeriesType](http://msdn.microsoft.com/en-us/library/bb241208.aspx).
    */
    enum XlDataSeriesType {
        xlAutoFill = 4 , /*!< Fill series according to AutoFill settings. */
        xlChronological = 3 , /*!< Fill with date values. */
        xlDataSeriesLinear = -4132 , /*!< Extend values, assuming an additive progression (for example, '1, 2' is extended as '3, 4, 5'). */
        xlGrowth = 2 , /*!< Extend values, assuming a multiplicative progression (for example, '1, 2' is extended as '4, 8, 16'). */
    };

    /*!  @brief Specifies how to shift cells to replace deleted cells. Since Excel 2007.

    [MSDN documentation for XlDeleteShiftDirection](http://msdn.microsoft.com/en-us/library/bb241209.aspx).
    */
    enum XlDeleteShiftDirection {
        xlShiftToLeft = -4159 , /*!< Cells are shifted to the left. */
        xlShiftUp = -4162 , /*!< Cells are shifted up. */
    };
    /*!  @brief Specifies the direction in which to move. Since Excel 2007.

    [MSDN documentation for XlDirection](http://msdn.microsoft.com/en-us/library/bb241212.aspx).
    */
    enum XlDirection {
        xlDown = -4121 , /*!< Down. */
        xlToLeft = -4159 , /*!< To left. */
        xlToRight = -4161 , /*!< To right. */
        xlUp = -4162 , /*!< Up. */
    };
    /*!  @brief Specifies how blank cells are plotted on a chart. Since Excel 2007.

    [MSDN documentation for XlDisplayBlanksAs](http://msdn.microsoft.com/en-us/library/bb241214.aspx).
    */
    enum XlDisplayBlanksAs {
        xlInterpolated = 3 , /*!< Values are interpolated into the chart. */
        xlNotPlotted = 1 , /*!< Blank cells are not plotted. */
        xlZero = 2 , /*!< Blanks are plotted as zero. */
    };
    /*!  @brief Specifies how shapes are displayed. Since Excel 2007.

    [MSDN documentation for XlDisplayDrawingObjects](http://msdn.microsoft.com/en-us/library/bb241216.aspx).
    */
    enum XlDisplayDrawingObjects {
        xlDisplayShapes = -4104 , /*!< Show all shapes. */
        xlHide = 3 , /*!< Hide all shapes. */
        xlPlaceholders = 2 , /*!< Show only placeholders. */
    };
    /*!  @brief Specifies the display unit label for an axis. Since Excel 2007.

    [MSDN documentation for XlDisplayUnit](http://msdn.microsoft.com/en-us/library/bb241219.aspx).
    */
    enum XlDisplayUnit {
        xlHundredMillions = -8 , /*!< Hundreds of millions. */
        xlHundreds = -2 , /*!< Hundreds. */
        xlHundredThousands = -5 , /*!< Hundreds of thousands. */
        xlMillionMillions = -10 , /*!< Millions of millions. */
        xlMillions = -6 , /*!< Millions. */
        xlTenMillions = -7 , /*!< Tens of millions. */
        xlTenThousands = -4 , /*!< Tens of thousands. */
        xlThousandMillions = -9 , /*!< Thousands of millions. */
        xlThousands = -3 , /*!< Thousands. */
    };
    /*!  @brief Specifies whether duplicate or unique values shoud be displayed. Since Excel 2007.

    [MSDN documentation for XlDupeUnique](http://msdn.microsoft.com/en-us/library/bb241223.aspx).
    */
    enum XlDupeUnique {
        xlDuplicate = 1 , /*!< Display duplicate values. */
        xlUnique = 0 , /*!< Display unique values. */
    };
    /*!  @brief Specifies the icon used in message boxes displayed during validation. Since Excel 2007.

    [MSDN documentation for XlDVAlertStyle](http://msdn.microsoft.com/en-us/library/bb241227.aspx).
    */
    enum XlDVAlertStyle {
        xlValidAlertInformation = 3 , /*!< Information icon. */
        xlValidAlertStop = 1 , /*!< Stop icon. */
        xlValidAlertWarning = 2 , /*!< Warning icon. */
    };
    /*!  @brief Specifies the type of validation test to be performed in conjunction with values. Since Excel 2007.

    [MSDN documentation for XlDVType](http://msdn.microsoft.com/en-us/library/bb241228.aspx).
    */
    enum XlDVType {
        xlValidateCustom = 7 , /*!< Data is validated using an arbitrary formula. */
        xlValidateDate = 4 , /*!< Date values. */
        xlValidateDecimal = 2 , /*!< Numeric values. */
        xlValidateInputOnly = 0 , /*!< Validate only when user changes the value. */
        xlValidateList = 3 , /*!< Value must be present in a specified list. */
        xlValidateTextLength = 6 , /*!< Length of text. */
        xlValidateTime = 5 , /*!< Time values. */
        xlValidateWholeNumber = 1 , /*!< Whole numeric values. */
    };
    /*!  @brief Specifies the filter criterion. Since Excel 2007.

    [MSDN documentation for XlDynamicFilterCriteria](http://msdn.microsoft.com/en-us/library/bb241234.aspx).
    */
    enum XlDynamicFilterCriteria {
        xlFilterAboveAverage = 33 , /*!< Filter all above-average values. */
        xlFilterAllDatesInPeriodApril = 24 , /*!< Filter all dates in April. */
        xlFilterAllDatesInPeriodAugust = 28 , /*!< Filter all dates in August. */
        xlFilterAllDatesInPeriodDecember = 32 , /*!< Filter all dates in December. */
        xlFilterAllDatesInPeriodFebruray = 22 , /*!< Filter all dates in February. */
        xlFilterAllDatesInPeriodJanuary = 21 , /*!< Filter all dates in January. */
        xlFilterAllDatesInPeriodJuly = 27 , /*!< Filter all dates in July. */
        xlFilterAllDatesInPeriodJune = 26 , /*!< Filter all dates in June. */
        xlFilterAllDatesInPeriodMarch = 23 , /*!< Filter all dates in March. */
        xlFilterAllDatesInPeriodMay = 25 , /*!< Filter all dates in May. */
        xlFilterAllDatesInPeriodNovember = 31 , /*!< Filter all dates in November. */
        xlFilterAllDatesInPeriodOctober = 30 , /*!< Filter all dates in October. */
        xlFilterAllDatesInPeriodQuarter1 = 17 , /*!< Filter all dates in Quarter1. */
        xlFilterAllDatesInPeriodQuarter2 = 18 , /*!< Filter all dates in Quarter2. */
        xlFilterAllDatesInPeriodQuarter3 = 19 , /*!< Filter all dates in Quarter3. */
        xlFilterAllDatesInPeriodQuarter4 = 20 , /*!< Filter all dates in Quarter4. */
        xlFilterAllDatesInPeriodSeptember = 29 , /*!< Filter all dates in September. */
        xlFilterBelowAverage = 34 , /*!< Filter all below-average values. */
        xlFilterLastMonth = 8 , /*!< Filter all values related to last month. */
        xlFilterLastQuarter = 11 , /*!< Filter all values related to last quarter. */
        xlFilterLastWeek = 5 , /*!< Filter all values related to last week. */
        xlFilterLastYear = 14 , /*!< Filter all values related to last year. */
        xlFilterNextMonth = 9 , /*!< Filter all values related to next month. */
        xlFilterNextQuarter = 12 , /*!< Filter all values related to next quarter. */
        xlFilterNextWeek = 6 , /*!< Filter all values related to next week. */
        xlFilterNextYear = 15 , /*!< Filter all values related to next  year. */
        xlFilterThisMonth = 7 , /*!< Filter all values related to the current month.  */
        xlFilterThisQuarter = 10 , /*!< Filter all values related to the current quarter. */
        xlFilterThisWeek = 4 , /*!< Filter all values related to the current week. */
        xlFilterThisYear = 13 , /*!< Filter all values related to the current year. */
        xlFilterToday = 1 , /*!< Filter all values related to the current date. */
        xlFilterTomorrow = 3 , /*!< Filter all values related to tomorrow. */
        xlFilterYearToDate = 16 , /*!< Filter all values from today until a year ago. */
        xlFilterYesterday = 2 , /*!< Filter all values related to yesterday. */
    };
    /*!  @brief Specifies the format of the published edition. This enumeration is only for Macintosh  and should not be used. Since Excel 2007.

    [MSDN documentation for XlEditionFormat](http://msdn.microsoft.com/en-us/library/bb241237.aspx).
    */
    enum XlEditionFormat {
        xlBIFF = 2 , /*!< Binary Interchange file format. */
        xlPICT = 1 , /*!< Metafile picture structure (.wmf). */
        xlRTF = 4 , /*!< Rich Text Format (.rtf). */
        xlVALU = 8 , /*!< VALU. */
    };
    /*!This enumeration is only for Macintosh  and should not be used. Since Excel 2007.

    [MSDN documentation for XlEditionOptionsOption](http://msdn.microsoft.com/en-us/library/bb241240.aspx).
    */
    enum XlEditionOptionsOption {
        xlAutomaticUpdate = 4 , /*!< Automatic update. */
        xlCancel = 1 , /*!< Cancel. */
        xlChangeAttributes = 6 , /*!< Change attributes. */
        xlManualUpdate = 5 , /*!< Manual update. */
        xlOpenSource = 3 , /*!< Open source. */
        xlSelect = 3 , /*!< Select. */
        xlSendPublisher = 2 , /*!< Send to Microsoft Publisher. */
        xlUpdateSubscriber = 2 , /*!< Update subscriber. */
    };
    /*!  @brief Specifies the type of edition to be changed. Since Excel 2007.

    [MSDN documentation for XlEditionType](http://msdn.microsoft.com/en-us/library/bb241245.aspx).
    */
    enum XlEditionType {
        xlPublisher = 1 , /*!< Publisher */
        xlSubscriber = 2 , /*!< Subscriber */
    };
    /*!  @brief Specifies how Microsoft Office Excel 2007 handles CTRL+BREAK (or ESC or COMMAND+PERIOD) user interruptions to the running procedure. Since Excel 2007.

    [MSDN documentation for XlEnableCancelKey](http://msdn.microsoft.com/en-us/library/bb241248.aspx).
    */
    enum XlEnableCancelKey {
        xlDisabled = 0 , /*!< Cancel key trapping is completely disabled. */
        xlErrorHandler = 2 , /*!< The interrupt is sent to the running procedure as an error, trappable by an error handler set up with an On Error GoTo statement. The trappable error code is 18. */
        xlInterrupt = 1 , /*!< The current procedure is interrupted, and the user can debug or end the procedure. */
    };
    /*!  @brief Specifies what can be selected on the sheet. Since Excel 2007.

    [MSDN documentation for XlEnableSelection](http://msdn.microsoft.com/en-us/library/bb241254.aspx).
    */
    enum XlEnableSelection {
        xlNoRestrictions = 0 , /*!< Anything can be selected. */
        xlNoSelection = -4142 , /*!< Nothing can be selected. */
        xlUnlockedCells = 1 , /*!< Only unlocked cells can be selected. */
    };
    /*!  @brief Specifies the end style for error bars. Since Excel 2007.

    [MSDN documentation for XlEndStyleCap](http://msdn.microsoft.com/en-us/library/bb241257.aspx).
    */
    enum XlEndStyleCap {
        xlCap = 1 , /*!< Caps applied. */
        xlNoCap = 2 , /*!< No caps applied. */
    };
    /*!  @brief Specifies which axis values are to receive error bars. Since Excel 2007.

    [MSDN documentation for XlErrorBarDirection](http://msdn.microsoft.com/en-us/library/bb241260.aspx).
    */
    enum XlErrorBarDirection {
        xlX = -4168 , /*!< Bars run parallel to the Y axis for X-axis values. */
        xlY = 1 , /*!< Bars run parallel to the X axis for Y-axis values. */
    };
    /*!  @brief Specifies which error-bar parts to include. Since Excel 2007.

    [MSDN documentation for XlErrorBarInclude](http://msdn.microsoft.com/en-us/library/bb241264.aspx).
    */
    enum XlErrorBarInclude {
        xlErrorBarIncludeBoth = 1 , /*!< Both positive and negative error range. */
        xlErrorBarIncludeMinusValues = 3 , /*!< Only negative error range. */
        xlErrorBarIncludeNone = -4142 , /*!< No error bar range. */
        xlErrorBarIncludePlusValues = 2 , /*!< Only positive error range. */
    };
    /*!  @brief Specifies the range marked by error bars. Since Excel 2007.

    [MSDN documentation for XlErrorBarType](http://msdn.microsoft.com/en-us/library/bb241269.aspx).
    */
    enum XlErrorBarType {
        xlErrorBarTypeCustom = -4114 , /*!< Range is set by fixed values or cell values. */
        xlErrorBarTypeFixedValue = 1 , /*!< Fixed-length error bars. */
        xlErrorBarTypePercent = 2 , /*!< Percentage of range to be covered by the error bars. */
        xlErrorBarTypeStDev = -4155 , /*!< Shows range for specified number of standard deviations. */
        xlErrorBarTypeStError = 4 , /*!< Shows standard error range. */
    };
    /*!  @brief Specifies the type of error object to be retrieved from the  Since Excel 2007.

    [MSDN documentation for XlErrorChecks](http://msdn.microsoft.com/en-us/library/bb241274.aspx).
    */
    enum XlErrorChecks {
        xlEmptyCellReferences = 7 , /*!< The cell contains a formula referring to empty cells. */
        xlEvaluateToError = 1 , /*!< The cell evaluates to an error value. */
        xlInconsistentFormula = 4 , /*!< The cell contains an inconsistent formula for a region. */
        xlInconsistentListFormula = 9 , /*!< The cell contains an inconsistent formula for a list. */
        xlListDataValidation = 8 , /*!< Data in the list contains a validation error. */
        xlNumberAsText = 3 , /*!< Number entered as text. */
        xlOmittedCells = 5 , /*!< Cells omitted. */
        xlTextDate = 2 , /*!< Date entered as text. */
        xlUnlockedFormulaCells = 6 , /*!< Formula cells are unlocked. */
    };
    /*!  @brief Specifies the new access mode for the object. Since Excel 2007.

    [MSDN documentation for XlFileAccess](http://msdn.microsoft.com/en-us/library/bb241275.aspx).
    */
    enum XlFileAccess {
        xlReadOnly = 3 , /*!< Read only. */
        xlReadWrite = 2 , /*!< Read/write. */
    };
    /*!  @brief Specifies the file format when saving the worksheet. Since Excel 2007.

    [MSDN documentation for XlFileFormat](http://msdn.microsoft.com/en-us/library/office/ff198017%28v=office.14%29.aspx).
    */
    enum XlFileFormat {
        xlAddIn = 18, /*!< Microsoft Excel 97-2003 Add-In */
        xlAddIn8 = 18, /*!< Microsoft Excel 97-2003 Add-In */
        xlCSV = 6, /*!< CSV */
        xlCSVMac = 22, /*!< Macintosh CSV */
        xlCSVMSDOS = 24, /*!< MSDOS CSV */
        xlCSVUTF8 = 62, /*!< UTF8 CSV */
        xlCSVWindows = 23, /*!< Windows CSV */
        xlCurrentPlatformText = -4158, /*!< Current Platform Text */
        xlDBF2 = 7, /*!< DBF2 */
        xlDBF3 = 8, /*!< DBF3 */
        xlDBF4 = 11, /*!< DBF4 */
        xlDIF = 9, /*!< DIF */
        xlExcel12 = 50, /*!< Excel12 */
        xlExcel2 = 16, /*!< Excel2 */
        xlExcel2FarEast = 27, /*!< Excel2 FarEast */
        xlExcel3 = 29, /*!< Excel3 */
        xlExcel4 = 33, /*!< Excel4 */
        xlExcel4Workbook = 35, /*!< Excel4 Workbook */
        xlExcel5 = 39, /*!< Excel5 */
        xlExcel7 = 39, /*!< Excel7 */
        xlExcel8 = 56, /*!< Excel8 */
        xlExcel9795 = 43, /*!< Excel9795 */
        xlHtml = 44, /*!< HTML format */
        xlIntlAddIn = 26, /*!< International Add-In */
        xlIntlMacro = 25, /*!< International Macro */
        xlOpenDocumentSpreadsheet = 60, /*!< OpenDocument Spreadsheet */
        xlOpenXMLAddIn = 55, /*!< Open XML Add-In */
        xlOpenXMLTemplate = 54, /*!< Open XML Template */
        xlOpenXMLTemplateMacroEnabled = 53, /*!< Open XML Template Macro Enabled */
        xlOpenXMLWorkbook = 51, /*!< Open XML Workbook */
        xlOpenXMLWorkbookMacroEnabled = 52, /*!< Open XML Workbook Macro Enabled */
        xlSYLK = 2, /*!< SYLK */
        xlTemplate = 17, /*!< Template */
        xlTemplate8 = 17, /*!< Template 8 */
        xlTextMac = 19, /*!< Macintosh Text */
        xlTextMSDOS = 21, /*!< MSDOS Text */
        xlTextPrinter = 36, /*!< Printer Text */
        xlTextWindows = 20, /*!< Windows Text */
        xlUnicodeText = 42, /*!< Unicode Text */
        xlWebArchive = 45, /*!< Web Archive */
        xlWJ2WD1 = 14, /*!< WJ2WD1 */
        xlWJ3 = 40, /*!< WJ3 */
        xlWJ3FJ3 = 41, /*!< WJ3FJ3 */
        xlWK1 = 5, /*!< WK1 */
        xlWK1ALL = 31, /*!< WK1ALL */
        xlWK1FMT = 30, /*!< WK1FMT */
        xlWK3 = 15, /*!< WK3 */
        xlWK3FM3 = 32, /*!< WK3FM3 */
        xlWK4 = 38, /*!< WK4 */
        xlWKS = 4, /*!< Worksheet */
        xlWorkbookDefault = 51, /*!< Workbook default */
        xlWorkbookNormal = -4143, /*!< Workbook normal */
        xlWorks2FarEast = 28, /*!< Works2 FarEast */
        xlWQ1 = 34, /*!< WQ1 */
        xlXMLSpreadsheet = 46, /*!< XML Spreadsheet */
    };
    /*!  @brief Specifies how to copy the range. Since Excel 2007.

    [MSDN documentation for XlFillWith](http://msdn.microsoft.com/en-us/library/bb241282.aspx).
    */
    enum XlFillWith {
        xlFillWithAll = -4104 , /*!< Copy contents and formats. */
        xlFillWithContents = 2 , /*!< Copy contents only. */
        xlFillWithFormats = -4122 , /*!< Copy formats only. */
    };
    /*!  @brief Specifies whether data is to be copied or left in place during a filter operation. Since Excel 2007.

    [MSDN documentation for XlFilterAction](http://msdn.microsoft.com/en-us/library/bb241284.aspx).
    */
    enum XlFilterAction {
        xlFilterCopy = 2 , /*!< Copy filtered data to new location. */
        xlFilterInPlace = 1 , /*!< Leave data in place. */
    };
    /*!  @brief Specifies how dates should be filtered in the specified period. Since Excel 2007.

    [MSDN documentation for XlFilterAllDatesInPeriod](http://msdn.microsoft.com/en-us/library/bb241286.aspx).
    */
    enum XlFilterAllDatesInPeriod {
        xlFilterAllDatesInPeriodDay = 2 , /*!< Filter all dates for the specified date. */
        xlFilterAllDatesInPeriodHour = 3 , /*!< Filter all dates for the specified hour. */
        xlFilterAllDatesInPeriodMinute = 4 , /*!< Filter all dates until the specified minute. */
        xlFilterAllDatesInPeriodMonth = 1 , /*!< Filter all dates for the specified month. */
        xlFilterAllDatesInPeriodSecond = 5 , /*!< Filter all dates until the specified second. */
        xlFilterAllDatesInPeriodYear = 0 , /*!< Filter all dates for the specified year. */
    };
    /*!  @brief Specifies the type of data to search. Since Excel 2007.

    [MSDN documentation for XlFindLookIn](http://msdn.microsoft.com/en-us/library/bb241289.aspx).
    */
    enum XlFindLookIn {
        xlComments = -4144 , /*!< Comments. */
        xlFormulas = -4123 , /*!< Formulas. */
        xlValues = -4163 , /*!< Values. */
    };
    /*!  @brief Specifies the quality of speadsheets saved in different fixed formats. Since Excel 2007.

    [MSDN documentation for XlFixedFormatQuality](http://msdn.microsoft.com/en-us/library/bb241292.aspx).
    */
    enum XlFixedFormatQuality {
        xlQualityMinimum = 1 , /*!< Minimum quality */
        xlQualityStandard = 0 , /*!< Standard quality */
    };
    /*!  @brief Specifies the type of file format. Since Excel 2007.

    [MSDN documentation for XlFixedFormatType](http://msdn.microsoft.com/en-us/library/bb241296.aspx).
    */
    enum XlFixedFormatType {
        xlTypePDF = 0 , /*!< "PDF" — Portable Document Format file (.pdf). */
        xlTypeXPS = 1 , /*!< "XPS" — XPS Document (.xps).
                        */
    };
    /*!  @brief Specifies the operator to use to compare a formula against the value in a cell or, for

    [MSDN documentation for XlFormatConditionOperator](http://msdn.microsoft.com/en-us/library/bb241299.aspx).
    */
    enum XlFormatConditionOperator {
        xlBetween = 1 , /*!< Between. Can be used only if two formulas are provided. */
        xlEqual = 3 , /*!< Equal. */
        xlGreater = 5 , /*!< Greater than. */
        xlGreaterEqual = 7 , /*!< Greater than or equal to. */
        xlLess = 6 , /*!< Less than. */
        xlLessEqual = 8 , /*!< Less than or equal to. */
        xlNotBetween = 2 , /*!< Not between. Can be used only if two formulas are provided. */
        xlNotEqual = 4 , /*!< Not equal. */
    };
    /*!  @brief Specifies whether the conditional format is based on a cell value or an expression. Since Excel 2007.

    [MSDN documentation for XlFormatConditionType](http://msdn.microsoft.com/en-us/library/bb241301.aspx).
    */
    enum XlFormatConditionType {
        xlAboveAverageCondition = 12 , /*!< Above average condition */
        xlBlanksCondition = 10 , /*!< Blanks condition */
        xlCellValue = 1 , /*!< Cell value */
        xlColorScale = 3 , /*!< Color scale */
        xlDatabar = 4 , /*!< Databar */
        xlErrorsCondition = 16 , /*!< Errors condition */
        xlExpression = 2 , /*!< Expression */
        XlIconSet = 6 , /*!< Icon set */
        xlNoBlanksCondition = 13 , /*!< No blanks condition */
        xlNoErrorsCondition = 17 , /*!< No errors condition */
        xlTextString = 9 , /*!< Text string */
        xlTimePeriod = 11 , /*!< Time period */
        xlTop10 = 5 , /*!< Top 10 values */
        xlUniqueValues = 8 , /*!< Unique values */
    };
    /*!  @brief Specifies the types of format filters. Since Excel 2007.

    [MSDN documentation for XlFormatFilterTypes](http://msdn.microsoft.com/en-us/library/bb241304.aspx).
    */
    enum XlFormatFilterTypes {
        FilterBottom = 0 , /*!< Bottom. */
        FilterBottomPercent = 2 , /*!< Bottom Percent. */
        FilterTop = 1 , /*!< Top. */
        FilterTopPercent = 3 , /*!< Top Percent. */
    };
    /*!  @brief Specifies the type of the form control. Since Excel 2007.

    [MSDN documentation for XlFormControl](http://msdn.microsoft.com/en-us/library/bb241306.aspx).
    */
    enum XlFormControl {
        xlButtonControl = 0 , /*!< Button. */
        xlCheckBox = 1 , /*!< Check box. */
        xlDropDown = 2 , /*!< Combo box. */
        xlEditBox = 3 , /*!< Text box. */
        xlGroupBox = 4 , /*!< Group box. */
        xlLabel = 5 , /*!< Label. */
        xlListBox = 6 , /*!< List box. */
        xlOptionButton = 7 , /*!< Option button. */
        xlScrollBar = 8 , /*!< Scroll bar. */
        xlSpinner = 9 , /*!< Spinner. */
    };
    /*!  @brief Specifies the formula label type for the specified range. Since Excel 2007.

    [MSDN documentation for XlFormulaLabel](http://msdn.microsoft.com/en-us/library/bb241310.aspx).
    */
    enum XlFormulaLabel {
        xlColumnLabels = 2 , /*!< Column labels only. */
        xlMixedLabels = 3 , /*!< Row and column labels. */
        xlNoLabels = -4142 , /*!< No labels. */
        xlRowLabels = 1 , /*!< Row labels only. */
    };
    /*!  @brief Specifies the type of table references. Since Excel 2007.

    [MSDN documentation for XlGenerateTableRefs](http://msdn.microsoft.com/en-us/library/bb241312.aspx).
    */
    enum XlGenerateTableRefs {
        xlA1TableRefs = 0 , /*!< A1 Table References. */
        xlTableNames = 1 , /*!< Table Names. */
    };
    /*!  @brief Specifies the type of  Since Excel 2007.

    [MSDN documentation for XlGradientFillType](http://msdn.microsoft.com/en-us/library/bb257010.aspx).
    */
    enum XlGradientFillType {
        GradientFillLinear = 0 , /*!< Gradient is filled in a straight line. */
        GradientFillPath = 1 , /*!< Gradient is filled in a non-linear or curved path. */
    };
    /*!  @brief Specifies the horizontal alignment for the object. Since Excel 2007.

    [MSDN documentation for XlHAlign](http://msdn.microsoft.com/en-us/library/bb241313.aspx).
    */
    enum XlHAlign {
        xlHAlignCenter = -4108 , /*!< Center. */
        xlHAlignCenterAcrossSelection = 7 , /*!< Center across selection. */
        xlHAlignDistributed = -4117 , /*!< Distribute. */
        xlHAlignFill = 5 , /*!< Fill. */
        xlHAlignGeneral = 1 , /*!< Align according to data type. */
        xlHAlignJustify = -4130 , /*!< Justify. */
        xlHAlignLeft = -4131 , /*!< Left. */
        xlHAlignRight = -4152 , /*!< Right. */
    };
    /*!  @brief Specifies the mode for the Hebrew spelling checker. Since Excel 2007.

    [MSDN documentation for XlHebrewModes](http://msdn.microsoft.com/en-us/library/bb241317.aspx).
    */
    enum XlHebrewModes {
        xlHebrewFullScript = 0 , /*!< The conventional script type as required by the Hebrew Language Academy when writing text without diacritics. */
        xlHebrewMixedAuthorizedScript = 3 , /*!< The Hebrew traditional script.  */
        xlHebrewMixedScript = 2 , /*!< In this mode the speller accepts any word recognized as Hebrew, whether in Full Script, Partial Script, or any unconventional spelling variation that is known to the speller.  */
        xlHebrewPartialScript = 1 , /*!< In this mode the speller accepts words both in Full Script and Partial Script. Some words will be flagged since this spelling is not authorized in either Full script or Partial script. */
    };
    /*!  @brief Specifies which set of changes is shown in a shared workbook. Since Excel 2007.

    [MSDN documentation for XlHighlightChangesTime](http://msdn.microsoft.com/en-us/library/bb241320.aspx).
    */
    enum XlHighlightChangesTime {
        xlAllChanges = 2 , /*!< Show all changes. */
        xlNotYetReviewed = 3 , /*!< Show only changes not yet reviewed. */
        xlSinceMyLastSave = 1 , /*!< Show changes made since last save by last user. */
    };
    /*!  @brief Specifies the type of HTML generated by Office Excel 2007 when you save the specified item to a Web page and whether the item is static or interactive. Since Excel 2007.

    [MSDN documentation for XlHtmlType](http://msdn.microsoft.com/en-us/library/bb241321.aspx).
    */
    enum XlHtmlType {
        xlHtmlCalc = 1 , /*!< Use the Spreadsheet component. Deprecated in Excel 2007. */
        xlHtmlChart = 3 , /*!< Use the Chart component. Deprecated in Excel 2007. */
        xlHtmlList = 2 , /*!< Use the PivotTable component. Deprecated in Excel 2007. */
        xlHtmlStatic = 0 , /*!< Use static (noninteractive) HTML for viewing only. */
    };

/*!  @brief Specifies the icon for a criterion in an icon set conditional formatting rule. Since Excel 2010.

    [MSDN documentation for XlIcon](http://msdn.microsoft.com/en-us/library/office/ff193828%28v=office.14%29.aspx).
    */
    enum XlIcon {
        xlIconGreenCheck = 22, /*!< Green Check . */
        xlIconGreenCheckSymbol = 19, /*!< Green Check Symbol . */
        xlIconGreenCircle = 10, /*!< Green Circle . */
        xlIconGreenFlag = 7, /*!< Green Flag . */
        xlIconGreenTrafficLight = 14, /*!< Green Traffic Light . */
        xlIconGreenUpArrow = 1, /*!< Green Up Arrow . */
        xlIconGreenUpTriangle = 45, /*!< Green Up Triangle . */
        xlIconHalfGoldStar = 43, /*!< Half Gold Star . */
        xlIconNoCellIcon = -1, /*!< No Cell Icon . */
        xlIconPinkCircle = 30, /*!< Pink Circle . */
        xlIconRedCircle = 29, /*!< Red Circle . */
        xlIconRedCircleWithBorder = 12, /*!< Red Circle With Border . */
        xlIconRedCross = 24, /*!< Red Cross . */
        xlIconRedCrossSymbol = 21, /*!< Red Cross Symbol . */
        xlIconRedDiamond = 18, /*!< Red Diamond . */
        xlIconRedDownArrow = 3, /*!< Red Down Arrow . */
        xlIconRedDownTriangle = 47, /*!< Red Down Triangle . */
        xlIconRedFlag = 9, /*!< Red Flag . */
        xlIconRedTrafficLight = 16, /*!< Red Traffic Light . */
        xlIconSilverStar = 44, /*!< Silver Star . */
        xlIconWhiteCircleAllWhiteQuarters = 36, /*!< White Circle (All White Quarters) . */
        xlIconYellowCircle = 11, /*!< Yellow Circle . */
        xlIconYellowDash = 46, /*!< Yellow Dash . */
        xlIconYellowDownInclineArrow = 26, /*!< Yellow Down Incline Arrow . */
        xlIconYellowExclamation = 23, /*!< Yellow Exclamation . */
        xlIconYellowExclamationSymbol = 20, /*!< Yellow Exclamation Symbol . */
        xlIconYellowFlag = 8, /*!< Yellow Flag . */
        xlIconYellowSideArrow = 2, /*!< Yellow Side Arrow . */
        xlIconYellowTrafficLight = 15, /*!< Yellow Traffic Light . */
        xlIconYellowTriangle = 17, /*!< Yellow Triangle . */
        xlIconYellowUpInclineArrow = 25, /*!< Yellow Up Incline Arrow . */
    };    
    
    
    /*!  @brief Specifies the type of icon set. Since Excel 2007.

    [MSDN documentation for XlIconSet](http://msdn.microsoft.com/en-us/library/office/ff197528%28v=office.14%29.aspx).
    */
    enum XlIconSetE {
        xl3Arrows = 1, /*!< 3 Arrows */
        xl3ArrowsGray = 2, /*!< 3 Arrows Gray */
        xl3Flags = 3, /*!< 3 Flags */
        xl3Signs = 6, /*!< 3 Signs */
        xl3Symbols = 7, /*!< 3 Symbols */
        xl3TrafficLights1 = 4, /*!< 3 Traffic Lights 1 */
        xl3TrafficLights2 = 5, /*!< 3 Traffic Lights 2 */
        xl4Arrows = 8, /*!< 4 Arrows */
        xl4ArrowsGray = 9, /*!< 4 Arrows Gray */
        xl4CRV = 11, /*!< 4 CRV */
        xl4RedToBlack = 10, /*!< 4 Red To Black */
        xl4TrafficLights = 12, /*!< 4 Traffic Lights */
        xl5Arrows = 13, /*!< 5 Arrows */
        xl5ArrowsGray = 14, /*!< 5 Arrows Gray */
        xl5CRV = 15, /*!< 5 CRV */
        xl5Quarters = 16, /*!< 5 Quarters */
    };
    /*!  @brief Specifies the description of the Japanese input rules. Since Excel 2007.

    [MSDN documentation for XlIMEMode](http://msdn.microsoft.com/en-us/library/bb241327.aspx).
    */
    enum XlIMEMode {
        xlIMEModeAlpha = 8 , /*!< Half-width alphanumeric. */
        xlIMEModeAlphaFull = 7 , /*!< Full-width alphanumeric. */
        xlIMEModeDisable = 3 , /*!< Disable. */
        xlIMEModeHangul = 10 , /*!< Hangul. */
        xlIMEModeHangulFull = 9 , /*!< Full-width Hangul. */
        xlIMEModeHiragana = 4 , /*!< Hiragana. */
        xlIMEModeKatakana = 5 , /*!< Katakana. */
        xlIMEModeKatakanaHalf = 6 , /*!< Half-width Katakana. */
        xlIMEModeNoControl = 0 , /*!< No control. */
        xlIMEModeOff = 2 , /*!< Off (English mode). */
        xlIMEModeOn = 1 , /*!< Mode on. */
    };
    /*!  @brief Specifies the format in which to return data from a database. Since Excel 2007.

    [MSDN documentation for XlImportDataAs](http://msdn.microsoft.com/en-us/library/bb241330.aspx).
    */
    enum XlImportDataAs {
        xlPivotTableReport = 1 , /*!< Returns the data as a PivotTable. */
        xlQueryTable = 0 , /*!< Returns the data as a QueryTable. */
    };
    /*!  @brief Specifies from where to copy the format for inserted rows. Since Excel 2007.

    [MSDN documentation for XlInsertFormatOrigin](http://msdn.microsoft.com/en-us/library/bb241334.aspx).
    */
    enum XlInsertFormatOrigin {
        xlFormatFromLeftOrAbove = 0 , /*!< Copy the format from cells above and/or to the left. */
        xlFormatFromRightOrBelow = 1 , /*!< Copy the format from cells below and/or to the right. */
    };
    /*!  @brief Specifies the direction in which to shift cells during an insertion. Since Excel 2007.

    [MSDN documentation for XlInsertShiftDirection](http://msdn.microsoft.com/en-us/library/bb241336.aspx).
    */
    enum XlInsertShiftDirection {
        xlShiftDown = -4121 , /*!< Shift cells down. */
        xlShiftToRight = -4161 , /*!< Shift cells to the right. */
    };
    /*!  @brief Specifies the way the specified PivotTable items appear—in table format or in outline format. Since Excel 2007.

    [MSDN documentation for XlLayoutFormType](http://msdn.microsoft.com/en-us/library/bb241338.aspx).
    */
    enum XlLayoutFormType {
        xlOutline = 1 , /*!< The LayoutSubtotalLocation property specifies where the subtotal appears in the PivotTable report. */
        xlTabular = 0 , /*!< Default. */
    };
    /*!  @brief Specifies the type of layout row. Since Excel 2007.

    [MSDN documentation for XlLayoutRowType](http://msdn.microsoft.com/en-us/library/bb241341.aspx).
    */
    enum XlLayoutRowType {
        xlCompactRow = 0 , /*!< Compact Row */
        xlOutlineRow = 2 , /*!< Outline Row */
        xlTabularRow = 1 , /*!< Tabular Row */
    };
    /*!  @brief Specifies the position of the legend on a chart.

    [MSDN documentation for XlLegendPosition](http://msdn.microsoft.com/en-us/library/bb241345.aspx).
    */
    enum XlLegendPosition {
        xlLegendPositionBottom = -4107 , /*!< Below the chart. */
        xlLegendPositionCorner = 2 , /*!< In the upper right-hand corner of the chart border. */
        xlLegendPositionCustom = -4161 , /*!< A custom position. */
        xlLegendPositionLeft = -4131 , /*!< Left of the chart. */
        xlLegendPositionRight = -4152 , /*!< Right of the chart. */
        xlLegendPositionTop = -4160 , /*!< Above the chart. */
    };
    /*!  @brief Specifies the line style for the border.

    [MSDN documentation for XlLineStyle](http://msdn.microsoft.com/en-us/library/bb241348.aspx).
    */
    enum XlLineStyle {
        xlContinuous = 1 , /*!< Continuous line. */
        xlDash = -4115 , /*!< Dashed line. */
        xlDashDot = 4 , /*!< Alternating dashes and dots. */
        xlDashDotDot = 5 , /*!< Dash followed by two dots. */
        xlDot = -4118 , /*!< Dotted line. */
        xlDouble = -4119 , /*!< Double line. */
        xlLineStyleNone = -4142 , /*!< No line. */
        xlSlantDashDot = 13 , /*!< Slanted dashes. */
    };
    /*!  @brief Specifies the type of link. Since Excel 2007.

    [MSDN documentation for XlLink](http://msdn.microsoft.com/en-us/library/bb241350.aspx).
    */
    enum XlLink {
        xlExcelLinks = 1 , /*!< The link is to an Excel worksheet. */
        xlOLELinks = 2 , /*!< The link is to an OLE source. */
        xlPublishers = 5 , /*!< Macintosh only. */
        xlSubscribers = 6 , /*!< Macintosh only. */
    };
    /*!  @brief Specifies the type of information the link will return. Since Excel 2007.

    [MSDN documentation for XlLinkInfo](http://msdn.microsoft.com/en-us/library/bb241352.aspx).
    */
    enum XlLinkInfo {
        xlEditionDate = 2 , /*!< Applies only to editions in the Macintosh operating system. */
        xlLinkInfoStatus = 3 , /*!< Returns the link status. */
        xlUpdateState = 1 , /*!< Specifies whether the link updates automatically or manually. */
    };
    /*!  @brief Specifies the type of link. Since Excel 2007.

    [MSDN documentation for XlLinkInfoType](http://msdn.microsoft.com/en-us/library/bb241353.aspx).
    */
    enum XlLinkInfoType {
        xlLinkInfoOLELinks = 2 , /*!< OLE or DDE server */
        xlLinkInfoPublishers = 5 , /*!< Publisher */
        xlLinkInfoSubscribers = 6 , /*!< Subscriber */
    };
    /*!  @brief Specifies the status of a link. Since Excel 2007.

    [MSDN documentation for XlLinkStatus](http://msdn.microsoft.com/en-us/library/bb241354.aspx).
    */
    enum XlLinkStatus {
        xlLinkStatusCopiedValues = 10 , /*!< Copied values. */
        xlLinkStatusIndeterminate = 5 , /*!< Unable to determine status. */
        xlLinkStatusInvalidName = 7 , /*!< Invalid name. */
        xlLinkStatusMissingFile = 1 , /*!< File missing. */
        xlLinkStatusMissingSheet = 2 , /*!< Sheet missing. */
        xlLinkStatusNotStarted = 6 , /*!< Not started. */
        xlLinkStatusOK = 0 , /*!< No errors. */
        xlLinkStatusOld = 3 , /*!< Status may be out of date. */
        xlLinkStatusSourceNotCalculated = 4 , /*!< Not yet calculated. */
        xlLinkStatusSourceNotOpen = 8 , /*!< Not open. */
        xlLinkStatusSourceOpen = 9 , /*!< Source document is open. */
    };
    /*!  @brief Specifies the type of link. Since Excel 2007.

    [MSDN documentation for XlLinkType](http://msdn.microsoft.com/en-us/library/bb241356.aspx).
    */
    enum XlLinkType {
        xlLinkTypeExcelLinks = 1 , /*!< A link to a Microsoft Office Excel source. */
        xlLinkTypeOLELinks = 2 , /*!< A link to an OLE source. */
    };
    /*!  @brief Specifies the conflict resolution options for updating a list on a Microsoft Windows SharePoint Services site with the changes made to a list in a Microsoft Office Excel worksheet. Since Excel 2007.

    [MSDN documentation for XlListConflict](http://msdn.microsoft.com/en-us/library/bb241358.aspx).
    */
    enum XlListConflict {
        xlListConflictDialog = 0 , /*!< Display a dialog box that allows the user to choose how to resolve conflicts. */
        xlListConflictDiscardAllConflicts = 2 , /*!< Accept the version of the data stored on the SharePoint site. */
        xlListConflictError = 3 , /*!< Raise an error if a conflict occurs. */
        xlListConflictRetryAllConflicts = 1 , /*!< Overwrite the version of the data stored on the SharePoint site. */
    };
    /*!  @brief Specifies the data type of a list column connected to a Microsoft Windows SharePoint Services site. Since Excel 2007.

    [MSDN documentation for XlListDataType](http://msdn.microsoft.com/en-us/library/bb241360.aspx).
    */
    enum XlListDataType {
        xlListDataTypeCheckbox = 9 , /*!< Check box. */
        xlListDataTypeChoice = 6 , /*!< Single-choice field. */
        xlListDataTypeChoiceMulti = 7 , /*!< Multiple-choice field. */
        xlListDataTypeCounter = 11 , /*!< Counter. */
        xlListDataTypeCurrency = 4 , /*!< Currency. */
        xlListDataTypeDateTime = 5 , /*!< Date/time. */
        xlListDataTypeHyperLink = 10 , /*!< Hyperlink. */
        xlListDataTypeListLookup = 8 , /*!< Lookup list. */
        xlListDataTypeMultiLineRichText = 12 , /*!< Rich text format with multiple lines. */
        xlListDataTypeMultiLineText = 2 , /*!< Plain text with multiple lines. */
        xlListDataTypeNone = 0 , /*!< Type not specified. */
        xlListDataTypeNumber = 3 , /*!< Numerical. */
        xlListDataTypeText = 1 , /*!< Plain text. */
    };
    /*!  @brief Specifies the current source of the list. Since Excel 2007.

    [MSDN documentation for XlListObjectSourceType](http://msdn.microsoft.com/en-us/library/bb241364.aspx).
    */
    enum XlListObjectSourceType {
        xlSrcExternal = 0 , /*!< External data source (Microsoft Windows SharePoint Services site). */
        xlSrcQuery = 3 , /*!< Query */
        xlSrcRange = 1 , /*!< Range */
        xlSrcXml = 2 , /*!< XML */
    };
    /*!  @brief Specifies the part of the PivotTable report that contains the upper-left corner of a range. Since Excel 2007.

    [MSDN documentation for XlLocationInTable](http://msdn.microsoft.com/en-us/library/bb241366.aspx).
    */
    enum XlLocationInTable {
        xlColumnHeader = -4110 , /*!< Column header */
        xlColumnItem = 5 , /*!< Column item */
        xlDataHeader = 3 , /*!< Data header */
        xlDataItem = 7 , /*!< Data item */
        xlPageHeader = 2 , /*!< Page header */
        xlPageItem = 6 , /*!< Page item */
        xlRowHeader = -4153 , /*!< Row header */
        xlRowItem = 4 , /*!< Row item */
        xlTableBody = 8 , /*!< Table body */
    };
    /*!  @brief Specifies whether a match is made against the whole of the search text or any part of the search text. Since Excel 2007.

    [MSDN documentation for XlLookAt](http://msdn.microsoft.com/en-us/library/bb241368.aspx).
    */
    enum XlLookAt {
        xlPart = 2 , /*!< Match against any part of the search text. */
        xlWhole = 1 , /*!< Match against the whole of the search text. */
    };
    /*!  @brief Specifies what to look for in searches. Since Excel 2007.

    [MSDN documentation for XlLookFor](http://msdn.microsoft.com/en-us/library/bb241370.aspx).
    */
    enum XlLookFor {
        LookForBlanks = 0 , /*!< Blanks */
        LookForErrors = 1 , /*!< Errors */
        LookForFormulas = 2 , /*!< Formulas */
    };
    /*!  @brief Specifies the mail system that is installed on the host computer. Since Excel 2007.

    [MSDN documentation for XlMailSystem](http://msdn.microsoft.com/en-us/library/bb241372.aspx).
    */
    enum XlMailSystem {
        xlMAPI = 1 , /*!< MAPI-complaint system */
        xlNoMailSystem = 0 , /*!< No mail system */
        xlPowerTalk = 2 , /*!< PowerTalk mail system */
    };
    /*!  @brief Specifies the marker style for a point or series in a line chart, scatter chart, or radar chart. Since Excel 2007.

    [MSDN documentation for XlMarkerStyle](http://msdn.microsoft.com/en-us/library/bb241374.aspx).
    */
    enum XlMarkerStyle {
        xlMarkerStyleAutomatic = -4105 , /*!< Automatic markers */
        xlMarkerStyleCircle = 8 , /*!< Circular markers */
        xlMarkerStyleDash = -4115 , /*!< Long bar markers */
        xlMarkerStyleDiamond = 2 , /*!< Diamond-shaped markers */
        xlMarkerStyleDot = -4118 , /*!< Short bar markers */
        xlMarkerStyleNone = -4142 , /*!< No markers */
        xlMarkerStylePicture = -4147 , /*!< Picture markers */
        xlMarkerStylePlus = 9 , /*!< Square markers with a plus sign */
        xlMarkerStyleSquare = 1 , /*!< Square markers */
        xlMarkerStyleStar = 5 , /*!< Square markers with an asterisk */
        xlMarkerStyleTriangle = 3 , /*!< Triangular markers */
        xlMarkerStyleX = -4168 , /*!< Square markers with an X */
    };
    /*!  @brief Specifies the measurement units. Since Excel 2007.

    [MSDN documentation for XlMeasurementUnits](http://msdn.microsoft.com/en-us/library/bb241375.aspx).
    */
    enum XlMeasurementUnits {
        xlCentimeters = 1 , /*!< Centimeters */
        xlInches = 0 , /*!< Inches */
        xlMillimeters = 2 , /*!< Millimeters */
    };
    /*!  @brief Specifies which mouse button was pressed. Since Excel 2007.

    [MSDN documentation for XlMouseButton](http://msdn.microsoft.com/en-us/library/bb241378.aspx).
    */
    enum XlMouseButton {
        xlNoButton = 0 , /*!< No button was pressed. */
        xlPrimaryButton = 1 , /*!< The primary button (normally the left mouse button) was pressed. */
        xlSecondaryButton = 2 , /*!< The secondary button (normally the right mouse button) was pressed. */
    };
    /*!  @brief Specifies the appearance of the mouse pointer in Excel 2007. Since Excel 2007.

    [MSDN documentation for XlMousePointer](http://msdn.microsoft.com/en-us/library/bb241380.aspx).
    */
    enum XlMousePointer {
        xlDefault = -4143 , /*!< The default pointer. */
        xlIBeam = 3 , /*!< The I-beam pointer. */
        xlNorthwestArrow = 1 , /*!< The northwest-arrow pointer. */
        xlWait = 2 , /*!< The hourglass pointer. */
    };
    /*!  @brief Specifies a Microsoft application. Since Excel 2007.

    [MSDN documentation for XlMSApplication](http://msdn.microsoft.com/en-us/library/bb241382.aspx).
    */
    enum XlMSApplication {
        xlMicrosoftAccess = 4 , /*!< Microsoft Office Access */
        xlMicrosoftFoxPro = 5 , /*!< Microsoft FoxPro */
        xlMicrosoftMail = 3 , /*!< Microsoft Office Outlook */
        xlMicrosoftPowerPoint = 2 , /*!< Microsoft Office PowerPoint */
        xlMicrosoftProject = 6 , /*!< Microsoft Office Project */
        xlMicrosoftSchedulePlus = 7 , /*!< Microsoft Schedule Plus */
        xlMicrosoftWord = 1 , /*!< Microsoft Office Word */
    };

        
    /*!  @brief Specifies the horizontal overflow setting for a text frame. Since Excel 2010.

    [MSDN documentation for XlOartHorizontalOverflow](http://msdn.microsoft.com/en-us/library/office/ff195402%28v=office.14%29.aspx).
    */
    enum XlOartHorizontalOverflow {
        xlOartHorizontalOverflowClip = 1, /*!< Hide text that does not fit horizontally in the text frame. . */
        xlOartHorizontalOverflowOverflow = 0, /*!< Allow text to overflow the text frame horizontally. . */
    };

        /*!  @brief Specifies the vertical overflow setting for a text frame. Since Excel 2010.

    [MSDN documentation for XlOartVerticalOverflow](http://msdn.microsoft.com/en-us/library/office/ff837847%28v=office.14%29.aspx).
    */
    enum XlOartVerticalOverflow {
        xlOartVerticalOverflowClip = 1, /*!< Hide text that does not fit vertically within the text frame. . */
        xlOartVerticalOverflowEllipsis = 2, /*!< Hide text that does not fit vertically within the text frame, and add an ellipsis (...) at the end of the visible text. . */
        xlOartVerticalOverflowOverflow = 0, /*!< Allow text to overflow the text frame vertically (can be from the top, bottom, or both depending on the text alignment). . */
    };

    
    /*!  @brief Specifies the way a chart is scaled to fit on a page. Since Excel 2007.

    [MSDN documentation for XlObjectSize](http://msdn.microsoft.com/en-us/library/bb241384.aspx).
    */
    enum XlObjectSize {
        xlFitToPage = 2 , /*!< Print the chart as large as possible, while retaining the chart's height-to-width ratio as shown on the screen. */
        xlFullPage = 3 , /*!< Print the chart to fit the page, adjusting the height-to-width ratio as necessary. */
        xlScreenSize = 1 , /*!< Print the chart the same size as it appears on the screen. */
    };
    /*!  @brief Specifies the OLE object type. Since Excel 2007.

    [MSDN documentation for XlOLEType](http://msdn.microsoft.com/en-us/library/bb241386.aspx).
    */
    enum XlOLEType {
        xlOLEControl = 2 , /*!< ActiveX control */
        xlOLEEmbed = 1 , /*!< Embedded OLE object */
        xlOLELink = 0 , /*!< Linked OLE object */
    };
    /*!  @brief Specifies the verb on which the server of the OLE object should act. Since Excel 2007.

    [MSDN documentation for XlOLEVerb](http://msdn.microsoft.com/en-us/library/bb241388.aspx).
    */
    enum XlOLEVerb {
        xlVerbOpen = 2 , /*!< Open the object. */
        xlVerbPrimary = 1 , /*!< Perform the primary action for the server. */
    };
    /*!  @brief Specifies the order in which cells are processed. Since Excel 2007.

    [MSDN documentation for XlOrder](http://msdn.microsoft.com/en-us/library/bb241390.aspx).
    */
    enum XlOrder {
        xlDownThenOver = 1 , /*!< Process down the rows before processing across pages or page fields to the right. */
        xlOverThenDown = 2 , /*!< Process across pages or page fields to the right before moving down the rows. */
    };
    /*!  @brief Specifies the text orientation. Since Excel 2007.

    [MSDN documentation for XlOrientation](http://msdn.microsoft.com/en-us/library/bb241392.aspx).
    */
    enum XlOrientation {
        xlDownward = -4170 , /*!< Text runs downward. */
        xlHorizontal = -4128 , /*!< Text runs horizontally. */
        xlUpward = -4171 , /*!< Text runs upward. */
        xlVertical = -4166 , /*!< Text runs downward and is centered in the cell. */
    };
    /*!  @brief Specifies page break location in the worksheet. Since Excel 2007.

    [MSDN documentation for XlPageBreak](http://msdn.microsoft.com/en-us/library/bb241393.aspx).
    */
    enum XlPageBreak {
        xlPageBreakAutomatic = -4105 , /*!< Excel will automatically add page breaks. */
        xlPageBreakManual = -4135 , /*!< Page breaks are manually inserted. */
        xlPageBreakNone = -4142 , /*!< Page breaks are not inserted in the worksheet. */
    };
    /*!  @brief Specifies whether a page break is full screen or applies only within the print area. Since Excel 2007.

    [MSDN documentation for XlPageBreakExtent](http://msdn.microsoft.com/en-us/library/bb241395.aspx).
    */
    enum XlPageBreakExtent {
        xlPageBreakFull = 1 , /*!< Full screen. */
        xlPageBreakPartial = 2 , /*!< Only within print area. */
    };
    /*!  @brief Specifies the page orientation when the worksheet is printed. Since Excel 2007.

    [MSDN documentation for XlPageOrientation](http://msdn.microsoft.com/en-us/library/bb241397.aspx).
    */
    enum XlPageOrientation {
        xlLandscape = 2 , /*!< Landscape mode. */
        xlPortrait = 1 , /*!< Portrait mode. */
    };
    /*!  @brief Specifies the size of the paper. Since Excel 2007.

    [MSDN documentation for XlPaperSize](http://msdn.microsoft.com/en-us/library/bb241398.aspx).
    */
    enum XlPaperSize {
        xlPaper10x14 = 16 , /*!< 10 in. x 14 in. */
        xlPaper11x17 = 17 , /*!< 11 in. x 17 in. */
        xlPaperA3 = 8 , /*!< A3 (297 mm x 420 mm) */
        xlPaperA4 = 9 , /*!< A4 (210 mm x 297 mm) */
        xlPaperA4Small = 10 , /*!< A4 Small (210 mm x 297 mm) */
        xlPaperA5 = 11 , /*!< A5 (148 mm x 210 mm) */
        xlPaperB4 = 12 , /*!< B4 (250 mm x 354 mm) */
        xlPaperB5 = 13 , /*!< A5 (148 mm x 210 mm) */
        xlPaperCsheet = 24 , /*!< C size sheet */
        xlPaperDsheet = 25 , /*!< D size sheet */
        xlPaperEnvelope10 = 20 , /*!< Envelope #10 (4-1/8 in. x 9-1/2 in.) */
        xlPaperEnvelope11 = 21 , /*!< Envelope #11 (4-1/2 in. x 10-3/8 in.) */
        xlPaperEnvelope12 = 22 , /*!< Envelope #12 (4-1/2 in. x 11 in.) */
        xlPaperEnvelope14 = 23 , /*!< Envelope #14 (5 in. x 11-1/2 in.) */
        xlPaperEnvelope9 = 19 , /*!< Envelope #9 (3-7/8 in. x 8-7/8 in.) */
        xlPaperEnvelopeB4 = 33 , /*!< Envelope B4 (250 mm x 353 mm) */
        xlPaperEnvelopeB5 = 34 , /*!< Envelope B5 (176 mm x 250 mm) */
        xlPaperEnvelopeB6 = 35 , /*!< Envelope B6 (176 mm x 125 mm) */
        xlPaperEnvelopeC3 = 29 , /*!< Envelope C3 (324 mm x 458 mm) */
        xlPaperEnvelopeC4 = 30 , /*!< Envelope C4 (229 mm x 324 mm) */
        xlPaperEnvelopeC5 = 28 , /*!< Envelope C5 (162 mm x 229 mm) */
        xlPaperEnvelopeC6 = 31 , /*!< Envelope C6 (114 mm x 162 mm) */
        xlPaperEnvelopeC65 = 32 , /*!< Envelope C65 (114 mm x 229 mm) */
        xlPaperEnvelopeDL = 27 , /*!< Envelope DL (110 mm x 220 mm) */
        xlPaperEnvelopeItaly = 36 , /*!< Envelope (110 mm x 230 mm) */
        xlPaperEnvelopeMonarch = 37 , /*!< Envelope Monarch (3-7/8 in. x 7-1/2 in.) */
        xlPaperEnvelopePersonal = 38 , /*!< Envelope (3-5/8 in. x 6-1/2 in.) */
        xlPaperEsheet = 26 , /*!< E size sheet */
        xlPaperExecutive = 7 , /*!< Executive (7-1/2 in. x 10-1/2 in.) */
        xlPaperFanfoldLegalGerman = 41 , /*!< German Legal Fanfold (8-1/2 in. x 13 in.) */
        xlPaperFanfoldStdGerman = 40 , /*!< German Legal Fanfold (8-1/2 in. x 13 in.) */
        xlPaperFanfoldUS = 39 , /*!< U.S. Standard Fanfold (14-7/8 in. x 11 in.) */
        xlPaperFolio = 14 , /*!< Folio (8-1/2 in. x 13 in.) */
        xlPaperLedger = 4 , /*!< Ledger (17 in. x 11 in.) */
        xlPaperLegal = 5 , /*!< Legal (8-1/2 in. x 14 in.) */
        xlPaperLetter = 1 , /*!< Letter (8-1/2 in. x 11 in.) */
        xlPaperLetterSmall = 2 , /*!< Letter Small (8-1/2 in. x 11 in.) */
        xlPaperNote = 18 , /*!< Note (8-1/2 in. x 11 in.) */
        xlPaperQuarto = 15 , /*!< Quarto (215 mm x 275 mm) */
        xlPaperStatement = 6 , /*!< Statement (5-1/2 in. x 8-1/2 in.) */
        xlPaperTabloid = 3 , /*!< Tabloid (11 in. x 17 in.) */
        xlPaperUser = 256 , /*!< User-defined */
    };
    /*!  @brief Specifies the data type of a query parameter. Since Excel 2007.

    [MSDN documentation for XlParameterDataType](http://msdn.microsoft.com/en-us/library/bb241400.aspx).
    */
    enum XlParameterDataType {
        xlParamTypeBigInt = -5 , /*!< Big integer. */
        xlParamTypeBinary = -2 , /*!< Binary. */
        xlParamTypeBit = -7 , /*!< Bit. */
        xlParamTypeChar = 1 , /*!< String. */
        xlParamTypeDate = 9 , /*!< Date. */
        xlParamTypeDecimal = 3 , /*!< Decimal. */
        xlParamTypeDouble = 8 , /*!< Double. */
        xlParamTypeFloat = 6 , /*!< Float. */
        xlParamTypeInteger = 4 , /*!< Integer. */
        xlParamTypeLongVarBinary = -4 , /*!< Long binary. */
        xlParamTypeLongVarChar = -1 , /*!< Long string. */
        xlParamTypeNumeric = 2 , /*!< Numeric. */
        xlParamTypeReal = 7 , /*!< Real. */
        xlParamTypeSmallInt = 5 , /*!< Small integer. */
        xlParamTypeTime = 10 , /*!< Time. */
        xlParamTypeTimestamp = 11 , /*!< Time stamp. */
        xlParamTypeTinyInt = -6 , /*!< Tiny integer. */
        xlParamTypeUnknown = 0 , /*!< Type unknown. */
        xlParamTypeVarBinary = -3 , /*!< Variable-length binary. */
        xlParamTypeVarChar = 12 , /*!< Variable-length string. */
        xlParamTypeWChar = -8 , /*!< Unicode character string. */
    };
    /*!  @brief Specifies how to determine the value of the parameter for the specified query table. Since Excel 2007.

    [MSDN documentation for XlParameterType](http://msdn.microsoft.com/en-us/library/bb241402.aspx).
    */
    enum XlParameterType {
        xlConstant = 1 , /*!< Uses the value specified by the Value argument. */
        xlPrompt = 0 , /*!< Displays a dialog box that prompts the user for the value. The Value argument specifies the text shown in the dialog box. */
        xlRange = 2 , /*!< Uses the value of the cell in the upper-left corner of the range. The Value argument specifies a Range object. */
    };
    
    /*!  @brief Constants passed to and returned by the Series.ParentDataLabelOption property. Since Excel 2016.

    [MSDN documentation for XlParentDataLabelOptions](https://msdn.microsoft.com/VBA/Excel-VBA/articles/xlparentdatalabeloptions-enumeration-excel).
    */
    enum XlParentDataLabelOptions {
        xlParentDataLabelOptionsNone = 0,       /*!< No parent labels are shown. */
        xlParentDataLabelOptionsBanner = 1,     /*!< The parent label layout is a banner above the category. */
        xlParentDataLabelOptionsOverlapping = 2,/*!< The parent label is laid out within the category. */
    } ;
    
    /*!  @brief Specifies how numeric data will be calculated with the destinations cells in the worksheet. Since Excel 2007.

    [MSDN documentation for XlPasteSpecialOperation](http://msdn.microsoft.com/en-us/library/bb241404.aspx).
    */
    enum XlPasteSpecialOperation {
        xlPasteSpecialOperationAdd = 2 , /*!< Copied data will be added with the value in the destination cell. */
        xlPasteSpecialOperationDivide = 5 , /*!< Copied data will be divided with the value in the destination cell. */
        xlPasteSpecialOperationMultiply = 4 , /*!< Copied data will be multiplied with the value in the destination cell. */
        xlPasteSpecialOperationNone = -4142 , /*!< No calculation will be done in the paste operation. */
        xlPasteSpecialOperationSubtract = 3 , /*!< Copied data will be subtracted with the value in the destination cell. */
    };
    /*!  @brief Specifies the part of the range to be pasted. Since Excel 2007.

    [MSDN documentation for XlPasteType](http://msdn.microsoft.com/en-us/library/bb241405.aspx).
    */
    enum XlPasteType {
        xlPasteAll = -4104 , /*!< Everything will be pasted. */
        xlPasteAllExceptBorders = 7 , /*!< Everything except borders will be pasted. */
        xlPasteAllUsingSourceTheme = 13 , /*!< Everything will be pasted using the source theme. */
        xlPasteColumnWidths = 8 , /*!< Copied column width is pasted. */
        xlPasteComments = -4144 , /*!< Comments are pasted. */
        xlPasteFormats = -4122 , /*!< Copied source format is pasted. */
        xlPasteFormulas = -4123 , /*!< Formulas are pasted. */
        xlPasteFormulasAndNumberFormats = 11 , /*!< Formulas and Number formats are pasted. */
        xlPasteValidation = 6 , /*!< Validations are pasted. */
        xlPasteValues = -4163 , /*!< Values are pasted. */
        xlPasteValuesAndNumberFormats = 12 , /*!< Values and Number formats are pasted. */
    };
    /*!  @brief Specifies the interior pattern of a chart or interior object. Since Excel 2007.

    [MSDN documentation for XlPattern](http://msdn.microsoft.com/en-us/library/bb241407.aspx).
    */
    enum XlPattern {
        xlPatternAutomatic = -4105 , /*!< Excel controls the pattern. */
        xlPatternChecker = 9 , /*!< Checkerboard. */
        xlPatternCrissCross = 16 , /*!< Criss-cross lines. */
        xlPatternDown = -4121 , /*!< Dark diagonal lines running from the upper left to the lower right. */
        xlPatternGray16 = 17 , /*!< 16% gray. */
        xlPatternGray25 = -4124 , /*!< 25% gray. */
        xlPatternGray50 = -4125 , /*!< 50% gray. */
        xlPatternGray75 = -4126 , /*!< 75% gray. */
        xlPatternGray8 = 18 , /*!< 8% gray. */
        xlPatternGrid = 15 , /*!< Grid. */
        xlPatternHorizontal = -4128 , /*!< Dark horizontal lines. */
        xlPatternLightDown = 13 , /*!< Light diagonal lines running from the upper left to the lower right. */
        xlPatternLightHorizontal = 11 , /*!< Light horizontal lines. */
        xlPatternLightUp = 14 , /*!< Light diagonal lines running from the lower left to the upper right. */
        xlPatternLightVertical = 12 , /*!< Light vertical bars. */
        xlPatternLinearGradient = 4000, /*!< Linear gradient. */
        xlPatternNone = -4142 , /*!< No pattern. */
        xlPatternRectangularGradient = 4001, /*!< Rectangular gradient. */
        xlPatternSemiGray75 = 10 , /*!< 75% dark moiré. */
        xlPatternSolid = 1 , /*!< Solid color. */
        xlPatternUp = -4162 , /*!< Dark diagonal lines running from the lower left to the upper right. */
        xlPatternVertical = -4166 , /*!< Dark vertical bars. */
    };
    /*!  @brief Specifies the alignment for phonetic text. Used with a  Since Excel 2007.

    [MSDN documentation for XlPhoneticAlignment](http://msdn.microsoft.com/en-us/library/bb241409.aspx).
    */
    enum XlPhoneticAlignment {
        xlPhoneticAlignCenter = 2 , /*!< Centered */
        xlPhoneticAlignDistributed = 3 , /*!< Distributed */
        xlPhoneticAlignLeft = 1 , /*!< Left aligned */
        xlPhoneticAlignNoControl = 0 , /*!< Excel controls alignment */
    };
    /*!  @brief Specifies the type of phonetic text in a cell. Since Excel 2007.

    [MSDN documentation for XlPhoneticCharacterType](http://msdn.microsoft.com/en-us/library/bb241411.aspx).
    */
    enum XlPhoneticCharacterType {
        xlHiragana = 2 , /*!< Hiragana */
        xlKatakana = 1 , /*!< Katakana */
        xlKatakanaHalf = 0 , /*!< Half-size Katakana */
        xlNoConversion = 3 , /*!< No conversion */
    };
    /*!  @brief Specifies how the picture should be copied. Since Excel 2007.

    [MSDN documentation for XlPictureAppearance](http://msdn.microsoft.com/en-us/library/bb241413.aspx).
    */
    enum XlPictureAppearance {
        xlPrinter = 2 , /*!< The picture is copied as it will look when it is printed. */
        xlScreen = 1 , /*!< The picture is copied to resemble its display on the screen as closely as possible. */
    };
    /*!  @brief Specifies how to convert a graphic. Since Excel 2007.

    [MSDN documentation for XlPictureConvertorType](http://msdn.microsoft.com/en-us/library/bb241416.aspx).
    */
    enum XlPictureConvertorType {
        xlBMP = 1 , /*!< Windows version 2.0–compatible bitmap */
        xlCGM = 7 , /*!< Computer Graphics Metafile */
        xlDRW = 4 , /*!< DRW */
        xlDXF = 5 , /*!< DXF */
        xlEPS = 8 , /*!< Encapsulated Postscript */
        xlHGL = 6 , /*!< HGL */
        xlPCT = 13 , /*!< Bitmap Graphic (Apple PICT format) */
        xlPCX = 10 , /*!< PC Paintbrush Bitmap Graphic */
        xlPIC = 11 , /*!< PIC */
        xlPLT = 12 , /*!< PLT */
        xlTIF = 9 , /*!< Tagged Image Format File */
        xlWMF = 2 , /*!< Windows Metafile */
        xlWPG = 3 , /*!< WordPerfect/DrawPerfect Graphic */
    };

    /*!  @brief Specifies which position on the slice to return the coordinate of. Since Excel 2010.

    [MSDN documentation for XlPieSliceIndex](http://msdn.microsoft.com/en-us/library/office/ff193239%28v=office.14%29.aspx).
    */
    enum XlPieSliceIndex {        
        xlCenterPoint = 5, /*!< The center point of a pie slice. . */
        xlInnerCenterPoint = 8, /*!< The innermost center point of a doughnut slice. . */
        xlInnerClockwisePoint = 7, /*!< The innermost point of the most clockwise radius of a doughnut slice. . */
        xlInnerCounterClockwisePoint = 9, /*!< The innermost point of the most counterclockwise radius of a doughnut slice. . */
        xlMidClockwiseRadiusPoint = 4, /*!< The midpoint of the most clockwise radius of a slice. . */
        xlMidCounterClockwiseRadiusPoint = 6, /*!< The midpoint of the most counterclockwise radius of a slice. . */
        xlOuterCenterPoint = 2, /*!< The outer center point of the circumference of a slice. . */
        xlOuterClockwisePoint = 3, /*!< The outermost clockwise point of the circumference of a slice. . */
        xlOuterCounterClockwisePoint = 1, /*!< The outermost counterclockwise point of the circumference of a slice. . */
    };

    /*!  @brief Specifies the horizontal or vertical position of a point on a pie chart, in points, from the top or left edge of the object to the top or left edge of the chart area. Since Excel 2010.

    [MSDN documentation for XlPieSliceLocation](http://msdn.microsoft.com/en-us/library/office/ff839356%28v=office.14%29.aspx).
    */
    enum XlPieSliceLocation {        
        xlHorizontalCoordinate = 1, /*!< The horizontal coordinate (x) . */
        xlVerticalCoordinate = 2, /*!< The vertical coordinate (y) . */
    };
    
    /*!  @brief Specifies the PivotTable entity to which the cell corresponds. Since Excel 2007.

    [MSDN documentation for XlPivotCellType](http://msdn.microsoft.com/en-us/library/bb241417.aspx).
    */
    enum XlPivotCellType {
        xlPivotCellBlankCell = 9 , /*!< A structural blank cell in the PivotTable. */
        xlPivotCellCustomSubtotal = 7 , /*!< A cell in the row or column area that is a custom subtotal. */
        xlPivotCellDataField = 4 , /*!< A data field label (not the Data button). */
        xlPivotCellDataPivotField = 8 , /*!< The Data button. */
        xlPivotCellGrandTotal = 3 , /*!< A cell in a row or column area that is a grand total. */
        xlPivotCellPageFieldItem = 6 , /*!< The cell that shows the selected item of a Page field. */
        xlPivotCellPivotField = 5 , /*!< The button for a field (not the Data button). */
        xlPivotCellPivotItem = 1 , /*!< A cell in the row or column area that is not a subtotal, grand total, custom subtotal, or blank line. */
        xlPivotCellSubtotal = 2 , /*!< A cell in the row or column area that is a subtotal. */
        xlPivotCellValue = 0 , /*!< Any cell in the data area (except a blank row). */
    };
    /*!This enumeration specifies the conditional formatting applied for filtering values from the  Since Excel 2007.

    [MSDN documentation for XlPivotConditionScope](http://msdn.microsoft.com/en-us/library/bb241419.aspx).
    */
    enum XlPivotConditionScope {
        xlDataFieldScope = 2 , /*!< Based on the data in the specified fields. */
        xlFieldsScope = 1 , /*!< Based on the specified fields. */
        xlSelectionScope = 0 , /*!< Based on the specified selection criteria. */
    };
    /*!  @brief Specifies the type of calculation performed by a data PivotField when a custom calculation is used. Since Excel 2007.

    [MSDN documentation for XlPivotFieldCalculation](http://msdn.microsoft.com/en-us/library/office/ff836192%28v=office.14%29.aspx).
    */
    enum XlPivotFieldCalculation {
       xlDifferenceFrom = 2, /*!< The difference from the value of the Base item in the Base field. */
        xlIndex = 9, /*!< Data calculated as ((value in cell) x (Grand Total of Grand Totals)) / ((Grand Row Total) x (Grand Column Total)). */
        xlNoAdditionalCalculation = -4143, /*!< No calculation. */
        xlPercentDifferenceFrom = 4, /*!< Percentage difference from the value of the Base item in the Base field. */
        xlPercentOf = 3, /*!< Percentage of the value of the Base item in the Base field. */
        xlPercentOfColumn = 7, /*!< Percentage of the total for the column or series. */
        xlPercentOfParent  = 12, /*!< Percentage of the total of the specified parent Base field. */
        xlPercentOfParentColumn  = 11, /*!< Percentage of the total of the parent column. */
        xlPercentOfParentRow  = 10, /*!< Percentage of the total of the parent row. */
        xlPercentOfRow = 6, /*!< Percentage of the total for the row or category. */
        xlPercentOfTotal = 8, /*!< Percentage of the grand total of all the data or data points in the report. */
        xlPercentRunningTotal  = 13, /*!< Percentatge of the running total of the specified Base field. */
        xlRankAscending  = 14, /*!< Rank smallest to largest. */
        xlRankDecending  = 15, /*!< Rank largest to smallest. */
        xlRunningTotal = 5, /*!< Data for successive items in the Base field as a running total. */
    };
    /*!  @brief Specifies the type of data in the  Since Excel 2007.

    [MSDN documentation for XlPivotFieldDataType](http://msdn.microsoft.com/en-us/library/bb241423.aspx).
    */
    enum XlPivotFieldDataType {
        xlDate = 2 , /*!< Contains a date. */
        xlNumber = -4145 , /*!< Contains a number. */
        xlText = -4158 , /*!< Contains text. */
    };
    /*!  @brief Specifies the location of the field in a PivotTable report. Since Excel 2007.

    [MSDN documentation for XlPivotFieldOrientation](http://msdn.microsoft.com/en-us/library/bb241425.aspx).
    */
    enum XlPivotFieldOrientation {
        xlColumnField = 2 , /*!< Column */
        xlDataField = 4 , /*!< Data */
        xlHidden = 0 , /*!< Hidden */
        xlPageField = 3 , /*!< Page */
        xlRowField = 1 , /*!< Row */
    };
    /*!	The type of filter applied. Since Excel 2007.

    [MSDN documentation for XlPivotFilterType](http://msdn.microsoft.com/en-us/library/bb241427.aspx).
    */
    enum XlPivotFilterType {
        xlBefore = 31 , /*!< Filters for all dates before a specified date */
        xlBeforeOrEqualTo = 32 , /*!< Filters for all dates on or before a specified date */
        xlAfter = 33 , /*!< Filters for all dates after a specified date */
        xlAfterOrEqualTo = 34 , /*!< Filters for all dates on or after a specified date */
        xlAllDatesInPeriodJanuary = 53 , /*!< Filters for all dates in January */
        xlAllDatesInPeriodFebruary = 54 , /*!< Filters for all dates in February */
        xlAllDatesInPeriodMarch = 55 , /*!< Filters for all dates in March */
        xlAllDatesInPeriodApril = 56 , /*!< Filters for all dates in April */
        xlAllDatesInPeriodMay = 57 , /*!< Filters for all dates in May */
        xlAllDatesInPeriodJune = 58 , /*!< Filters for all dates in June */
        xlAllDatesInPeriodJuly = 59 , /*!< Filters for all dates in July */
        xlAllDatesInPeriodAugust = 60 , /*!< Filters for all dates in August */
        xlAllDatesInPeriodSeptember = 61 , /*!< Filters for all dates in September */
        xlAllDatesInPeriodOctober = 62 , /*!< Filters for all dates in October */
        xlAllDatesInPeriodNovember = 63 , /*!< Filters for all dates in November */
        xlAllDatesInPeriodDecember = 64 , /*!< Filters for all dates in December */
        xlAllDatesInPeriodQuarter1 = 49 , /*!< Filters for all dates in Quarter1 */
        xlAllDatesInPeriodQuarter2 = 50 , /*!< Filters for all dates in Quarter2 */
        xlAllDatesInPeriodQuarter3 = 51 , /*!< Filters for all dates in Quarter3 */
        xlAllDatesInPeriodQuarter4 = 52 , /*!< Filters for all dates in Quarter 4 */
        xlBottomCount = 2 , /*!< Filters for the specified number of values from the bottom of a list */
        xlBottomPercent = 4 , /*!< Filters for the specified percentage of values from the bottom of a list */
        xlBottomSum = 6 , /*!< Sum of the values from the bottom of the list */
        xlCaptionBeginsWith = 17 , /*!< Filters for all captions beginning with the specified string */
        xlCaptionContains = 21 , /*!< Filters for all captions that contain the specified string */
        xlCaptionDoesNotBeginWith = 18 , /*!< Filters for all captions that do not begin with the specified string */
        xlCaptionDoesNotContain = 22 , /*!< Filters for all captions that do not contain the specified string */
        xlCaptionDoesNotEndWith = 20 , /*!< Filters for all captions that do not end with the specified string */
        xlCaptionDoesNotEqual = 16 , /*!< Filters for all captions that do not match the specified string */
        xlCaptionEndsWith = 19 , /*!< Filters for all captions that end with the specified string */
        xlCaptionEquals = 15 , /*!< Filters for all captions that match the specified string */
        xlCaptionIsBetween = 27 , /*!< Filters for all captions that are between a specified range of values */
        xlCaptionIsGreaterThan = 23 , /*!< Filters for all captions that are greater than the specified value */
        xlCaptionIsGreaterThanOrEqualTo = 24 , /*!< Filters for all captions that are greater than or match the specified value */
        xlCaptionIsLessThan = 25 , /*!< Filters for all captions that are less than the specified value */
        xlCaptionIsLessThanOrEqualTo = 26 , /*!< Filters for all captions that are less than or match the specified value */
        xlCaptionIsNotBetween = 28 , /*!< Filters for all captions that are not between a specified range of values */
        xlDateBetween = 32 , /*!< Filters for all dates that are between a specified range of dates */
        xlDateLastMonth = 41 , /*!< Filters for all dates that apply to the previous month */
        xlDateLastQuarter = 44 , /*!< Filters for all dates that apply to the previous quarter */
        xlDateLastWeek = 38 , /*!< Filters for all dates that apply to the previous week */
        xlDateLastYear = 47 , /*!< Filters for all dates that apply to the previous year */
        xlDateNextMonth = 39 , /*!< Filters for all dates that apply to the next month */
        xlDateNextQuarter = 42 , /*!< Filters for all dates that apply to the next quarter */
        xlDateNextWeek = 36 , /*!< Filters for all dates that apply to the next week */
        xlDateNextYear = 45 , /*!< Filters for all dates that apply to the next year */
        xlDateThisMonth = 40 , /*!< Filters for all dates that apply to the current month */
        xlDateThisQuarter = 43 , /*!< Filters for all dates that apply to the current quarter */
        xlDateThisWeek = 37 , /*!< Filters for all dates that apply to the current week */
        xlDateThisYear = 46 , /*!< Filters for all dates that apply to the current year */
        xlDateToday = 34 , /*!< Filters for all dates that apply to the current date */
        xlDateTomorrow = 33 , /*!< Filters for all dates that apply to the next day */
        xlDateYesterday = 35 , /*!< Filters for all dates that apply to the previous day */
        xlNotSpecificDate = 30 , /*!< Filters for all dates that do not match a specified date */
        xlSpecificDate = 29 , /*!< Filters for all dates that match a specified date */
        xlTopCount = 1 , /*!< Filters for the specified number of values from the top of a list */
        xlTopPercent = 3 , /*!< Filters for the specified percentage of values from a list */
        xlTopSum = 5 , /*!< Sum of the values from the top of the list */
        xlValueDoesNotEqual = 8 , /*!< Filters for all values that do not match the specified value */
        xlValueEquals = 7 , /*!< Filters for all values that match the specified value */
        xlValueIsBetween = 13 , /*!< Filters for all values that are between a specified range of values */
        xlValueIsGreaterThan = 9 , /*!< Filters for all values that are greater than the specified value */
        xlValueIsGreaterThanOrEqualTo = 10 , /*!< Filters for all values that are greater than or match the specified value */
        xlValueIsLessThan = 11 , /*!< Filters for all values that are less than the specified value */
        xlValueIsLessThanOrEqualTo = 12 , /*!< Filters for all values that are less than or match the specified value */
        xlValueIsNotBetween = 14 , /*!< Filters for all values that are not between a specified range of values */
        xlYearToDate = 48 , /*!< Filters for all values that are within one year of a specified date */
    };
    /*!  @brief Specifies the type of report formatting to be applied to the specified PivotTable report. Since Excel 2007.

    [MSDN documentation for XlPivotFormatType](http://msdn.microsoft.com/en-us/library/bb241429.aspx).
    */
    enum XlPivotFormatType {
        xlPTClassic = 20 , /*!< PivotTable classic format. */
        xlPTNone = 21 , /*!< Does not apply formatting to the PivotTable report. */
        xlReport1 = 0 , /*!< Use the xlReport1 formatting for the PivotTable. */
        xlReport10 = 9 , /*!< Use the xlReport10 formatting for the PivotTable. */
        xlReport2 = 1 , /*!< Use the xlReport2 formatting for the PivotTable. */
        xlReport3 = 2 , /*!< Use the xlReport3 formatting for the PivotTable. */
        xlReport4 = 3 , /*!< Use the xlReport4 formatting for the PivotTable. */
        xlReport5 = 4 , /*!< Use the xlReport5 formatting for the PivotTable. */
        xlReport6 = 5 , /*!< Use the xlReport6 formatting for the PivotTable. */
        xlReport7 = 6 , /*!< Use the xlReport7 formatting for the PivotTable. */
        xlReport8 = 7 , /*!< Use the xlReport8 formatting for the PivotTable. */
        xlReport9 = 8 , /*!< Use the xlReport9 formatting for the PivotTable. */
        xlTable1 = 10 , /*!< Use the xlTable1 formatting for the PivotTable. */
        xlTable10 = 19 , /*!< Use the xlTable10 formatting for the PivotTable. */
        xlTable2 = 11 , /*!< Use the xlTable2 formatting for the PivotTable. */
        xlTable3 = 12 , /*!< Use the xlTable3 formatting for the PivotTable. */
        xlTable4 = 13 , /*!< Use the xlTable4 formatting for the PivotTable. */
        xlTable5 = 14 , /*!< Use the xlTable5 formatting for the PivotTable. */
        xlTable6 = 15 , /*!< Use the xlTable6 formatting for the PivotTable. */
        xlTable7 = 16 , /*!< Use the xlTable7 formatting for the PivotTable. */
        xlTable8 = 17 , /*!< Use the xlTable8 formatting for the PivotTable. */
        xlTable9 = 18 , /*!< Use the xlTable9 formatting for the PivotTable. */
    };
    /*!  @brief Specifies the type of the PivotLine. Since Excel 2007.

    [MSDN documentation for XlPivotLineType](http://msdn.microsoft.com/en-us/library/bb241431.aspx).
    */
    enum XlPivotLineType {
        xlPivotLineBlank = 3 , /*!< Blank line after each group. */
        xlPivotLineGrandTotal = 2 , /*!< Grand Total line. */
        xlPivotLineRegular = 0 , /*!< Regular PivotLine with pivot items. */
        xlPivotLineSubtotal = 1 , /*!< Subtotal line. */
    };
    /*!  @brief Specifies the maximum number of unique items allowed per PivotField. Since Excel 2007.

    [MSDN documentation for XlPivotTableMissingItems](http://msdn.microsoft.com/en-us/library/bb241433.aspx).
    */
    enum XlPivotTableMissingItems {
        xlMissingItemsDefault = -1 , /*!< The default number of unique items per PivotField allowed. */
        xlMissingItemsMax = 32500 , /*!< The maximum number of unique items per PivotField allowed (32,500) for  a pre-Excel 2007  PivotTable. */
        xlMissingItemsMax2 = 1048576 , /*!< The maximum number of unique items per PivotField allowed (10,48,576) for a pre-Excel 2007 PivotTable. */
        xlMissingItemsNone = 0 , /*!< No unique items per PivotField allowed (zero). */
    };
    /*!  @brief Specifies the source of the report data. Since Excel 2007.

    [MSDN documentation for XlPivotTableSourceType](http://msdn.microsoft.com/en-us/library/bb241435.aspx).
    */
    enum XlPivotTableSourceType {
        xlConsolidation = 3 , /*!< Multiple consolidation ranges. */
        xlDatabase = 1 , /*!< Microsoft Excel list or database. */
        xlExternal = 2 , /*!< Data from another application. */
        xlPivotTable = -4148 , /*!< Same source as another PivotTable report. */
        xlScenario = 4 , /*!< Data is based on scenarios created using the Scenario Manager. */
    };
    /*!  @brief Specifies the version of a PivotTable or a PivotCache. Creating PivotTables with a specific version ensures that tables created in Microsoft Office Excel 2010 behave in the same manner as they did in the corresponding version of  Excel. Since Excel 2007.

    [MSDN documentation for XlPivotTableVersionList](http://msdn.microsoft.com/en-us/library/office/ff837609%28v=office.14%29.aspx).
    */
    enum XlPivotTableVersionList {
        xlPivotTableVersion2000 = 0, /*!< Excel 2000 */
        xlPivotTableVersion10 = 1, /*!< Excel 2002 */
        xlPivotTableVersion11 = 2, /*!< Excel 2003 */
        xlPivotTableVersion12 = 3, /*!< Excel 2007 */
        xlPivotTableVersion14 = 4, /*!< Excel 2010 */
        xlPivotTableVersionCurrent = -1, /*!< Provided only for backward compatibility */
    };
    /*!  @brief Specifies the way that an object is attached to its underlying cells. Since Excel 2007.

    [MSDN documentation for XlPlacement](http://msdn.microsoft.com/en-us/library/bb241440.aspx).
    */
    enum XlPlacement {
        xlFreeFloating = 3 , /*!< Object is free floating. */
        xlMove = 2 , /*!< Object is moved with the cells. */
        xlMoveAndSize = 1 , /*!< Object is moved and sized with the cells. */
    };
    /*!  @brief Specifies the platform on which a text file originated. Since Excel 2007.

    [MSDN documentation for XlPlatform](http://msdn.microsoft.com/en-us/library/bb241441.aspx).
    */
    enum XlPlatform {
        xlMacintosh = 1 , /*!< Macintosh */
        xlMSDOS = 3 , /*!< MS-DOS */
        xlWindows = 2 , /*!< Microsoft Windows */
    };
    /*!  @brief Specifies the type of print error displayed. Since Excel 2007.

    [MSDN documentation for XlPrintErrors](http://msdn.microsoft.com/en-us/library/bb241443.aspx).
    */
    enum XlPrintErrors {
        xlPrintErrorsBlank = 1 , /*!< Print errors are blank. */
        xlPrintErrorsDash = 2 , /*!< Print errors are displayed as dashes. */
        xlPrintErrorsDisplayed = 0 , /*!< All print errors are displayed. */
        xlPrintErrorsNA = 3 , /*!< Print errors are displayed as not available. */
    };
    /*!  @brief Specifies the way that comments are printed with the sheet. Since Excel 2007.

    [MSDN documentation for XlPrintLocation](http://msdn.microsoft.com/en-us/library/bb241446.aspx).
    */
    enum XlPrintLocation {
        xlPrintInPlace = 16 , /*!< Comments will be printed where they were inserted in the worksheet. */
        xlPrintNoComments = -4142 , /*!< Comments will not be printed. */
        xlPrintSheetEnd = 1 , /*!< Comments will be printed as end notes at the end of the worksheet. */
    };
    /*!  @brief Specifies the priority of a SendMailer message. Since Excel 2007.

    [MSDN documentation for XlPriority](http://msdn.microsoft.com/en-us/library/bb241448.aspx).
    */
    enum XlPriority {
        xlPriorityHigh = -4127 , /*!< High */
        xlPriorityLow = -4134 , /*!< Low */
        xlPriorityNormal = -4143 , /*!< Normal */
    };

    /*!  @brief Specifies the mode for checking the spelling of Portuguese. Since Excel 2010.

    [MSDN documentation for XlPortugueseReform](http://msdn.microsoft.com/en-us/library/office/ff196529%28v=office.14%29.aspx).
    */
    enum XlPortugueseReform {
        xlPortugueseBoth  = 3, /*!< The spelling checker recognizes both pre-reform and post-reform spellings. . */
        xlPortuguesePostReform  = 2, /*!< The spelling checker recognizes only post-reform spellings. . */
        xlPortuguesePreReform  = 1, /*!< The spelling checker recognizes only pre-reform spellings. . */
    };    
    
    /*!  @brief Specifies where to display the property.

    [MSDN documentation for XlPropertyDisplayedIn](http://msdn.microsoft.com/en-us/library/bb241450.aspx).
    */
    enum XlPropertyDisplayedIn {
        xlDisplayPropertyInPivotTable = 1 , /*!< Displays member property in the PivotTable only. This is the default value. */
        xlDisplayPropertyInPivotTableAndTooltip = 3 , /*!< Displays member property in the tooltip only. */
        xlDisplayPropertyInTooltip = 2 , /*!< Displays member property in both the tooltip and the PivotTable.  */
    };

    /*!  @brief Specifies how the Protected View window was closed. Since Excel 2010.

    [MSDN documentation for XlProtectedViewCloseReason](http://msdn.microsoft.com/en-us/library/office/ff839651%28v=office.14%29.aspx).
    */
    enum XlProtectedViewCloseReason {
        xlProtectedViewCloseEdit = 1, /*!< The window was closed when the user clicked the Enable Editing button. . */
        xlProtectedViewCloseForced = 2, /*!< The window was closed because the application shut it down forcefully or stopped responding. . */
        xlProtectedViewCloseNormal = 0, /*!< The window was closed normally. . */
    };

    /*!  @brief Specifies the state of the Protected View window. Since Excel 2010.

    [MSDN documentation for XlProtectedViewWindowState](http://msdn.microsoft.com/en-us/library/office/ff194129%28v=office.14%29.aspx).
    */
    enum XlProtectedViewWindowState {
        xlProtectedViewWindowMaximized = 2, /*!< Maximized. */
        xlProtectedViewWindowMinimized = 1, /*!< Minimized. */
        xlProtectedViewWindowNormal = 0, /*!< Normal. */        
    };

    /*!  @brief Specifies what can be selected in a PivotTable during a structured selection. These constants can be combined to select multiple types. Since Excel 2007.

    [MSDN documentation for XlPTSelectionMode](http://msdn.microsoft.com/en-us/library/bb241533.aspx).
    */
    enum XlPTSelectionMode {
        xlBlanks = 4 , /*!< Blanks */
        xlButton = 15 , /*!< Buttons */
        xlDataAndLabel = 0 , /*!< Data and labels */
        xlDataOnly = 2 , /*!< Data */
        xlFirstRow = 256 , /*!< First row */
        xlLabelOnly = 1 , /*!< Label */
        xlOrigin = 3 , /*!< Origin */
    };
    /*!  @brief Specifies the type of query used by Microsoft Office Excel to populate the query table or PivotTable cache. Since Excel 2007.

    [MSDN documentation for XlQueryType](http://msdn.microsoft.com/en-us/library/bb241536.aspx).
    */
    enum XlQueryType {
        xlADORecordset = 7 , /*!< Based on an ADO recordset query */
        xlDAORecordset = 2 , /*!< Based on a DAO recordset query, for query tables only */
        xlODBCQuery = 1 , /*!< Based on an ODBC data source */
        xlOLEDBQuery = 5 , /*!< Based on an OLE DB query, including OLAP data sources */
        xlTextImport = 6 , /*!< Based on a text file, for query tables only */
        xlWebQuery = 4 , /*!< Based on a Web page, for query tables only */
    };
    /*!  @brief Specifies the predefined format when a range is automatically formatted. Since Excel 2007.

    [MSDN documentation for XlRangeAutoFormat](http://msdn.microsoft.com/en-us/library/bb241539.aspx).
    */
    enum XlRangeAutoFormat {
        xlRangeAutoFormat3DEffects1 = 13 , /*!< 3-D effects 1. */
        xlRangeAutoFormat3DEffects2 = 14 , /*!< 3-D effects 2. */
        xlRangeAutoFormatAccounting1 = 4 , /*!< Accounting 1. */
        xlRangeAutoFormatAccounting2 = 5 , /*!< Accounting 2. */
        xlRangeAutoFormatAccounting3 = 6 , /*!< Accounting 3. */
        xlRangeAutoFormatAccounting4 = 17 , /*!< Accounting 4. */
        xlRangeAutoFormatClassic1 = 1 , /*!< Classic 1. */
        xlRangeAutoFormatClassic2 = 2 , /*!< Classic 2. */
        xlRangeAutoFormatClassic3 = 3 , /*!< Classic 3. */
        xlRangeAutoFormatClassicPivotTable = 31 , /*!< Classic pivot table. */
        xlRangeAutoFormatColor1 = 7 , /*!< Color 1. */
        xlRangeAutoFormatColor2 = 8 , /*!< Color 2. */
        xlRangeAutoFormatColor3 = 9 , /*!< Color 3. */
        xlRangeAutoFormatList1 = 10 , /*!< List 1. */
        xlRangeAutoFormatList2 = 11 , /*!< List 2. */
        xlRangeAutoFormatList3 = 12 , /*!< List 3. */
        xlRangeAutoFormatLocalFormat1 = 15 , /*!< Local Format 1. */
        xlRangeAutoFormatLocalFormat2 = 16 , /*!< Local Format 2. */
        xlRangeAutoFormatLocalFormat3 = 19 , /*!< Local Format 3. */
        xlRangeAutoFormatLocalFormat4 = 20 , /*!< Local Format 4. */
        xlRangeAutoFormatNone = -4142 , /*!< No specified format. */
        xlRangeAutoFormatPTNone = 42 , /*!< No specified pivot table format. */
        xlRangeAutoFormatReport1 = 21 , /*!< Report 1. */
        xlRangeAutoFormatReport10 = 30 , /*!< Report 10. */
        xlRangeAutoFormatReport2 = 22 , /*!< Report 2. */
        xlRangeAutoFormatReport3 = 23 , /*!< Report 3. */
        xlRangeAutoFormatReport4 = 24 , /*!< Report 4. */
        xlRangeAutoFormatReport5 = 25 , /*!< Report 5. */
        xlRangeAutoFormatReport6 = 26 , /*!< Report 6. */
        xlRangeAutoFormatReport7 = 27 , /*!< Report 7. */
        xlRangeAutoFormatReport8 = 28 , /*!< Report 8. */
        xlRangeAutoFormatReport9 = 29 , /*!< Report 9. */
        xlRangeAutoFormatSimple = -4154 , /*!< Simple. */
        xlRangeAutoFormatTable1 = 32 , /*!< Table 1. */
        xlRangeAutoFormatTable10 = 41 , /*!< Table 10. */
        xlRangeAutoFormatTable2 = 33 , /*!< Table 2. */
        xlRangeAutoFormatTable3 = 34 , /*!< Table 3. */
        xlRangeAutoFormatTable4 = 35 , /*!< Table 4. */
        xlRangeAutoFormatTable5 = 36 , /*!< Table 5. */
        xlRangeAutoFormatTable6 = 37 , /*!< Table 6. */
        xlRangeAutoFormatTable7 = 38 , /*!< Table 7. */
        xlRangeAutoFormatTable8 = 39 , /*!< Table 8. */
        xlRangeAutoFormatTable9 = 40 , /*!< Table 9. */
    };
    /*!  @brief Specifies the range value data type. Since Excel 2007.

    [MSDN documentation for XlRangeValueDataType](http://msdn.microsoft.com/en-us/library/bb241542.aspx).
    */
    enum XlRangeValueDataType {
        xlRangeValueDefault = 10 , /*!< Default. If the specified Range object is empty, returns the value Empty (use the IsEmpty function to test for this case). If the Range object contains more than one cell, returns an array of values (use the IsArray function to test for this case). */
        xlRangeValueMSPersistXML = 12 , /*!< Returns the recordset representation of the specified Range object in an XML format. */
        xlRangeValueXMLSpreadsheet = 11 , /*!< Returns the values, formatting, formulas, and names of the specified Range object in the XML Spreadsheet format. */
    };
    /*!  @brief Specifies the reference style. Since Excel 2007.

    [MSDN documentation for XlReferenceStyle](http://msdn.microsoft.com/en-us/library/bb241546.aspx).
    */
    enum XlReferenceStyle {
        xlA1 = 1 , /*!< Default. Use xlA1 to return an A1-style reference. */
        xlR1C1 = -4150 , /*!< Use xlR1C1 to return an R1C1-style reference. */
    };
    /*!  @brief Specifies cell reference style when a formula is being converted. Since Excel 2007.

    [MSDN documentation for XlReferenceType](http://msdn.microsoft.com/en-us/library/bb241550.aspx).
    */
    enum XlReferenceType {
        xlAbsolute = 1 , /*!< Convert to absolute row and column style. */
        xlAbsRowRelColumn = 2 , /*!< Convert to absolute row and relative column style. */
        xlRelative = 4 , /*!< Convert to relative row and column style. */
        xlRelRowAbsColumn = 3 , /*!< Convert to relative row and absolute column style. */
    };
    /*!  @brief Specifies the type information to be removed from the document information. Since Excel 2007.

    [MSDN documentation for XlRemoveDocInfoType](http://msdn.microsoft.com/en-us/library/bb257016.aspx).
    */
    enum XlRemoveDocInfoType {
        xlRDIAll = 99 , /*!< Removes all documentation information. */
        xlRDIComments = 1 , /*!< Removes comments from the document information. */
        xlRDIContentType = 16 , /*!< Removes content type data from the document information. */
        xlRDIDefinedNameComments = 18 , /*!< Removes defined nameauthor="telias" time="20060816T160514-0800" data="s" comments from the documentation information. */
        xlRDIDocumentManagementPolicy = 15 , /*!< Removes document management policy data from the document information. */
        xlRDIDocumentProperties = 8 , /*!< Removes document properties from the document information. */
        xlRDIDocumentServerProperties = 14 , /*!< Removes server properties from the document information. */
        xlRDIDocumentWorkspace = 10 , /*!< Removes workspace data from the document information. */
        xlRDIEmailHeader = 5 , /*!< Removes email headers from the document information. */
        xlRDIInactiveDataConnections = 19 , /*!< Removes inactive data connection data from the document information. */
        xlRDIInkAnnotations = 11 , /*!< Removes ink annotations from the document information. */
        xlRDIPublishInfo = 13 , /*!< Removes the pubish information data from the document information. */
        xlRDIRemovePersonalInformation = 4 , /*!< Removes personal information from the document information. */
        xlRDIRoutingSlip = 6 , /*!< Removes routing slip information from the document information. */
        xlRDIScenarioComments = 12 , /*!< Removes scenario comments from the document information. */
        xlRDISendForReview = 7 , /*!< Removes the send for review information from the document information. */
    };
    /*!  @brief Specifies the RGB color. Since Excel 2007.

    [MSDN documentation for XlRgbColor](http://msdn.microsoft.com/en-us/library/bb241561.aspx).
    */
    enum XlRgbColor {
        rgbAliceBlue = 16775408 , /*!< Alice Blue */
        rgbAntiqueWhite = 14150650 , /*!< Antique White */
        rgbAqua = 16776960 , /*!< Aqua */
        rgbAquamarine = 13959039 , /*!< Aquamarine */
        rgbAzure = 16777200 , /*!< Azure */
        rgbBeige = 14480885 , /*!< Beige */
        rgbBisque = 12903679 , /*!< Bisque */
        rgbBlack = 0 , /*!< Black */
        rgbBlanchedAlmond = 13495295 , /*!< Blanched Almond */
        rgbBlue = 16711680 , /*!< Blue */
        rgbBlueViolet = 14822282 , /*!< Blue Violet */
        rgbBrown = 2763429 , /*!< Brown */
        rgbBurlyWood = 8894686 , /*!< Burly Wood */
        rgbCadetBlue = 10526303 , /*!< Cadet Blue */
        rgbChartreuse = 65407 , /*!< Chartreuse */
        rgbCoral = 5275647 , /*!< Coral */
        rgbCornflowerBlue = 15570276 , /*!< Cornflower Blue */
        rgbCornsilk = 14481663 , /*!< Cornsilk */
        rgbCrimson = 3937500 , /*!< Crimson */
        rgbDarkBlue = 9109504 , /*!< Dark Blue */
        rgbDarkCyan = 9145088 , /*!< Dark Cyan */
        rgbDarkGoldenrod = 755384 , /*!< Dark Goldenrod */
        rgbDarkGray = 11119017 , /*!< Dark Gray */
        rgbDarkGreen = 25600 , /*!< Dark Green */
        rgbDarkGrey = 11119017 , /*!< Dark Grey */
        rgbDarkKhaki = 7059389 , /*!< Dark Khaki */
        rgbDarkMagenta = 9109643 , /*!< Dark Magenta */
        rgbDarkOliveGreen = 3107669 , /*!< Dark Olive Green */
        rgbDarkOrange = 36095 , /*!< Dark Orange */
        rgbDarkOrchid = 13382297 , /*!< Dark Orchid */
        rgbDarkRed = 139 , /*!< Dark Red */
        rgbDarkSalmon = 8034025 , /*!< Dark Salmon */
        rgbDarkSeaGreen = 9419919 , /*!< Dark Sea Green */
        rgbDarkSlateBlue = 9125192 , /*!< Dark Slate Blue */
        rgbDarkSlateGray = 5197615 , /*!< Dark Slate Gray */
        rgbDarkSlateGrey = 5197615 , /*!< Dark Slate Grey */
        rgbDarkTurquoise = 13749760 , /*!< Dark Turquoise */
        rgbDarkViolet = 13828244 , /*!< Dark Violet */
        rgbDeepPink = 9639167 , /*!< Deep Pink */
        rgbDeepSkyBlue = 16760576 , /*!< Deep Sky Blue */
        rgbDimGray = 6908265 , /*!< Dim Gray */
        rgbDimGrey = 6908265 , /*!< Dim Grey */
        rgbDodgerBlue = 16748574 , /*!< Dodger Blue */
        rgbFireBrick = 2237106 , /*!< Fire Brick */
        rgbFloralWhite = 15792895 , /*!< Floral White */
        rgbForestGreen = 2263842 , /*!< Forest Green */
        rgbFuchsia = 16711935 , /*!< Fuchsia */
        rgbGainsboro = 14474460 , /*!< Gainsboro */
        rgbGhostWhite = 16775416 , /*!< Ghost White */
        rgbGold = 55295 , /*!< Gold */
        rgbGoldenrod = 2139610 , /*!< Goldenrod */
        rgbGray = 8421504 , /*!< Gray */
        rgbGreen = 32768 , /*!< Green */
        rgbGreenYellow = 3145645 , /*!< Green Yellow */
        rgbGrey = 8421504 , /*!< Grey */
        rgbHoneydew = 15794160 , /*!< Honeydew */
        rgbHotPink = 11823615 , /*!< Hot Pink */
        rgbIndianRed = 6053069 , /*!< Indian Red */
        rgbIndigo = 8519755 , /*!< Indigo */
        rgbIvory = 15794175 , /*!< Ivory */
        rgbKhaki = 9234160 , /*!< Khaki */
        rgbLavender = 16443110 , /*!< Lavender */
        rgbLavenderBlush = 16118015 , /*!< Lavender Blush */
        rgbLawnGreen = 64636 , /*!< Lawn Green */
        rgbLemonChiffon = 13499135 , /*!< Lemon Chiffon */
        rgbLightBlue = 15128749 , /*!< Light Blue */
        rgbLightCoral = 8421616 , /*!< Light Coral */
        rgbLightCyan = 9145088 , /*!< Light Cyan */
        rgbLightGoldenrodYellow = 13826810 , /*!< LightGoldenrodYellow */
        rgbLightGray = 13882323 , /*!< Light Gray */
        rgbLightGreen = 9498256 , /*!< Light Green */
        rgbLightGrey = 13882323 , /*!< Light Grey */
        rgbLightPink = 12695295 , /*!< Light Pink */
        rgbLightSalmon = 8036607 , /*!< Light Salmon */
        rgbLightSeaGreen = 11186720 , /*!< Light Sea Green */
        rgbLightSkyBlue = 16436871 , /*!< Light Sky Blue */
        rgbLightSlateGray = 10061943 , /*!< Light Slate Gray */
        rgbLightSteelBlue = 14599344 , /*!< Light Steel Blue */
        rgbLightYellow = 14745599 , /*!< Light Yellow */
        rgbLime = 65280 , /*!< Lime */
        rgbLimeGreen = 3329330 , /*!< Lime Green */
        rgbLinen = 15134970 , /*!< Linen */
        rgbMaroon = 128 , /*!< Maroon */
        rgbMediumAquamarine = 11206502 , /*!< Medium Aquamarine */
        rgbMediumBlue = 13434880 , /*!< Medium Blue */
        rgbMediumOrchid = 13850042 , /*!< Medium Orchid */
        rgbMediumPurple = 14381203 , /*!< Medium Purple */
        rgbMediumSeaGreen = 7451452 , /*!< Medium Sea Green */
        rgbMediumSlateBlue = 15624315 , /*!< Medium Slate Blue */
        rgbMediumSpringGreen = 10156544 , /*!< Medium Spring Green */
        rgbMediumTurquoise = 13422920 , /*!< Medium Turquoise */
        rgbMediumVioletRed = 8721863 , /*!< Medium Violet Red */
        rgbMidnightBlue = 7346457 , /*!< Midnight Blue */
        rgbMintCream = 16449525 , /*!< Mint Cream */
        rgbMistyRose = 14804223 , /*!< Misty Rose */
        rgbMoccasin = 11920639 , /*!< Moccasin */
        rgbNavajoWhite = 11394815 , /*!< Navajo White */
        rgbNavy = 8388608 , /*!< Navy */
        rgbNavyBlue = 8388608 , /*!< Navy Blue */
        rgbOldLace = 15136253 , /*!< Old Lace */
        rgbOlive = 32896 , /*!< Olive */
        rgbOliveDrab = 2330219 , /*!< Olive Drab */
        rgbOrange = 42495 , /*!< Orange */
        rgbOrangeRed = 17919 , /*!< Orange Red */
        rgbOrchid = 14053594 , /*!< Orchid */
        rgbPaleGoldenrod = 7071982 , /*!< Pale Goldenrod */
        rgbPaleGreen = 10025880 , /*!< Pale Green */
        rgbPaleTurquoise = 15658671 , /*!< Pale Turquoise */
        rgbPaleVioletRed = 9662683 , /*!< Pale Violet Red */
        rgbPapayaWhip = 14020607 , /*!< Papaya Whip */
        rgbPeachPuff = 12180223 , /*!< Peach Puff */
        rgbPeru = 4163021 , /*!< Peru */
        rgbPink = 13353215 , /*!< Pink */
        rgbPlum = 14524637 , /*!< Plum */
        rgbPowderBlue = 15130800 , /*!< Powder Blue */
        rgbPurple = 8388736 , /*!< Purple */
        rgbRed = 255 , /*!< Red */
        rgbRosyBrown = 9408444 , /*!< Rosy Brown */
        rgbRoyalBlue = 14772545 , /*!< Royal Blue */
        rgbSalmon = 7504122 , /*!< Salmon */
        rgbSandyBrown = 6333684 , /*!< Sandy Brown */
        rgbSeaGreen = 5737262 , /*!< Sea Green */
        rgbSeashell = 15660543 , /*!< Seashell */
        rgbSienna = 2970272 , /*!< Sienna */
        rgbSilver = 12632256 , /*!< Silver */
        rgbSkyBlue = 15453831 , /*!< Sky Blue */
        rgbSlateBlue = 13458026 , /*!< Slate Blue */
        rgbSlateGray = 9470064 , /*!< Slate Gray */
        rgbSnow = 16448255 , /*!< Snow */
        rgbSpringGreen = 8388352 , /*!< Spring Green */
        rgbSteelBlue = 11829830 , /*!< Steel Blue */
        rgbTan = 9221330 , /*!< Tan */
        rgbTeal = 8421376 , /*!< Teal */
        rgbThistle = 14204888 , /*!< Thistle */
        rgbTomato = 4678655 , /*!< Tomato */
        rgbTurquoise = 13688896 , /*!< Turquoise */
        rgbViolet = 15631086 , /*!< Violet */
        rgbWheat = 11788021 , /*!< Wheat */
        rgbWhite = 16777215 , /*!< White */
        rgbWhiteSmoke = 16119285 , /*!< White Smoke */
        rgbYellow = 65535 , /*!< Yellow */
        rgbYellowGreen = 3329434 , /*!< Yellow Green */
    };
    /*!  @brief Specifies how the PivotTable cache connects to its data source. Since Excel 2007.

    [MSDN documentation for XlRobustConnect](http://msdn.microsoft.com/en-us/library/bb241562.aspx).
    */
    enum XlRobustConnect {
        xlAlways = 1 , /*!< The cache always uses external source information (as defined by the SourceConnectionFile or SourceDataFile property) to reconnect. */
        xlAsRequired = 0 , /*!< The cache uses external source information to reconnect using the Connection property. */
        xlNever = 2 , /*!< The cache never uses source information to reconnect. */
    };
    /*!  @brief Specifies the routing delivery method. Since Excel 2007.

    [MSDN documentation for XlRoutingSlipDelivery](http://msdn.microsoft.com/en-us/library/bb241565.aspx).
    */
    enum XlRoutingSlipDelivery {
        xlAllAtOnce = 2 , /*!< Deliver to all recipients simultaneously. */
        xlOneAfterAnother = 1 , /*!< Deliver to recipients one after another. */
    };
    /*!  @brief Specifies the status of the routing slip. Since Excel 2007.

    [MSDN documentation for XlRoutingSlipStatus](http://msdn.microsoft.com/en-us/library/bb241566.aspx).
    */
    enum XlRoutingSlipStatus {
        xlNotYetRouted = 0 , /*!< Routing slip has not yet been sent. */
        xlRoutingComplete = 2 , /*!< Routing is complete. */
        xlRoutingInProgress = 1 , /*!< Routing is in progress. */
    };
    /*!  @brief Specifies whether the values corresponding to a particular data series are in rows or columns. Since Excel 2007.

    [MSDN documentation for XlRowCol](http://msdn.microsoft.com/en-us/library/bb241571.aspx).
    */
    enum XlRowCol {
        xlColumns = 2 , /*!< Data series is in a row. */
        xlRows = 1 , /*!< Data series is in a column. */
    };
    /*!  @brief Specifies the automatic macro to run. Since Excel 2007.

    [MSDN documentation for XlRunAutoMacro](http://msdn.microsoft.com/en-us/library/bb241573.aspx).
    */
    enum XlRunAutoMacro {
        xlAutoActivate = 3 , /*!< Auto_Activate macros */
        xlAutoClose = 2 , /*!< Auto_Close macros */
        xlAutoDeactivate = 4 , /*!< Auto_Deactivate macros */
        xlAutoOpen = 1 , /*!< Auto_Open macros */
    };
    /*!  @brief Specifies during file close if the file will be saved. Since Excel 2007.

    [MSDN documentation for XlSaveAction](http://msdn.microsoft.com/en-us/library/bb241577.aspx).
    */
    enum XlSaveAction {
        xlDoNotSaveChanges = 2 , /*!< Changes will not be saved. */
        xlSaveChanges = 1 , /*!< Changes will be saved. */
    };
    /*!  @brief Specifies the access mode for the Save As function. Since Excel 2007.

    [MSDN documentation for XlSaveAsAccessMode](http://msdn.microsoft.com/en-us/library/bb241580.aspx).
    */
    enum XlSaveAsAccessMode {
        xlExclusive = 3 , /*!< Exclusive mode */
        xlNoChange = 1 , /*!< Default (does not change the access mode) */
        xlShared = 2 , /*!< Share list */
    };
    /*!  @brief Specifies the way that conflicts are to be resolved whenever a shared workbook is updated. Since Excel 2007.

    [MSDN documentation for XlSaveConflictResolution](http://msdn.microsoft.com/en-us/library/bb241584.aspx).
    */
    enum XlSaveConflictResolution {
        xlLocalSessionChanges = 2 , /*!< The local user's changes are always accepted. */
        xlOtherSessionChanges = 3 , /*!< The local user's changes are always rejected. */
        xlUserResolution = 1 , /*!< A dialog box asks the user to resolve the conflict. */
    };
    /*!  @brief Specifies the scale type of the value axis. Since Excel 2007.

    [MSDN documentation for XlScaleType](http://msdn.microsoft.com/en-us/library/bb241587.aspx).
    */
    enum XlScaleType {
        xlScaleLinear = -4132 , /*!< Linear */
        xlScaleLogarithmic = -4133 , /*!< Logarithmic */
    };
    /*!  @brief Specifies the search direction when searching a range. Since Excel 2007.

    [MSDN documentation for XlSearchDirection](http://msdn.microsoft.com/en-us/library/bb241589.aspx).
    */
    enum XlSearchDirection {
        xlNext = 1 , /*!< Search for next matching value in range. */
        xlPrevious = 2 , /*!< Search for previous matching value in range. */
    };
    /*!  @brief Specifies the order in which to search the range. Since Excel 2007.

    [MSDN documentation for XlSearchOrder](http://msdn.microsoft.com/en-us/library/bb241592.aspx).
    */
    enum XlSearchOrder {
        xlByColumns = 2 , /*!< Searches down through a column, then moves to the next column. */
        xlByRows = 1 , /*!< Searches across a row, then moves to the next row. */
    };
    /*!  @brief Specifies the extent of the search for the range. Since Excel 2007.

    [MSDN documentation for XlSearchWithin](http://msdn.microsoft.com/en-us/library/bb241593.aspx).
    */
    enum XlSearchWithin {
        xlWithinSheet = 1 , /*!< Limit search to current sheet. */
        xlWithinWorkbook = 2 , /*!< Search whole workbook. */
    };
    /*!  @brief Specifies the worksheet type. Since Excel 2007.

    [MSDN documentation for XlSheetType](http://msdn.microsoft.com/en-us/library/bb241596.aspx).
    */
    enum XlSheetType {
        xlChart = -4109 , /*!< Chart */
        xlDialogSheet = -4116 , /*!< Dialog sheet */
        xlExcel4IntlMacroSheet = 4 , /*!< Excel version 4 international macro sheet */
        xlExcel4MacroSheet = 3 , /*!< Excel version 4 macro sheet */
        xlWorksheet = -4167 , /*!< Worksheet */
    };
    /*!  @brief Specifies whether the object is visible. Since Excel 2007.

    [MSDN documentation for XlSheetVisibility](http://msdn.microsoft.com/en-us/library/bb241599.aspx).
    */
    enum XlSheetVisibility {
        xlSheetHidden = 0 , /*!< Hides the worksheet which the user can unhide via menu. */
        xlSheetVeryHidden = 2 , /*!< Hides the object so that the only way for you to make it visible again is by setting this property to True (the user cannot make the object visible). */
        xlSheetVisible = -1 , /*!< Displays the sheet. */
    };
    /*!  @brief Specifies what the bubble size represents on a bubble chart. Since Excel 2007.

    [MSDN documentation for XlSizeRepresents](http://msdn.microsoft.com/en-us/library/bb241601.aspx).
    */
    enum XlSizeRepresents {
        xlSizeIsArea = 1 , /*!< Area of the bubble. */
        xlSizeIsWidth = 2 , /*!< Width of the bubble. */
    };
    /*!  @brief Specifies the type of Smart Document control displayed in the

    [MSDN documentation for XlSmartTagControlType](http://msdn.microsoft.com/en-us/library/bb241603.aspx).
    */
    enum XlSmartTagControlType {
        xlSmartTagControlActiveX = 13 , /*!< ActiveX control. */
        xlSmartTagControlButton = 6 , /*!< Button. */
        xlSmartTagControlCheckbox = 9 , /*!< Check box. */
        xlSmartTagControlCombo = 12 , /*!< Combo box. */
        xlSmartTagControlHelp = 3 , /*!< Help text. */
        xlSmartTagControlHelpURL = 4 , /*!< Absolute URL to Help file. */
        xlSmartTagControlImage = 8 , /*!< Image. */
        xlSmartTagControlLabel = 7 , /*!< Label. */
        xlSmartTagControlLink = 2 , /*!< Link. */
        xlSmartTagControlListbox = 11 , /*!< List box. */
        xlSmartTagControlRadioGroup = 14 , /*!< Radio button (option button) group. */
        xlSmartTagControlSeparator = 5 , /*!< Separator. */
        xlSmartTagControlSmartTag = 1 , /*!< Smart tag. */
        xlSmartTagControlTextbox = 10 , /*!< Text box. */
    };
    /*!  @brief Specifies the display features for smart tags.

    [MSDN documentation for XlSmartTagDisplayMode](http://msdn.microsoft.com/en-us/library/bb241605.aspx).
    */
    enum XlSmartTagDisplayMode {
        xlButtonOnly = 2 , /*!< Displays only the button for smart tags. */
        xlDisplayNone = 1 , /*!< Displays  nothing for smart tags. */
        xlIndicatorAndButton = 0 , /*!< Displays the indicator and button for smart tags. */
    };
    /*!  @brief Specifies how to sort text. Since Excel 2007.

    [MSDN documentation for XlSortDataOption](http://msdn.microsoft.com/en-us/library/bb241609.aspx).
    */
    enum XlSortDataOption {
        xlSortNormal = 0 , /*!< default. Sorts numeric and text data separately. */
        xlSortTextAsNumbers = 1 , /*!< Treat text as numeric data for the sort. */
    };
    /*!  @brief Specifies the type of sort.

    [MSDN documentation for XlSortMethod](http://msdn.microsoft.com/en-us/library/bb241613.aspx).
    */
    enum XlSortMethod {
        xlPinYin = 1 , /*!< Phonetic Chinese sort order for characters. This is the default value.  */
        xlStroke = 2 , /*!< Sort by the quantity of strokes in each character. */
    };
    /*!  @brief Specifies how to sort when using East Asian sorting methods.

    [MSDN documentation for XlSortMethodOld](http://msdn.microsoft.com/en-us/library/bb241619.aspx).
    */
    enum XlSortMethodOld {
        xlCodePage = 2 , /*!< Sort by code page. */
        xlSyllabary = 1 , /*!< Sort phonetically. */
    };
    /*!  @brief Specifies the parameter on which the data should be sorted. Since Excel 2007.

    [MSDN documentation for XlSortOn](http://msdn.microsoft.com/en-us/library/bb241621.aspx).
    */
    enum XlSortOn {
        SortOnCellColor = 1 , /*!< Cell color. */
        SortOnFontColor = 2 , /*!< Font color. */
        SortOnIcon = 3 , /*!< Icon. */
        SortOnValues = 0 , /*!< Values. */
    };
    /*!  @brief Specifies the sort order for the specified field or range.

    [MSDN documentation for XlSortOrder](http://msdn.microsoft.com/en-us/library/bb241626.aspx).
    */
    enum XlSortOrder {
        xlAscending = 1 , /*!< Sorts the specified field in ascending order. This is the default value. */
        xlDescending = 2 , /*!< Sorts the specified field in descending order. */
    };
    /*!  @brief Specifies the sort orientation.

    [MSDN documentation for XlSortOrientation](http://msdn.microsoft.com/en-us/library/bb241629.aspx).
    */
    enum XlSortOrientation {
        xlSortColumns = 1 , /*!< Sorts by column. */
        xlSortRows = 2 , /*!< Sorts by row. This is the default value. */
    };
    /*!  @brief Specifies which elements are to be sorted. Use this argument only when sorting PivotTable reports.

    [MSDN documentation for XlSortType](http://msdn.microsoft.com/en-us/library/bb241632.aspx).
    */
    enum XlSortType {
        xlSortLabels = 2 , /*!< Sorts the PivotTable report by labels. */
        xlSortValues = 1 , /*!< Sorts the PivotTable report by values. */
    };
    /*!Identifies the source object.

    [MSDN documentation for XlSourceType](http://msdn.microsoft.com/en-us/library/bb241636.aspx).
    */
    enum XlSourceType {
        xlSourceAutoFilter = 3 , /*!< An AutoFilter range */
        xlSourcePivotTable = 6 , /*!< A PivotTable report */
        xlSourcePrintArea = 2 , /*!< A range of cells selected for printing */
        xlSourceQuery = 7 , /*!< A query table (external data range) */
        xlSourceRange = 4 , /*!< A range of cells */
        xlSourceSheet = 1 , /*!< An entire worksheet */
        xlSourceWorkbook = 0 , /*!< A workbook */
    };
    
    /*!  @brief Specifies the order in which the cells are spoken.

    [MSDN documentation for XlSpeakDirection](http://msdn.microsoft.com/en-us/library/bb216303.aspx).
    */
    enum XlSpeakDirection {
        xlSpeakByColumns = 1 , /*!< Reads down a column, then moves to the next column. */
        xlSpeakByRows = 0 , /*!< Reads across a row, then moves to the next row. */
    };
    
    /*!  @brief Specifies the mode for checking the spelling of Spanish. Since Excel 2010.

    [MSDN documentation for XlSpanishModes](http://msdn.microsoft.com/en-us/library/office/ff836842%28v=office.14%29.aspx).
    */
    enum XlSpanishModes {
        xlSpanishTuteoAndVoseo = 1, /*!< Tuteo and Voseo verb forms. . */
        xlSpanishTuteoOnly = 0, /*!< Tuteo verb forms only. . */
        xlSpanishVoseoOnly = 2, /*!< Voseo verb forms only. . */
    };

    /*!  @brief Specifies how to plot the sparkline when the data on which it is based is in a square-shaped range. Since Excel 2010.

    [MSDN documentation for XlSparklineRowCol](http://msdn.microsoft.com/en-us/library/office/ff194726%28v=office.14%29.aspx).
    */
    enum XlSparklineRowCol {
        xlSparklineColumnsSquare = 2, /*!< Plot the data by columns. */
        xlSparklineNonSquare = 0, /*!< The sparkline is not bound to data in a square-shaped range. */
        xlSparklineRowsSquare = 1, /*!< Plot the data by rows. */        
    };

    /*!  @brief Specifies how the minimum or maximum value of the vertical axis of the sparkline is scaled relative to other sparklines in the group. Since Excel 2010.

    [MSDN documentation for XlSparkScale](http://msdn.microsoft.com/en-us/library/office/ff822790%28v=office.14%29.aspx).
    */
    enum XlSparkScale {
        xlSparkScaleCustom = 3, /*!< The minimum or maximum value for the vertical axis of the sparkline has a user-defined value. . */
        xlSparkScaleGroup = 1, /*!< The minimum or maximum value for the vertical axes of all of the sparklines in the group have the same value. . */
        xlSparkScaleSingle = 2, /*!< The minimum or maximum value for the vertical axis of each sparkline in the group is automatically set to its own calculated value. . */
    };


    /*!  @brief Specifies the type of sparkline. Since Excel 2010.

    [MSDN documentation for XlSparkType](http://msdn.microsoft.com/en-us/library/office/ff841146%28v=office.14%29.aspx).
    */
    enum XlSparkType  {
        xlSparkColumn = 2, /*!< A column chart sparkline. . */
        xlSparkColumnStacked100 = 3, /*!< A win/loss chart sparkline. . */
        xlSparkLine = 1, /*!< A line chart sparkline. . */       
    };


    /*!  @brief Specifies cells with a particular type of value to include in the result. Since Excel 2007.

    [MSDN documentation for XlSpecialCellsValue](http://msdn.microsoft.com/en-us/library/bb216309.aspx).
    */
    enum XlSpecialCellsValue {
        xlErrors = 16 , /*!< Cells with errors. */
        xlLogical = 4 , /*!< Cells with logical values. */
        xlNumbers = 1 , /*!< Cells with numeric values. */
        xlTextValues = 2 , /*!< Cells with text. */
    };
    /*!  @brief Specifies the standard color scale. Since Excel 2007.

    [MSDN documentation for XlStdColorScale](http://msdn.microsoft.com/en-us/library/bb216312.aspx).
    */
    enum XlStdColorScale {
        ColorScaleBlackWhite = 3 , /*!< Black over White. */
        ColorScaleGYR = 2 , /*!< GYR. */
        ColorScaleRYG = 1 , /*!< RYG. */
        ColorScaleWhiteBlack = 4 , /*!< White over Black. */
    };
    /*!  @brief Specifies the format to use when subscribing to a published edition. Since Excel 2007.

    [MSDN documentation for XlSubscribeToFormat](http://msdn.microsoft.com/en-us/library/bb216318.aspx).
    */
    enum XlSubscribeToFormat {
        xlSubscribeToPicture = -4147 , /*!< Picture */
        xlSubscribeToText = -4158 , /*!< Text */
    };
    /*!  @brief Specifies where the subtotal will be displayed on the worksheet. Since Excel 2007.

    [MSDN documentation for XlSubtototalLocationType](http://msdn.microsoft.com/en-us/library/bb216321.aspx).
    */
    enum XlSubtototalLocationType {
        xlAtBottom = 2 , /*!< Subtotal will be at the bottom. */
        xlAtTop = 1 , /*!< Subtotal will be at the top. */
    };
    /*!  @brief Specifies the location of the summary columns in the outline. Since Excel 2007.

    [MSDN documentation for XlSummaryColumn](http://msdn.microsoft.com/en-us/library/bb216326.aspx).
    */
    enum XlSummaryColumn {
        xlSummaryOnLeft = -4131 , /*!< The summary column will be positioned to the left of the detail columns in the outline. */
        xlSummaryOnRight = -4152 , /*!< The summary column will be positioned to the right of the detail columns in the outline. */
    };
    /*!  @brief Specifies the type of summary to be created for scenarios. Since Excel 2007.

    [MSDN documentation for XlSummaryReportType](http://msdn.microsoft.com/en-us/library/bb216330.aspx).
    */
    enum XlSummaryReportType {
        xlStandardSummary = 1 , /*!< List scenarios side by side. */
        xlSummaryPivotTable = -4148 , /*!< Display scenarios in a PivotTable report. */
    };
    /*!  @brief Specifies the location of the summary rows in the outline. Since Excel 2007.

    [MSDN documentation for XlSummaryRow](http://msdn.microsoft.com/en-us/library/bb216336.aspx).
    */
    enum XlSummaryRow {
        xlSummaryAbove = 0 , /*!< The summary row will be positioned above the detail rows in the outline. */
        xlSummaryBelow = 1 , /*!< The summary row will be positioned below the detail rows in the outline. */
    };
    /*!  @brief Specifies the table style element used. Since Excel 2007.

    [MSDN documentation for XlTableStyleElementType](http://msdn.microsoft.com/en-us/library/office/ff196883%28v=office.14%29.aspx).
    */
    enum XlTableStyleElementType {
        xlBlankRow = 19, /*!< Blank row */
        xlColumnStripe1 = 7, /*!< Column Stripe1 */
        xlColumnStripe2 = 8, /*!< Column Stripe2 */
        xlColumnSubheading1 = 20, /*!< Column Subheading1 */
        xlColumnSubheading2 = 21, /*!< Column Subheading2 */
        xlColumnSubheading3 = 22, /*!< Column Subheading3 */
        xlFirstColumn = 3, /*!< First column */
        xlFirstHeaderCell = 9, /*!< First header cell */
        xlFirstTotalCell = 11, /*!< First total cell */
        xlGrandTotalColumn = 4, /*!< Grand total column */
        xlGrandTotalRow = 2, /*!< Grand total row */
        xlHeaderRow = 1, /*!< Header row */
        xlLastColumn = 4, /*!< Last column */
        xlLastHeaderCell = 10, /*!< Last header cell */
        xlLastTotalCell = 12, /*!< Last total cell */
        xlPageFieldLabels = 26, /*!< Page field labels */
        xlPageFieldValues = 27, /*!< Page field values */
        xlRowStripe1 = 5, /*!< Row Stripe1 */
        xlRowStripe2 = 6, /*!< Row Stripe2 */
        xlRowSubheading1 = 23, /*!< Row Subheading1 */
        xlRowSubheading2 = 24, /*!< Row Subheading2 */
        xlRowSubheading3 = 25, /*!< Row Subheading3 */
        xlSlicerHoveredSelectedItemWithData = 33, /*!< A selected item, hovered over by the user, that contains data. */
        xlSlicerHoveredSelectedItemWithNoData = 35, /*!< A selected item, hovered over by the user, that does not contain data. */
        xlSlicerHoveredUnselectedItemWithData = 32, /*!< An item, hovered over by the user, that is not selected and that contains data. */
        xlSlicerHoveredUnselectedItemWithNoData = 34, /*!< A selected item, hovered over by the user, that is not selected and that does not contain data. */
        xlSlicerSelectedItemWithData = 30, /*!< A selected item that contains data. */
        xlSlicerSelectedItemWithNoData = 31, /*!< A selected item that does not contain data. */
        xlSlicerUnselectedItemWithData = 28, /*!< An item that is not selected that contains data. */
        xlSlicerUnselectedItemWithNoData = 29, /*!< An item that is not selected that does not contain data. */
        xlSubtotalColumn1 = 13, /*!< Subtotal Column1 */
        xlSubtotalColumn2 = 14, /*!< Subtotal Column2 */
        xlSubtotalColumn3 = 15, /*!< Subtotal Column3 */
        xlSubtotalRow1 = 16, /*!< Subtotal Row1 */
        xlSubtotalRow2 = 17, /*!< Subtotal Row2 */
        xlSubtotalRow3 = 18, /*!< Subtotal Row3 */
        xlTotalRow = 2, /*!< Total Row */
        xlWholeTable = 0, /*!< Whole Table */
    };
    /*!  @brief Specifies the first or last tab position. Since Excel 2007.

    [MSDN documentation for XlTabPosition](http://msdn.microsoft.com/en-us/library/bb216343.aspx).
    */
    enum XlTabPosition {
        xlTabPositionFirst = 0 , /*!< First tab position. */
        xlTabPositionLast = 1 , /*!< Last tab position. */
    };
    /*!  @brief Specifies the column format for the data in the text file that you are importing into a query table. Since Excel 2007.

    [MSDN documentation for XlTextParsingType](http://msdn.microsoft.com/en-us/library/bb216347.aspx).
    */
    enum XlTextParsingType {
        xlDelimited = 1 , /*!< Default. Indicates that the file is delimited by delimiter characters. */
        xlFixedWidth = 2 , /*!< Indicates that the data in the file is arranged in columns of fixed widths. */
    };
    /*!  @brief Specifies the delimiter to use to specify text. Since Excel 2007.

    [MSDN documentation for XlTextQualifier](http://msdn.microsoft.com/en-us/library/bb216350.aspx).
    */
    enum XlTextQualifier {
        xlTextQualifierDoubleQuote = 1 , /*!< Double quotation mark ("). */
        xlTextQualifierNone = -4142 , /*!< No delimiter. */
        xlTextQualifierSingleQuote = 2 , /*!< Single quotation mark ('). */
    };
    /*!  @brief Specifies whether the visual layout of the text being imported is left-to-right or right-to-left. Since Excel 2007.

    [MSDN documentation for XlTextVisualLayoutType](http://msdn.microsoft.com/en-us/library/bb216354.aspx).
    */
    enum XlTextVisualLayoutType {
        xlTextVisualLTR = 1 , /*!< Left-to-right */
        xlTextVisualRTL = 2 , /*!< Right-to-left */
    };
    /*!  @brief Specifies the theme color to be used. Since Excel 2007.

    [MSDN documentation for XlThemeColor](http://msdn.microsoft.com/en-us/library/bb216356.aspx).
    */
    enum XlThemeColor {
        xlThemeColorAccent1 = 5 , /*!< Accent1 */
        xlThemeColorAccent2 = 6 , /*!< Accent2 */
        xlThemeColorAccent3 = 7 , /*!< Accent3 */
        xlThemeColorAccent4 = 8 , /*!< Accent4 */
        xlThemeColorAccent5 = 9 , /*!< Accent5 */
        xlThemeColorAccent6 = 10 , /*!< Accent6 */
        xlThemeColorDark1 = 1 , /*!< Dark1 */
        xlThemeColorDark2 = 3 , /*!< Dark2 */
        xlThemeColorFollowedHyperlink = 12 , /*!< Followed hyperlink */
        xlThemeColorHyperlink = 11 , /*!< Hyperlink */
        xlThemeColorLight1 = 2 , /*!< Light1 */
        xlThemeColorLight2 = 4 , /*!< Light2 */
    };
    /*!  @brief Specifies the theme font to use. Since Excel 2007.

    [MSDN documentation for XlThemeFont](http://msdn.microsoft.com/en-us/library/bb216361.aspx).
    */
    enum XlThemeFont {
        xlThemeFontMajor = 2 , /*!< Major. */
        xlThemeFontMinor = 1 , /*!< Minor. */
        xlThemeFontNone = 0 , /*!< Do not use any theme font. */
    };
    /*!  @brief Specifies the control over the multi-threaded calculation mode. Since Excel 2007.

    [MSDN documentation for XlThreadMode](http://msdn.microsoft.com/en-us/library/bb257018.aspx).
    */
    enum XlThreadMode {
        xlThreadModeAutomatic = 0 , /*!< Multi-threaded calculation mode is automatic. */
        xlThreadModeManual = 1 , /*!< Multi-threaded calculation mode is manual. */
    };
    /*!  @brief Specifies the text orientation for tick-mark labels.

    [MSDN documentation for XlTickLabelOrientation](http://msdn.microsoft.com/en-us/library/bb216366.aspx).
    */
    enum XlTickLabelOrientation {
        xlTickLabelOrientationAutomatic = -4105 , /*!< Text orientation set by Excel. */
        xlTickLabelOrientationDownward = -4170 , /*!< Text runs down. */
        xlTickLabelOrientationHorizontal = -4128 , /*!< Characters run horizontally. */
        xlTickLabelOrientationUpward = -4171 , /*!< Text runs up. */
        xlTickLabelOrientationVertical = -4166 , /*!< Characters run vertically. */
    };
    /*!  @brief Specifies the position of tick-mark labels on the specified axis.

    [MSDN documentation for XlTickLabelPosition](http://msdn.microsoft.com/en-us/library/bb216368.aspx).
    */
    enum XlTickLabelPosition {
        xlTickLabelPositionHigh = -4127 , /*!< Top or right side of the chart. */
        xlTickLabelPositionLow = -4134 , /*!< Bottom or left side of the chart. */
        xlTickLabelPositionNextToAxis = 4 , /*!< Next to axis (where axis is not at either side of the chart). */
        xlTickLabelPositionNone = -4142 , /*!< No tick marks. */
    };
    /*!  @brief Specifies the position of major and minor tick marks for an axis. Since Excel 2007.

    [MSDN documentation for XlTickMark](http://msdn.microsoft.com/en-us/library/bb216376.aspx).
    */
    enum XlTickMark {
        xlTickMarkCross = 4 , /*!< Crosses the axis */
        xlTickMarkInside = 2 , /*!< Inside the axis */
        xlTickMarkNone = -4142 , /*!< No mark */
        xlTickMarkOutside = 3 , /*!< Outside the axis */
    };
    /*!  @brief Specifies the time period. Since Excel 2007.

    [MSDN documentation for XlTimePeriods](http://msdn.microsoft.com/en-us/library/bb216381.aspx).
    */
    enum XlTimePeriods {
        xlLast7Days = 2 , /*!< Last 7 days */
        xlLastMonth = 5 , /*!< Last month */
        xlLastWeek = 4 , /*!< Last week */
        xlNextMonth = 8 , /*!< Next month */
        xlNextWeek = 7 , /*!< Next week */
        xlThisMonth = 9 , /*!< This month */
        xlThisWeek = 3 , /*!< This week */
        xlToday = 0 , /*!< Today */
        xlTomorrow = 6 , /*!< Tomorrow */
        xlYesterday = 1 , /*!< Yesterday */
    };
    /*!  @brief Specifies the unit of time for chart axes and data series. Since Excel 2007.

    [MSDN documentation for XlTimeUnit](http://msdn.microsoft.com/en-us/library/bb216385.aspx).
    */
    enum XlTimeUnit {
        xlDays = 0 , /*!< Days */
        xlMonths = 1 , /*!< Months */
        xlYears = 2 , /*!< Years */
    };
    /*!  @brief Specifies which properties of a toolbar are restricted. Options can be combined using Or. Since Excel 2007.

    [MSDN documentation for XlToolbarProtection](http://msdn.microsoft.com/en-us/library/bb216391.aspx).
    */
    enum XlToolbarProtection {
        xlNoButtonChanges = 1 , /*!< No button changes permitted. */
        xlNoChanges = 4 , /*!< No changes of any kind. */
        xlNoDockingChanges = 3 , /*!< No changes to toolbar's docking position. */
        xlNoShapeChanges = 2 , /*!< No changes to toolbar shape. */
        xlToolbarProtectionNone = -4143 , /*!< All changes permitted. */
    };
    /*!  @brief Specifies the top 10 values from the top or bottom of a series of values. Since Excel 2007.

    [MSDN documentation for XlTopBottom](http://msdn.microsoft.com/en-us/library/bb216393.aspx).
    */
    enum XlTopBottom {
        xlTop10Bottom = 0 , /*!< Top 10 bottom values */
        xlTop10Top = 1 , /*!< Top 10 values */
    };
    /*!  @brief Specifies the type of calculation in the Totals row of a list column. Since Excel 2007.

    [MSDN documentation for XlTotalsCalculation](http://msdn.microsoft.com/en-us/library/bb216397.aspx).
    */
    enum XlTotalsCalculation {
        xlTotalsCalculationAverage = 2 , /*!< Average */
        xlTotalsCalculationCount = 3 , /*!< Count of non-empty cells */
        xlTotalsCalculationCountNums = 4 , /*!< Count of cells with numeric values */
        xlTotalsCalculationCustom = 9 , /*!< Custom calculation */
        xlTotalsCalculationMax = 6 , /*!< Maximum value in the list */
        xlTotalsCalculationMin = 5 , /*!< Minimum value in the list */
        xlTotalsCalculationNone = 0 , /*!< No calculation */
        xlTotalsCalculationStdDev = 7 , /*!< Standard deviation value */
        xlTotalsCalculationSum = 1 , /*!< Sum of all values in the list column */
        xlTotalsCalculationVar = 8 , /*!< Variable */
    };
    /*!  @brief Specifies how the trendline that smoothes out fluctuations in the data is calculated.

    [MSDN documentation for XlTrendlineType](http://msdn.microsoft.com/en-us/library/bb216402.aspx).
    */
    enum XlTrendlineType {
        xlExponential = 5 , /*!< Uses an equation to calculate the least squares fit through points, for example, y=ab^x . */
        xlLinear = -4132 , /*!< Uses the linear equation y = mx + b to calculate the least squares fit through points. */
        xlLogarithmic = -4133 , /*!< Uses the equation y = c ln x + b to calculate the least squares fit through points. */
        xlMovingAvg = 6 , /*!< Uses a sequence of averages computed from parts of the data series. The number of points equals the total number of points in the series less the number specified for the period. */
        xlPolynomial = 3 , /*!< Uses an equation to calculate the least squares fit through points, for example, y = ax^6 + bx^5 + cx^4 + dx^3 + ex^2 + fx + g. */
        xlPower = 4 , /*!< Uses an equation to calculate the least squares fit through points, for example, y = ax^b. */
    };
    /*!  @brief Specifies the type of underline applied to a font.

    [MSDN documentation for XlUnderlineStyle](http://msdn.microsoft.com/en-us/library/bb216406.aspx).
    */
    enum XlUnderlineStyle {
        xlUnderlineStyleDouble = -4119 , /*!< Double thick underline. */
        xlUnderlineStyleDoubleAccounting = 5 , /*!< Two thin underlines placed close together. */
        xlUnderlineStyleNone = -4142 , /*!< No underlining. */
        xlUnderlineStyleSingle = 2 , /*!< Single underlining. */
        xlUnderlineStyleSingleAccounting = 4 , /*!< Not supported. */
    };
    /*!  @brief Specifies a workbook's setting for updating embedded OLE links. Since Excel 2007.

    [MSDN documentation for XlUpdateLinks](http://msdn.microsoft.com/en-us/library/bb216410.aspx).
    */
    enum XlUpdateLinks {
        xlUpdateLinksAlways = 3 , /*!< Embedded OLE links are always updated for the specified workbook. */
        xlUpdateLinksNever = 2 , /*!< Embedded OLE links are never updated for the specified workbook. */
        xlUpdateLinksUserSetting = 1 , /*!< Embedded OLE links are updated according to the user's settings for the specified workbook. */
    };
    /*!  @brief Specifies the vertical alignment for the object. Since Excel 2007.

    [MSDN documentation for XlVAlign](http://msdn.microsoft.com/en-us/library/bb216413.aspx).
    */
    enum XlVAlign {
        xlVAlignBottom = -4107 , /*!< Bottom */
        xlVAlignCenter = -4108 , /*!< Center */
        xlVAlignDistributed = -4117 , /*!< Distributed */
        xlVAlignJustify = -4130 , /*!< Justify */
        xlVAlignTop = -4160 , /*!< Top */
    };
    /*!  @brief Specifies the type of workbook to create. The new workbook contains a single sheet of the specified type. Since Excel 2007.

    [MSDN documentation for XlWBATemplate](http://msdn.microsoft.com/en-us/library/bb216417.aspx).
    */
    enum XlWBATemplate {
        xlWBATChart = -4109 , /*!< Chart */
        xlWBATExcel4IntlMacroSheet = 4 , /*!< Excel version 4 macro */
        xlWBATExcel4MacroSheet = 3 , /*!< Excel version 4 international macro */
        xlWBATWorksheet = -4167 , /*!< Worksheet */
    };
    /*!  @brief Specifies how much formatting from a Web page, if any, is applied when a Web page is imported into a query table. Since Excel 2007.

    [MSDN documentation for XlWebFormatting](http://msdn.microsoft.com/en-us/library/bb216421.aspx).
    */
    enum XlWebFormatting {
        xlWebFormattingAll = 1 , /*!< All formatting is imported. */
        xlWebFormattingNone = 3 , /*!< No formatting is imported. */
        xlWebFormattingRTF = 2 , /*!< Rich Text Format–compatible formatting is imported. */
    };
    /*!  @brief Specifies whether an entire Web page, all tables on the Web page, or only a specific table is imported into a query table. Since Excel 2007.

    [MSDN documentation for XlWebSelectionType](http://msdn.microsoft.com/en-us/library/bb216424.aspx).
    */
    enum XlWebSelectionType {
        xlAllTables = 2 , /*!< All tables */
        xlEntirePage = 1 , /*!< Entire page */
        xlSpecifiedTables = 3 , /*!< Specified tables */
    };
    /*!  @brief Specifies the state of the window.

    [MSDN documentation for XlWindowState](http://msdn.microsoft.com/en-us/library/bb216429.aspx).
    */
    enum XlWindowState {
        xlMaximized = -4137 , /*!< Maximized */
        xlMinimized = -4140 , /*!< Minimized */
        xlNormal = -4143 , /*!< Normal */
    };
    /*!  @brief Specifies how the chart is displayed.

    [MSDN documentation for XlWindowType](http://msdn.microsoft.com/en-us/library/bb216432.aspx).
    */
    enum XlWindowType {
        xlChartAsWindow = 5 , /*!< The chart will open in a new window. */
        xlChartInPlace = 4 , /*!< The chart will be displayed on the current worksheet. */
        xlClipboard = 3 , /*!< The chart is copied to the clipboard. */
        xlInfo = -4129 , /*!< This constant has been deprecated. */
        xlWorkbook = 1 , /*!< This constant applies to Macintosh only. */
    };
    /*!  @brief Specifies the view showing in the window. Since Excel 2007.

    [MSDN documentation for XlWindowView](http://msdn.microsoft.com/en-us/library/bb216434.aspx).
    */
    enum XlWindowView {
        xlNormalView = 1 , /*!< Normal. */
        xlPageBreakPreview = 2 , /*!< Page break preview. */
        xlPageLayoutView = 3 , /*!< Page layout view. */
    };
    /*!  @brief Specifies, in a Microsoft Excel version 4 macro worksheet, what type of macro a name refers to or whether the name refers to a macro. Since Excel 2007.

    [MSDN documentation for XlXLMMacroType](http://msdn.microsoft.com/en-us/library/bb216437.aspx).
    */
    enum XlXLMMacroType {
        xlCommand = 2 , /*!< Custom command. */
        xlFunction = 1 , /*!< Custom function. */
        xlNotXLM = 3 , /*!< Not a macro. */
    };
    /*!  @brief Specifies the results of the save or export operation.

    [MSDN documentation for XlXmlExportResult](http://msdn.microsoft.com/en-us/library/bb216440.aspx).
    */
    enum XlXmlExportResult {
        xlXmlExportSuccess = 0 , /*!< The XML data file was successfully exported. */
        xlXmlExportValidationFailed = 1 , /*!< The contents of the XML data file do not match the specified schema map. */
    };
    /*!  @brief Specifies the results of the refresh or import operation.

    [MSDN documentation for XlXmlImportResult](http://msdn.microsoft.com/en-us/library/bb216442.aspx).
    */
    enum XlXmlImportResult {
        xlXmlImportElementsTruncated = 1 , /*!< The contents of the specified XML data file have been truncated because the XML data file is too large for the worksheet. */
        xlXmlImportSuccess = 0 , /*!< The XML data file was successfully imported. */
        xlXmlImportValidationFailed = 2 , /*!< The contents of the XML data file do not match the specified schema map. */
    };
    /*!  @brief Specifies how Excel opens the XML data file.

    [MSDN documentation for XlXmlLoadOption](http://msdn.microsoft.com/en-us/library/bb216445.aspx).
    */
    enum XlXmlLoadOption {
        xlXmlLoadImportToList = 2 , /*!< Places the contents of the XML data file in an XML table. */
        xlXmlLoadMapXml = 3 , /*!< Displays the schema of the XML data file in the XML Structure task pane. */
        xlXmlLoadOpenXml = 1 , /*!< Opens the XML data file. The contents of the file will be flattened. */
        xlXmlLoadPromptUser = 0 , /*!< Prompts the user to choose how to open the file. */
    };
    /*!  @brief Specifies whether or not the first row contains headers. Cannot be used when sorting PivotTable reports. Since Excel 2007.

    [MSDN documentation for XlYesNoGuess](http://msdn.microsoft.com/en-us/library/bb216447.aspx).
    */
    enum XlYesNoGuess {
        xlGuess = 0 , /*!< Excel determines whether there is a header, and where it is, if there is one. */
        xlNo = 2 , /*!< Default. The entire range should be sorted. */
        xlYes = 1 , /*!< The entire range should not be sorted. */
    };


} // namespace wxAutoExcel


#endif //_WXAUTOEXCEL_ENUMS_H
