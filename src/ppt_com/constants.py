"""PowerPoint and MSO COM Automation Constants.

Central repository of all enumeration constants used across the MCP server.
All constants are defined as module-level integers grouped by enumeration type.
"""

# ==============================================================================
# MsoTriState
# ==============================================================================
msoTrue = -1
msoFalse = 0
msoCTrue = 1
msoTriStateToggle = -3
msoTriStateMixed = -2

# ==============================================================================
# PpWindowState
# ==============================================================================
ppWindowNormal = 1
ppWindowMinimized = 2
ppWindowMaximized = 3

# ==============================================================================
# PpViewType
# ==============================================================================
ppViewSlide = 1
ppViewSlideMaster = 2
ppViewNotesPage = 3
ppViewHandoutMaster = 4
ppViewNotesMaster = 5
ppViewOutline = 6
ppViewSlideSorter = 7
ppViewTitleMaster = 8
ppViewNormal = 9
ppViewPrintPreview = 10
ppViewThumbnails = 11
ppViewMasterThumbnails = 12

# ==============================================================================
# PpSlideLayout
# ==============================================================================
ppLayoutTitle = 1
ppLayoutText = 2
ppLayoutTwoColumnText = 3
ppLayoutTable = 4
ppLayoutTextAndChart = 5
ppLayoutChartAndText = 6
ppLayoutOrgchart = 7
ppLayoutChart = 8
ppLayoutTextAndClipArt = 9
ppLayoutClipArtAndText = 10
ppLayoutTitleOnly = 11
ppLayoutBlank = 12
ppLayoutTextAndObject = 13
ppLayoutObjectAndText = 14
ppLayoutLargeObject = 15
ppLayoutObject = 16
ppLayoutTextAndMediaClip = 17
ppLayoutMediaClipAndText = 18
ppLayoutObjectOverText = 19
ppLayoutTextOverObject = 20
ppLayoutTextAndTwoObjects = 21
ppLayoutTwoObjectsAndText = 22
ppLayoutTwoObjectsOverText = 23
ppLayoutFourObjects = 24
ppLayoutVerticalText = 25
ppLayoutClipArtAndVerticalText = 26
ppLayoutVerticalTitleAndText = 27
ppLayoutVerticalTitleAndTextOverChart = 28
ppLayoutTwoObjects = 29
ppLayoutObjectAndTwoObjects = 30
ppLayoutTwoObjectsAndObject = 31
ppLayoutCustom = 32
ppLayoutSectionHeader = 33
ppLayoutComparison = 34
ppLayoutContentWithCaption = 35
ppLayoutPictureWithCaption = 36
ppLayoutMixed = -2

# ==============================================================================
# PpSaveAsFileType
# ==============================================================================
ppSaveAsPresentation = 1
ppSaveAsTemplate = 5
ppSaveAsRTF = 6
ppSaveAsShow = 7
ppSaveAsAddIn = 8
ppSaveAsDefault = 11
ppSaveAsGIF = 16
ppSaveAsJPG = 17
ppSaveAsPNG = 18
ppSaveAsBMP = 19
ppSaveAsTIF = 21
ppSaveAsOpenXMLPresentation = 24
ppSaveAsOpenXMLPresentationMacroEnabled = 25
ppSaveAsOpenXMLTemplate = 26
ppSaveAsOpenXMLShow = 28
ppSaveAsOpenXMLAddin = 30
ppSaveAsOpenXMLTheme = 31
ppSaveAsPDF = 32
ppSaveAsXPS = 33
ppSaveAsOpenDocumentPresentation = 35
ppSaveAsWMV = 37
ppSaveAsMP4 = 39

# ==============================================================================
# MsoShapeType
# ==============================================================================
msoAutoShape = 1
msoCallout = 2
msoChart = 3
msoComment = 4
msoFreeform = 5
msoGroup = 6
msoEmbeddedOLEObject = 7
msoFormControl = 8
msoLine = 9
msoLinkedOLEObject = 10
msoLinkedPicture = 11
msoOLEControlObject = 12
msoPicture = 13
msoPlaceholder = 14
msoTextEffect = 15
msoMedia = 16
msoTextBox = 17
msoTable = 19
msoSmartArt = 24

# ==============================================================================
# MsoAutoShapeType (commonly used)
# ==============================================================================
msoShapeRectangle = 1
msoShapeParallelogram = 2
msoShapeTrapezoid = 3
msoShapeDiamond = 4
msoShapeRoundedRectangle = 5
msoShapeOctagon = 6
msoShapeIsoscelesTriangle = 7
msoShapeRightTriangle = 8
msoShapeOval = 9
msoShapeHexagon = 10
msoShapeCross = 11
msoShapeRegularPentagon = 12
msoShapeCan = 13
msoShapeCube = 14
msoShapeSmileyFace = 17
msoShapeDonut = 18
msoShapeNoSymbol = 19
msoShapeHeart = 21
msoShapeLightningBolt = 22
msoShapeSun = 23
msoShapeMoon = 24
msoShapeArc = 25
msoShapeRightArrow = 33
msoShapeLeftArrow = 34
msoShapeUpArrow = 35
msoShapeDownArrow = 36
msoShapeLeftRightArrow = 37
msoShapeUpDownArrow = 38
msoShapeQuadArrow = 39
msoShapePentagon = 51
msoShapeChevron = 52
msoShapeFlowchartProcess = 61
msoShapeFlowchartDecision = 63
msoShapeFlowchartData = 64
msoShapeFlowchartDocument = 67
msoShapeFlowchartTerminator = 69
msoShapeFlowchartConnector = 73
msoShapeExplosion1 = 89
msoShapeExplosion2 = 90
msoShape4pointStar = 91
msoShape5pointStar = 92
msoShape8pointStar = 93
msoShape16pointStar = 94
msoShape24pointStar = 95
msoShape32pointStar = 96
msoShapeCloud = 179

# ==============================================================================
# PpPlaceholderType
# ==============================================================================
ppPlaceholderMixed = -2
ppPlaceholderTitle = 1
ppPlaceholderBody = 2
ppPlaceholderCenterTitle = 3
ppPlaceholderSubtitle = 4
ppPlaceholderVerticalTitle = 5
ppPlaceholderVerticalBody = 6
ppPlaceholderObject = 7
ppPlaceholderChart = 8
ppPlaceholderBitmap = 9
ppPlaceholderMediaClip = 10
ppPlaceholderOrgChart = 11
ppPlaceholderTable = 12
ppPlaceholderSlideNumber = 13
ppPlaceholderHeader = 14
ppPlaceholderFooter = 15
ppPlaceholderDate = 16
ppPlaceholderVerticalObject = 17
ppPlaceholderPicture = 18

# ==============================================================================
# PpParagraphAlignment
# ==============================================================================
ppAlignLeft = 1
ppAlignCenter = 2
ppAlignRight = 3
ppAlignJustify = 4
ppAlignDistribute = 5

# ==============================================================================
# MsoTextOrientation
# ==============================================================================
msoTextOrientationHorizontal = 1
msoTextOrientationUpward = 2
msoTextOrientationDownward = 3
msoTextOrientationVertical = 5
msoTextOrientationVerticalFarEast = 6

# ==============================================================================
# PpAutoSize
# ==============================================================================
ppAutoSizeNone = 0
ppAutoSizeShapeToFitText = 1
ppAutoSizeMixed = -2

# ==============================================================================
# MsoZOrderCmd
# ==============================================================================
msoBringToFront = 0
msoSendToBack = 1
msoBringForward = 2
msoSendBackward = 3

# ==============================================================================
# MsoFlipCmd
# ==============================================================================
msoFlipHorizontal = 0
msoFlipVertical = 1

# ==============================================================================
# MsoFillType
# ==============================================================================
msoFillSolid = 1
msoFillPatterned = 2
msoFillGradient = 3
msoFillTextured = 4
msoFillBackground = 5
msoFillPicture = 6

# ==============================================================================
# MsoGradientStyle
# ==============================================================================
msoGradientHorizontal = 1
msoGradientVertical = 2
msoGradientDiagonalUp = 3
msoGradientDiagonalDown = 4
msoGradientFromCorner = 5
msoGradientFromCenter = 7

# ==============================================================================
# MsoLineDashStyle
# ==============================================================================
msoLineSolid = 1
msoLineRoundDot = 2
msoLineDot = 3
msoLineDash = 4
msoLineDashDot = 5
msoLineDashDotDot = 6
msoLineLongDash = 7
msoLineLongDashDot = 8

# ==============================================================================
# MsoArrowheadStyle
# ==============================================================================
msoArrowheadNone = 1
msoArrowheadTriangle = 2
msoArrowheadOpen = 3
msoArrowheadStealth = 4
msoArrowheadDiamond = 5
msoArrowheadOval = 6

# ==============================================================================
# MsoThemeColorIndex
# ==============================================================================
msoThemeColorDark1 = 1
msoThemeColorLight1 = 2
msoThemeColorDark2 = 3
msoThemeColorLight2 = 4
msoThemeColorAccent1 = 5
msoThemeColorAccent2 = 6
msoThemeColorAccent3 = 7
msoThemeColorAccent4 = 8
msoThemeColorAccent5 = 9
msoThemeColorAccent6 = 10
msoThemeColorHyperlink = 11
msoThemeColorFollowedHyperlink = 12

# ==============================================================================
# PpBulletType
# ==============================================================================
ppBulletNone = 0
ppBulletUnnumbered = 1
ppBulletNumbered = 2
ppBulletPicture = 3

# ==============================================================================
# PpSelectionType
# ==============================================================================
ppSelectionNone = 0
ppSelectionSlides = 1
ppSelectionShapes = 2
ppSelectionText = 3

# ==============================================================================
# PpSlideShowState
# ==============================================================================
ppSlideShowRunning = 1
ppSlideShowPaused = 2
ppSlideShowBlackScreen = 3
ppSlideShowWhiteScreen = 4
ppSlideShowDone = 5

# ==============================================================================
# PpFixedFormatType
# ==============================================================================
ppFixedFormatTypePDF = 2
ppFixedFormatTypeXPS = 1

# ==============================================================================
# Shape type name mapping for human-readable output
# ==============================================================================
SHAPE_TYPE_NAMES = {
    1: "AutoShape",
    2: "Callout",
    3: "Chart",
    4: "Comment",
    5: "Freeform",
    6: "Group",
    7: "EmbeddedOLEObject",
    8: "FormControl",
    9: "Line",
    10: "LinkedOLEObject",
    11: "LinkedPicture",
    12: "OLEControlObject",
    13: "Picture",
    14: "Placeholder",
    15: "TextEffect",
    16: "Media",
    17: "TextBox",
    19: "Table",
    24: "SmartArt",
}

PLACEHOLDER_TYPE_NAMES = {
    1: "Title",
    2: "Body",
    3: "CenterTitle",
    4: "Subtitle",
    5: "VerticalTitle",
    6: "VerticalBody",
    7: "Object",
    8: "Chart",
    9: "Bitmap",
    10: "MediaClip",
    11: "OrgChart",
    12: "Table",
    13: "SlideNumber",
    14: "Header",
    15: "Footer",
    16: "Date",
    17: "VerticalObject",
    18: "Picture",
}

WINDOW_STATE_NAMES = {
    1: "normal",
    2: "minimized",
    3: "maximized",
}
