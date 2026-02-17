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
ppAutoSizeTextToFitShape = 2  # MsoAutoSize (TextFrame2 only) — shrink text on overflow
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
# PpSlideShowType
# ==============================================================================
ppShowTypeSpeaker = 1
ppShowTypeWindow = 2
ppShowTypeKiosk = 3

# ==============================================================================
# PpSlideShowAdvanceMode
# ==============================================================================
ppSlideShowManualAdvance = 1
ppSlideShowUseSlideTimings = 2

# ==============================================================================
# PpSlideShowRangeType
# ==============================================================================
ppShowAll = 1
ppShowSlideRange = 2
ppShowNamedSlideShow = 3

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

SLIDESHOW_STATE_NAMES = {
    1: "running",
    2: "paused",
    3: "black_screen",
    4: "white_screen",
    5: "done",
}

SHOW_TYPE_NAMES = {
    1: "speaker",
    2: "window",
    3: "kiosk",
}

# ==============================================================================
# MsoConnectorType
# ==============================================================================
msoConnectorStraight = 1
msoConnectorElbow = 2
msoConnectorCurve = 3

# ==============================================================================
# PpActionType
# ==============================================================================
ppActionNone = 0
ppActionNextSlide = 1
ppActionPreviousSlide = 2
ppActionFirstSlide = 3
ppActionLastSlide = 4
ppActionEndShow = 6
ppActionHyperlink = 7

# ==============================================================================
# PpMouseActivation
# ==============================================================================
ppMouseClick = 1
ppMouseOver = 2

# ==============================================================================
# XlChartType (commonly used)
# ==============================================================================
xlArea = 1
xlLine = 4
xlPie = 5
xlBubble = 15
xlColumnClustered = 51
xlColumnStacked = 52
xlColumnStacked100 = 53
xl3DColumnClustered = 54
xlBarClustered = 57
xlBarStacked = 58
xlLineStacked = 63
xlLineMarkers = 65
xlPieExploded = 69
xlXYScatterLines = 74
xlAreaStacked = 76
xlStockHLC = 88
xlXYScatter = -4169
xlDoughnut = -4120
xl3DPie = -4102
xl3DLine = -4101
xlRadar = -4151

# ==============================================================================
# XlAxisType
# ==============================================================================
xlCategory = 1
xlValue = 2
xlSeriesAxis = 3

# ==============================================================================
# XlLegendPosition
# ==============================================================================
xlLegendPositionBottom = -4107
xlLegendPositionLeft = -4131
xlLegendPositionRight = -4152
xlLegendPositionTop = -4160
xlLegendPositionCorner = 2

# ==============================================================================
# MsoAnimEffect (commonly used)
# ==============================================================================
msoAnimEffectAppear = 1
msoAnimEffectFly = 2
msoAnimEffectBlinds = 3
msoAnimEffectBox = 4
msoAnimEffectCheckerboard = 5
msoAnimEffectCircle = 6
msoAnimEffectDiamond = 8
msoAnimEffectDissolve = 9
msoAnimEffectFade = 10
msoAnimEffectSplit = 16
msoAnimEffectWipe = 22
msoAnimEffectZoom = 23
msoAnimEffectBounce = 26
msoAnimEffectFloat = 56
msoAnimEffectGrowAndTurn = 57
msoAnimEffectSpin = 61
msoAnimEffectTransparency = 62

# ==============================================================================
# MsoAnimTriggerType
# ==============================================================================
msoAnimTriggerNone = 0
msoAnimTriggerOnPageClick = 1
msoAnimTriggerWithPrevious = 2
msoAnimTriggerAfterPrevious = 3
msoAnimTriggerOnShapeClick = 4

# ==============================================================================
# PpEntryEffect (slide transitions)
# ==============================================================================
ppEffectNone = 0
ppEffectCut = 257
ppEffectFade = 3844
ppEffectPush = 3845
ppEffectWipe = 3846
ppEffectSplit = 3847
ppEffectReveal = 3848
ppEffectRandom = 513
ppEffectBlindsHorizontal = 769
ppEffectBlindsVertical = 770
ppEffectDissolve = 1537

# ==============================================================================
# PpTransitionSpeed
# ==============================================================================
ppTransitionSpeedFast = 1
ppTransitionSpeedMedium = 2
ppTransitionSpeedSlow = 3

# ==============================================================================
# PpMediaType
# ==============================================================================
ppMediaTypeMixed = -2
ppMediaTypeOther = 0
ppMediaTypeSound = 1
ppMediaTypeMovie = 3

# ==============================================================================
# PpDateTimeFormat (for HeadersFooters)
# ==============================================================================
ppDateTimeMdyy = 1
ppDateTimeddddMMMMddyyyy = 2
ppDateTimedMMMMyyyy = 3
ppDateTimeMMMMdyyyy = 4
ppDateTimedMMMyy = 5
ppDateTimeMMMMyy = 6
ppDateTimeMMyy = 7
ppDateTimeMddyy = 8
ppDateTimeHmm = 9
ppDateTimeHmmss = 10
ppDateTimehmmAMPM = 11
ppDateTimehmmssAMPM = 12

# ==============================================================================
# Name lookup maps for Phase 3
# ==============================================================================
CONNECTOR_TYPE_NAMES = {
    1: "straight",
    2: "elbow",
    3: "curve",
}

ANIMATION_EFFECT_NAMES = {
    1: "appear", 2: "fly", 3: "blinds", 4: "box",
    5: "checkerboard", 6: "circle", 8: "diamond",
    9: "dissolve", 10: "fade", 16: "split", 22: "wipe",
    23: "zoom", 26: "bounce", 56: "float",
    57: "grow_and_turn", 61: "spin", 62: "transparency",
}

ANIMATION_TRIGGER_NAMES = {
    0: "none", 1: "on_click", 2: "with_previous",
    3: "after_previous", 4: "on_shape_click",
}

# ==============================================================================
# Phase 4 Constants
# ==============================================================================

# MsoAlignCmd
msoAlignLefts = 0
msoAlignCenters = 1
msoAlignRights = 2
msoAlignTops = 3
msoAlignMiddles = 4
msoAlignBottoms = 5

# MsoDistributeCmd
msoDistributeHorizontally = 0
msoDistributeVertically = 1

# MsoFlipCmd
msoFlipHorizontal = 0
msoFlipVertical = 1

# MsoMergeCmd
msoMergeUnion = 1
msoMergeCombine = 2
msoMergeIntersect = 3
msoMergeSubtract = 4
msoMergeFragment = 5

# PpSlideSizeType
ppSlideSizeOnScreen = 1
ppSlideSizeLetterPaper = 2
ppSlideSizeA4Paper = 3
ppSlideSize35MM = 4
ppSlideSizeOverhead = 5
ppSlideSizeBanner = 6
ppSlideSizeCustom = 7
ppSlideSizeOnScreen16x9 = 8
ppSlideSizeOnScreen16x10 = 9
ppSlideSizeWidescreen = 10

# MsoOrientation
msoOrientationHorizontal = 1
msoOrientationVertical = 2
msoOrientationMixed = -2

# PpSelectionType
ppSelectionNone = 0
ppSelectionSlides = 1
ppSelectionShapes = 2
ppSelectionText = 3

# PpPasteDataType
ppPasteDefault = 0
ppPasteBitmap = 1
ppPasteEnhancedMetafile = 2
ppPasteMetafilePicture = 3
ppPasteGIF = 4
ppPasteJPG = 5
ppPastePNG = 6
ppPasteText = 7
ppPasteHTML = 8
ppPasteRTF = 9
ppPasteOLEObject = 10
ppPasteShape = 11

# PpShapeFormatType (for Shape.Export)
ppShapeFormatGIF = 0
ppShapeFormatJPG = 1
ppShapeFormatPNG = 2
ppShapeFormatBMP = 3
ppShapeFormatWMF = 4
ppShapeFormatEMF = 5

# MsoGradientStyle (for backgrounds)
msoGradientHorizontal = 1
msoGradientVertical = 2
msoGradientDiagonalUp = 3
msoGradientDiagonalDown = 4
msoGradientFromCorner = 5
msoGradientFromTitle = 6
msoGradientFromCenter = 7

# Friendly name maps for Phase 4
ALIGN_CMD_MAP = {
    "left": 0, "center": 1, "right": 2,
    "top": 3, "middle": 4, "bottom": 5,
}

DISTRIBUTE_CMD_MAP = {
    "horizontal": 0, "vertical": 1,
}

FLIP_CMD_MAP = {
    "horizontal": 0, "vertical": 1,
}

MERGE_CMD_MAP = {
    "union": 1, "combine": 2, "intersect": 3,
    "subtract": 4, "fragment": 5,
}

SLIDE_SIZE_MAP = {
    "4:3": 1, "letter": 2, "a4": 3, "35mm": 4,
    "overhead": 5, "banner": 6, "custom": 7,
    "a3": 8, "16:9": 9, "16:10": 10, "widescreen": 9,
}

GRADIENT_STYLE_MAP = {
    "horizontal": 1, "vertical": 2,
    "diagonal_up": 3, "diagonal_down": 4,
    "from_corner": 5, "from_title": 6, "from_center": 7,
}

SHAPE_FORMAT_MAP = {
    "gif": 0, "jpg": 1, "png": 2, "bmp": 3,
    "wmf": 4, "emf": 5,
}

VIEW_TYPE_MAP = {
    "normal": 1, "slide_master": 2, "notes_page": 3,
    "handout_master": 4, "notes_master": 5,
    "outline": 6, "slide_sorter": 7,
    "title_master": 8, "reading": 10,
}

VIEW_TYPE_NAMES = {v: k for k, v in VIEW_TYPE_MAP.items()}
