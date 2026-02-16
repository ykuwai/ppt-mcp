# Module: Core Infrastructure & Utilities

## Overview

This module provides the foundational infrastructure for the PowerPoint MCP server: COM connection management, unit conversion utilities, color helpers, PowerPoint/MSO constants, and the application-level connection wrapper. Every other module depends on this one. It handles COM initialization, connection lifecycle, error recovery, and provides reusable utility functions for the entire codebase.

## Dependencies

- **External packages**: `pywin32` (provides `win32com.client`, `pythoncom`, `pywintypes`)
- **Standard library**: `gc`, `logging`, `threading`
- No internal module dependencies (this is the base module).

## File Structure

```
ppt_com_mcp/
  utils/
    __init__.py
    com_wrapper.py      # COM connection management
    units.py            # Unit conversion functions
    color.py            # Color helper functions
  ppt_com/
    __init__.py
    constants.py        # All PowerPoint/MSO enumeration constants
    app.py              # Application-level connection and info
```

---

## File: `utils/com_wrapper.py` - COM Connection Management

### Purpose

Manage the COM lifecycle for the PowerPoint Application object. Handles CoInitialize, Dispatch, GetActiveObject, error recovery, reconnection, and cleanup. Since PowerPoint supports only a single instance, this module provides a singleton-like access pattern.

### Constants & Threading Model

```python
import pythoncom
import win32com.client
import pywintypes
import gc
import logging

logger = logging.getLogger(__name__)

# COM apartment model: PowerPoint requires STA (Single-Threaded Apartment)
# Any thread accessing COM must call CoInitializeEx first.
# COM objects CANNOT be passed between threads.
```

### Class: `PowerPointCOMWrapper`

```python
class PowerPointCOMWrapper:
    """
    Manages the lifecycle of a PowerPoint COM Application object.

    PowerPoint only supports a single running instance (unlike Word/Excel),
    so Dispatch("PowerPoint.Application") always returns the same instance
    if PowerPoint is already running.
    """

    def __init__(self):
        self._app = None
        self._initialized = False

    def initialize(self) -> None:
        """
        Initialize COM for the current thread.
        Must be called once per thread before any COM operations.
        Uses STA (Single-Threaded Apartment) model.
        """
        pythoncom.CoInitializeEx(pythoncom.COINIT_APARTMENTTHREADED)
        self._initialized = True

    def connect(self, visible: bool = True) -> "COMObject":
        """
        Connect to an existing PowerPoint instance or create a new one.

        Strategy:
        1. Try GetActiveObject to connect to running instance
        2. If that fails, use Dispatch to create/connect
        3. Set Visible property

        Returns:
            PowerPoint.Application COM object

        Raises:
            ConnectionError: If unable to connect to PowerPoint
        """
        if not self._initialized:
            self.initialize()

        try:
            # Try to connect to existing instance first
            self._app = win32com.client.GetActiveObject("PowerPoint.Application")
            logger.info("Connected to existing PowerPoint instance")
        except Exception:
            try:
                # Dispatch will connect to existing or create new
                self._app = win32com.client.Dispatch("PowerPoint.Application")
                logger.info("Created/connected PowerPoint instance via Dispatch")
            except pywintypes.com_error as e:
                raise ConnectionError(
                    f"Failed to connect to PowerPoint: {e.strerror}"
                ) from e

        self._app.Visible = visible  # msoTrue = -1 when True
        return self._app

    def get_app(self) -> "COMObject":
        """
        Get the current Application object, reconnecting if necessary.

        Returns:
            PowerPoint.Application COM object

        Raises:
            ConnectionError: If no connection exists and reconnect fails
        """
        if self._app is None:
            return self.connect()

        # Test if the connection is still alive
        try:
            _ = self._app.Name  # Simple test call
            return self._app
        except (pywintypes.com_error, AttributeError):
            logger.warning("COM connection lost, attempting reconnect...")
            self._app = None
            return self.connect()

    def ensure_presentation(self) -> "COMObject":
        """
        Ensure at least one presentation is open.

        Returns:
            The active Presentation object

        Raises:
            RuntimeError: If no presentation is open
        """
        app = self.get_app()
        if app.Presentations.Count == 0:
            raise RuntimeError("No presentation is open in PowerPoint")
        return app.ActivePresentation

    def cleanup(self) -> None:
        """
        Release COM references and uninitialize COM.
        Call this when shutting down the MCP server.

        IMPORTANT: Do NOT call app.Quit() unless explicitly requested.
        The user's PowerPoint should remain open.
        """
        if self._app is not None:
            self._app = None
        gc.collect()  # Force garbage collection to release COM refs

        if self._initialized:
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass  # May fail if already uninitialized
            self._initialized = False
```

### Function: `handle_com_error`

```python
def handle_com_error(e: pywintypes.com_error) -> dict:
    """
    Parse a COM error into a structured dict for MCP error responses.

    Args:
        e: The pywintypes.com_error exception

    Returns:
        dict with keys: hresult, message, source, description
    """
    result = {
        "hresult": e.hresult,
        "message": str(e.strerror) if e.strerror else "Unknown COM error",
        "source": None,
        "description": None,
    }
    if e.excepinfo:
        result["source"] = e.excepinfo[1] if len(e.excepinfo) > 1 else None
        result["description"] = e.excepinfo[2] if len(e.excepinfo) > 2 else None
    return result
```

### Function: `safe_com_call`

```python
def safe_com_call(func, *args, **kwargs):
    """
    Execute a COM call with error handling.

    Wraps any COM function call and catches pywintypes.com_error,
    converting it to a structured error response.

    Args:
        func: The callable to execute
        *args, **kwargs: Arguments to pass to func

    Returns:
        The result of func(*args, **kwargs)

    Raises:
        pywintypes.com_error: Re-raised with additional context
    """
    try:
        return func(*args, **kwargs)
    except pywintypes.com_error as e:
        error_info = handle_com_error(e)
        logger.error(
            f"COM Error: HRESULT={error_info['hresult']:#x}, "
            f"Message={error_info['message']}, "
            f"Description={error_info['description']}"
        )
        raise
```

### Error Cases

| Error | Cause | Recovery |
|-------|-------|----------|
| `pywintypes.com_error` with "PowerPoint is not running" | GetActiveObject fails | Fall back to Dispatch |
| `pywintypes.com_error` with RPC errors | PowerPoint crashed | Reconnect via connect() |
| `AttributeError` on COM object | Stale COM reference | Reconnect via get_app() |
| `pythoncom.com_error` during CoInitialize | Thread already initialized | Ignore, already initialized |

### Implementation Notes

- **Thread safety**: COM objects are thread-bound. The MCP server should route all COM calls through a single thread. Use a queue or asyncio bridge if the MCP framework uses async.
- **Single instance**: PowerPoint only supports one running instance. `Dispatch("PowerPoint.Application")` called multiple times returns the same instance.
- **DisplayAlerts**: Consider setting `app.DisplayAlerts = False` (ppAlertsNone = 0) to prevent dialog popups during automated operations. However, this setting is `ppAlertsAll = -1` by default and should be reset on cleanup.
- **Visible property**: Operations on shapes and text generally work with `Visible = True`. Some window-related operations fail when `Visible = False` or when presentations are opened with `WithWindow=False`.

---

## File: `utils/units.py` - Unit Conversion Functions

### Purpose

PowerPoint COM uses **points** as the native unit for all positions and sizes (Left, Top, Width, Height). This module provides bidirectional conversion between points, inches, centimeters, and EMUs (English Metric Units).

### Constants

```python
# PowerPoint's native unit is points
POINTS_PER_INCH = 72.0
CM_PER_INCH = 2.54
POINTS_PER_CM = POINTS_PER_INCH / CM_PER_INCH  # ~28.3465
EMU_PER_POINT = 12700
EMU_PER_INCH = 914400
EMU_PER_CM = 360000

# Standard slide sizes in points
SLIDE_WIDTH_16_9 = 960.0    # 13.333 inches
SLIDE_HEIGHT_16_9 = 540.0   # 7.5 inches
SLIDE_WIDTH_4_3 = 720.0     # 10 inches
SLIDE_HEIGHT_4_3 = 540.0    # 7.5 inches
SLIDE_WIDTH_A4_LANDSCAPE = 842.0   # ~11.693 inches
SLIDE_HEIGHT_A4_LANDSCAPE = 595.0  # ~8.268 inches
```

### Conversion Functions

```python
def inches_to_points(inches: float) -> float:
    """Convert inches to points. 1 inch = 72 points."""
    return inches * POINTS_PER_INCH

def points_to_inches(points: float) -> float:
    """Convert points to inches."""
    return points / POINTS_PER_INCH

def cm_to_points(cm: float) -> float:
    """Convert centimeters to points. 1 cm ~ 28.35 points."""
    return cm * POINTS_PER_CM

def points_to_cm(points: float) -> float:
    """Convert points to centimeters."""
    return points / POINTS_PER_CM

def emu_to_points(emu: int) -> float:
    """Convert EMU (English Metric Units) to points. 1 point = 12700 EMU."""
    return emu / EMU_PER_POINT

def points_to_emu(points: float) -> int:
    """Convert points to EMU."""
    return int(round(points * EMU_PER_POINT))

def inches_to_emu(inches: float) -> int:
    """Convert inches to EMU. 1 inch = 914400 EMU."""
    return int(round(inches * EMU_PER_INCH))

def emu_to_inches(emu: int) -> float:
    """Convert EMU to inches."""
    return emu / EMU_PER_INCH

def cm_to_emu(cm: float) -> int:
    """Convert centimeters to EMU. 1 cm = 360000 EMU."""
    return int(round(cm * EMU_PER_CM))

def emu_to_cm(emu: int) -> float:
    """Convert EMU to centimeters."""
    return emu / EMU_PER_CM

def inches_to_cm(inches: float) -> float:
    """Convert inches to centimeters."""
    return inches * CM_PER_INCH

def cm_to_inches(cm: float) -> float:
    """Convert centimeters to inches."""
    return cm / CM_PER_INCH
```

---

## File: `utils/color.py` - Color Helper Functions

### Purpose

PowerPoint COM uses BGR-ordered integers for RGB color values. The formula is `R + (G * 256) + (B * 65536)`. This is the opposite of the typical `0xRRGGBB` hex notation. This module provides conversions between standard RGB, hex strings, and the PowerPoint BGR integer format.

### Functions

```python
def rgb_to_int(r: int, g: int, b: int) -> int:
    """
    Convert RGB values (0-255 each) to PowerPoint's BGR integer format.

    PowerPoint COM RGB = R + (G << 8) + (B << 16)
    This is equivalent to VBA's RGB(r, g, b) function.

    IMPORTANT: 0x0000FF in this format means RED (R=255, G=0, B=0),
    NOT blue. This is the opposite of standard HTML hex colors.

    Args:
        r: Red component (0-255)
        g: Green component (0-255)
        b: Blue component (0-255)

    Returns:
        Integer color value for PowerPoint COM

    Examples:
        rgb_to_int(255, 0, 0)   -> 255       (red)
        rgb_to_int(0, 255, 0)   -> 65280     (green)
        rgb_to_int(0, 0, 255)   -> 16711680  (blue)
        rgb_to_int(255, 255, 255) -> 16777215 (white)
    """
    return r + (g << 8) + (b << 16)

def int_to_rgb(color_int: int) -> tuple[int, int, int]:
    """
    Convert PowerPoint's BGR integer to an (R, G, B) tuple.

    Args:
        color_int: PowerPoint COM color integer

    Returns:
        Tuple of (red, green, blue) each 0-255
    """
    r = color_int & 0xFF
    g = (color_int >> 8) & 0xFF
    b = (color_int >> 16) & 0xFF
    return (r, g, b)

def hex_to_rgb(hex_str: str) -> tuple[int, int, int]:
    """
    Convert a hex color string to (R, G, B) tuple.

    Accepts formats: "#RRGGBB", "RRGGBB", "#RGB"

    Args:
        hex_str: Hex color string

    Returns:
        Tuple of (red, green, blue) each 0-255
    """
    hex_str = hex_str.lstrip("#")
    if len(hex_str) == 3:
        hex_str = "".join(c * 2 for c in hex_str)
    if len(hex_str) != 6:
        raise ValueError(f"Invalid hex color: {hex_str}")
    r = int(hex_str[0:2], 16)
    g = int(hex_str[2:4], 16)
    b = int(hex_str[4:6], 16)
    return (r, g, b)

def hex_to_int(hex_str: str) -> int:
    """
    Convert a hex color string directly to PowerPoint's BGR integer.

    Args:
        hex_str: Hex color string like "#FF0000" (red)

    Returns:
        Integer color value for PowerPoint COM

    Example:
        hex_to_int("#FF0000") -> 255 (red in BGR format)
    """
    r, g, b = hex_to_rgb(hex_str)
    return rgb_to_int(r, g, b)

def int_to_hex(color_int: int) -> str:
    """
    Convert PowerPoint's BGR integer to a "#RRGGBB" hex string.

    Args:
        color_int: PowerPoint COM color integer

    Returns:
        Hex color string like "#FF0000"
    """
    r, g, b = int_to_rgb(color_int)
    return f"#{r:02X}{g:02X}{b:02X}"

# Theme color name-to-constant mapping
THEME_COLOR_MAP = {
    "dark1": 1,             # msoThemeColorDark1
    "light1": 2,            # msoThemeColorLight1
    "dark2": 3,             # msoThemeColorDark2
    "light2": 4,            # msoThemeColorLight2
    "accent1": 5,           # msoThemeColorAccent1
    "accent2": 6,           # msoThemeColorAccent2
    "accent3": 7,           # msoThemeColorAccent3
    "accent4": 8,           # msoThemeColorAccent4
    "accent5": 9,           # msoThemeColorAccent5
    "accent6": 10,          # msoThemeColorAccent6
    "hyperlink": 11,        # msoThemeColorHyperlink
    "followed_hyperlink": 12,  # msoThemeColorFollowedHyperlink
}

def get_theme_color_index(name: str) -> int:
    """
    Convert a theme color name to its MsoThemeColorIndex constant value.

    Args:
        name: Theme color name (case-insensitive), e.g. "accent1", "dark1"

    Returns:
        Integer constant value for ObjectThemeColor property

    Raises:
        ValueError: If name is not a valid theme color
    """
    key = name.lower().replace(" ", "_").replace("-", "_")
    if key not in THEME_COLOR_MAP:
        raise ValueError(
            f"Unknown theme color: '{name}'. "
            f"Valid names: {list(THEME_COLOR_MAP.keys())}"
        )
    return THEME_COLOR_MAP[key]
```

---

## File: `ppt_com/constants.py` - PowerPoint/MSO Constants

### Purpose

Central repository of all PowerPoint and MSO enumeration constants used across the entire MCP server. Each constant is defined with its integer value and grouped by enumeration. Agents for other modules import constants from here.

### Complete Constants Definition

```python
"""
PowerPoint and MSO COM Automation Constants.

All constants are defined as module-level integers.
Grouped by enumeration type for readability.
"""

# ============================================================
# MsoTriState
# ============================================================
msoTrue = -1
msoFalse = 0
msoCTrue = 1
msoTriStateToggle = -3
msoTriStateMixed = -2

# ============================================================
# PpWindowState
# ============================================================
ppWindowNormal = 1
ppWindowMinimized = 2
ppWindowMaximized = 3

# ============================================================
# PpViewType
# ============================================================
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

# ============================================================
# PpSlideLayout
# ============================================================
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

# ============================================================
# PpSaveAsFileType
# ============================================================
ppSaveAsPresentation = 1
ppSaveAsTemplate = 5
ppSaveAsRTF = 6
ppSaveAsShow = 7
ppSaveAsAddIn = 8
ppSaveAsDefault = 11
ppSaveAsMetaFile = 15
ppSaveAsGIF = 16
ppSaveAsJPG = 17
ppSaveAsPNG = 18
ppSaveAsBMP = 19
ppSaveAsTIF = 21
ppSaveAsEMF = 23
ppSaveAsOpenXMLPresentation = 24
ppSaveAsOpenXMLPresentationMacroEnabled = 25
ppSaveAsOpenXMLTemplate = 26
ppSaveAsOpenXMLTemplateMacroEnabled = 27
ppSaveAsOpenXMLShow = 28
ppSaveAsOpenXMLShowMacroEnabled = 29
ppSaveAsOpenXMLAddin = 30
ppSaveAsOpenXMLTheme = 31
ppSaveAsPDF = 32
ppSaveAsXPS = 33
ppSaveAsXMLPresentation = 34
ppSaveAsOpenDocumentPresentation = 35
ppSaveAsOpenXMLPicturePresentation = 36
ppSaveAsWMV = 37
ppSaveAsStrictOpenXMLPresentation = 38
ppSaveAsMP4 = 39
ppSaveAsAnimatedGIF = 40

# ============================================================
# MsoShapeType
# ============================================================
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
msoScriptAnchor = 18
msoTable = 19
msoCanvas = 20
msoDiagram = 21
msoInk = 22
msoInkComment = 23
msoSmartArt = 24
msoSlicer = 25
msoWebVideo = 26
msoContentApp = 27
msoGraphic = 28
msoLinkedGraphic = 29
mso3DModel = 30
msoShapeTypeMixed = -2

# ============================================================
# MsoAutoShapeType (commonly used shapes)
# ============================================================
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
msoShapeFunnel = 174
msoShapeCloud = 179

# ============================================================
# PpPlaceholderType
# ============================================================
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

# ============================================================
# PpParagraphAlignment
# ============================================================
ppAlignLeft = 1
ppAlignCenter = 2
ppAlignRight = 3
ppAlignJustify = 4
ppAlignDistribute = 5
ppAlignmentMixed = -2

# ============================================================
# MsoTextOrientation
# ============================================================
msoTextOrientationHorizontal = 1
msoTextOrientationUpward = 2
msoTextOrientationDownward = 3
msoTextOrientationVertical = 5
msoTextOrientationVerticalFarEast = 6

# ============================================================
# PpAutoSize
# ============================================================
ppAutoSizeNone = 0
ppAutoSizeShapeToFitText = 1
ppAutoSizeMixed = -2

# ============================================================
# MsoZOrderCmd
# ============================================================
msoBringToFront = 0
msoSendToBack = 1
msoBringForward = 2
msoSendBackward = 3

# ============================================================
# MsoFlipCmd
# ============================================================
msoFlipHorizontal = 0
msoFlipVertical = 1

# ============================================================
# MsoConnectorType
# ============================================================
msoConnectorStraight = 1
msoConnectorElbow = 2
msoConnectorCurve = 3

# ============================================================
# MsoFillType
# ============================================================
msoFillSolid = 1
msoFillPatterned = 2
msoFillGradient = 3
msoFillTextured = 4
msoFillBackground = 5
msoFillPicture = 6
msoFillMixed = -2

# ============================================================
# MsoGradientStyle
# ============================================================
msoGradientHorizontal = 1
msoGradientVertical = 2
msoGradientDiagonalUp = 3
msoGradientDiagonalDown = 4
msoGradientFromCorner = 5
msoGradientFromTitle = 6
msoGradientFromCenter = 7

# ============================================================
# MsoLineDashStyle
# ============================================================
msoLineSolid = 1
msoLineRoundDot = 2
msoLineDot = 3
msoLineDash = 4
msoLineDashDot = 5
msoLineDashDotDot = 6
msoLineLongDash = 7
msoLineLongDashDot = 8

# ============================================================
# MsoLineStyle
# ============================================================
msoLineSingle = 1
msoLineThinThin = 2
msoLineThinThick = 3
msoLineThickThin = 4
msoLineThickBetweenThin = 5

# ============================================================
# MsoArrowheadStyle
# ============================================================
msoArrowheadNone = 1
msoArrowheadTriangle = 2
msoArrowheadOpen = 3
msoArrowheadStealth = 4
msoArrowheadDiamond = 5
msoArrowheadOval = 6

# ============================================================
# MsoArrowheadLength / Width
# ============================================================
msoArrowheadShort = 1
msoArrowheadLengthMedium = 2
msoArrowheadLong = 3
msoArrowheadNarrow = 1
msoArrowheadWidthMedium = 2
msoArrowheadWide = 3

# ============================================================
# MsoThemeColorIndex
# ============================================================
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

# ============================================================
# PpBulletType
# ============================================================
ppBulletNone = 0
ppBulletUnnumbered = 1
ppBulletNumbered = 2
ppBulletPicture = 3
ppBulletMixed = -2

# ============================================================
# PpNumberedBulletStyle
# ============================================================
ppBulletArabicParenRight = 2
ppBulletArabicPeriod = 3
ppBulletArabicParenBoth = 12
ppBulletRomanUCPeriod = 4
ppBulletRomanLCPeriod = 5
ppBulletAlphaUCPeriod = 6
ppBulletAlphaLCPeriod = 7

# ============================================================
# PpSlideShowState
# ============================================================
ppSlideShowRunning = 1
ppSlideShowPaused = 2
ppSlideShowBlackScreen = 3
ppSlideShowWhiteScreen = 4
ppSlideShowDone = 5

# ============================================================
# PpSlideShowType
# ============================================================
ppShowTypeSpeaker = 1
ppShowTypeWindow = 2
ppShowTypeKiosk = 3

# ============================================================
# PpSlideShowAdvanceMode
# ============================================================
ppSlideShowManualAdvance = 1
ppSlideShowUseSlideTimings = 2
ppSlideShowRehearseNewTimings = 3

# ============================================================
# PpSlideShowRangeType
# ============================================================
ppShowAll = 1
ppShowSlideRange = 2
ppShowNamedSlideShow = 3

# ============================================================
# PpSlideShowPointerType
# ============================================================
ppSlideShowPointerNone = 0
ppSlideShowPointerArrow = 1
ppSlideShowPointerPen = 2
ppSlideShowPointerAlwaysHidden = 3
ppSlideShowPointerAutoArrow = 4
ppSlideShowPointerEraser = 5

# ============================================================
# PpTransitionSpeed
# ============================================================
ppTransitionSpeedFast = 1
ppTransitionSpeedMedium = 2
ppTransitionSpeedSlow = 3
ppTransitionSpeedMixed = -2

# ============================================================
# PpFixedFormatType
# ============================================================
ppFixedFormatTypeXPS = 1
ppFixedFormatTypePDF = 2

# ============================================================
# PpFixedFormatIntent
# ============================================================
ppFixedFormatIntentScreen = 1
ppFixedFormatIntentPrint = 2

# ============================================================
# PpSelectionType
# ============================================================
ppSelectionNone = 0
ppSelectionSlides = 1
ppSelectionShapes = 2
ppSelectionText = 3

# ============================================================
# PpActionType
# ============================================================
ppActionNone = 0
ppActionNextSlide = 1
ppActionPreviousSlide = 2
ppActionFirstSlide = 3
ppActionLastSlide = 4
ppActionLastSlideViewed = 5
ppActionEndShow = 6
ppActionHyperlink = 7
ppActionRunMacro = 8
ppActionRunProgram = 9
ppActionNamedSlideShow = 10
ppActionOLEVerb = 11

# ============================================================
# PpMouseActivation
# ============================================================
ppMouseClick = 1
ppMouseOver = 2

# ============================================================
# MsoSegmentType / MsoEditingType (for freeform)
# ============================================================
msoSegmentLine = 0
msoSegmentCurve = 1
msoEditingAuto = 0
msoEditingCorner = 1
msoEditingSmooth = 2
msoEditingSymmetric = 3

# ============================================================
# MsoCalloutType
# ============================================================
msoCalloutOne = 1
msoCalloutTwo = 2
msoCalloutThree = 3
msoCalloutFour = 4

# ============================================================
# Table Border constants
# ============================================================
ppBorderTop = 1
ppBorderLeft = 2
ppBorderBottom = 3
ppBorderRight = 4
ppBorderDiagonalDown = 5
ppBorderDiagonalUp = 6

# ============================================================
# XlChartType (commonly used)
# ============================================================
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
xlDoughnut = -4120
xl3DLine = -4101
xl3DPie = -4102
xlRadar = -4151
xlXYScatter = -4169

# ============================================================
# XlLegendPosition
# ============================================================
xlLegendPositionBottom = -4107
xlLegendPositionLeft = -4131
xlLegendPositionRight = -4152
xlLegendPositionTop = -4160
xlLegendPositionCorner = 2

# ============================================================
# XlAxisType
# ============================================================
xlCategory = 1
xlValue = 2
xlSeriesAxis = 3

# ============================================================
# MsoAnimEffect (commonly used)
# ============================================================
msoAnimEffectAppear = 1
msoAnimEffectFly = 2
msoAnimEffectBlinds = 3
msoAnimEffectBox = 4
msoAnimEffectCheckerboard = 5
msoAnimEffectDiamond = 8
msoAnimEffectDissolve = 9
msoAnimEffectFade = 10
msoAnimEffectFlashOnce = 11
msoAnimEffectWipe = 22
msoAnimEffectZoom = 23
msoAnimEffectSpin = 61

# ============================================================
# MsoAnimTriggerType
# ============================================================
msoAnimTriggerNone = 0
msoAnimTriggerOnPageClick = 1
msoAnimTriggerWithPrevious = 2
msoAnimTriggerAfterPrevious = 3
msoAnimTriggerOnShapeClick = 4

# ============================================================
# PpEntryEffect (commonly used transition effects)
# ============================================================
ppEffectNone = 0
ppEffectCut = 257
ppEffectRandom = 513
ppEffectBlindsHorizontal = 769
ppEffectBlindsVertical = 770
ppEffectCheckerboardAcross = 1025
ppEffectCoverRight = 1281
ppEffectCoverDown = 1284
ppEffectDissolve = 1537
ppEffectStripsDownLeft = 2305
ppEffectFade = 3844
ppEffectPush = 3845
ppEffectWipe = 3846
ppEffectSplit = 3847
ppEffectReveal = 3848

# ============================================================
# PpDateTimeFormat
# ============================================================
ppDateTimeMdyy = 1
ppDateTimeddddMMMMddyyyy = 2
ppDateTimedMMMMyyyy = 3
ppDateTimeMMMMdyyyy = 4
ppDateTimedMMMyy = 5
ppDateTimeMMMMyy = 6
ppDateTimeMMyy = 7
ppDateTimeMMddyyHmm = 8
ppDateTimeMMddyyhmmAMPM = 9
ppDateTimeHmm = 10
ppDateTimeHmmss = 11
ppDateTimehmmAMPM = 12
ppDateTimehmmssAMPM = 13
ppDateTimeFigureOut = 14

# ============================================================
# MsoShadowStyle
# ============================================================
msoShadowStyleInnerShadow = 1
msoShadowStyleOuterShadow = 2

# ============================================================
# MsoSoftEdgeType
# ============================================================
msoSoftEdgeTypeNone = 0
msoSoftEdgeType1 = 1
msoSoftEdgeType2 = 2
msoSoftEdgeType3 = 3
msoSoftEdgeType4 = 4
msoSoftEdgeType5 = 5
msoSoftEdgeType6 = 6

# ============================================================
# MsoReflectionType
# ============================================================
msoReflectionTypeNone = 0
msoReflectionType1 = 1
msoReflectionType2 = 2
msoReflectionType3 = 3
msoReflectionType4 = 4
msoReflectionType5 = 5
msoReflectionType6 = 6
msoReflectionType7 = 7
msoReflectionType8 = 8
msoReflectionType9 = 9

# ============================================================
# MsoBevelType (commonly used)
# ============================================================
msoBevelNone = 1
msoBevelRelaxedInset = 2
msoBevelCircle = 3
msoBevelCross = 4
msoBevelSlope = 5
msoBevelAngle = 6
msoBevelSoftRound = 7
msoBevelConvex = 8
msoBevelCoolSlant = 9
msoBevelDivot = 10
msoBevelRiblet = 11
msoBevelHardEdge = 12
msoBevelArtDeco = 13

# ============================================================
# MsoPresetMaterial (commonly used)
# ============================================================
msoMaterialMatte = 1
msoMaterialPlastic = 2
msoMaterialMetal = 3
msoMaterialWireFrame = 4
msoMaterialMatte2 = 5
msoMaterialMetal2 = 6
msoMaterialDarkEdge = 7
msoMaterialSoftEdge = 8
msoMaterialFlat = 9

# ============================================================
# PpMediaType
# ============================================================
ppMediaTypeMixed = -2
ppMediaTypeOther = 1
ppMediaTypeSound = 2
ppMediaTypeMovie = 3
```

---

## File: `ppt_com/app.py` - Application Connection & Info

### Purpose

High-level wrapper around the PowerPoint Application object. Provides MCP tool implementations for connecting to PowerPoint, getting application info, and managing windows.

### MCP Tools

### Tool: `get_app_info`
- **Description**: Get information about the connected PowerPoint application
- **Parameters**: None
- **Returns**:
  ```json
  {
    "name": "Microsoft PowerPoint",
    "version": "16.0",
    "path": "C:\\Program Files\\...",
    "visible": true,
    "window_state": "maximized",
    "presentations_count": 2,
    "active_presentation": "MyFile.pptx"
  }
  ```
- **COM Implementation**:
  ```python
  def get_app_info(app):
      info = {
          "name": app.Name,
          "version": app.Version,
          "path": app.Path,
          "visible": bool(app.Visible),
          "window_state": {1: "normal", 2: "minimized", 3: "maximized"}.get(
              app.WindowState, "unknown"
          ),
          "presentations_count": app.Presentations.Count,
      }
      if app.Presentations.Count > 0:
          info["active_presentation"] = app.ActivePresentation.Name
      else:
          info["active_presentation"] = None
      return info
  ```
- **Error Cases**: COM connection lost (reconnect), no window available

### Tool: `list_presentations`
- **Description**: List all open presentations
- **Parameters**: None
- **Returns**:
  ```json
  {
    "presentations": [
      {
        "index": 1,
        "name": "MyPresentation.pptx",
        "full_name": "C:\\path\\to\\MyPresentation.pptx",
        "path": "C:\\path\\to",
        "slides_count": 15,
        "read_only": false,
        "saved": true
      }
    ]
  }
  ```
- **COM Implementation**:
  ```python
  def list_presentations(app):
      result = []
      for i in range(1, app.Presentations.Count + 1):
          pres = app.Presentations(i)
          result.append({
              "index": i,
              "name": pres.Name,
              "full_name": pres.FullName,
              "path": pres.Path,
              "slides_count": pres.Slides.Count,
              "read_only": bool(pres.ReadOnly),
              "saved": bool(pres.Saved),
          })
      return {"presentations": result}
  ```

### Tool: `activate_presentation`
- **Description**: Activate (bring to front) a specific presentation window
- **Parameters**:
  - `index` (int, optional): Presentation index (1-based)
  - `name` (str, optional): Presentation filename
- **Returns**: `{"success": true, "active_presentation": "name.pptx"}`
- **COM Implementation**:
  ```python
  def activate_presentation(app, index=None, name=None):
      if index is not None:
          pres = app.Presentations(index)
      elif name is not None:
          pres = None
          for i in range(1, app.Presentations.Count + 1):
              if app.Presentations(i).Name == name:
                  pres = app.Presentations(i)
                  break
          if pres is None:
              raise ValueError(f"Presentation not found: {name}")
      else:
          raise ValueError("Must specify index or name")

      # Find and activate the window for this presentation
      for i in range(1, app.Windows.Count + 1):
          if app.Windows(i).Presentation.Name == pres.Name:
              app.Windows(i).Activate()
              break
      return {"success": True, "active_presentation": pres.Name}
  ```

## Implementation Notes

1. **COM Threading**: All COM calls must happen on the same thread that called `CoInitializeEx`. If the MCP server is async, use a dedicated COM thread with a synchronous queue.
2. **Error Recovery**: Always wrap COM calls with try/except for `pywintypes.com_error`. On RPC_E_DISCONNECTED or similar errors, clear the cached app reference and reconnect.
3. **Color Format**: Always remember PowerPoint's BGR integer format. The `utils/color.py` module must be used consistently. Never pass standard 0xRRGGBB hex values directly to `.RGB` properties.
4. **Unit System**: All position/size values in PowerPoint COM are in **points** (72 points = 1 inch). Use `utils/units.py` for conversions.
5. **1-Based Indexing**: All PowerPoint COM collections use 1-based indexing. `Slides(1)` is the first slide, `Shapes(1)` is the first shape.
6. **MsoTriState**: Boolean properties in COM use MsoTriState where `True` = -1 (msoTrue) and `False` = 0 (msoFalse). Python's `True`/`False` work in most cases because win32com handles the conversion, but explicitly use `-1`/`0` when setting values to be safe.
