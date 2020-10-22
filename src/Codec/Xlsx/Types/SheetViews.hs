{-# LANGUAGE CPP #-}
{-# LANGUAGE OverloadedStrings #-}
{-# LANGUAGE RecordWildCards   #-}
{-# LANGUAGE TemplateHaskell   #-}
{-# LANGUAGE DeriveGeneric #-}
module Codec.Xlsx.Types.SheetViews (
    -- * Structured type to construct 'SheetViews'
    SheetView(..)
  , Selection(..)
  , Pane(..)
  , SheetViewType(..)
  , PaneType(..)
  , PaneState(..)
    -- * Lenses
    -- ** SheetView
  , sheetViewColorId
  , sheetViewDefaultGridColor
  , sheetViewRightToLeft
  , sheetViewShowFormulas
  , sheetViewShowGridLines
  , sheetViewShowOutlineSymbols
  , sheetViewShowRowColHeaders
  , sheetViewShowRuler
  , sheetViewShowWhiteSpace
  , sheetViewShowZeros
  , sheetViewTabSelected
  , sheetViewTopLeftCell
  , sheetViewType
  , sheetViewWindowProtection
  , sheetViewWorkbookViewId
  , sheetViewZoomScale
  , sheetViewZoomScaleNormal
  , sheetViewZoomScalePageLayoutView
  , sheetViewZoomScaleSheetLayoutView
  , sheetViewPane
  , sheetViewSelection
    -- ** Selection
  , selectionActiveCell
  , selectionActiveCellId
  , selectionPane
  , selectionSqref
    -- ** Pane
  , paneActivePane
  , paneState
  , paneTopLeftCell
  , paneXSplit
  , paneYSplit
  ) where

import GHC.Generics (Generic)

#ifdef USE_MICROLENS
import Lens.Micro.TH (makeLenses)
#else
import Control.Lens (makeLenses)
#endif
import Control.DeepSeq (NFData)
import Data.Default
import Data.Maybe (catMaybes, maybeToList, listToMaybe)
import Text.XML
import Text.XML.Cursor
import qualified Data.Map  as Map

import Codec.Xlsx.Types.Common
import Codec.Xlsx.Parser.Internal
import Codec.Xlsx.Writer.Internal

{-------------------------------------------------------------------------------
  Main types
-------------------------------------------------------------------------------}

-- | Worksheet view
--
-- A single sheet view definition. When more than one sheet view is defined in
-- the file, it means that when opening the workbook, each sheet view
-- corresponds to a separate window within the spreadsheet application, where
-- each window is showing the particular sheet containing the same
-- workbookViewId value, the last sheetView definition is loaded, and the others
-- are discarded. When multiple windows are viewing the same sheet, multiple
-- sheetView elements (with corresponding workbookView entries) are saved.
--
-- TODO: The @pivotSelection@ and @extLst@ child elements are unsupported.
--
-- See Section 18.3.1.87 "sheetView (Worksheet View)" (p. 1880)
data SheetView = SheetView {
    -- | Index to the color value for row/column text headings and gridlines.
    -- This is an 'index color value' (ICV) rather than rgb value.
    _sheetViewColorId :: Maybe Int

    -- | Flag indicating that the consuming application should use the default
    -- grid lines color (system dependent). Overrides any color specified in
    -- colorId.
  , _sheetViewDefaultGridColor :: Maybe Bool

    -- | Flag indicating whether the sheet is in 'right to left' display mode.
    -- When in this mode, Column A is on the far right, Column B ;is one column
    -- left of Column A, and so on. Also, information in cells is displayed in
    -- the Right to Left format.
  , _sheetViewRightToLeft :: Maybe Bool

    -- | Flag indicating whether this sheet should display formulas.
  , _sheetViewShowFormulas :: Maybe Bool

    -- | Flag indicating whether this sheet should display gridlines.
  , _sheetViewShowGridLines :: Maybe Bool

    -- | Flag indicating whether the sheet has outline symbols visible. This
    -- flag shall always override SheetPr element's outlinePr child element
    -- whose attribute is named showOutlineSymbols when there is a conflict.
  , _sheetViewShowOutlineSymbols :: Maybe Bool

    -- | Flag indicating whether the sheet should display row and column headings.
  , _sheetViewShowRowColHeaders :: Maybe Bool

    -- | Show the ruler in Page Layout View.
  , _sheetViewShowRuler :: Maybe Bool

    -- | Flag indicating whether page layout view shall display margins. False
    -- means do not display left, right, top (header), and bottom (footer)
    -- margins (even when there is data in the header or footer).
  , _sheetViewShowWhiteSpace :: Maybe Bool

    -- | Flag indicating whether the window should show 0 (zero) in cells
    -- containing zero value. When false, cells with zero value appear blank
    -- instead of showing the number zero.
  , _sheetViewShowZeros :: Maybe Bool

    -- | Flag indicating whether this sheet is selected. When only 1 sheet is
    -- selected and active, this value should be in synch with the activeTab
    -- value. In case of a conflict, the Start Part setting wins and sets the
    -- active sheet tab.
    --
    -- Multiple sheets can be selected, but only one sheet shall be active at
    -- one time.
  , _sheetViewTabSelected :: Maybe Bool

    -- | Location of the top left visible cell Location of the top left visible
    -- cell in the bottom right pane (when in Left-to-Right mode).
  , _sheetViewTopLeftCell :: Maybe CellRef

    -- | Indicates the view type.
  , _sheetViewType :: Maybe SheetViewType

    -- | Flag indicating whether the panes in the window are locked due to
    -- workbook protection. This is an option when the workbook structure is
    -- protected.
  , _sheetViewWindowProtection :: Maybe Bool

    -- | Zero-based index of this workbook view, pointing to a workbookView
    -- element in the bookViews collection.
    --
    -- NOTE: This attribute is required.
  , _sheetViewWorkbookViewId :: Int

    -- | Window zoom magnification for current view representing percent values.
    -- This attribute is restricted to values ranging from 10 to 400. Horizontal &
    -- Vertical scale together.
  , _sheetViewZoomScale :: Maybe Int

    -- | Zoom magnification to use when in normal view, representing percent
    -- values. This attribute is restricted to values ranging from 10 to 400.
    -- Horizontal & Vertical scale together.
  , _sheetViewZoomScaleNormal :: Maybe Int

    -- | Zoom magnification to use when in page layout view, representing
    -- percent values. This attribute is restricted to values ranging from 10 to
    -- 400. Horizontal & Vertical scale together.
  , _sheetViewZoomScalePageLayoutView :: Maybe Int

    -- | Zoom magnification to use when in page break preview, representing
    -- percent values. This attribute is restricted to values ranging from 10 to
    -- 400. Horizontal & Vertical scale together.
  , _sheetViewZoomScaleSheetLayoutView :: Maybe Int

    -- | Worksheet view pane
  , _sheetViewPane :: Maybe Pane

    -- | Worksheet view selection
    --
    -- Minimum of 0, maximum of 4 elements
  , _sheetViewSelection :: [Selection]
  }
  deriving (Eq, Ord, Show, Generic)
instance NFData SheetView

-- | Worksheet view selection.
--
-- Section 18.3.1.78 "selection (Selection)" (p. 1864)
data Selection = Selection {
    -- | Location of the active cell
    _selectionActiveCell :: Maybe CellRef

    -- | 0-based index of the range reference (in the array of references listed
    -- in sqref) containing the active cell. Only used when the selection in
    -- sqref is not contiguous. Therefore, this value needs to be aware of the
    -- order in which the range references are written in sqref.
    --
    -- When this value is out of range then activeCell can be used.
  , _selectionActiveCellId :: Maybe Int

    -- | The pane to which this selection belongs.
  , _selectionPane :: Maybe PaneType

    -- | Range of the selection. Can be non-contiguous set of ranges.
  , _selectionSqref :: Maybe SqRef
  }
  deriving (Eq, Ord, Show, Generic)
instance NFData Selection

-- | Worksheet view pane
--
-- Section 18.3.1.66 "pane (View Pane)" (p. 1843)
data Pane = Pane {
    -- | The pane that is active.
    _paneActivePane :: Maybe PaneType

    -- | Indicates whether the pane has horizontal / vertical splits, and
    -- whether those splits are frozen.
  , _paneState :: Maybe PaneState

    -- | Location of the top left visible cell in the bottom right pane (when in
    -- Left-To-Right mode).
  , _paneTopLeftCell :: Maybe CellRef

    -- | Horizontal position of the split, in 1/20th of a point; 0 (zero) if
    -- none. If the pane is frozen, this value indicates the number of columns
    -- visible in the top pane.
  , _paneXSplit :: Maybe Double

    -- | Vertical position of the split, in 1/20th of a point; 0 (zero) if none.
    -- If the pane is frozen, this value indicates the number of rows visible in
    -- the left pane.
  , _paneYSplit :: Maybe Double
  }
  deriving (Eq, Ord, Show, Generic)
instance NFData Pane

{-------------------------------------------------------------------------------
  Enumerations
-------------------------------------------------------------------------------}

-- | View setting of the sheet
--
-- Section 18.18.69 "ST_SheetViewType (Sheet View Type)" (p. 2726)
data SheetViewType =
    -- | Normal view
    SheetViewTypeNormal

    -- | Page break preview
  | SheetViewTypePageBreakPreview

    -- | Page layout view
  | SheetViewTypePageLayout
  deriving (Eq, Ord, Show, Generic)
instance NFData SheetViewType

-- | Pane type
--
-- Section 18.18.52 "ST_Pane (Pane Types)" (p. 2710)
data PaneType =
    -- | Bottom left pane, when both vertical and horizontal splits are applied.
    --
    -- This value is also used when only a horizontal split has been applied,
    -- dividing the pane into upper and lower regions. In that case, this value
    -- specifies the bottom pane.
    PaneTypeBottomLeft

    -- Bottom right pane, when both vertical and horizontal splits are applied.
  | PaneTypeBottomRight

    -- | Top left pane, when both vertical and horizontal splits are applied.
    --
    -- This value is also used when only a horizontal split has been applied,
    -- dividing the pane into upper and lower regions. In that case, this value
    -- specifies the top pane.
    --
    -- This value is also used when only a vertical split has been applied,
    -- dividing the pane into right and left regions. In that case, this value
    -- specifies the left pane
  | PaneTypeTopLeft

    -- | Top right pane, when both vertical and horizontal splits are applied.
    --
    -- This value is also used when only a vertical split has been applied,
    -- dividing the pane into right and left regions. In that case, this value
    -- specifies the right pane.
  | PaneTypeTopRight
  deriving (Eq, Ord, Show, Generic)
instance NFData PaneType

-- | State of the sheet's pane.
--
-- Section 18.18.53 "ST_PaneState (Pane State)" (p. 2711)
data PaneState =
    -- | Panes are frozen, but were not split being frozen. In this state, when
    -- the panes are unfrozen again, a single pane results, with no split. In
    -- this state, the split bars are not adjustable.
    PaneStateFrozen

    -- | Panes are frozen and were split before being frozen. In this state,
    -- when the panes are unfrozen again, the split remains, but is adjustable.
  | PaneStateFrozenSplit

    -- | Panes are split, but not frozen. In this state, the split bars are
    -- adjustable by the user.
  | PaneStateSplit
  deriving (Eq, Ord, Show, Generic)
instance NFData PaneState

{-------------------------------------------------------------------------------
  Lenses
-------------------------------------------------------------------------------}

makeLenses ''SheetView
makeLenses ''Selection
makeLenses ''Pane

{-------------------------------------------------------------------------------
  Default instances
-------------------------------------------------------------------------------}

-- | NOTE: The 'Default' instance for 'SheetView' sets the required attribute
-- '_sheetViewWorkbookViewId' to @0@.
instance Default SheetView where
  def = SheetView {
      _sheetViewColorId                  = Nothing
    , _sheetViewDefaultGridColor         = Nothing
    , _sheetViewRightToLeft              = Nothing
    , _sheetViewShowFormulas             = Nothing
    , _sheetViewShowGridLines            = Nothing
    , _sheetViewShowOutlineSymbols       = Nothing
    , _sheetViewShowRowColHeaders        = Nothing
    , _sheetViewShowRuler                = Nothing
    , _sheetViewShowWhiteSpace           = Nothing
    , _sheetViewShowZeros                = Nothing
    , _sheetViewTabSelected              = Nothing
    , _sheetViewTopLeftCell              = Nothing
    , _sheetViewType                     = Nothing
    , _sheetViewWindowProtection         = Nothing
    , _sheetViewWorkbookViewId           = 0
    , _sheetViewZoomScale                = Nothing
    , _sheetViewZoomScaleNormal          = Nothing
    , _sheetViewZoomScalePageLayoutView  = Nothing
    , _sheetViewZoomScaleSheetLayoutView = Nothing
    , _sheetViewPane                     = Nothing
    , _sheetViewSelection                = []
    }

instance Default Selection where
  def = Selection {
      _selectionActiveCell   = Nothing
    , _selectionActiveCellId = Nothing
    , _selectionPane         = Nothing
    , _selectionSqref        = Nothing
    }

instance Default Pane where
  def = Pane {
      _paneActivePane  = Nothing
    , _paneState       = Nothing
    , _paneTopLeftCell = Nothing
    , _paneXSplit      = Nothing
    , _paneYSplit      = Nothing
    }

{-------------------------------------------------------------------------------
  Rendering
-------------------------------------------------------------------------------}

-- | See @CT_SheetView@, p. 3913
instance ToElement SheetView where
  toElement nm SheetView{..} = Element {
      elementName       = nm
    , elementNodes      = map NodeElement . concat $ [
          map (toElement "pane")      (maybeToList _sheetViewPane)
        , map (toElement "selection") _sheetViewSelection
          -- TODO: pivotSelection
          -- TODO: extLst
        ]
    , elementAttributes = Map.fromList . catMaybes $ [
          "windowProtection"         .=? _sheetViewWindowProtection
        , "showFormulas"             .=? _sheetViewShowFormulas
        , "showGridLines"            .=? _sheetViewShowGridLines
        , "showRowColHeaders"        .=? _sheetViewShowRowColHeaders
        , "showZeros"                .=? _sheetViewShowZeros
        , "rightToLeft"              .=? _sheetViewRightToLeft
        , "tabSelected"              .=? _sheetViewTabSelected
        , "showRuler"                .=? _sheetViewShowRuler
        , "showOutlineSymbols"       .=? _sheetViewShowOutlineSymbols
        , "defaultGridColor"         .=? _sheetViewDefaultGridColor
        , "showWhiteSpace"           .=? _sheetViewShowWhiteSpace
        , "view"                     .=? _sheetViewType
        , "topLeftCell"              .=? _sheetViewTopLeftCell
        , "colorId"                  .=? _sheetViewColorId
        , "zoomScale"                .=? _sheetViewZoomScale
        , "zoomScaleNormal"          .=? _sheetViewZoomScaleNormal
        , "zoomScaleSheetLayoutView" .=? _sheetViewZoomScaleSheetLayoutView
        , "zoomScalePageLayoutView"  .=? _sheetViewZoomScalePageLayoutView
        , Just $ "workbookViewId"    .=  _sheetViewWorkbookViewId
        ]
    }

-- | See @CT_Selection@, p. 3914
instance ToElement Selection where
  toElement nm Selection{..} = Element {
      elementName       = nm
    , elementNodes      = []
    , elementAttributes = Map.fromList . catMaybes $ [
          "pane"         .=? _selectionPane
        , "activeCell"   .=? _selectionActiveCell
        , "activeCellId" .=? _selectionActiveCellId
        , "sqref"        .=? _selectionSqref
        ]
    }

-- | See @CT_Pane@, p. 3913
instance ToElement Pane where
  toElement nm Pane{..} = Element {
      elementName       = nm
    , elementNodes      = []
    , elementAttributes = Map.fromList . catMaybes $ [
          "xSplit"      .=? _paneXSplit
        , "ySplit"      .=? _paneYSplit
        , "topLeftCell" .=? _paneTopLeftCell
        , "activePane"  .=? _paneActivePane
        , "state"       .=? _paneState
        ]
    }

-- | See @ST_SheetViewType@, p. 3913
instance ToAttrVal SheetViewType where
  toAttrVal SheetViewTypeNormal           = "normal"
  toAttrVal SheetViewTypePageBreakPreview = "pageBreakPreview"
  toAttrVal SheetViewTypePageLayout       = "pageLayout"

-- | See @ST_Pane@, p. 3914
instance ToAttrVal PaneType where
  toAttrVal PaneTypeBottomRight = "bottomRight"
  toAttrVal PaneTypeTopRight    = "topRight"
  toAttrVal PaneTypeBottomLeft  = "bottomLeft"
  toAttrVal PaneTypeTopLeft     = "topLeft"

-- | See @ST_PaneState@, p. 3929
instance ToAttrVal PaneState where
  toAttrVal PaneStateSplit       = "split"
  toAttrVal PaneStateFrozen      = "frozen"
  toAttrVal PaneStateFrozenSplit = "frozenSplit"

{-------------------------------------------------------------------------------
  Parsing
-------------------------------------------------------------------------------}
-- | See @CT_SheetView@, p. 3913
instance FromCursor SheetView where
  fromCursor cur = do
    _sheetViewWindowProtection         <- maybeAttribute "windowProtection" cur
    _sheetViewShowFormulas             <- maybeAttribute "showFormulas" cur
    _sheetViewShowGridLines            <- maybeAttribute "showGridLines" cur
    _sheetViewShowRowColHeaders        <- maybeAttribute "showRowColHeaders"cur
    _sheetViewShowZeros                <- maybeAttribute "showZeros" cur
    _sheetViewRightToLeft              <- maybeAttribute "rightToLeft" cur
    _sheetViewTabSelected              <- maybeAttribute "tabSelected" cur
    _sheetViewShowRuler                <- maybeAttribute "showRuler" cur
    _sheetViewShowOutlineSymbols       <- maybeAttribute "showOutlineSymbols" cur
    _sheetViewDefaultGridColor         <- maybeAttribute "defaultGridColor" cur
    _sheetViewShowWhiteSpace           <- maybeAttribute "showWhiteSpace" cur
    _sheetViewType                     <- maybeAttribute "view" cur
    _sheetViewTopLeftCell              <- maybeAttribute "topLeftCell" cur
    _sheetViewColorId                  <- maybeAttribute "colorId" cur
    _sheetViewZoomScale                <- maybeAttribute "zoomScale" cur
    _sheetViewZoomScaleNormal          <- maybeAttribute "zoomScaleNormal" cur
    _sheetViewZoomScaleSheetLayoutView <- maybeAttribute "zoomScaleSheetLayoutView" cur
    _sheetViewZoomScalePageLayoutView  <- maybeAttribute "zoomScalePageLayoutView" cur
    _sheetViewWorkbookViewId           <- fromAttribute "workbookViewId" cur
    let _sheetViewPane = listToMaybe $ cur $/ element (n_ "pane") >=> fromCursor
        _sheetViewSelection = cur $/ element (n_ "selection") >=> fromCursor
    return SheetView{..}

instance FromXenoNode SheetView where
  fromXenoNode root = parseAttributes root $ do
    _sheetViewWindowProtection         <- maybeAttr "windowProtection"
    _sheetViewShowFormulas             <- maybeAttr "showFormulas"
    _sheetViewShowGridLines            <- maybeAttr "showGridLines"
    _sheetViewShowRowColHeaders        <- maybeAttr "showRowColHeaders"
    _sheetViewShowZeros                <- maybeAttr "showZeros"
    _sheetViewRightToLeft              <- maybeAttr "rightToLeft"
    _sheetViewTabSelected              <- maybeAttr "tabSelected"
    _sheetViewShowRuler                <- maybeAttr "showRuler"
    _sheetViewShowOutlineSymbols       <- maybeAttr "showOutlineSymbols"
    _sheetViewDefaultGridColor         <- maybeAttr "defaultGridColor"
    _sheetViewShowWhiteSpace           <- maybeAttr "showWhiteSpace"
    _sheetViewType                     <- maybeAttr "view"
    _sheetViewTopLeftCell              <- maybeAttr "topLeftCell"
    _sheetViewColorId                  <- maybeAttr "colorId"
    _sheetViewZoomScale                <- maybeAttr "zoomScale"
    _sheetViewZoomScaleNormal          <- maybeAttr "zoomScaleNormal"
    _sheetViewZoomScaleSheetLayoutView <- maybeAttr "zoomScaleSheetLayoutView"
    _sheetViewZoomScalePageLayoutView  <- maybeAttr "zoomScalePageLayoutView"
    _sheetViewWorkbookViewId           <- fromAttr "workbookViewId"
    (_sheetViewPane, _sheetViewSelection) <-
      toAttrParser . collectChildren root $
      (,) <$> maybeFromChild "pane" <*> fromChildList "selection"
    return SheetView {..}

-- | See @CT_Pane@, p. 3913
instance FromCursor Pane where
  fromCursor cur = do
    _paneXSplit      <- maybeAttribute "xSplit" cur
    _paneYSplit      <- maybeAttribute "ySplit" cur
    _paneTopLeftCell <- maybeAttribute "topLeftCell" cur
    _paneActivePane  <- maybeAttribute "activePane" cur
    _paneState       <- maybeAttribute "state" cur
    return Pane{..}

instance FromXenoNode Pane where
  fromXenoNode root =
    parseAttributes root $ do
      _paneXSplit <- maybeAttr "xSplit"
      _paneYSplit <- maybeAttr "ySplit"
      _paneTopLeftCell <- maybeAttr "topLeftCell"
      _paneActivePane <- maybeAttr "activePane"
      _paneState <- maybeAttr "state"
      return Pane {..}

-- | See @CT_Selection@, p. 3914
instance FromCursor Selection where
  fromCursor cur = do
    _selectionPane         <- maybeAttribute "pane" cur
    _selectionActiveCell   <- maybeAttribute "activeCell" cur
    _selectionActiveCellId <- maybeAttribute "activeCellId" cur
    _selectionSqref        <- maybeAttribute "sqref" cur
    return Selection{..}

instance FromXenoNode Selection where
  fromXenoNode root =
    parseAttributes root $ do
      _selectionPane <- maybeAttr "pane"
      _selectionActiveCell <- maybeAttr "activeCell"
      _selectionActiveCellId <- maybeAttr "activeCellId"
      _selectionSqref <- maybeAttr "sqref"
      return Selection {..}

-- | See @ST_SheetViewType@, p. 3913
instance FromAttrVal SheetViewType where
  fromAttrVal "normal"           = readSuccess SheetViewTypeNormal
  fromAttrVal "pageBreakPreview" = readSuccess SheetViewTypePageBreakPreview
  fromAttrVal "pageLayout"       = readSuccess SheetViewTypePageLayout
  fromAttrVal t                  = invalidText "SheetViewType" t

instance FromAttrBs SheetViewType where
  fromAttrBs "normal"           = return SheetViewTypeNormal
  fromAttrBs "pageBreakPreview" = return SheetViewTypePageBreakPreview
  fromAttrBs "pageLayout"       = return SheetViewTypePageLayout
  fromAttrBs x                  = unexpectedAttrBs "SheetViewType" x

-- | See @ST_Pane@, p. 3914
instance FromAttrVal PaneType where
  fromAttrVal "bottomRight" = readSuccess PaneTypeBottomRight
  fromAttrVal "topRight"    = readSuccess PaneTypeTopRight
  fromAttrVal "bottomLeft"  = readSuccess PaneTypeBottomLeft
  fromAttrVal "topLeft"     = readSuccess PaneTypeTopLeft
  fromAttrVal t             = invalidText "PaneType" t

instance FromAttrBs PaneType where
  fromAttrBs "bottomRight" = return PaneTypeBottomRight
  fromAttrBs "topRight"    = return PaneTypeTopRight
  fromAttrBs "bottomLeft"  = return PaneTypeBottomLeft
  fromAttrBs "topLeft"     = return PaneTypeTopLeft
  fromAttrBs x             = unexpectedAttrBs "PaneType" x

-- | See @ST_PaneState@, p. 3929
instance FromAttrVal PaneState where
  fromAttrVal "split"       = readSuccess PaneStateSplit
  fromAttrVal "frozen"      = readSuccess PaneStateFrozen
  fromAttrVal "frozenSplit" = readSuccess PaneStateFrozenSplit
  fromAttrVal t             = invalidText "PaneState" t

instance FromAttrBs PaneState where
  fromAttrBs "split"       = return PaneStateSplit
  fromAttrBs "frozen"      = return PaneStateFrozen
  fromAttrBs "frozenSplit" = return PaneStateFrozenSplit
  fromAttrBs x             = unexpectedAttrBs "PaneState" x
