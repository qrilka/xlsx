-- | Higher level interface for creating styled worksheets
{-# LANGUAGE CPP      #-}
{-# LANGUAGE RankNTypes      #-}
{-# LANGUAGE RecordWildCards #-}
{-# LANGUAGE TemplateHaskell #-}
{-# LANGUAGE DeriveGeneric #-}
module Codec.Xlsx.Formatted
  ( FormattedCell(..)
  , Formatted(..)
  , Format(..)
  , formatted
  , formatWorkbook
  , toFormattedCells
  , CondFormatted(..)
  , conditionallyFormatted
    -- * Lenses
    -- ** Format
  , formatAlignment
  , formatBorder
  , formatFill
  , formatFont
  , formatNumberFormat
  , formatProtection
  , formatPivotButton
  , formatQuotePrefix
    -- ** FormattedCell
  , formattedCell
  , formattedFormat
  , formattedColSpan
  , formattedRowSpan
    -- ** FormattedCondFmt
  , condfmtCondition
  , condfmtDxf
  , condfmtPriority
  , condfmtStopIfTrue
  ) where

#ifdef USE_MICROLENS
import Lens.Micro
import Lens.Micro.Mtl
import Lens.Micro.TH
import Lens.Micro.GHC ()
#else
import Control.Lens
#endif
import Control.Monad.State hiding (forM_, mapM)
import Data.Default
import Data.Foldable (asum, forM_)
import Data.Function (on)
import Data.List (foldl', groupBy, sortBy, sortBy)
import Data.Map (Map)
import qualified Data.Map as M
import Data.Ord (comparing)
import Data.Text (Text)
import Data.Traversable (mapM)
import Data.Tuple (swap)
import GHC.Generics (Generic)
import Prelude hiding (mapM)
import Safe (headNote, fromJustNote)

import Codec.Xlsx.Types
import Codec.Xlsx.Types.Cell

{-------------------------------------------------------------------------------
  Internal: formatting state
-------------------------------------------------------------------------------}

data FormattingState = FormattingState {
    _formattingBorders :: Map Border Int
  , _formattingCellXfs :: Map CellXf Int
  , _formattingFills   :: Map Fill   Int
  , _formattingFonts   :: Map Font   Int
  , _formattingNumFmts :: Map Text   Int
  , _formattingMerges  :: [Range] -- ^ In reverse order
  }

makeLenses ''FormattingState

stateFromStyleSheet :: StyleSheet -> FormattingState
stateFromStyleSheet StyleSheet{..} = FormattingState{
      _formattingBorders = fromValueList _styleSheetBorders
    , _formattingCellXfs = fromValueList _styleSheetCellXfs
    , _formattingFills   = fromValueList _styleSheetFills
    , _formattingFonts   = fromValueList _styleSheetFonts
    , _formattingNumFmts = M.fromList . map swap $ M.toList _styleSheetNumFmts
    , _formattingMerges  = []
    }

fromValueList :: Ord a => [a] -> Map a Int
fromValueList = M.fromList . (`zip` [0..])

toValueList :: Map a Int -> [a]
toValueList = map snd . sortBy (comparing fst) . map swap . M.toList

updateStyleSheetFromState :: StyleSheet -> FormattingState -> StyleSheet
updateStyleSheetFromState sSheet FormattingState{..} = sSheet
    { _styleSheetBorders = toValueList _formattingBorders
    , _styleSheetCellXfs = toValueList _formattingCellXfs
    , _styleSheetFills   = toValueList _formattingFills
    , _styleSheetFonts   = toValueList _formattingFonts
    , _styleSheetNumFmts = M.fromList . map swap $ M.toList _formattingNumFmts
    }

getId :: Ord a => Lens' FormattingState (Map a Int) -> a -> State FormattingState Int
getId = getId' 0

getId' :: Ord a
       => Int
       -> Lens' FormattingState (Map a Int)
       -> a
       -> State FormattingState Int
getId' k f v = do
    aMap <- use f
    case M.lookup v aMap of
      Just anId -> return anId
      Nothing  -> do let anId = k + M.size aMap
                     f %= M.insert v anId
                     return anId

{-------------------------------------------------------------------------------
  Unwrapped cell conditional formatting
-------------------------------------------------------------------------------}

data FormattedCondFmt = FormattedCondFmt
    { _condfmtCondition  :: Condition
    , _condfmtDxf        :: Dxf
    , _condfmtPriority   :: Int
    , _condfmtStopIfTrue :: Maybe Bool
    } deriving (Eq, Show, Generic)

makeLenses ''FormattedCondFmt

{-------------------------------------------------------------------------------
  Cell with formatting
-------------------------------------------------------------------------------}

-- | Formatting options used to format cells
--
-- TODOs:
--
-- * Add a number format ('_cellXfApplyNumberFormat', '_cellXfNumFmtId')
-- * Add references to the named style sheets ('_cellXfId')
data Format = Format
    { _formatAlignment    :: Maybe Alignment
    , _formatBorder       :: Maybe Border
    , _formatFill         :: Maybe Fill
    , _formatFont         :: Maybe Font
    , _formatNumberFormat :: Maybe NumberFormat
    , _formatProtection   :: Maybe Protection
    , _formatPivotButton  :: Maybe Bool
    , _formatQuotePrefix  :: Maybe Bool
    } deriving (Eq, Show, Generic)

makeLenses ''Format

-- | Cell with formatting. '_cellStyle' property of '_formattedCell' is ignored
--
-- See 'formatted' for more details.
data FormattedCell = FormattedCell
    { _formattedCell    :: Cell
    , _formattedFormat  :: Format
    , _formattedColSpan :: Int
    , _formattedRowSpan :: Int
    } deriving (Eq, Show, Generic)

makeLenses ''FormattedCell

{-------------------------------------------------------------------------------
  Default instances
-------------------------------------------------------------------------------}

instance Default FormattedCell where
  def = FormattedCell
        { _formattedCell    = def
        , _formattedFormat  = def
        , _formattedColSpan = 1
        , _formattedRowSpan = 1
        }

instance Default Format where
  def = Format
        { _formatAlignment    = Nothing
        , _formatBorder       = Nothing
        , _formatFill         = Nothing
        , _formatFont         = Nothing
        , _formatNumberFormat = Nothing
        , _formatProtection   = Nothing
        , _formatPivotButton  = Nothing
        , _formatQuotePrefix  = Nothing
        }

instance Default FormattedCondFmt where
  def = FormattedCondFmt ContainsBlanks def topCfPriority Nothing

{-------------------------------------------------------------------------------
  Client-facing API
-------------------------------------------------------------------------------}

-- | Result of formatting
--
-- See 'formatted'
data Formatted = Formatted {
    -- | The final 'CellMap'; see '_wsCells'
    formattedCellMap    :: CellMap

    -- | The final stylesheet; see '_xlStyles' (and 'renderStyleSheet')
  , formattedStyleSheet :: StyleSheet

    -- | The final list of cell merges; see '_wsMerges'
  , formattedMerges     :: [Range]
  } deriving (Eq, Show, Generic)

-- | Higher level API for creating formatted documents
--
-- Creating formatted Excel spreadsheets using the 'Cell' datatype directly,
-- even with the support for the 'StyleSheet' datatype, is fairly painful.
-- This has a number of causes:
--
-- * The 'Cell' datatype wants an 'Int' for the style, which is supposed to
--   point into the '_styleSheetCellXfs' part of a stylesheet. However, this can
--   be difficult to work with, as it requires manual tracking of cell style
--   IDs, which in turns requires manual tracking of font IDs, border IDs, etc.
-- * Row-span and column-span properties are set on the worksheet as a whole
--   ('wsMerges') rather than on individual cells.
-- * Excel does not correctly deal with borders on cells that span multiple
--   columns or rows. Instead, these rows must be set on all the edge cells
--   in the block. Again, this means that this becomes a global property of
--   the spreadsheet rather than properties of individual cells.
--
-- This function deals with all these problems. Given a map of 'FormattedCell's,
-- which refer directly to 'Font's, 'Border's, etc. (rather than font IDs,
-- border IDs, etc.), and an initial stylesheet, it recovers all possible
-- sharing, constructs IDs, and then constructs the final 'CellMap', as well as
-- the final stylesheet and list of merges.
--
-- If you don't already have a 'StyleSheet' you want to use as starting point
-- then 'minimalStyleSheet' is a good choice.
formatted :: Map (RowIndex, ColIndex) FormattedCell -> StyleSheet -> Formatted
formatted cs styleSheet =
   let initSt         = stateFromStyleSheet styleSheet
       (cs', finalSt) = runState (mapM (uncurry formatCell) (M.toList cs)) initSt
       styleSheet'    = updateStyleSheetFromState styleSheet finalSt
   in Formatted {
          formattedCellMap    = toNested $ M.fromList (concat cs')
        , formattedStyleSheet = styleSheet'
        , formattedMerges     = reverse (finalSt ^. formattingMerges)
        }

formatWorkbook :: [(Text, Map (Int, Int) FormattedCell)] -> StyleSheet -> Xlsx
formatWorkbook nfcss initStyle = extract go
  where
    initSt = stateFromStyleSheet initStyle
    go = flip runState initSt $
      forM nfcss $ \(name, fcs) -> do
        cs' <- forM (M.toList fcs) $ \(rc, fc) -> formatCell rc fc
        merges <- reverse . _formattingMerges <$> get
        return ( name
               , def & wsCells  .~ toNested (M.fromList (concat cs'))
                     & wsMerges .~ merges)
    extract (sheets, st) =
      def & xlSheets .~ sheets
          & xlStyles .~ renderStyleSheet (updateStyleSheetFromState initStyle st)

-- | reverse to 'formatted' which allows to get a map of formatted cells
-- from an existing worksheet and its workbook's style sheet
toFormattedCells :: CellMap -> [Range] -> StyleSheet -> Map (Int, Int) FormattedCell
toFormattedCells m merges StyleSheet{..} = applyMerges $
  unNested $ fmap toFormattedCell <$> m
  where
    toFormattedCell :: Cell -> FormattedCell
    toFormattedCell cell@Cell{..} =
        FormattedCell
        { _formattedCell    = cell{ _cellStyle = Nothing } -- just to remove confusion
        , _formattedFormat  = maybe def formatFromStyle $ flip M.lookup cellXfs =<< _cellStyle
        , _formattedColSpan = 1
        , _formattedRowSpan = 1 }
    formatFromStyle cellXf =
        Format
        { _formatAlignment    = applied _cellXfApplyAlignment _cellXfAlignment cellXf
        , _formatBorder       = flip M.lookup borders =<<
                                applied _cellXfApplyBorder _cellXfBorderId cellXf
        , _formatFill         = flip M.lookup fills =<<
                                applied _cellXfApplyFill _cellXfFillId cellXf
        , _formatFont         = flip M.lookup fonts =<<
                                applied _cellXfApplyFont _cellXfFontId cellXf
        , _formatNumberFormat = lookupNumFmt =<<
                                applied _cellXfApplyNumberFormat _cellXfNumFmtId cellXf
        , _formatProtection   = _cellXfProtection  cellXf
        , _formatPivotButton  = _cellXfPivotButton cellXf
        , _formatQuotePrefix  = _cellXfQuotePrefix cellXf }
    idMapped :: [a] -> Map Int a
    idMapped = M.fromList . zip [0..]
    cellXfs = idMapped _styleSheetCellXfs
    borders = idMapped _styleSheetBorders
    fills = idMapped _styleSheetFills
    fonts = idMapped _styleSheetFonts
    lookupNumFmt fId = asum
        [ StdNumberFormat <$> idToStdNumberFormat fId
        , UserNumberFormat <$> M.lookup fId _styleSheetNumFmts]
    applied :: (CellXf -> Maybe Bool) -> (CellXf -> Maybe a) -> CellXf -> Maybe a
    applied applyProp prop cXf = do
        apply <- applyProp cXf
        if apply then prop cXf else fail "not applied"
    applyMerges cells = foldl' onlyTopLeft cells merges
    onlyTopLeft cells range = flip execState cells $ do
        let ((r1, c1), (r2, c2)) = fromJustNote "fromRange" $ fromRange range
            nonTopLeft = tail [(r, c) | r<-[r1..r2], c<-[c1..c2]]
        forM_ nonTopLeft (modify . M.delete)
        at (r1, c1) . non def . formattedRowSpan .= (r2 - r1 +1)
        at (r1, c1) . non def . formattedColSpan .= (c2 - c1 +1)

data CondFormatted = CondFormatted {
    -- | The resulting stylesheet
    condformattedStyleSheet    :: StyleSheet
    -- | The final map of conditional formatting rules applied to ranges
    , condformattedFormattings :: Map SqRef ConditionalFormatting
    } deriving (Eq, Show, Generic)

conditionallyFormatted :: Map CellRef [FormattedCondFmt] -> StyleSheet -> CondFormatted
conditionallyFormatted cfs styleSheet = CondFormatted
    { condformattedStyleSheet  = styleSheet & styleSheetDxfs .~ finalDxfs
    , condformattedFormattings = fmts
    }
  where
    (cellFmts, dxf2id) = runState (mapM (mapM mapDxf) cfs) dxf2id0
    dxf2id0 = fromValueList (styleSheet ^. styleSheetDxfs)
    fmts = M.fromList . map mergeSqRef . groupBy ((==) `on` snd) .
           sortBy (comparing snd) $ M.toList cellFmts
    mergeSqRef cellRefs2fmt =
        (SqRef (map fst cellRefs2fmt),
         headNote "fmt group should not be empty" (map snd cellRefs2fmt))
    finalDxfs = toValueList dxf2id

{-------------------------------------------------------------------------------
  Implementation details
-------------------------------------------------------------------------------}

-- | Format a cell with (potentially) rowspan or colspan
formatCell :: (Int, Int) -> FormattedCell -> State FormattingState [((Int, Int), Cell)]
formatCell (row, col) cell = do
    let (block, mMerge) = cellBlock (row, col) cell
    forM_ mMerge $ \merge -> formattingMerges %= (:) merge
    mapM go block
  where
    go :: ((Int, Int), FormattedCell) -> State FormattingState ((Int, Int), Cell)
    go (pos, c@FormattedCell{..}) = do
      styleId <- cellStyleId c
      return (pos, _formattedCell{_cellStyle = styleId})

-- | Cell block corresponding to a single 'FormattedCell'
--
-- A single 'FormattedCell' might have a colspan or rowspan greater than 1.
-- Although Excel obviously supports cell merges, it does not correctly apply
-- borders to the cells covered by the rowspan or colspan. Therefore we create
-- a block of cells in this function; the top-left is the cell proper, and the
-- remaining cells are the cells covered by the rowspan/colspan.
--
-- Also returns the cell merge instruction, if any.
cellBlock :: (Int, Int) -> FormattedCell
          -> ([((Int, Int), FormattedCell)], Maybe Range)
cellBlock (row, col) cell@FormattedCell{..} = (block, merge)
  where
    block :: [((Int, Int), FormattedCell)]
    block = [ ((row', col'), cellAt (row', col'))
            | row' <- [topRow  .. bottomRow]
            , col' <- [leftCol .. rightCol]
            ]

    merge :: Maybe Range
    merge = do guard (topRow /= bottomRow || leftCol /= rightCol)
               return $ mkRange (topRow, leftCol) (bottomRow, rightCol)

    cellAt :: (Int, Int) -> FormattedCell
    cellAt (row', col') =
      if row' == row && col == col'
        then cell
        else def & formattedFormat . formatBorder ?~ borderAt (row', col')

    border = _formatBorder _formattedFormat

    borderAt :: (Int, Int) -> Border
    borderAt (row', col') = def
      & borderTop    .~ do guard (row' == topRow)    ; _borderTop    =<< border
      & borderBottom .~ do guard (row' == bottomRow) ; _borderBottom =<< border
      & borderLeft   .~ do guard (col' == leftCol)   ; _borderLeft   =<< border
      & borderRight  .~ do guard (col' == rightCol)  ; _borderRight  =<< border

    topRow, bottomRow, leftCol, rightCol :: Int
    topRow    = row
    bottomRow = row + _formattedRowSpan - 1
    leftCol   = col
    rightCol  = col + _formattedColSpan - 1

cellStyleId :: FormattedCell -> State FormattingState (Maybe Int)
cellStyleId c = mapM (getId formattingCellXfs) =<< constructCellXf c

constructCellXf :: FormattedCell -> State FormattingState (Maybe CellXf)
constructCellXf FormattedCell{_formattedFormat=Format{..}} = do
    mBorderId <- getId formattingBorders `mapM` _formatBorder
    mFillId   <- getId formattingFills   `mapM` _formatFill
    mFontId   <- getId formattingFonts   `mapM` _formatFont
    let getFmtId :: Lens' FormattingState (Map Text Int) -> NumberFormat -> State FormattingState Int
        getFmtId _ (StdNumberFormat  fmt) = return (stdNumberFormatId fmt)
        getFmtId l (UserNumberFormat fmt) = getId' firstUserNumFmtId l fmt
    mNumFmtId <- getFmtId formattingNumFmts `mapM` _formatNumberFormat
    let xf = CellXf {
            _cellXfApplyAlignment    = apply _formatAlignment
          , _cellXfApplyBorder       = apply mBorderId
          , _cellXfApplyFill         = apply mFillId
          , _cellXfApplyFont         = apply mFontId
          , _cellXfApplyNumberFormat = apply _formatNumberFormat
          , _cellXfApplyProtection   = apply _formatProtection
          , _cellXfBorderId          = mBorderId
          , _cellXfFillId            = mFillId
          , _cellXfFontId            = mFontId
          , _cellXfNumFmtId          = mNumFmtId
          , _cellXfPivotButton       = _formatPivotButton
          , _cellXfQuotePrefix       = _formatQuotePrefix
          , _cellXfId                = Nothing -- TODO
          , _cellXfAlignment         = _formatAlignment
          , _cellXfProtection        = _formatProtection
          }
    return $ if xf == def then Nothing else Just xf
  where
    -- If we have formatting instructions, we want to set the corresponding
    -- applyXXX properties
    apply :: Maybe a -> Maybe Bool
    apply Nothing  = Nothing
    apply (Just _) = Just True

mapDxf :: FormattedCondFmt -> State (Map Dxf Int) CfRule
mapDxf FormattedCondFmt{..} = do
    dxf2id <- get
    dxfId <- case M.lookup _condfmtDxf dxf2id of
                 Just i ->
                     return i
                 Nothing -> do
                     let newId = M.size dxf2id
                     modify $ M.insert _condfmtDxf newId
                     return newId
    return CfRule
        { _cfrCondition  = _condfmtCondition
        , _cfrDxfId      = Just dxfId
        , _cfrPriority   = _condfmtPriority
        , _cfrStopIfTrue = _condfmtStopIfTrue
        }
