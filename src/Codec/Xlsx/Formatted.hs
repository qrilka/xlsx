-- | Higher level interface for creating styled worksheets
{-# LANGUAGE CPP             #-}
{-# LANGUAGE RecordWildCards #-}
{-# LANGUAGE TemplateHaskell #-}
{-# LANGUAGE RankNTypes      #-}
{-# OPTIONS_GHC -Wall        #-}
module Codec.Xlsx.Formatted (
    FormattedCell(..)
  , Formatted(..)
  , formatted
    -- * Lenses
    -- ** FormattedCell
  , formattedAlignment
  , formattedBorder
  , formattedFill
  , formattedFont
  , formattedNumberFormat
  , formattedProtection
  , formattedPivotButton
  , formattedQuotePrefix
  , formattedValue
  , formattedColSpan
  , formattedRowSpan
  ) where

import Prelude hiding (mapM)
import Control.Lens
import Control.Monad.State hiding (mapM, forM_)
import Data.Default
import Data.Foldable (forM_)
import Data.List (sortBy)
import Data.Map (Map)
import Data.Ord (comparing)
import Data.Traversable (mapM)
import Data.Tuple (swap)
import qualified Data.Map as Map

import Codec.Xlsx.Types

{-------------------------------------------------------------------------------
  Internal: formatting state
-------------------------------------------------------------------------------}

data FormattingState = FormattingState {
    _formattingBorders :: Map Border Int
  , _formattingCellXfs :: Map CellXf Int
  , _formattingFills   :: Map Fill   Int
  , _formattingFonts   :: Map Font   Int
  -- TODO: dxfs
  , _formattingMerges  :: [Range]         -- ^ In reverse order
  }

makeLenses ''FormattingState

stateFromStyleSheet :: StyleSheet -> FormattingState
stateFromStyleSheet StyleSheet{..} = FormattingState{
      _formattingBorders = fromValueList _styleSheetBorders
    , _formattingCellXfs = fromValueList _styleSheetCellXfs
    , _formattingFills   = fromValueList _styleSheetFills
    , _formattingFonts   = fromValueList _styleSheetFonts
    , _formattingMerges  = []
    }
  where
    fromValueList :: Ord a => [a] -> Map a Int
    fromValueList = Map.fromList . (`zip` [0..])

stateToStyleSheet :: FormattingState -> StyleSheet
stateToStyleSheet FormattingState{..} = StyleSheet{
      _styleSheetBorders = toList  _formattingBorders
    , _styleSheetCellXfs = toList  _formattingCellXfs
    , _styleSheetFills   = toList  _formattingFills
    , _styleSheetFonts   = toList  _formattingFonts
    , _styleSheetDxfs    = []
    }
  where
    toList :: Map a Int -> [a]
    toList = map snd . sortBy (comparing fst) . map swap . Map.toList

getId :: Ord a => Lens' FormattingState (Map a Int) -> a -> State FormattingState Int
getId f a = do
    aMap <- use f
    case Map.lookup a aMap of
      Just aId -> return aId
      Nothing  -> do let aId = Map.size aMap
                     f %= Map.insert a aId
                     return aId

{-------------------------------------------------------------------------------
  Cell with formatting
-------------------------------------------------------------------------------}

-- | Cell with formatting
--
-- See 'formatted' for more details.
--
-- TODOs:
--
-- * Add a number format ('_cellXfApplyNumberFormat', '_cellXfNumFmtId')
-- * Add references to the named style sheets ('_cellXfId')
data FormattedCell = FormattedCell {
    _formattedAlignment    :: Maybe Alignment
  , _formattedBorder       :: Maybe Border
  , _formattedFill         :: Maybe Fill
  , _formattedFont         :: Maybe Font
  , _formattedNumberFormat :: Maybe NumberFormat
  , _formattedProtection   :: Maybe Protection
  , _formattedPivotButton  :: Maybe Bool
  , _formattedQuotePrefix  :: Maybe Bool
  , _formattedValue        :: Maybe CellValue
  , _formattedColSpan      :: Int
  , _formattedRowSpan      :: Int
  }
  deriving (Show, Eq)

makeLenses ''FormattedCell

instance Default FormattedCell where
  def = FormattedCell {
      _formattedAlignment    = Nothing
    , _formattedBorder       = Nothing
    , _formattedFill         = Nothing
    , _formattedFont         = Nothing
    , _formattedNumberFormat = Nothing
    , _formattedProtection   = Nothing
    , _formattedPivotButton  = Nothing
    , _formattedQuotePrefix  = Nothing
    , _formattedValue        = Nothing
    , _formattedColSpan      = 1
    , _formattedRowSpan      = 1
    }

{-------------------------------------------------------------------------------
  Client-facing API
-------------------------------------------------------------------------------}

-- | Result of formatting
--
-- See 'formatted'
data Formatted = Formatted {
    -- | The final 'CellMap'; see '_wsCells'
    formattedCellMap  :: CellMap

    -- | The final stylesheet; see '_xlStyles' (and 'renderStyleSheet')
  , formattedStyleSheet :: StyleSheet

    -- | The final list of cell merges; see '_wsMerges'
  , formattedMerges :: [Range]
  }

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
formatted :: Map (Int, Int) FormattedCell -> StyleSheet -> Formatted
formatted cs styleSheet =
   let initSt         = stateFromStyleSheet styleSheet
       (cs', finalSt) = runState (mapM (uncurry formatCell) (Map.toList cs)) initSt
       styleSheet'    = stateToStyleSheet finalSt
   in Formatted {
          formattedCellMap    = Map.fromList (concat cs')
        , formattedStyleSheet = styleSheet'
        , formattedMerges     = reverse (finalSt ^. formattingMerges)
        }

-- | Format a cell with (potentially) rowspan or colspan
formatCell :: (Int, Int) -> FormattedCell -> State FormattingState [((Int, Int), Cell)]
formatCell (row, col) cell = do
    let (block, mMerge) = cellBlock (row, col) cell
    forM_ mMerge $ \merge -> formattingMerges %= (:) merge
    mapM go block
  where
    go :: ((Int, Int), FormattedCell) -> State FormattingState ((Int, Int), Cell)
    go (pos, c) = do
      styleId <- cellStyleId c
      return (pos, Cell styleId (_formattedValue c) Nothing)

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
        else def & formattedBorder .~ Just (borderAt (row', col'))

    borderAt :: (Int, Int) -> Border
    borderAt (row', col') = def
      & borderTop    .~ do guard (row' == topRow)    ; _borderTop    =<< _formattedBorder
      & borderBottom .~ do guard (row' == bottomRow) ; _borderBottom =<< _formattedBorder
      & borderLeft   .~ do guard (col' == leftCol)   ; _borderLeft   =<< _formattedBorder
      & borderRight  .~ do guard (col' == rightCol)  ; _borderRight  =<< _formattedBorder

    topRow, bottomRow, leftCol, rightCol :: Int
    topRow    = row
    bottomRow = row + _formattedRowSpan - 1
    leftCol   = col
    rightCol  = col + _formattedColSpan - 1

cellStyleId :: FormattedCell -> State FormattingState (Maybe Int)
cellStyleId c = mapM (getId formattingCellXfs) =<< cellXf c

cellXf :: FormattedCell -> State FormattingState (Maybe CellXf)
cellXf FormattedCell{..} = do
    mBorderId <- getId formattingBorders `mapM` _formattedBorder
    mFillId   <- getId formattingFills   `mapM` _formattedFill
    mFontId   <- getId formattingFonts   `mapM` _formattedFont
    let mNumFmtId = fmap numberFormatId _formattedNumberFormat
    let xf = CellXf {
            _cellXfApplyAlignment    = apply _formattedAlignment
          , _cellXfApplyBorder       = apply mBorderId
          , _cellXfApplyFill         = apply mFillId
          , _cellXfApplyFont         = apply mFontId
          , _cellXfApplyNumberFormat = apply _formattedNumberFormat
          , _cellXfApplyProtection   = apply _formattedProtection
          , _cellXfBorderId          = mBorderId
          , _cellXfFillId            = mFillId
          , _cellXfFontId            = mFontId
          , _cellXfNumFmtId          = mNumFmtId
          , _cellXfPivotButton       = _formattedPivotButton
          , _cellXfQuotePrefix       = _formattedQuotePrefix
          , _cellXfId                = Nothing -- TODO
          , _cellXfAlignment         = _formattedAlignment
          , _cellXfProtection        = _formattedProtection
          }
    return $ if xf == def then Nothing else Just xf
  where
    -- If we have formatting instructions, we want to set the corresponding
    -- applyXXX properties
    apply :: Maybe a -> Maybe Bool
    apply Nothing  = Nothing
    apply (Just _) = Just True
