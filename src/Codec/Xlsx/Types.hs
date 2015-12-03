{-# LANGUAGE RecordWildCards #-}
{-# LANGUAGE NoMonomorphismRestriction #-}
{-# LANGUAGE OverloadedStrings #-}
{-# LANGUAGE TemplateHaskell #-}
module Codec.Xlsx.Types
    ( Xlsx(..), xlSheets, xlStyles, xlDefinedNames
    , def
    , Styles(..)
    , emptyStyles
    , renderStyleSheet
    , DefinedNames(..)
    , ColumnsWidth(..)
    , PageSetup(..)
    , Worksheet(..), wsColumns, wsRowPropertiesMap, wsCells, wsMerges, wsSheetViews, wsPageSetup
    , CellMap
    , CellValue(..)
    , Cell(..), cellValue, cellStyle
    , RowProperties (..)
    , Range
    , int2col
    , col2int
    , mkCellRef
    , mkRange
    , toRows
    , fromRows
    , module X
    ) where

import           Control.Lens.TH
import qualified Data.ByteString.Lazy as L
import           Data.Char
import           Data.Default
import           Data.Function (on)
import           Data.List (groupBy)
import           Data.Map (Map)
import qualified Data.Map as M
import           Data.Text (Text)
import qualified Data.Text as T
import           Text.XML (renderLBS)
import           Text.XML.Cursor


import           Codec.Xlsx.PageSetup as X
import           Codec.Xlsx.StyleSheet as X
import           Codec.Xlsx.Types.Common as X
import           Codec.Xlsx.RichText as X
import           Codec.Xlsx.SheetViews as X
import           Codec.Xlsx.Parser.Internal
import           Codec.Xlsx.Writer.Internal

-- | Cell values include text, numbers and booleans,
-- standard includes date format also but actually dates
-- are represented by numbers with a date format assigned
-- to a cell containing it
data CellValue = CellText   Text
               | CellDouble Double
               | CellBool   Bool
               | CellRich   [RichTextRun]
               deriving (Eq, Show)

-- | Currently cell details include only cell values and style ids
-- (e.g. formulas from @\<f\>@ and inline strings from @\<is\>@
-- subelements are ignored)
data Cell = Cell
    { _cellStyle  :: Maybe Int
    , _cellValue  :: Maybe CellValue
    } deriving (Eq, Show)

makeLenses ''Cell

instance Default Cell where
    def = Cell Nothing Nothing

-- | Map containing cell values which are indexed by row and column
-- if you need to use more traditional (x,y) indexing please you could
-- use corresponding accessors from ''Codec.Xlsx.Lens''
type CellMap = Map (Int, Int) Cell

data RowProperties = RowProps { rowHeight :: Maybe Double, rowStyle::Maybe Int}
                   deriving (Read,Eq,Show,Ord)

-- | Column range (from cwMin to cwMax) width
data ColumnsWidth = ColumnsWidth
    { cwMin :: Int
    , cwMax :: Int
    , cwWidth :: Double
    , cwStyle :: Int
    } deriving (Eq, Show)

instance FromCursor ColumnsWidth where
    fromCursor c = do
      cwMin <- decimal =<< attribute "min" c
      cwMax <- decimal =<< attribute "max" c
      cwWidth <- rational =<< attribute "width" c
      cwStyle <- decimal =<< attribute "style" c
      return ColumnsWidth{..}

-- | Excel range (e.g. @D13:H14@)
type Range = Text

-- | Xlsx worksheet
data Worksheet = Worksheet
    { _wsColumns          :: [ColumnsWidth]         -- ^ column widths
    , _wsRowPropertiesMap :: Map Int RowProperties  -- ^ custom row properties (height, style) map
    , _wsCells            :: CellMap                -- ^ data mapped by (row, column) pairs
    , _wsMerges           :: [Range]                -- ^ list of cell merges
    , _wsSheetViews       :: Maybe [SheetView]
    , _wsPageSetup        :: Maybe PageSetup
    } deriving (Eq, Show)

makeLenses ''Worksheet

instance Default Worksheet where
    def = Worksheet [] M.empty M.empty [] Nothing Nothing

newtype Styles = Styles {unStyles :: L.ByteString}
            deriving (Eq, Show)

-- | Structured representation of Xlsx file (currently a subset of its contents)
data Xlsx = Xlsx
    { _xlSheets       :: Map Text Worksheet
    , _xlStyles       :: Styles
    , _xlDefinedNames :: DefinedNames
    } deriving (Eq, Show)

-- | Defined names
--
-- Each defined name consists of a name, an optional local sheet ID, and a value.
--
-- This element defines the collection of defined names for this workbook.
-- Defined names are descriptive names to represent cells, ranges of cells,
-- formulas, or constant values. Defined names can be used to represent a range
-- on any worksheet.
--
-- Excel also defines a number of reserved names with a special interpretation:
--
-- * @_xlnm.Print_Area@ specifies the workbook's print area.
--   Example value: @SheetName!$A:$A,SheetName!$1:$4@
-- * @_xlnm.Print_Titles@ specifies the row(s) or column(s) to repeat
--   at the top of each printed page.
-- * @_xlnm.Sheet_Title@:refers to a sheet title.
--
-- and others. See Section 18.2.6, "definedNames (Defined Names)" (p. 1728) of
-- the spec (second edition).
--
-- NOTE: Right now this is only a minimal implementation of defined names.
newtype DefinedNames = DefinedNames [(Text, Maybe Text, Text)]
  deriving (Eq, Show)

makeLenses ''Xlsx

instance Default Xlsx where
    def = Xlsx M.empty emptyStyles def

instance Default DefinedNames where
    def = DefinedNames []

emptyStyles :: Styles
emptyStyles = Styles "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\
\<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"></styleSheet>"

-- | Render 'StyleSheet'
--
-- This is used to render a structured 'StyleSheet' into a raw XML 'Styles'
-- document. Actually /replacing/ 'Styles' with 'StyleSheet' would mean we
-- would need to write a /parser/ for 'StyleSheet' as well (and would moreover
-- require that we support the full style sheet specification, which is still
-- quite a bit of work).
renderStyleSheet :: StyleSheet -> Styles
renderStyleSheet = Styles . renderLBS def . toDocument

-- | convert column number (starting from 1) to its textual form (e.g. 3 -> \"C\")
int2col :: Int -> Text
int2col = T.pack . reverse . map int2let . base26
    where
        int2let 0 = 'Z'
        int2let x = chr $ (x - 1) + ord 'A'
        base26  0 = []
        base26  i = let i' = (i `mod` 26)
                        i'' = if i' == 0 then 26 else i'
                    in seq i' (i' : base26 ((i - i'') `div` 26))

-- | reverse to 'int2col'
col2int :: Text -> Int
col2int = T.foldl' (\i c -> i * 26 + let2int c) 0
    where
        let2int c = 1 + ord c - ord 'A'

-- | converts cells mapped by (row, column) into rows which contain
-- row index and cells as pairs of column indices and cell values
toRows :: CellMap -> [(Int, [(Int, Cell)])]
toRows cells =
    map extractRow $ groupBy ((==) `on` (fst . fst)) $ M.toList cells
  where
    extractRow row@(((x,_),_):_) =
        (x, map (\((_,y),v) -> (y,v)) row)
    extractRow _ = error "invalid CellMap row"

-- | reverse to 'toRows'
fromRows :: [(Int, [(Int, Cell)])] -> CellMap
fromRows rows = M.fromList $ concatMap mapRow rows
  where
    mapRow (r, cells) = map (\(c, v) -> ((r, c), v)) cells

-- | Render position in @(row, col)@ format to an Excel reference.
--
-- > mkCellRef (2, 4) == "D2"
mkCellRef :: (Int, Int) -> CellRef
mkCellRef (row, col) = T.concat [int2col col, T.pack (show row)]

-- | Render range
--
-- > mkRange (2, 4) (6, 8) == "D2:H6"
mkRange :: (Int, Int) -> (Int, Int) -> Range
mkRange fr to = T.concat [mkCellRef fr, T.pack ":", mkCellRef to]
