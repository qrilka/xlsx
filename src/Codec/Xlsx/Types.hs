{-# LANGUAGE NoMonomorphismRestriction #-}
{-# LANGUAGE OverloadedStrings #-}
{-# LANGUAGE TemplateHaskell #-}
module Codec.Xlsx.Types
    ( Xlsx(..), xlSheets, xlStyles
    , Styles(..)
    , emptyStyles
    , ColumnsWidth(..)
    , Worksheet(..), wsColumns, wsRowPropertiesMap, wsCells, wsMerges
    , CellMap
    , CellValue(..)
    , Cell(..), cellValue, cellStyle
    , RowProperties (..)
    , int2col
    , col2int
    , toRows
    , fromRows
    ) where

import           Control.Lens.TH
import qualified Data.ByteString.Lazy as L
import           Data.Char
import           Data.Function (on)
import           Data.List (groupBy)
import           Data.Map (Map)
import qualified Data.Map as M
import           Data.Text (Text)
import qualified Data.Text as T


-- | Cell values include text, numbers and booleans,
-- standard includes date format also but actually dates
-- are represented by numbers with a date format assigned
-- to a cell containing it
data CellValue = CellText   Text
               | CellDouble Double
               | CellBool   Bool
               deriving (Eq, Show)

data Cell = Cell
    { _cellStyle  :: Maybe Int
    , _cellValue  :: Maybe CellValue
    } deriving (Eq, Show)

makeLenses ''Cell

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

-- | Xlsx worksheet
data Worksheet = Worksheet
    { _wsColumns          :: [ColumnsWidth]         -- ^ column widths
    , _wsRowPropertiesMap :: Map Int RowProperties  -- ^ custom row properties (height, style) map
    , _wsCells            :: CellMap                -- ^ data mapped by (row, column) pairs
    , _wsMerges           :: [Text]
    } deriving (Eq, Show)

makeLenses ''Worksheet

newtype Styles = Styles {unStyles :: L.ByteString}
            deriving (Eq, Show)

-- | Structured representation of Xlsx file (currently a subset of its contents)
data Xlsx = Xlsx
    { _xlSheets :: Map Text Worksheet
    , _xlStyles :: Styles
    } deriving (Eq, Show)

makeLenses ''Xlsx

emptyStyles :: Styles
emptyStyles = Styles "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\
\<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"></styleSheet>"

-- | convert column number (starting from 1) to its textual form (e.g. 3 -> "C")
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
