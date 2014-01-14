{-# LANGUAGE NoMonomorphismRestriction #-}
{-# LANGUAGE OverloadedStrings #-}
module Codec.Xlsx.Types
    ( Xlsx(..)
    , WorksheetFile(..)
    , Styles(..)
    , emptyStyles
    , ColumnsWidth(..)
    , RowHeights
    , Worksheet(..)
    , CellMap
    , CellValue(..)
    , Cell(..)
    , CellData(..)
    , MappedSheet(..)
    , FullyIndexedCellValue (..)
    , RowProperties (..)
    , int2col
    , col2int
    , toRows
    , fromRows
    , xlsxLensNames
    , worksheetFileLensNames
    , worksheetLensNames
    , cellLensNames
    , cellDataLensNames
    , mappedSheetLensNames
    ) where

import qualified Data.ByteString.Lazy as L
import           Data.Char
import           Data.Function (on)
import           Data.IntMap (IntMap)
import           Data.List (groupBy)
import           Data.Map (Map)
import qualified Data.Map as M
import           Data.Text (Text)
import qualified Data.Text as T


data Xlsx = Xlsx
    { xlSheets :: Map Text Worksheet
    , xlStyles :: Styles
    } deriving (Eq, Show)

xlsxLensNames :: [ (String,String)]
xlsxLensNames = [ ("xlArchive"        , "lensXlArchive"       )
                , ("xlSharedStrings"  , "lensXlSharedStrings" )
                , ("xlStyles"         , "lensXlStyles"        )
                , ("xlWorksheetFiles" , "lensXlWorksheetFiles")]


newtype Styles = Styles {unStyles :: L.ByteString}
            deriving (Eq, Show)

emptyStyles :: Styles
emptyStyles = Styles "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\
\<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"></styleSheet>"

data WorksheetFile = WorksheetFile { wfName :: Text
                                   , wfPath :: FilePath
                                   }
                   deriving Show

worksheetFileLensNames :: [ (String,String)]
worksheetFileLensNames = [("wfName","lensWfName"),("wfPath","lensWfPath")]



-- | Column range (from cwMin to cwMax) width
data ColumnsWidth = ColumnsWidth
    { cwMin :: Int
    , cwMax :: Int
    , cwWidth :: Double
    , cwStyle :: Int
    } deriving (Eq, Show)

type RowHeights = Map Int RowProperties

data RowProperties = RowProps { rowHeight :: Maybe Double, rowStyle::Maybe Int}
                   deriving (Read,Eq,Show,Ord)

newtype MappedSheet = MappedSheet { unMappedSheet :: (IntMap  Worksheet )}

mappedSheetLensNames :: [(String,String)]
mappedSheetLensNames = [("unMappedSheet","lensMappedSheet")]

data Worksheet = Worksheet
    { wsColumns    :: [ColumnsWidth] -- ^ column widths
    , wsRowHeights :: RowHeights     -- ^ custom row height map
    , wsCells      :: CellMap        -- ^ data mapped by (row, column) pairs
    , wsMerges     :: [Text]
    } deriving (Eq, Show)

type CellMap = Map (Int, Int) CellData

worksheetLensNames:: [ (String,String)]
worksheetLensNames = [ ("wsColumns"    , "lensWsColumns"    )
                     , ("wsRowHeights" , "lensWsRowHeights" )
                     , ("wsCells"      , "lensWsCells"      )]

data FullyIndexedCellValue = FICV { ficvSheetIdx :: Int , ficvColIdx ::Int, ficvRowIdx :: Int , ficvValue :: CellValue}
                             deriving (Show)

data CellValue = CellText   Text
               | CellDouble Double
               | CellBool   Bool
               deriving (Eq, Show)

data Cell = Cell { cellIx   :: (Text, Int)
                 , cellData :: CellData
                 }
            deriving Show

cellLensNames :: [ (String,String)]
cellLensNames = [("cellIx","lensCellIx"),("cellData","lensCellData")]

data CellData = CellData
    { cdStyle  :: Maybe Int
    , cdValue  :: Maybe CellValue
    } deriving (Eq, Show)


cellDataLensNames :: [ (String,String)]
cellDataLensNames = [("cdStyle","lensCdStyle"),("cdValue","lensCdValue")]


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

toRows :: CellMap -> [(Int, [(Int, CellData)])]
toRows cells = 
    map extractRow $ groupBy ((==) `on` (fst . fst)) $ M.toList cells
  where
    extractRow row@(((x,_),_):_) =
        (x, map (\((_,y),v) -> (y,v)) row)
    extractRow _ = error "invalid CellMap row"

fromRows :: [(Int, [(Int, CellData)])] -> CellMap
fromRows rows = M.fromList $ concatMap mapRow rows
  where
    mapRow (r, cells) = map (\(c, v) -> ((r, c), v)) cells
