{-# LANGUAGE NoMonomorphismRestriction #-}
{-# LANGUAGE OverloadedStrings #-}
{-# LANGUAGE BangPatterns #-}
{-# Language TemplateHaskell #-}


module Codec.Xlsx(
  Xlsx(..),
  WorksheetFile(..),
  Styles(..),
  ColumnsWidth(..),
  RowHeights,
  Worksheet(..),
  CellValue_ (..),
  CellValue(..),
  Cell(..),
  CellData(..),
  int2col,
  col2int,
  foldRows,
  toList,
  fromList,
  xlsxLensNames,
  worksheetFileLensNames,
  worksheetLensNames,
  cellLensNames,
  cellDataLensNames
  
  ) where

import           Control.Arrow
import           Data.Char
import           Data.Map (Map)
import qualified Data.Map as Map
import           Data.IntMap (IntMap)
import           Data.Maybe
import           Data.Text (Text)
import qualified Data.Text as T
import           Data.Time.LocalTime (LocalTime)
import qualified Codec.Archive.Zip as Zip
import qualified Data.ByteString.Lazy as L


data Xlsx = Xlsx{ xlArchive :: Zip.Archive
                , xlSharedStrings :: IntMap Text
                , xlStyles :: Styles
                , xlWorksheetFiles :: [WorksheetFile]
                }
xlsxLensNames :: [ (String,String)]
xlsxLensNames = [  ("xlArchive"        , "lensXlArchive"       )
                , ("xlSharedStrings"  , "lensXlSharedStrings" )
                , ("xlStyles"         , "lensXlStyles"        )
                , ("xlWorksheetFiles" , "lensXlWorksheetFiles")]

            

newtype Styles = Styles {unStyles :: L.ByteString}
            deriving Show


data WorksheetFile = WorksheetFile { wfName :: Text
                                   , wfPath :: FilePath
                                   }
                   deriving Show
worksheetFileLensNames :: [ (String,String)]
worksheetFileLensNames = [("wfName","lensWfName"),("wfPath","lensWfPath")]




-- | Column range (from cwMin to cwMax) width
data ColumnsWidth = ColumnsWidth { cwMin :: Int
                                 , cwMax :: Int
                                 , cwWidth :: Double
                                 }
                  deriving Show

type RowHeights = Map Int Double

data Worksheet = Worksheet { wsName       :: Text                   -- ^ worksheet name
                           , wsMinX       :: Int                    -- ^ minimum non-empty column number (1-based)
                           , wsMaxX       :: Int                    -- ^ maximum non-empty column number (1-based)
                           , wsMinY       :: Int                    -- ^ minimum non-empty row number (1-based)
                           , wsMaxY       :: Int                    -- ^ maximum non-empty row number (1-based)
                           , wsColumns    :: [ColumnsWidth]         -- ^ column widths
                           , wsRowHeights :: RowHeights             -- ^ custom row height map
                           , wsCells      :: Map (Int,Int) CellData -- ^ data mapped by (column, row) pairs
                           }
               deriving Show
worksheetLensNames:: [ (String,String)]
worksheetLensNames =[
   ("wsName"         , "lensWsName"       )
  ,("wsMinX"         , "lensWsMinX"       )
  ,("wsMaxX"         , "lensWsMaxX"       )
  ,("wsMinY"         , "lensWsMinY"       )
  ,("wsMaxY"         , "lensWsMaxY"       )
  ,("wsColumns"      , "lensWsColumns"    )
  ,("wsRowHeights"   , "lensWsRowHeights" )
  ,("wsCells"        , "lensWsCells"      )]



  

data CellValue_ t d l = CellText {unCellText :: t} | CellDouble {unCellDouble :: d} | CellLocalTime {unCellLocalTime ::l}
               deriving Show

type CellValue = CellValue_ Text Double LocalTime

data Cell = Cell { cellIx   :: (Text, Int)
                 , cellData :: CellData
                 } 
          deriving Show

cellLensNames :: [ (String,String)]
cellLensNames = [("cellIx","lensCellIx"),("cellData","lensCellData")]

data CellData = CellData { cdStyle  :: Maybe Int
                         , cdValue  :: Maybe CellValue
                         }
              deriving Show


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

-- | fold worksheet by row, then by column, so for region A1:B2 you'll get fold order like A1, A2, B1, B2
foldRows :: (Int -> Int -> Maybe CellData -> a -> a) -> a -> Worksheet -> a
foldRows f i Worksheet{wsMinX=minX, wsMaxX=maxX,
                       wsMinY=minY, wsMaxY=maxY, wsCells=cells} = foldr foldRow i [minX..maxX]
  where
    foldRow x acc = foldr (\y -> f x y $ Map.lookup (x,y) cells) acc [minY..maxY]

toList :: Worksheet -> [[Maybe CellData]]
toList Worksheet{wsMinX=minX, wsMaxX=maxX,
                 wsMinY=minY, wsMaxY=maxY, wsCells=cells} =
  [[Map.lookup (x,y) cells | x <- [minX..maxX]] | y <- [minY..maxY]]

fromList :: Text -> [ColumnsWidth] -> RowHeights -> [[Maybe CellData]] -> Worksheet
fromList sName cw rh d = Worksheet sName 1 maxX 1 maxY cw rh $ Map.fromList cellList
  where
    maxY = max (length d + 1) (maximum' $ Map.keys rh)
    maximum' l = if null l then minBound else maximum l
    maxX = maximum' $ map length d
    filterMap f p = map f . filter p
    cellList = filterMap (second fromJust) (isJust . snd)
               $ concatMap (\(y,ds) -> map (\(x,v) -> ((x,y),v)) (zip [1..] ds))
               $ zip [1..] d
