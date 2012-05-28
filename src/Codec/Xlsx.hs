module Codec.Xlsx(
  ColumnsWidth(..),
  RowHeights(..),
  Worksheet(..),
  CellValue(..),
  Cell(..),
  CellData(..),
  cell2cd,
  int2col,
  col2int,
  foldRows,
  toList,
  fromList
  ) where

import           Data.Char
import           Data.List
import           Data.Map (Map)
import qualified Data.Map as Map
import           Data.Maybe
import           Data.Text (Text)
import qualified Data.Text as T
import           Data.Time.LocalTime (LocalTime)


data ColumnsWidth = ColumnsWidth { cwMin :: Int
                                 , cwMax :: Int
                                 , cwWidth :: Double
                                 }
                  deriving Show

type RowHeights = Map Int Double

data Worksheet = Worksheet { wsName       :: Text
                           , wsMinX       :: Int
                           , wsMaxX       :: Int
                           , wsMinY       :: Int
                           , wsMaxY       :: Int
                           , wsColumns    :: [ColumnsWidth]
                           , wsRowHeights :: RowHeights
                           , wsCells      :: Map (Int,Int) CellData
                           }
               deriving Show

data CellValue = CellText Text | CellDouble Double | CellLocalTime LocalTime
               deriving Show

data Cell = Cell { cellIx :: (Text, Int)
                 , cellStyle  :: Maybe Int
                 , cellValue  :: Maybe CellValue
                 }
          deriving Show

data CellData = CellData { cdStyle  :: Maybe Int
                         , cdValue  :: Maybe CellValue
                         }
              deriving Show

cell2cd Cell{cellStyle=s,cellValue=v} = CellData{cdStyle=s, cdValue=v}

int2col :: Int -> Text
int2col = T.pack . reverse . map int2let . base26
    where 
        int2let 0 = 'Z'
        int2let x = chr $ (x - 1) + ord 'A'
        base26  0 = []
        base26  i = let i' = (i `mod` 26) 
                        i'' = if i' == 0 then 26 else i'
                    in seq i' (i' : base26 ((i - i'') `div` 26))

col2int :: Text -> Int
col2int t = T.foldl' (\i c -> i * 26 + let2int c) 0 t
    where 
        let2int c = 1 + (ord c) - (ord 'A')

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
fromList sName cw rh d = Worksheet sName 1 maxX 1 maxY cw rh cells
  where
    maxY = max (length d + 1) (maximum' $ Map.keys rh)
    maximum' l = if null l then minBound else maximum l
    (maxX, cells) = foldr foldRow (1, Map.empty) (zip [1..] d)
    foldRow (y, row) (maxX, cells) = (max maxX (length row), foldr (foldCell y) cells (zip [1..] row))
    foldCell y (x, Nothing) acc = acc
    foldCell y (x, (Just cd)) m = Map.insert (x,y) cd m
