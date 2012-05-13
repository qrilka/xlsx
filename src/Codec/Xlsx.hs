module Codec.Xlsx(
     CellValue(..),
     Cell(..),
     CellData(..),
     cell2cd,
     int2col,
     col2Int
     ) where

import           Data.Char
import           Data.List
import           Data.Maybe
import           Data.Text (Text)
import qualified Data.Text as T
import           Data.Time.LocalTime (LocalTime)


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

col2Int :: Text -> Int
col2Int t = T.foldl' (\i c -> i * 26 + let2int c) 0 t
    where 
        let2int c = 1 + (ord c) - (ord 'A')
