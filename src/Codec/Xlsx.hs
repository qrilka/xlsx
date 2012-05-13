module Codec.Xlsx(
     CellValue(..),
     Cell(..),
     CellData(..),
     cell2cd,
     xlsxCol2int,
     int2xlsxCol,
     int2Col,
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
                         , cdValue  :: CellValue
                         }
              deriving Show

cell2cd Cell{cellStyle=s,cellValue=v} = CellData{cdStyle=s, cdValue=fromJust v}

xlsxCol2int :: Text -> Int
xlsxCol2int t = prev + cur
  where
    prev = foldl' (\x y -> x * 26 + y) 0 $! replicate (T.length t - 1) 1 ++ [0]
    cur = T.foldl' (\a ch -> a * 26 + letter2int ch) 0 t
    letter2int x = ord x - ord 'A'

int2xlsxCol :: Int -> Text
int2xlsxCol i = alphabet n (i - prefixes)
  where
    alphabet n = T.pack . map (\x -> chr $ x + ord 'A') . padZeroes n . base26
    padZeroes n l = replicate (n - length l) 0 ++ l
    base26 = reverse . base26'
    base26' 0 = []
    base26' x = (x `mod` 26) : base26' (x `div` 26)
    (n, prefixes) = if i < 26 then (1, 0) else inner 2 26 26
      where inner t s p | s + p * 26 < i = inner (t+1) (s + p * 26) (p * 26)
                        | otherwise = (t,s)

int2Col :: Int -> Text
int2Col = T.pack . reverse . map int2let . base26
    where 
        int2let 0 = 'Z'
        int2let x = chr $ (x - 1) + ord 'A'
        base26  0 = []
        base26  i = let i' = (i `mod` 26) 
                        i'' = if i'==0 then 26 else i'
                    in seq i' (i': base26 ((i-i'') `div` 26))

col2Int :: Text -> Int
col2Int t = T.foldl' (\i c -> i*26+let2int c) 0 t
    where 
        let2int c = 1 + (ord c) - (ord 'A')
