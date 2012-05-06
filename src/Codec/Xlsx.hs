module Codec.Xlsx(
     CellValue(..),
     Cell(..)
  ) where

import Data.Text (Text)
import Data.Time.LocalTime (LocalTime)

data CellValue = CellText Text | CellDouble Double | CellLocalTime LocalTime
               deriving Show

data Cell = Cell { cellIx :: (Text, Int)
                 , cellStyle  :: Maybe Int
                 , cellValue  :: Maybe CellValue
                 }
  deriving Show
