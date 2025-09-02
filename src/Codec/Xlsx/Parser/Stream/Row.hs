{-# LANGUAGE TemplateHaskell #-}
{-# LANGUAGE DerivingStrategies #-}
{-# LANGUAGE DeriveGeneric #-}
{-# LANGUAGE DeriveAnyClass #-}

module Codec.Xlsx.Parser.Stream.Row where

import Codec.Xlsx.LensCompat
import Codec.Xlsx.Types.Cell (Cell)
import Codec.Xlsx.Types.Common ( RowIndex)
import Data.IntMap.Strict (IntMap)
import GHC.Generics (Generic)
import Control.DeepSeq (NFData)

type CellRow = IntMap Cell

data Row = MkRow
  { _ri_row_index   :: !RowIndex  -- ^ Row number
  , _ri_cell_row    :: CellRow  -- ^ Row itself
  } deriving stock (Generic, Show)
    deriving anyclass NFData

makeLenses 'MkRow
