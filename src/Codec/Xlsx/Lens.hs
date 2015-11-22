{-# LANGUAGE RankNTypes #-}
module Codec.Xlsx.Lens
    ( ixSheet
    , atSheet
    , ixCell
    , atCell
    , cellValueAt
 ) where

import Codec.Xlsx.Types
import Control.Lens
import Data.Text

ixSheet :: Text -> Traversal' Xlsx Worksheet
ixSheet s = xlSheets . ix s

atSheet :: Text -> Lens' Xlsx (Maybe Worksheet)
atSheet s = xlSheets . at s

ixCell :: (Int, Int) -> Traversal' Worksheet Cell
ixCell i = wsCells . ix i

atCell :: (Int, Int) -> Lens' Worksheet (Maybe Cell)
atCell i = wsCells . at i

cellValueAt :: (Int, Int) -> Lens' Worksheet (Maybe CellValue)
cellValueAt i = atCell i . non def . cellValue
