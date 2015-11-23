{-# LANGUAGE RankNTypes #-}
module Codec.Xlsx.Lens
    ( ixSheet
    , atSheet
    , ixCell
    , ixCellRC
    , ixCellXY
    , atCell
    , atCellRC
    , atCellXY
    , cellValueAt
    , cellValueAtRC
    , cellValueAtXY
 ) where

import Codec.Xlsx.Types
import Control.Lens
import Data.Text
import Data.Tuple (swap)

ixSheet :: Text -> Traversal' Xlsx Worksheet
ixSheet s = xlSheets . ix s

atSheet :: Text -> Lens' Xlsx (Maybe Worksheet)
atSheet s = xlSheets . at s

ixCell :: (Int, Int) -> Traversal' Worksheet Cell
ixCell = ixCellRC

ixCellRC :: (Int, Int) -> Traversal' Worksheet Cell
ixCellRC i = wsCells . ix i

ixCellXY :: (Int, Int) -> Traversal' Worksheet Cell
ixCellXY = ixCellRC . swap

atCell :: (Int, Int) -> Lens' Worksheet (Maybe Cell)
atCell = atCellRC

atCellRC :: (Int, Int) -> Lens' Worksheet (Maybe Cell)
atCellRC i = wsCells . at i

atCellXY :: (Int, Int) -> Lens' Worksheet (Maybe Cell)
atCellXY = atCellRC . swap

cellValueAt :: (Int, Int) -> Lens' Worksheet (Maybe CellValue)
cellValueAt = cellValueAtRC

cellValueAtRC :: (Int, Int) -> Lens' Worksheet (Maybe CellValue)
cellValueAtRC i = atCell i . non def . cellValue

cellValueAtXY :: (Int, Int) -> Lens' Worksheet (Maybe CellValue)
cellValueAtXY = cellValueAtRC . swap
