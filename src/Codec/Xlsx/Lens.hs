-- | lenses to access sheets, cells and values of 'Xlsx'
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

-- | lens giving access to a worksheet from 'Xlsx' object
-- by its name
ixSheet :: Text -> Traversal' Xlsx Worksheet
ixSheet s = xlSheets . ix s

-- | 'Control.Lens.At' variant of 'ixSheet' lens
atSheet :: Text -> Lens' Xlsx (Maybe Worksheet)
atSheet s = xlSheets . at s

-- | lens giving access to a cell in some worksheet
-- by its position, by default row+column index is used
-- so this lens is a synonym of 'ixCellRC'
ixCell :: (Int, Int) -> Traversal' Worksheet Cell
ixCell = ixCellRC

-- | lens to access cell in a worksheet
ixCellRC :: (Int, Int) -> Traversal' Worksheet Cell
ixCellRC i = wsCells . ix i

-- | lens to access cell in a worksheet using more traditional
-- x+y coordinates
ixCellXY :: (Int, Int) -> Traversal' Worksheet Cell
ixCellXY = ixCellRC . swap

-- | accessor that can read, write or delete cell in a worksheet
-- synonym of 'atCellRC' so uses row+column index
atCell :: (Int, Int) -> Lens' Worksheet (Maybe Cell)
atCell = atCellRC

-- | lens to read, write or delete cell in a worksheet
atCellRC :: (Int, Int) -> Lens' Worksheet (Maybe Cell)
atCellRC i = wsCells . at i

-- | lens to read, write or delete cell in a worksheet
-- using more traditional x+y or row+column index
atCellXY :: (Int, Int) -> Lens' Worksheet (Maybe Cell)
atCellXY = atCellRC . swap

-- | lens to read, write or delete cell value in a worksheet
-- with row+column coordinates, synonym for 'cellValueRC'
cellValueAt :: (Int, Int) -> Lens' Worksheet (Maybe CellValue)
cellValueAt = cellValueAtRC

-- | lens to read, write or delete cell value in a worksheet
-- using row+column coordinates of that cell
cellValueAtRC :: (Int, Int) -> Lens' Worksheet (Maybe CellValue)
cellValueAtRC i = atCell i . non def . cellValue

-- | lens to read, write or delete cell value in a worksheet
-- using traditional x+y coordinates
cellValueAtXY :: (Int, Int) -> Lens' Worksheet (Maybe CellValue)
cellValueAtXY = cellValueAtRC . swap
