{-# LANGUAGE CPP   #-}
{-# LANGUAGE RankNTypes   #-}
{-# LANGUAGE TypeFamilies #-}
{-# LANGUAGE DeriveGeneric #-}

-- | lenses to access sheets, cells and values of 'Xlsx'
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
  , atRow
  , ixRow
  ) where

import Codec.Xlsx.Types
import Codec.Xlsx.Types.Cell
#ifdef USE_MICROLENS
import Lens.Micro
import Lens.Micro.Internal
import Lens.Micro.GHC ()
#else
import Control.Lens
#endif
import Data.Function (on)
import Data.List (deleteBy)
import Data.Text
import Data.Tuple (swap)
import GHC.Generics (Generic)

newtype SheetList = SheetList{ unSheetList :: [(Text, Worksheet)] }
    deriving (Eq, Show, Generic)

type instance IxValue (SheetList) = Worksheet
type instance Index (SheetList) = Text

instance Ixed SheetList where
    ix k f sl@(SheetList l) = case lookup k l of
        Just v  -> f v <&> \v' -> SheetList (upsert k v' l)
        Nothing -> pure sl
    {-# INLINE ix #-}

instance At SheetList where
  at k f (SheetList l) = f mv <&> \r -> case r of
      Nothing -> SheetList $ maybe l (\v -> deleteBy ((==) `on` fst) (k,v) l) mv
      Just v' -> SheetList $ upsert k v' l
    where
      mv = lookup k l
  {-# INLINE at #-}

upsert :: (Eq k) => k -> v -> [(k,v)] -> [(k,v)]
upsert k v [] = [(k,v)]
upsert k v ((k1,v1):r) =
    if k == k1
    then (k,v):r
    else (k1,v1):upsert k v r

-- | lens giving access to a worksheet from 'Xlsx' object
-- by its name
ixSheet :: Text -> Traversal' Xlsx Worksheet
ixSheet s = xlSheets . \f -> fmap unSheetList . ix s f . SheetList

-- | 'Control.Lens.At' variant of 'ixSheet' lens
--
-- /Note:/ if there is no such sheet in this workbook then new sheet will be
-- added as the last one to the sheet list
atSheet :: Text -> Lens' Xlsx (Maybe Worksheet)
atSheet s = xlSheets . \f -> fmap unSheetList . at s f . SheetList

-- | lens giving access to a cell in some worksheet
-- by its position, by default row+column index is used
-- so this lens is a synonym of 'ixCellRC'
ixCell :: (Int, Int) -> Traversal' Worksheet Cell
ixCell = ixCellRC

-- | lens to access cell in a worksheet
ixCellRC :: (Int, Int) -> Traversal' Worksheet Cell
ixCellRC (y, x) = wsCells . ix y . ix x

-- | lens to access cell in a worksheet using more traditional
-- x+y coordinates
ixCellXY :: (Int, Int) -> Traversal' Worksheet Cell
ixCellXY i = ixCellRC $ swap i

-- | accessor that can read, write or delete cell in a worksheet
-- synonym of 'atCellRC' so uses row+column index
atCell :: (Int, Int) -> Lens' Worksheet (Maybe Cell)
atCell = atCellRC

atRow :: Int -> Lens' Worksheet (Maybe CellRow)
atRow y = wsCells . at y

ixRow :: Int -> Traversal' Worksheet CellRow
ixRow y = wsCells . ix y

-- | lens to read, write or delete cell in a worksheet
atCellRC :: (Int, Int) -> Lens' Worksheet (Maybe Cell)
atCellRC (y,x) = atRow y . non mempty . at x

-- | lens to read, write or delete cell in a worksheet
-- using more traditional x+y or row+column index
atCellXY :: (Int, Int) -> Lens' Worksheet (Maybe Cell)
atCellXY i = atCellRC $ swap i

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
cellValueAtXY i = cellValueAtRC $ swap i
