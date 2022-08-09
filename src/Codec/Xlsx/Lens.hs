{-# LANGUAGE RankNTypes   #-}
{-# LANGUAGE TypeFamilies #-}
{-# LANGUAGE DeriveGeneric #-}
{-# LANGUAGE ViewPatterns #-}
{-# LANGUAGE TupleSections #-}

-- | lenses to access sheets, cells and values of 'Xlsx'
module Codec.Xlsx.Lens
  ( ixSheet
  , ixSheet'
  , atSheet
  , atSheet'
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
import Control.Applicative ((<|>))
import Control.Lens
import Control.Monad.Zip (mzip)
import Data.Function (on)
import Data.List as L (deleteBy, find)
import Data.Text
import Data.Tuple (swap)
import GHC.Generics (Generic)
import Codec.Xlsx.Types.SheetState as SheetState (SheetState(..))
import Prelude hiding (lookup)

newtype SheetList = SheetList{ unSheetList :: [(Text, SheetState, Worksheet)] }
    deriving (Eq, Show, Generic)

type instance IxValue (SheetList) = (SheetState, Worksheet)
type instance Index (SheetList) = Text

instance Ixed SheetList where
    ix k f sl@(SheetList l) = case lookup k l of
        Just v2  -> f v2 <&> \v2' -> SheetList (upsert k v2' l)
        Nothing -> pure sl
    {-# INLINE ix #-}

instance At SheetList where
  at k f (SheetList l) = f mv2 <&> \r -> case r of
      Nothing -> SheetList $ maybe l (\v2 -> deleteBy ((==) `on` view _1) (t2to3 k v2) l) mv2
      Just v2' -> SheetList $ upsert k v2' l
    where
      mv2 = lookup k l
  {-# INLINE at #-}

t2to3 :: k -> (v, v') -> (k, v, v')
t2to3 k = uncurry (k,,)
{-# INLINE t2to3 #-}

t3to2 :: (k, v, v') -> (v, v')
t3to2 (_, v, v') = (v, v')
{-# INLINE t3to2 #-}

lookup :: (Eq k) => k -> [(k, v, v')] -> Maybe (v, v')
lookup k l = L.find (views _1 (== k)) l <&> t3to2

upsert :: (Eq k) => k -> (v, v') -> [(k, v, v')] -> [(k, v, v')]
upsert k v2 [] = [t2to3 k v2]
upsert k v2 (t3@(view _1 -> k'):r) =
    if k == k'
    then t2to3 k v2:r
    else t3:upsert k v2 r

sheetList :: Iso' [(Text, SheetState, Worksheet)] SheetList
sheetList = iso SheetList unSheetList

-- | lens giving access to a worksheet from 'Xlsx' object
-- by its name
ixSheet :: Text -> Traversal' Xlsx Worksheet
ixSheet s = ixSheet' s . _2

-- | Lens to the worksheet and its visibility state.
ixSheet' :: Text -> Traversal' Xlsx (SheetState, Worksheet)
ixSheet' s = xlSheets . sheetList . ix s

-- | 'Control.Lens.At' variant of 'ixSheet' lens
--
-- /Note:/ if there is no such sheet in this workbook then new sheet will be
-- added as the last one to the sheet list; set Nothing to delete
--
-- > sans "some sheet" xlsx ≡ at "some sheet" .~ Nothing xlsx {- delete -}
-- > xlsx & at "some sheet" ?~ worksheet {- upsert -}
atSheet :: Text -> Lens' Xlsx (Maybe Worksheet)
atSheet s = lens viewSheet setSheet
  where
    viewSheet :: Xlsx -> Maybe Worksheet
    viewSheet = views (atSheet' s) (fmap snd)
    setSheet :: Xlsx -> Maybe Worksheet -> Xlsx
    setSheet xlsx ws =
      let state = views (atSheet' s) (fmap fst) xlsx <|> Just SheetState.Visible
          pair = state `mzip` ws
          in set (atSheet' s) pair xlsx

-- | lens to the worksheet and its visibility state
atSheet' :: Text -> Lens' Xlsx (Maybe (SheetState, Worksheet))
atSheet' s = xlSheets . sheetList . at s

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
