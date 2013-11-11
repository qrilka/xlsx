{-# LANGUAGE NoMonomorphismRestriction #-}
{-# LANGUAGE OverloadedStrings #-}
{-# LANGUAGE BangPatterns #-}
{-# LANGUAGE Rank2Types #-}
{-# LANGUAGE FlexibleContexts #-}
{-# Language TemplateHaskell #-}

module Codec.Xlsx.Lens where

import Control.Lens
import Control.Lens.Lens
import Control.Lens.Prism
import Control.Lens.TH
import qualified Data.Foldable as F
import Codec.Xlsx
import Codec.Xlsx.Parser
import Codec.Xlsx.Writer
import Data.Text


-- |Lens into the worksheet file stuff


{-|
xlsxLensNames :: [ (String,String)]
xlsxLensNames = [  ("xlArchive"        , "lensXlArchive"       )
                , ("xlSharedStrings"  , "lensXlSharedStrings" )
                , ("xlStyles"         , "lensXlStyles"        )
                , ("xlWorksheetFiles" , "lensXlWorksheetFiles")]


|-}
makeLensesFor xlsxLensNames ''Xlsx

{-|
worksheetFileLensNames :: [ (String,String)]
worksheetFileLensNames = [("wfName","lensWfName"),("wfPath","lensWfPath")]
|-}
makeLensesFor worksheetFileLensNames ''WorksheetFile

{-|
worksheetLensNames:: [ (String,String)]
worksheetLensNames =[
   ("wsName"         , "lensWsName"       )
  ,("wsMinX"         , "lensWsMinX"       )
  ,("wsMaxX"         , "lensWsMaxX"       )
  ,("wsMinY"         , "lensWsMinY"       )
  ,("wsMaxY"         , "lensWsMaxY"       )
  ,("wsColumns"      , "lensWsColumns"    )
  ,("wsRowHeights"   , "lensWsRowHeights" )
  ,("wsCells"        , "lensWsCells"      )]

|-}

makeLensesFor worksheetLensNames ''Worksheet

{-|
cellDataLensNames :: [ (String,String)]
cellDataLensNames = [("cdStyle","lensCdStyle"),("cdValue","lensCdValue")]
|-}
makeLensesFor cellDataLensNames ''CellData

{-|
cellLensNames :: [ (String,String)]
cellLensNames = [("cellIx","lensCellIx"),("cellData","lensCellData")]
|-}

makeLensesFor cellLensNames ''Cell


{-|
mappedSheetLensNames :: [(String,String)]
mappedSheetLensNames :: [("unMappedSheet","lensMappedSheet")]
|-}

makeLensesFor mappedSheetLensNames  ''MappedSheet

{-

 -- bar :: Lens' (Foo a) Int
 -- baz :: Lens (Foo a) (Foo b) a b
 quux :: Functor f => (a -> f b) -> Foo a -> f (Foo b)
 quux f (Foo a b c) = fmap (Foo a b) (f c)
-}


-- |Lens Compositions
-- p is the ~ (->) applier that is stored by declaring something Indexable
-- in this case it takes (Maybe CellData) to (f (Maybe CellData))
-- lensSheetCell :: (Indexable (Int,Int) p) => (Int,Int) -> IndexedLens (Int,Int) Worksheet Worksheet (Maybe CellData) (Maybe CellData)
lensSheetCell i = lensWsCells . at i

traverseSheetCellData i = (lensSheetCell i)._Just. lensCdValue

lensSheetMap i = lensMappedSheet. at i


traverseMappedSheetCellData (FICV sheetIndex rowIndex colIndex _) = lensSheetMap sheetIndex . _Just .
                                                                  traverseSheetCellData (rowIndex,colIndex)


setMappedSheetCellData :: MappedSheet ->FullyIndexedCellValue -> MappedSheet
setMappedSheetCellData ms  i = ms & (traverseMappedSheetCellData i) ?~ (ficvValue i)

setMultiMappedSheetCellData :: (F.Foldable f ) => MappedSheet -> f FullyIndexedCellValue -> MappedSheet
setMultiMappedSheetCellData ms fFICLV = F.foldl' (\a b -> setMappedSheetCellData a b) ms fFICLV

-- Lens s s (m a) (m a)
-- Lens a a b b
-- Lens (m a) (m a) a a

-- Lens CellData CellData CellValue CellValue
-- Lens (Maybe CellData) (Maybe CellData) (CellData) (CellData)
-- lensCdValue :: Functor f =>  (Maybe CellValue -> f (Maybe CellValue)) -> CellData -> f CellData


-- type IndexedLens i s t a b = forall f p. (Indexable i p, Functor f) => p a (f b) -> s -> f t
-- type IndexedLens' i s s a a = forall f p. (Indexable i p, Functor f) => p a (f a) -> s -> f s
-- type Lens s t a b = forall f. Functor f => (a -> f b) -> s -> f t
-- type Lens' s a = Lens s s a a

--lensSheetCellDataValue i = (lensSheetCell i) . lensCdValue

-- Lens'


--  IndexedLens' i s a = IndexedLens i s s a a
-- lensSheetCell
--   :: (Functor f,
--       Control.Lens.Internal.Indexed.Indexable (Int, Int) p) =>
--      (Int, Int)                                
--      -> p (Maybe Codec.Xlsx.CellData) (f (Maybe Codec.Xlsx.CellData))
--      -> Codec.Xlsx.Worksheet  
--      -> f Codec.Xlsx.Worksheet

-- lensCdValue
--   :: Functor f =>
--      (Maybe Codec.Xlsx.CellValue -> f (Maybe Codec.Xlsx.CellValue))
--      -> Codec.Xlsx.CellData -> f Codec.Xlsx.CellData


-- at :: Index m -> IndexedLens' (Index m) m (Maybe (IxValue m))
-- 



