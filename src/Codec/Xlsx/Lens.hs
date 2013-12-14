{-# LANGUAGE NoMonomorphismRestriction #-}
{-# LANGUAGE OverloadedStrings #-}
{-# LANGUAGE BangPatterns #-}
{-# LANGUAGE Rank2Types #-}
{-# LANGUAGE FlexibleContexts #-}
{-# Language TemplateHaskell #-}

module Codec.Xlsx.Lens where

import Control.Lens
import Control.Applicative (Applicative)
import qualified Data.Foldable as F

import Codec.Xlsx.Types


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
lensSheetCell :: (Functor f, Indexable (Int, Int) p)
              => (Int, Int)
              -> p (Maybe CellData) (f (Maybe CellData))
              -> Worksheet
              -> f Worksheet
lensSheetCell i = lensWsCells . at i

traverseSheetCellData :: Applicative f
                      => (Int, Int)
                      -> (Maybe CellValue -> f (Maybe CellValue))
                      -> Worksheet
                      -> f Worksheet
traverseSheetCellData i = (lensSheetCell i) . _Just . lensCdValue

lensSheetMap :: (Functor f, Indexable Int p)
             => Int
             -> p (Maybe Worksheet) (f (Maybe Worksheet))
             -> MappedSheet
             -> f MappedSheet
lensSheetMap i = lensMappedSheet. at i

traverseMappedSheetCellData :: Applicative f
                            => FullyIndexedCellValue
                            -> (Maybe CellValue -> f (Maybe CellValue))
                            -> MappedSheet
                            -> f MappedSheet
traverseMappedSheetCellData (FICV sheetIndex rowIndex colIndex _) = lensSheetMap sheetIndex . _Just .
                                                                  traverseSheetCellData (rowIndex, colIndex)


setMappedSheetCellData :: MappedSheet -> FullyIndexedCellValue -> MappedSheet
setMappedSheetCellData ms  i = ms & (traverseMappedSheetCellData i) ?~ (ficvValue i)

setMultiMappedSheetCellData :: (F.Foldable f ) => MappedSheet -> f FullyIndexedCellValue -> MappedSheet
setMultiMappedSheetCellData ms fFICLV = F.foldl' (\a b -> setMappedSheetCellData a b) ms fFICLV

