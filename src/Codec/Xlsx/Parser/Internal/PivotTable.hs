{-# LANGUAGE OverloadedStrings #-}
{-# LANGUAGE RecordWildCards #-}
module Codec.Xlsx.Parser.Internal.PivotTable
  ( parsePivotTable
  , parseCache
  ) where

import Control.Applicative
import Control.Arrow ((&&&))
import Data.ByteString.Lazy (ByteString)
import Data.Maybe (listToMaybe, maybeToList)
import Data.Text (Text)
import Text.XML
import Text.XML.Cursor

import Codec.Xlsx.Parser.Internal
import Codec.Xlsx.Types.Common
import Codec.Xlsx.Types.PivotTable
import Codec.Xlsx.Types.PivotTable.Internal

parsePivotTable
  :: (CacheId -> Maybe (Text, Range, [CacheField]))
  -> ByteString
  -> Maybe PivotTable
parsePivotTable srcByCacheId bs =
  listToMaybe . parse . fromDocument $ parseLBS_ def bs
  where
    parse cur = do
      cacheId <- fromAttribute "cacheId" cur
      case srcByCacheId cacheId of
        Nothing -> fail "no such cache"
        Just (_pvtSrcSheet, _pvtSrcRef, cacheFields) -> do
          _pvtDataCaption <- attribute "dataCaption" cur
          _pvtName <- attribute "name" cur
          _pvtLocation <- cur $/ element (n_ "location") >=> fromAttribute "ref"
          _pvtRowGrandTotals <- fromAttributeDef "rowGrandTotals" True cur
          _pvtColumnGrandTotals <- fromAttributeDef "colGrandTotals" True cur
          _pvtOutline <- fromAttributeDef "outline" False cur
          _pvtOutlineData <- fromAttributeDef "outlineData" False cur
          let name2CacheField = map (cfName &&& id) cacheFields
              _pvtFields =
                cur $/ element (n_ "pivotFields") &/ element (n_ "pivotField") >=> \c -> do
                  _pfiName <- fromAttribute "name" c
                  _pfiSortType <- fromAttributeDef "sortType" FieldSortManual c
                  _pfiOutline <- fromAttributeDef "outline" True c
                  let hidden =
                        c $/ element (n_ "items") &/ element (n_ "item") >=>
                        attrValIs "h" True >=> fromAttribute "x"
                      items = maybe [] cfItems $ lookup _pfiName name2CacheField
                      _pfiHiddenItems =
                        [item | (n, item) <- zip [(0 :: Int) ..] items, n `elem` hidden]
                  return PivotFieldInfo {..}
              nToFieldName = zip [0 ..] $ map _pfiName _pvtFields
              fieldNameList fld = maybeToList $ lookup fld nToFieldName
              _pvtRowFields =
                cur $/ element (n_ "rowFields") &/ element (n_ "field") >=>
                fromAttribute "x" >=> fieldPosition
              _pvtColumnFields =
                cur $/ element (n_ "colFields") &/ element (n_ "field") >=>
                fromAttribute "x" >=> fieldPosition
              _pvtDataFields =
                cur $/ element (n_ "dataFields") &/ element (n_ "dataField") >=> \c -> do
                  fld <- fromAttribute "fld" c
                  _dfField <- fieldNameList fld
                  -- TOFIX
                  _dfName <- fromAttributeDef "name" "" c
                  _dfFunction <- fromAttributeDef "subtotal" ConsolidateSum c
                  return DataField {..}
              fieldPosition :: Int -> [PositionedField]
              fieldPosition (-2) = return DataPosition
              fieldPosition n =
                FieldPosition <$> fieldNameList n
          return PivotTable {..}

parseCache :: ByteString -> Maybe (Text, CellRef, [CacheField])
parseCache bs = listToMaybe . parse . fromDocument $ parseLBS_ def bs
  where
    parse cur = do
      (sheet, ref) <-
        cur $/ element (n_ "cacheSource") &/ element (n_ "worksheetSource") >=>
        liftA2 (,) <$> attribute "sheet" <*> fromAttribute "ref"
      let fields =
            cur $/ element (n_ "cacheFields") &/ element (n_ "cacheField") >=>
            fromCursor
      return (sheet, ref, fields)
