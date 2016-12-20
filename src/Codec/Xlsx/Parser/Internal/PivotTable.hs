{-# LANGUAGE OverloadedStrings #-}
{-# LANGUAGE RecordWildCards #-}
module Codec.Xlsx.Parser.Internal.PivotTable (
   parsePivotTable
  , parseCache
) where

import Control.Applicative
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
  :: (CacheId -> Maybe (Text, Range, [PivotFieldName]))
  -> ByteString
  -> Maybe PivotTable
parsePivotTable srcByCacheId bs =
  listToMaybe . parse . fromDocument $ parseLBS_ def bs
  where
    parse cur = do
      cacheId <- fromAttribute "cacheId" cur
      case srcByCacheId cacheId of
        Nothing -> fail "no such cache"
        Just (_pvtSrcSheet, _pvtSrcRef, fieldNames) -> do
          _pvtDataCaption <- attribute "dataCaption" cur
          _pvtName <- attribute "name" cur
          _pvtLocation <- cur $/ element (n_ "location") >=> fromAttribute "ref"
          _pvtRowGrandTotals <- fromAttributeDef "rowGrandTotals" True cur
          _pvtColumnGrandTotals <- fromAttributeDef "colGrandTotals" True cur
          _pvtOutline <- fromAttributeDef "outline" False cur
          _pvtOutlineData <- fromAttributeDef "outlineData" False cur
          let nToField = zip [0 ..] fieldNames
              outlines =
                cur $/ element (n_ "pivotFields") &/ element (n_ "pivotField") >=>
                fromAttributeDef "outline" True
              _pvtFields =
                [ PivotFieldInfo {_pfiName = name, _pfiOutline = outline}
                | (name, outline) <- zip fieldNames outlines
                ]
              _pvtRowFields =
                cur $/ element (n_ "rowFields") &/ element (n_ "field") >=>
                fromAttribute "x" >=> fieldPosition
              _pvtColumnFields =
                cur $/ element (n_ "colFields") &/ element (n_ "field") >=>
                fromAttribute "x" >=> fieldPosition
              _pvtDataFields =
                cur $/ element (n_ "dataFields") &/ element (n_ "dataField") >=> \c -> do
                  fld <- fromAttribute "fld" c
                  _dfField <- maybeToList $ lookup fld nToField
                  -- TOFIX
                  _dfName <- fromAttributeDef "name" "" c
                  _dfFunction <- fromAttributeDef "subtotal" ConsolidateSum c
                  return DataField {..}
              fieldPosition :: Int -> [PositionedField]
              fieldPosition (-2) = return DataPosition
              fieldPosition n =
                FieldPosition <$> maybeToList (lookup n nToField)
          return PivotTable {..}

parseCache :: ByteString -> Maybe (Text, CellRef, [PivotFieldName])
parseCache bs = listToMaybe . parse . fromDocument $ parseLBS_ def bs
  where
    parse cur = do
      (sheet, ref) <-
        cur $/ element (n_ "cacheSource") &/ element (n_ "worksheetSource") >=>
        liftA2 (,) <$> attribute "sheet" <*> fromAttribute "ref"
      let fieldNames =
            cur $/ element (n_ "cacheFields") &/ element (n_ "cacheField") >=>
            fromAttribute "name"
      return (sheet, ref, fieldNames)
