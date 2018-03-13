{-# LANGUAGE OverloadedStrings #-}
{-# LANGUAGE RecordWildCards #-}
{-# LANGUAGE DeriveGeneric #-}
module Codec.Xlsx.Types.PivotTable.Internal
  ( CacheId(..)
  , CacheField(..)
  , CacheRecordValue(..)
  , CacheRecord
  , recordValueFromNode
  ) where

import GHC.Generics (Generic)

import Control.Arrow (first)
import Data.Maybe (catMaybes)
import Data.Text (Text)
import Text.XML
import Text.XML.Cursor

import Codec.Xlsx.Parser.Internal
import Codec.Xlsx.Types.Common
import Codec.Xlsx.Types.PivotTable
import Codec.Xlsx.Writer.Internal

newtype CacheId = CacheId Int deriving (Eq, Generic)

data CacheField = CacheField
  { cfName :: PivotFieldName
  , cfItems :: [CellValue]
  } deriving (Eq, Show, Generic)

data CacheRecordValue
  = CacheText Text
  | CacheNumber Double
  | CacheIndex Int
  deriving (Eq, Show, Generic)

type CacheRecord = [CacheRecordValue]

{-------------------------------------------------------------------------------
  Parsing
-------------------------------------------------------------------------------}

instance FromAttrVal CacheId where
  fromAttrVal = fmap (first CacheId) . fromAttrVal

instance FromCursor CacheField where
  fromCursor cur = do
    cfName <- fromAttribute "name" cur
    let cfItems =
          cur $/ element (n_ "sharedItems") &/ anyElement >=>
          cellValueFromNode . node
    return CacheField {..}

cellValueFromNode :: Node -> [CellValue]
cellValueFromNode n
  | n `nodeElNameIs` (n_ "s") = CellText <$> attributeV
  | n `nodeElNameIs` (n_ "n") = CellDouble <$> attributeV
  | otherwise = fail "no matching shared item"
  where
    cur = fromNode n
    attributeV :: FromAttrVal a => [a]
    attributeV = fromAttribute "v" cur

recordValueFromNode :: Node -> [CacheRecordValue]
recordValueFromNode n
  | n `nodeElNameIs` (n_ "s") = CacheText <$> attributeV
  | n `nodeElNameIs` (n_ "n") = CacheNumber <$> attributeV
  | n `nodeElNameIs` (n_ "x") = CacheIndex <$> attributeV
  | otherwise = fail "not valid cache record value"
  where
    cur = fromNode n
    attributeV :: FromAttrVal a => [a]
    attributeV = fromAttribute "v" cur

{-------------------------------------------------------------------------------
  Rendering
-------------------------------------------------------------------------------}

instance ToElement CacheField where
  toElement nm CacheField {..} =
    elementList nm ["name" .= cfName] [sharedItems]
    where
      -- Excel doesn't like embedded integer sharedImes in cache
      sharedItems = elementList "sharedItems" typeAttrs $
        if containsString then map cvToItem cfItems else []
      cvToItem (CellText t) = leafElement "s" ["v" .= t]
      cvToItem (CellDouble n) = leafElement "n" ["v" .= n]
      cvToItem _ = error "Only string and number values are currently supported"
      typeAttrs =
        catMaybes
          [ "containsNumber" .=? justTrue containsNumber
          , "containsString" .=? justFalse containsString
          , "containsSemiMixedTypes" .=? justFalse containsString
          , "containsMixedTypes" .=? justTrue (containsNumber && containsString)
          ]
      containsNumber = any isNumber cfItems
      isNumber (CellDouble _) = True
      isNumber _ = False
      containsString = any isString cfItems
      isString (CellText _) = True
      isString _ = False
