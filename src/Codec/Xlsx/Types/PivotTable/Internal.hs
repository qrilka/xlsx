{-# LANGUAGE CPP #-}
{-# LANGUAGE OverloadedStrings #-}
{-# LANGUAGE RecordWildCards #-}
{-# LANGUAGE DeriveGeneric #-}
module Codec.Xlsx.Types.PivotTable.Internal
  ( CacheId(..)
  , CacheField(..)
  ) where

import Control.Arrow (first)
import Data.Maybe (catMaybes)
import Text.XML
import Text.XML.Cursor

#if !MIN_VERSION_base(4,8,0)
import Control.Applicative
#endif

import Codec.Xlsx.Parser.Internal
import Codec.Xlsx.Types.Common
import Codec.Xlsx.Types.PivotTable
import Codec.Xlsx.Writer.Internal

newtype CacheId = CacheId Int deriving (Eq, Generic)

data CacheField = CacheField
  { cfName :: PivotFieldName
  , cfItems :: [CellValue]
  } deriving (Eq, Show, Generic)

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

{-------------------------------------------------------------------------------
  Rendering
-------------------------------------------------------------------------------}

instance ToElement CacheField where
  toElement nm CacheField {..} =
    elementList nm ["name" .= cfName] [sharedItems]
    where
      sharedItems = elementList "sharedItems" typeAttrs $ map cvToItem cfItems
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
