module Codec.Xlsx.Types.PivotTable.Internal (
   CacheId(..)
) where

import Control.Arrow (first)

import Codec.Xlsx.Parser.Internal

newtype CacheId = CacheId Int deriving Eq

{-------------------------------------------------------------------------------
  Parsing
-------------------------------------------------------------------------------}

instance FromAttrVal CacheId where
  fromAttrVal = fmap (first CacheId) . fromAttrVal
