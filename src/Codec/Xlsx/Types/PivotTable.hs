{-# LANGUAGE OverloadedStrings #-}
{-# LANGUAGE RecordWildCards #-}
module Codec.Xlsx.Types.PivotTable
  ( PivotTable(..)
  , PivotFieldName(..)
  , PivotFieldInfo(..)
  , PositionedField(..)
  , DataField(..)
  , ConsolidateFunction(..)
  ) where

import Control.Applicative
import Control.Arrow (first)
import Data.ByteString.Lazy (ByteString)
import qualified Data.Map as M
import Data.Maybe (catMaybes, listToMaybe, maybeToList)
import Data.Text (Text)
import Safe (fromJustNote)
import Text.XML
import Text.XML.Cursor

import Codec.Xlsx.Types.Common
import Codec.Xlsx.Parser.Internal
import Codec.Xlsx.Writer.Internal

data PivotTable = PivotTable
  { _pvtName :: Text
  , _pvtDataCaption :: Text
  , _pvtRowFields :: [PositionedField]
  , _pvtColumnFields :: [PositionedField]
  , _pvtDataFields :: [DataField]
  , _pvtFields :: [PivotFieldInfo]
  , _pvtRowGrandTotals :: Bool
  , _pvtColumnGrandTotals :: Bool
  , _pvtOutline :: Bool
  , _pvtOutlineData :: Bool
  , _pvtLocation :: CellRef
  , _pvtSrcSheet :: Text
  , _pvtSrcRef :: Range
  } deriving (Eq, Show)

data PivotFieldInfo = PivotFieldInfo
  { _pfiName :: PivotFieldName
  , _pfiOutline :: Bool
  } deriving (Eq, Show)

newtype PivotFieldName =
  PivotFieldName Text
  deriving (Eq, Ord, Show)

data PositionedField
  = DataPosition
  | FieldPosition PivotFieldName
  deriving (Eq, Ord, Show)

data DataField = DataField
  { _dfField :: PivotFieldName
  , _dfName :: Text
  , _dfFunction :: ConsolidateFunction
  } deriving (Eq, Show)

-- | Data consolidation functions specified by the user and used to
-- consolidate ranges of data
--
-- See 18.18.17 "ST_DataConsolidateFunction (Data Consolidation
-- Functions)" (p.  2447)
data ConsolidateFunction
  = ConsolidateAverage
    -- ^ The average of the values.
  | ConsolidateCount
    -- ^ The number of data values. The Count consolidation function
    -- works the same as the COUNTA worksheet function.
  | ConsolidateCountNums
    -- ^ The number of data values that are numbers. The Count Nums
    -- consolidation function works the same as the COUNT worksheet
    -- function.
  | ConsolidateMaximum
    -- ^ The largest value.
  | ConsolidateMinimum
    -- ^ The smallest value.
  | ConsolidateProduct
    -- ^ The product of the values.
  | ConsolidateStdDev
    -- ^ An estimate of the standard deviation of a population, where
    -- the sample is a subset of the entire population.
  | ConsolidateStdDevP
    -- ^ The standard deviation of a population, where the population
    -- is all of the data to be summarized.
  | ConsolidateSum
    -- ^ The sum of the values.
  | ConsolidateVariance
    -- ^ An estimate of the variance of a population, where the sample
    -- is a subset of the entire population.
  | ConsolidateVarP
    -- ^ The variance of a population, where the population is all of
    -- the data to be summarized.
  deriving (Eq, Show)

{-------------------------------------------------------------------------------
  Rendering
-------------------------------------------------------------------------------}

instance ToAttrVal ConsolidateFunction where
  toAttrVal ConsolidateAverage = "average"
  toAttrVal ConsolidateCount = "count"
  toAttrVal ConsolidateCountNums = "countNums"
  toAttrVal ConsolidateMaximum = "max"
  toAttrVal ConsolidateMinimum = "min"
  toAttrVal ConsolidateProduct = "product"
  toAttrVal ConsolidateStdDev = "stdDev"
  toAttrVal ConsolidateStdDevP = "stdDevp"
  toAttrVal ConsolidateSum = "sum"
  toAttrVal ConsolidateVariance = "var"
  toAttrVal ConsolidateVarP = "varp"

instance ToAttrVal PivotFieldName where
  toAttrVal (PivotFieldName n) = toAttrVal n

{-------------------------------------------------------------------------------
  Parsing
-------------------------------------------------------------------------------}

instance FromAttrVal ConsolidateFunction where
  fromAttrVal "average" = readSuccess ConsolidateAverage
  fromAttrVal "count" = readSuccess ConsolidateCount
  fromAttrVal "countNums" = readSuccess ConsolidateCountNums
  fromAttrVal "max" = readSuccess ConsolidateMaximum
  fromAttrVal "min" = readSuccess ConsolidateMinimum
  fromAttrVal "product" = readSuccess ConsolidateProduct
  fromAttrVal "stdDev" = readSuccess ConsolidateStdDev
  fromAttrVal "stdDevp" = readSuccess ConsolidateStdDevP
  fromAttrVal "sum" = readSuccess ConsolidateSum
  fromAttrVal "var" = readSuccess ConsolidateVariance
  fromAttrVal "varp" = readSuccess ConsolidateVarP
  fromAttrVal t = invalidText "ConsolidateFunction" t

instance FromAttrVal PivotFieldName where
  fromAttrVal = fmap (first PivotFieldName) . fromAttrVal
