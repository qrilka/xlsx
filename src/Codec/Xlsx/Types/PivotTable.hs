{-# LANGUAGE OverloadedStrings #-}
{-# LANGUAGE RecordWildCards #-}
{-# LANGUAGE DeriveGeneric #-}
module Codec.Xlsx.Types.PivotTable
  ( PivotTable(..)
  , PivotFieldName(..)
  , PivotFieldInfo(..)
  , FieldSortType(..)
  , PositionedField(..)
  , DataField(..)
  , ConsolidateFunction(..)
  ) where

import Control.Arrow (first)
import Control.DeepSeq (NFData)
import Data.Text (Text)
import GHC.Generics (Generic)

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
  } deriving (Eq, Show, Generic)
instance NFData PivotTable

data PivotFieldInfo = PivotFieldInfo
  { _pfiName :: Maybe PivotFieldName
  , _pfiOutline :: Bool
  , _pfiSortType :: FieldSortType
  , _pfiHiddenItems :: [CellValue]
  } deriving (Eq, Show, Generic)
instance NFData PivotFieldInfo

-- | Sort orders that can be applied to fields in a PivotTable
--
-- See 18.18.28 "ST_FieldSortType (Field Sort Type)" (p. 2454)
data FieldSortType
  = FieldSortAscending
  | FieldSortDescending
  | FieldSortManual
  deriving (Eq, Ord, Show, Generic)
instance NFData FieldSortType

newtype PivotFieldName =
  PivotFieldName Text
  deriving (Eq, Ord, Show, Generic)
instance NFData PivotFieldName

data PositionedField
  = DataPosition
  | FieldPosition PivotFieldName
  deriving (Eq, Ord, Show, Generic)
instance NFData PositionedField

data DataField = DataField
  { _dfField :: PivotFieldName
  , _dfName :: Text
  , _dfFunction :: ConsolidateFunction
  } deriving (Eq, Show, Generic)
instance NFData DataField

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
  deriving (Eq, Show, Generic)
instance NFData ConsolidateFunction

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

instance ToAttrVal FieldSortType where
  toAttrVal FieldSortManual = "manual"
  toAttrVal FieldSortAscending = "ascending"
  toAttrVal FieldSortDescending = "descending"

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

instance FromAttrVal FieldSortType where
  fromAttrVal "manual" = readSuccess FieldSortManual
  fromAttrVal "ascending" = readSuccess FieldSortAscending
  fromAttrVal "descending" = readSuccess FieldSortDescending
  fromAttrVal t = invalidText "FieldSortType" t
