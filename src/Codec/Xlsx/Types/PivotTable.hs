{-# LANGUAGE OverloadedStrings #-}
{-# LANGUAGE RecordWildCards #-}
module Codec.Xlsx.Types.PivotTable
  ( PivotTable(..)
  , PivotFieldName(..)
  , PivotFieldInfo(..)
  , PositionedField(..)
  , DataField(..)
  , ConsolidateFunction(..)
  , PivotTableFiles(..)
  , CacheId(..)
  , parsePivotTable
  , parseCache
  , renderPivotTableFiles
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

import Codec.Xlsx.Parser.Internal
import Codec.Xlsx.Types.Common
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

data CacheDefinition = CacheDefinition
  { cdSourceRef :: CellRef
  , cdSourceSheet :: Text
  , cdFields :: [CacheField]
  } deriving (Eq, Show)

newtype CacheId = CacheId Int deriving Eq

newtype CacheField =
  CacheField Text
  deriving (Eq, Show)

{-------------------------------------------------------------------------------
  Parsing
-------------------------------------------------------------------------------}

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

instance FromAttrVal CacheId where
  fromAttrVal = fmap (first CacheId) . fromAttrVal

instance FromAttrVal PivotFieldName where
  fromAttrVal = fmap (first PivotFieldName) . fromAttrVal

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

{-------------------------------------------------------------------------------
  Rendering
-------------------------------------------------------------------------------}

data PivotTableFiles = PivotTableFiles
  { pvtfTable :: ByteString
  , pvtfCacheDefinition :: ByteString
  } deriving (Eq, Show)

renderPivotTableFiles :: Int -> PivotTable -> PivotTableFiles
renderPivotTableFiles cacheId t = PivotTableFiles {..}
  where
    pvtfTable = renderLBS def $ ptDefinitionDocument cacheId t
    cacheDefinition = generateCache t
    pvtfCacheDefinition = renderLBS def $ toDocument cacheDefinition

ptDefinitionDocument :: Int -> PivotTable -> Document
ptDefinitionDocument cacheId t =
    documentFromElement "Pivot table generated by xlsx" $
    ptDefinitionElement "pivotTableDefinition" cacheId t

ptDefinitionElement :: Name -> Int -> PivotTable -> Element
ptDefinitionElement nm cacheId PivotTable {..} =
  elementList nm attrs elements
  where
    attrs =
      catMaybes
        [ "colGrandTotals" .=? justFalse _pvtColumnGrandTotals
        , "rowGrandTotals" .=? justFalse _pvtRowGrandTotals
        , "outline" .=? justTrue _pvtOutline
        , "outlineData" .=? justTrue  _pvtOutlineData
        ] ++
      [ "name" .= _pvtName
      , "dataCaption" .= _pvtDataCaption
      , "cacheId" .= cacheId
      , "dataOnRows" .= (DataPosition `elem` _pvtRowFields)
      ]
    elements = [location, pivotFields, rowFields, colFields, dataFields]
    location =
      leafElement
        "location"
        [ "ref" .= _pvtLocation
          -- TODO : set proper
        , "firstHeaderRow" .= (1 :: Int)
        , "firstDataRow" .= (2 :: Int)
        , "firstDataCol" .= (1 :: Int)
        ]
    name2x = M.fromList $ zip (map _pfiName _pvtFields) [0 ..]
    mapFieldToX f = fromJustNote "no field" $ M.lookup f name2x
    pivotFields =
      elementListSimple "pivotFields" $ map pFieldEl _pvtFields
    pFieldEl PivotFieldInfo{_pfiName=fName, _pfiOutline=outline}
      | FieldPosition fName `elem` _pvtRowFields =
        pFieldEl' fName outline ("axisRow" :: Text)
      | FieldPosition fName `elem` _pvtColumnFields =
        pFieldEl' fName outline ("axisCol" :: Text)
      | otherwise =
        leafElement
          "pivotField"
          [ "name" .= fName
          , "dataField" .= True
          , "showAll" .= False
          , "outline" .= outline
          ]
    pFieldEl' fName outline axis =
      elementList
        "pivotField"
        [ "name" .= fName
        , "axis" .= axis
        , "showAll" .= False
        , "outline" .= outline
        ]
        [ elementListSimple "items" $
          [leafElement "item" ["t" .= ("default" :: Text)]]
        ]
    rowFields =
      elementListSimple "rowFields" . map fieldEl $
      if length _pvtDataFields > 1
        then _pvtRowFields
        else filter (/= DataPosition) _pvtRowFields
    colFields = elementListSimple "colFields" $ map fieldEl _pvtColumnFields
    fieldEl p = leafElement "field" ["x" .= fieldPos p]
    fieldPos DataPosition = (-2) :: Int
    fieldPos (FieldPosition f) = mapFieldToX f
    dataFields = elementListSimple "dataFields" $ map dFieldEl _pvtDataFields
    dFieldEl DataField {..} =
      leafElement "dataField" $
      catMaybes
        [ "name" .=? Just _dfName
        , "fld" .=? Just (mapFieldToX _dfField)
        , "subtotal" .=? justNonDef ConsolidateSum _dfFunction
        ]

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

generateCache :: PivotTable -> CacheDefinition
generateCache PivotTable {..} =
  CacheDefinition
  { cdSourceRef = _pvtSrcRef
  , cdSourceSheet = _pvtSrcSheet
  , cdFields = cachedFields
  }
  where
    cachedFields = map (cache . _pfiName) _pvtFields
    cache (PivotFieldName name) = CacheField name

instance ToDocument CacheDefinition where
  toDocument =
    documentFromElement "Pivot cache definition generated by xlsx" .
    toElement "pivotCacheDefinition"

instance ToElement CacheDefinition where
 toElement nm CacheDefinition {..} = elementList nm attrs elements
  where
    attrs = ["invalid" .= True, "refreshOnLoad" .= True]
    elements = [worksheetSource, cacheFields]
    worksheetSource =
      elementList
        "cacheSource"
        ["type" .= ("worksheet" :: Text)]
        [ leafElement
            "worksheetSource"
            ["ref" .= cdSourceRef, "sheet" .= cdSourceSheet]
        ]
    cacheFields =
      elementListSimple "cacheFields" $ map (toElement "cacheField") cdFields

instance ToElement CacheField where
  toElement nm (CacheField fieldName) = leafElement nm ["name" .= fieldName]
