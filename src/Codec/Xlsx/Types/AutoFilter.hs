{-# LANGUAGE OverloadedStrings #-}
{-# LANGUAGE RecordWildCards #-}
{-# LANGUAGE TemplateHaskell #-}
{-# LANGUAGE DeriveGeneric #-}
module Codec.Xlsx.Types.AutoFilter where

import Control.Lens (makeLenses)
import Data.Default
import Data.Map (Map)
import Data.Maybe (catMaybes)
import qualified Data.Map as M
import Data.Text (Text)
import Text.XML
import Text.XML.Cursor

import Codec.Xlsx.Parser.Internal
import Codec.Xlsx.Types.Common
import Codec.Xlsx.Writer.Internal

-- | The filterColumn collection identifies a particular column in the
-- AutoFilter range and specifies filter information that has been
-- applied to this column. If a column in the AutoFilter range has no
-- criteria specified, then there is no corresponding filterColumn
-- collection expressed for that column.
--
-- Section 18.3.2.7 "filterColumn (AutoFilter Column)" (p. 1717)
data FilterColumn
  = Filters { _fltValues :: [Text]}
  | ACustomFilter CustomFilter
  | CustomFiltersOr CustomFilter
                    CustomFilter
  | CustomFiltersAnd CustomFilter
                     CustomFilter
  deriving (Eq, Show, Generic)

data CustomFilter = CustomFilter
  { cfltOperator :: CustomFilterOperator
  , cfltValue :: Text
  } deriving (Eq, Show, Generic)

data CustomFilterOperator
  = FltrEqual
    -- ^ Show results which are equal to criteria.
  | FltrGreaterThan
    -- ^ Show results which are greater than criteria.
  | FltrGreaterThanOrEqual
    -- ^ Show results which are greater than or equal to criteria.
  | FltrLessThan
    -- ^ Show results which are less than criteria.
  | FltrLessThanOrEqual
    -- ^ Show results which are less than or equal to criteria.
  | FltrNotEqual
    -- ^ Show results which are not equal to criteria.
  deriving (Eq, Show, Generic)

-- | AutoFilter temporarily hides rows based on a filter criteria,
-- which is applied column by column to a table of datain the
-- worksheet.
--
-- TODO: sortState, extList
--
-- Section 18.3.1.2 "autoFilter (AutoFilter Settings)" (p. 1596)
data AutoFilter = AutoFilter
  { _afRef :: Maybe CellRef
  , _afFilterColumns :: Map Int FilterColumn
  } deriving (Eq, Show, Generic)

makeLenses ''AutoFilter


{-------------------------------------------------------------------------------
  Default instances
-------------------------------------------------------------------------------}

instance Default AutoFilter where
    def = AutoFilter Nothing M.empty

{-------------------------------------------------------------------------------
  Parsing
-------------------------------------------------------------------------------}

instance FromCursor AutoFilter where
  fromCursor cur = do
    _afRef <- maybeAttribute "ref" cur
    let _afFilterColumns = M.fromList $ cur $/ element (n_ "filterColumn") >=> \c -> do
          colId <- fromAttribute "colId" c
          fcol <- c $/ anyElement >=> fltColFromNode . node
          return (colId, fcol)
    return AutoFilter {..}

fltColFromNode :: Node -> [FilterColumn]
fltColFromNode n | n `nodeElNameIs` (n_ "filters") = do
                     let _fltValues = cur $/ element (n_ "filter") >=> fromAttribute "val"
                     return Filters{..}
                 | n `nodeElNameIs` (n_ "customFilters") = do
                     isAnd <- fromAttributeDef "and" False cur
                     let cFilters = cur $/ element (n_ "customFilter") >=> \c -> do
                           op <- fromAttributeDef "operator" FltrEqual c
                           val <- fromAttribute "val" c
                           return $ CustomFilter op val
                     case cFilters of
                       [f] ->
                         return $ ACustomFilter f
                       [f1, f2] ->
                         if isAnd
                           then return $ CustomFiltersAnd f1 f2
                           else return $ CustomFiltersOr f1 f2
                       _ ->
                         fail "bad custom filter"
                 | otherwise = fail "no matching nodes"
  where
    cur = fromNode n

instance FromAttrVal CustomFilterOperator where
  fromAttrVal "equal" = readSuccess FltrEqual
  fromAttrVal "greaterThan" = readSuccess FltrGreaterThan
  fromAttrVal "greaterThanOrEqual" = readSuccess FltrGreaterThanOrEqual
  fromAttrVal "lessThan" = readSuccess FltrLessThan
  fromAttrVal "lessThanOrEqual" = readSuccess FltrLessThanOrEqual
  fromAttrVal "notEqual" = readSuccess FltrNotEqual
  fromAttrVal t = invalidText "CustomFilterOperator" t

{-------------------------------------------------------------------------------
  Rendering
-------------------------------------------------------------------------------}

instance ToElement AutoFilter where
  toElement nm AutoFilter {..} =
    elementList
      nm
      (catMaybes ["ref" .=? _afRef])
      [ elementList
        (n_ "filterColumn")
        ["colId" .= colId]
        [fltColToElement fCol]
      | (colId, fCol) <- M.toList _afFilterColumns
      ]

fltColToElement :: FilterColumn -> Element
fltColToElement Filters {..} =
  elementListSimple
    (n_ "filters")
    [leafElement (n_ "filter") ["val" .= v] | v <- _fltValues]
fltColToElement (ACustomFilter f) =
  elementListSimple (n_ "customFilters") [toElement (n_ "customFilter") f]
fltColToElement (CustomFiltersOr f1 f2) =
  elementListSimple
    (n_ "customFilters")
    [toElement (n_ "customFilter") f | f <- [f1, f2]]
fltColToElement (CustomFiltersAnd f1 f2) =
  elementList
    (n_ "customFilters")
    ["and" .= True]
    [toElement (n_ "customFilter") f | f <- [f1, f2]]

instance ToElement CustomFilter where
  toElement nm CustomFilter {..} =
    leafElement nm ["operator" .= cfltOperator, "val" .= cfltValue]

instance ToAttrVal CustomFilterOperator where
  toAttrVal FltrEqual = "equal"
  toAttrVal FltrGreaterThan = "greaterThan"
  toAttrVal FltrGreaterThanOrEqual = "greaterThanOrEqual"
  toAttrVal FltrLessThan = "lessThan"
  toAttrVal FltrLessThanOrEqual = "lessThanOrEqual"
  toAttrVal FltrNotEqual = "notEqual"
