{-# LANGUAGE CPP #-}
{-# LANGUAGE FlexibleInstances #-}
{-# LANGUAGE OverloadedStrings #-}
{-# LANGUAGE RecordWildCards #-}
{-# LANGUAGE ScopedTypeVariables #-}
{-# LANGUAGE TemplateHaskell #-}
{-# LANGUAGE DeriveGeneric #-}
module Codec.Xlsx.Types.AutoFilter where

import Control.Arrow (first)
import Control.DeepSeq (NFData)
#ifdef USE_MICROLENS
import Lens.Micro.TH (makeLenses)
#else
import Control.Lens (makeLenses)
#endif
import Data.Bool (bool)
import Data.ByteString (ByteString)
import Data.Default
import Data.Foldable (asum)
import Data.Map (Map)
import qualified Data.Map as M
import Data.Maybe (catMaybes)
import Data.Monoid ((<>))
import Data.Text (Text)
import qualified Data.Text as T
import GHC.Generics (Generic)
import Text.XML
import Text.XML.Cursor hiding (bool)
import qualified Xeno.DOM as Xeno

import Codec.Xlsx.Parser.Internal
import Codec.Xlsx.Types.Common
import Codec.Xlsx.Types.ConditionalFormatting (IconSetType)
import Codec.Xlsx.Writer.Internal

-- | The filterColumn collection identifies a particular column in the
-- AutoFilter range and specifies filter information that has been
-- applied to this column. If a column in the AutoFilter range has no
-- criteria specified, then there is no corresponding filterColumn
-- collection expressed for that column.
--
-- See 18.3.2.7 "filterColumn (AutoFilter Column)" (p. 1717)
data FilterColumn
  = Filters FilterByBlank [FilterCriterion]
  | ColorFilter ColorFilterOptions
  | ACustomFilter CustomFilter
  | CustomFiltersOr CustomFilter CustomFilter
  | CustomFiltersAnd CustomFilter CustomFilter
  | DynamicFilter DynFilterOptions
  | IconFilter (Maybe Int) IconSetType
  -- ^ Specifies the icon set and particular icon within that set to
  -- filter by. Icon is specified using zero-based index of an icon in
  -- an icon set. 'Nothing' means "no icon"
  | BottomNFilter EdgeFilterOptions
  -- ^ Specifies the bottom N (percent or number of items) to filter by
  | TopNFilter EdgeFilterOptions
  -- ^ Specifies the top N (percent or number of items) to filter by
  deriving (Eq, Show, Generic)
instance NFData FilterColumn

data FilterByBlank
  = FilterByBlank
  | DontFilterByBlank
  deriving (Eq, Show, Generic)
instance NFData FilterByBlank

data FilterCriterion
  = FilterValue Text
  | FilterDateGroup DateGroup
  deriving (Eq, Show, Generic)
instance NFData FilterCriterion

-- | Used to express a group of dates or times which are used in an
-- AutoFilter criteria
--
-- Section 18.3.2.4 "dateGroupItem (Date Grouping)" (p. 1714)
data DateGroup
  = DateGroupByYear Int
  | DateGroupByMonth Int Int
  | DateGroupByDay Int Int Int
  | DateGroupByHour Int Int Int Int
  | DateGroupByMinute Int Int Int Int Int
  | DateGroupBySecond Int Int Int Int Int Int
  deriving (Eq, Show, Generic)
instance NFData DateGroup

data CustomFilter = CustomFilter
  { cfltOperator :: CustomFilterOperator
  , cfltValue :: Text
  } deriving (Eq, Show, Generic)
instance NFData CustomFilter

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
instance NFData CustomFilterOperator

data EdgeFilterOptions = EdgeFilterOptions
  { _efoUsePercents :: Bool
  -- ^ Flag indicating whether or not to filter by percent value of
  -- the column. A false value filters by number of items.
  , _efoVal :: Double
  -- ^ Top or bottom value to use as the filter criteria.
  -- Example: "Filter by Top 10 Percent" or "Filter by Top 5 Items"
  , _efoFilterVal :: Maybe Double
  -- ^ The actual cell value in the range which is used to perform the
  -- comparison for this filter.
  } deriving (Eq, Show, Generic)
instance NFData EdgeFilterOptions

-- | Specifies the color to filter by and whether to use the cell's
-- fill or font color in the filter criteria. If the cell's font or
-- fill color does not match the color specified in the criteria, the
-- rows corresponding to those cells are hidden from view.
--
-- See 18.3.2.1 "colorFilter (Color Filter Criteria)" (p. 1712)
data ColorFilterOptions = ColorFilterOptions
  { _cfoCellColor :: Bool
  -- ^ Flag indicating whether or not to filter by the cell's fill
  -- color. 'True' indicates to filter by cell fill. 'False' indicates
  -- to filter by the cell's font color.
  --
  -- For rich text in cells, if the color specified appears in the
  -- cell at all, it shall be included in the filter.
  , _cfoDxfId :: Maybe Int
  -- ^ Id of differential format record (dxf) in the Styles Part (see
  -- '_styleSheetDxfs') which expresses the color value to filter by.
  } deriving (Eq, Show, Generic)
instance NFData ColorFilterOptions

-- | Specifies dynamic filter criteria. These criteria are considered
-- dynamic because they can change, either with the data itself (e.g.,
-- "above average") or with the current system date (e.g., show values
-- for "today"). For any cells whose values do not meet the specified
-- criteria, the corresponding rows shall be hidden from view when the
-- filter is applied.
--
-- '_dfoMaxVal' shall be required for 'DynFilterTday',
-- 'DynFilterYesterday', 'DynFilterTomorrow', 'DynFilterNextWeek',
-- 'DynFilterThisWeek', 'DynFilterLastWeek', 'DynFilterNextMonth',
-- 'DynFilterThisMonth', 'DynFilterLastMonth', 'DynFilterNextQuarter',
-- 'DynFilterThisQuarter', 'DynFilterLastQuarter',
-- 'DynFilterNextYear', 'DynFilterThisYear', 'DynFilterLastYear', and
-- 'DynFilterYearToDate.
--
-- The above criteria are based on a value range; that is, if today's
-- date is September 22nd, then the range for thisWeek is the values
-- greater than or equal to September 17 and less than September
-- 24. In the thisWeek range, the lower value is expressed
-- '_dfoval'. The higher value is expressed using '_dfoMmaxVal'.
--
-- These dynamic filters shall not require '_dfoVal or '_dfoMaxVal':
-- 'DynFilterQ1', 'DynFilterQ2', 'DynFilterQ3', 'DynFilterQ4',
-- 'DynFilterM1', 'DynFilterM2', 'DynFilterM3', 'DynFilterM4',
-- 'DynFilterM5', 'DynFilterM6', 'DynFilterM7', 'DynFilterM8',
-- 'DynFilterM9', 'DynFilterM10', 'DynFilterM11' and 'DynFilterM12'.
--
-- The above criteria shall not specify the range using valIso and
-- maxValIso because Q1 always starts from M1 to M3, and M1 is always
-- January.
--
-- These types of dynamic filters shall use valIso and shall not use
-- '_dfoMaxVal': 'DynFilterAboveAverage' and 'DynFilterBelowAverage'
--
-- /Note:/ Specification lists 'valIso' and 'maxIso' to store datetime
-- values but it appears that Excel doesn't use them and stored them
-- as numeric values (as it does for datetimes in cell values)
--
-- See 18.3.2.5 "dynamicFilter (Dynamic Filter)" (p. 1715)
data DynFilterOptions = DynFilterOptions
  { _dfoType :: DynFilterType
  , _dfoVal :: Maybe Double
  -- ^ A minimum numeric value for dynamic filter.
  , _dfoMaxVal :: Maybe Double
  -- ^ A maximum value for dynamic filter.
  } deriving (Eq, Show, Generic)
instance NFData DynFilterOptions

-- | Specifies concrete type of dynamic filter used
--
-- See 18.18.26 "ST_DynamicFilterType (Dynamic Filter)" (p. 2452)
data DynFilterType
  = DynFilterAboveAverage
  -- ^ Shows values that are above average.
  | DynFilterBelowAverage
  -- ^ Shows values that are below average.
  | DynFilterLastMonth
  -- ^ Shows last month's dates.
  | DynFilterLastQuarter
  -- ^ Shows last calendar quarter's dates.
  | DynFilterLastWeek
  -- ^ Shows last week's dates, using Sunday as the first weekday.
  | DynFilterLastYear
  -- ^ Shows last year's dates.
  | DynFilterM1
  -- ^ Shows the dates that are in January, regardless of year.
  | DynFilterM10
  -- ^ Shows the dates that are in October, regardless of year.
  | DynFilterM11
  -- ^ Shows the dates that are in November, regardless of year.
  | DynFilterM12
  -- ^ Shows the dates that are in December, regardless of year.
  | DynFilterM2
  -- ^ Shows the dates that are in February, regardless of year.
  | DynFilterM3
  -- ^ Shows the dates that are in March, regardless of year.
  | DynFilterM4
  -- ^ Shows the dates that are in April, regardless of year.
  | DynFilterM5
  -- ^ Shows the dates that are in May, regardless of year.
  | DynFilterM6
  -- ^ Shows the dates that are in June, regardless of year.
  | DynFilterM7
  -- ^ Shows the dates that are in July, regardless of year.
  | DynFilterM8
  -- ^ Shows the dates that are in August, regardless of year.
  | DynFilterM9
  -- ^ Shows the dates that are in September, regardless of year.
  | DynFilterNextMonth
  -- ^ Shows next month's dates.
  | DynFilterNextQuarter
  -- ^ Shows next calendar quarter's dates.
  | DynFilterNextWeek
  -- ^ Shows next week's dates, using Sunday as the first weekday.
  | DynFilterNextYear
  -- ^ Shows next year's dates.
  | DynFilterNull
  -- ^ Common filter type not available.
  | DynFilterQ1
  -- ^ Shows the dates that are in the 1st calendar quarter,
  -- regardless of year.
  | DynFilterQ2
  -- ^ Shows the dates that are in the 2nd calendar quarter,
  -- regardless of year.
  | DynFilterQ3
  -- ^ Shows the dates that are in the 3rd calendar quarter,
  -- regardless of year.
  | DynFilterQ4
  -- ^ Shows the dates that are in the 4th calendar quarter,
  -- regardless of year.
  | DynFilterThisMonth
  -- ^ Shows this month's dates.
  | DynFilterThisQuarter
  -- ^ Shows this calendar quarter's dates.
  | DynFilterThisWeek
  -- ^ Shows this week's dates, using Sunday as the first weekday.
  | DynFilterThisYear
  -- ^ Shows this year's dates.
  | DynFilterToday
  -- ^ Shows today's dates.
  | DynFilterTomorrow
  -- ^ Shows tomorrow's dates.
  | DynFilterYearToDate
  -- ^ Shows the dates between the beginning of the year and today, inclusive.
  | DynFilterYesterday
  -- ^ Shows yesterday's dates.
  deriving (Eq, Show, Generic)
instance NFData DynFilterType

-- | AutoFilter temporarily hides rows based on a filter criteria,
-- which is applied column by column to a table of datain the
-- worksheet.
--
-- TODO: sortState, extList
--
-- See 18.3.1.2 "autoFilter (AutoFilter Settings)" (p. 1596)
data AutoFilter = AutoFilter
  { _afRef :: Maybe CellRef
  , _afFilterColumns :: Map Int FilterColumn
  } deriving (Eq, Show, Generic)
instance NFData AutoFilter

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

instance FromXenoNode AutoFilter where
  fromXenoNode root = do
    _afRef <- parseAttributes root $ maybeAttr "ref"
    _afFilterColumns <-
      fmap M.fromList . collectChildren root $ fromChildList "filterColumn"
    return AutoFilter {..}

instance FromXenoNode (Int, FilterColumn) where
  fromXenoNode root = do
    colId <- parseAttributes root $ fromAttr "colId"
    fCol <-
      collectChildren root $ asum [filters, color, custom, dynamic, icon, top10]
    return (colId, fCol)
    where
      filters =
        requireAndParse "filters" $ \node -> do
          filterBlank <-
            parseAttributes node $ fromAttrDef "blank" DontFilterByBlank
          filterCriteria <- childListAny node
          return $ Filters filterBlank filterCriteria
      color =
        requireAndParse "colorFilter" $ \node ->
          parseAttributes node $ do
            _cfoCellColor <- fromAttrDef "cellColor" True
            _cfoDxfId <- maybeAttr "dxfId"
            return $ ColorFilter ColorFilterOptions {..}
      custom =
        requireAndParse "customFilters" $ \node -> do
          isAnd <- parseAttributes node $ fromAttrDef "and" False
          cfilters <- collectChildren node $ fromChildList "customFilter"
          case cfilters of
            [f] -> return $ ACustomFilter f
            [f1, f2] ->
              if isAnd
                then return $ CustomFiltersAnd f1 f2
                else return $ CustomFiltersOr f1 f2
            _ ->
              Left $
              "expected 1 or 2 custom filters but found " <>
              T.pack (show $ length cfilters)
      dynamic =
        requireAndParse "dynamicFilter" . flip parseAttributes $ do
          _dfoType <- fromAttr "type"
          _dfoVal <- maybeAttr "val"
          _dfoMaxVal <- maybeAttr "maxVal"
          return $ DynamicFilter DynFilterOptions {..}
      icon =
        requireAndParse "iconFilter" . flip parseAttributes $
        IconFilter <$> maybeAttr "iconId" <*> fromAttr "iconSet"
      top10 =
        requireAndParse "top10" . flip parseAttributes $ do
          top <- fromAttrDef "top" True
          percent <- fromAttrDef "percent" False
          val <- fromAttr "val"
          filterVal <- maybeAttr "filterVal"
          let opts = EdgeFilterOptions percent val filterVal
          if top
            then return $ TopNFilter opts
            else return $ BottomNFilter opts

instance FromXenoNode CustomFilter where
  fromXenoNode root =
    parseAttributes root $
    CustomFilter <$> fromAttrDef "operator" FltrEqual <*> fromAttr "val"

fltColFromNode :: Node -> [FilterColumn]
fltColFromNode n | n `nodeElNameIs` (n_ "filters") = do
                     let filterCriteria = cur $/ anyElement >=> fromCursor
                     filterBlank <- fromAttributeDef "blank" DontFilterByBlank cur
                     return $ Filters filterBlank filterCriteria
                 | n `nodeElNameIs` (n_ "colorFilter") = do
                     _cfoCellColor <- fromAttributeDef "cellColor" True cur
                     _cfoDxfId <- maybeAttribute "dxfId" cur
                     return $ ColorFilter ColorFilterOptions {..}
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
                 | n `nodeElNameIs` (n_ "dynamicFilter") = do
                     _dfoType <- fromAttribute "type" cur
                     _dfoVal <- maybeAttribute "val" cur
                     _dfoMaxVal <- maybeAttribute "maxVal" cur
                     return $ DynamicFilter DynFilterOptions{..}
                 | n `nodeElNameIs` (n_ "iconFilter") = do
                     iconId <- maybeAttribute "iconId" cur
                     iconSet <- fromAttribute "iconSet" cur
                     return $ IconFilter iconId iconSet
                 | n `nodeElNameIs` (n_ "top10") = do
                     top <- fromAttributeDef "top" True cur
                     let percent = fromAttributeDef "percent" False cur
                         val = fromAttribute "val" cur
                         filterVal = maybeAttribute "filterVal" cur
                     if top
                       then fmap TopNFilter $
                            EdgeFilterOptions <$> percent <*> val <*> filterVal
                       else fmap BottomNFilter $
                            EdgeFilterOptions <$> percent <*> val <*> filterVal
                 | otherwise = fail "no matching nodes"
  where
    cur = fromNode n

instance FromCursor FilterCriterion where
  fromCursor = filterCriterionFromNode . node

instance FromXenoNode FilterCriterion where
  fromXenoNode root =
    case Xeno.name root of
      "filter" -> parseAttributes root $ do FilterValue <$> fromAttr "val"
      "dateGroupItem" ->
        parseAttributes root $ do
          grouping <- fromAttr "dateTimeGrouping"
          group <- case grouping of
            ("year" :: ByteString) ->
              DateGroupByYear <$> fromAttr "year"
            "month" ->
              DateGroupByMonth <$> fromAttr "year"
                               <*> fromAttr "month"
            "day" ->
              DateGroupByDay <$> fromAttr "year"
                             <*> fromAttr "month"
                             <*> fromAttr "day"
            "hour" ->
              DateGroupByHour <$> fromAttr "year"
                              <*> fromAttr "month"
                              <*> fromAttr "day"
                              <*> fromAttr "hour"
            "minute" ->
              DateGroupByMinute <$> fromAttr "year"
                                <*> fromAttr "month"
                                <*> fromAttr "day"
                                <*> fromAttr "hour"
                                <*> fromAttr "minute"
            "second" ->
              DateGroupBySecond <$> fromAttr "year"
                                <*> fromAttr "month"
                                <*> fromAttr "day"
                                <*> fromAttr "hour"
                                <*> fromAttr "minute"
                                <*> fromAttr "second"
            _ -> toAttrParser . Left $ "Unexpected date grouping"
          return $ FilterDateGroup group
      _ -> Left "Bad FilterCriterion"

-- TODO: follow the spec about the fact that dategroupitem always go after filter
filterCriterionFromNode :: Node -> [FilterCriterion]
filterCriterionFromNode n
  | n `nodeElNameIs` (n_ "filter") = do
    v <- fromAttribute "val" cur
    return $ FilterValue v
  | n `nodeElNameIs` (n_ "dateGroupItem") = do
    g <- fromAttribute "dateTimeGrouping" cur
    let year = fromAttribute "year" cur
        month = fromAttribute "month" cur
        day = fromAttribute "day" cur
        hour = fromAttribute "hour" cur
        minute = fromAttribute "minute" cur
        second = fromAttribute "second" cur
    FilterDateGroup <$>
      case g of
        "year" -> DateGroupByYear <$> year
        "month" -> DateGroupByMonth <$> year <*> month
        "day" -> DateGroupByDay <$> year <*> month <*> day
        "hour" -> DateGroupByHour <$> year <*> month <*> day <*> hour
        "minute" ->
          DateGroupByMinute <$> year <*> month <*> day <*> hour <*> minute
        "second" ->
          DateGroupBySecond <$> year <*> month <*> day <*> hour <*> minute <*>
          second
        _ -> fail $ "unexpected dateTimeGrouping " ++ show (g :: Text)
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

instance FromAttrBs CustomFilterOperator where
  fromAttrBs "equal" = return FltrEqual
  fromAttrBs "greaterThan" = return FltrGreaterThan
  fromAttrBs "greaterThanOrEqual" = return FltrGreaterThanOrEqual
  fromAttrBs "lessThan" = return FltrLessThan
  fromAttrBs "lessThanOrEqual" = return FltrLessThanOrEqual
  fromAttrBs "notEqual" = return FltrNotEqual
  fromAttrBs x = unexpectedAttrBs "CustomFilterOperator" x

instance FromAttrVal FilterByBlank where
  fromAttrVal =
    fmap (first $ bool DontFilterByBlank FilterByBlank) . fromAttrVal

instance FromAttrBs FilterByBlank where
  fromAttrBs = fmap (bool DontFilterByBlank FilterByBlank) . fromAttrBs

instance FromAttrVal DynFilterType where
  fromAttrVal "aboveAverage" = readSuccess DynFilterAboveAverage
  fromAttrVal "belowAverage" = readSuccess DynFilterBelowAverage
  fromAttrVal "lastMonth" = readSuccess DynFilterLastMonth
  fromAttrVal "lastQuarter" = readSuccess DynFilterLastQuarter
  fromAttrVal "lastWeek" = readSuccess DynFilterLastWeek
  fromAttrVal "lastYear" = readSuccess DynFilterLastYear
  fromAttrVal "M1" = readSuccess DynFilterM1
  fromAttrVal "M10" = readSuccess DynFilterM10
  fromAttrVal "M11" = readSuccess DynFilterM11
  fromAttrVal "M12" = readSuccess DynFilterM12
  fromAttrVal "M2" = readSuccess DynFilterM2
  fromAttrVal "M3" = readSuccess DynFilterM3
  fromAttrVal "M4" = readSuccess DynFilterM4
  fromAttrVal "M5" = readSuccess DynFilterM5
  fromAttrVal "M6" = readSuccess DynFilterM6
  fromAttrVal "M7" = readSuccess DynFilterM7
  fromAttrVal "M8" = readSuccess DynFilterM8
  fromAttrVal "M9" = readSuccess DynFilterM9
  fromAttrVal "nextMonth" = readSuccess DynFilterNextMonth
  fromAttrVal "nextQuarter" = readSuccess DynFilterNextQuarter
  fromAttrVal "nextWeek" = readSuccess DynFilterNextWeek
  fromAttrVal "nextYear" = readSuccess DynFilterNextYear
  fromAttrVal "null" = readSuccess DynFilterNull
  fromAttrVal "Q1" = readSuccess DynFilterQ1
  fromAttrVal "Q2" = readSuccess DynFilterQ2
  fromAttrVal "Q3" = readSuccess DynFilterQ3
  fromAttrVal "Q4" = readSuccess DynFilterQ4
  fromAttrVal "thisMonth" = readSuccess DynFilterThisMonth
  fromAttrVal "thisQuarter" = readSuccess DynFilterThisQuarter
  fromAttrVal "thisWeek" = readSuccess DynFilterThisWeek
  fromAttrVal "thisYear" = readSuccess DynFilterThisYear
  fromAttrVal "today" = readSuccess DynFilterToday
  fromAttrVal "tomorrow" = readSuccess DynFilterTomorrow
  fromAttrVal "yearToDate" = readSuccess DynFilterYearToDate
  fromAttrVal "yesterday" = readSuccess DynFilterYesterday
  fromAttrVal t = invalidText "DynFilterType" t

instance FromAttrBs DynFilterType where
  fromAttrBs "aboveAverage" = return DynFilterAboveAverage
  fromAttrBs "belowAverage" = return DynFilterBelowAverage
  fromAttrBs "lastMonth" = return DynFilterLastMonth
  fromAttrBs "lastQuarter" = return DynFilterLastQuarter
  fromAttrBs "lastWeek" = return DynFilterLastWeek
  fromAttrBs "lastYear" = return DynFilterLastYear
  fromAttrBs "M1" = return DynFilterM1
  fromAttrBs "M10" = return DynFilterM10
  fromAttrBs "M11" = return DynFilterM11
  fromAttrBs "M12" = return DynFilterM12
  fromAttrBs "M2" = return DynFilterM2
  fromAttrBs "M3" = return DynFilterM3
  fromAttrBs "M4" = return DynFilterM4
  fromAttrBs "M5" = return DynFilterM5
  fromAttrBs "M6" = return DynFilterM6
  fromAttrBs "M7" = return DynFilterM7
  fromAttrBs "M8" = return DynFilterM8
  fromAttrBs "M9" = return DynFilterM9
  fromAttrBs "nextMonth" = return DynFilterNextMonth
  fromAttrBs "nextQuarter" = return DynFilterNextQuarter
  fromAttrBs "nextWeek" = return DynFilterNextWeek
  fromAttrBs "nextYear" = return DynFilterNextYear
  fromAttrBs "null" = return DynFilterNull
  fromAttrBs "Q1" = return DynFilterQ1
  fromAttrBs "Q2" = return DynFilterQ2
  fromAttrBs "Q3" = return DynFilterQ3
  fromAttrBs "Q4" = return DynFilterQ4
  fromAttrBs "thisMonth" = return DynFilterThisMonth
  fromAttrBs "thisQuarter" = return DynFilterThisQuarter
  fromAttrBs "thisWeek" = return DynFilterThisWeek
  fromAttrBs "thisYear" = return DynFilterThisYear
  fromAttrBs "today" = return DynFilterToday
  fromAttrBs "tomorrow" = return DynFilterTomorrow
  fromAttrBs "yearToDate" = return DynFilterYearToDate
  fromAttrBs "yesterday" = return DynFilterYesterday
  fromAttrBs x = unexpectedAttrBs "DynFilterType" x

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
fltColToElement (Filters filterBlank filterCriteria) =
  let attrs = catMaybes ["blank" .=? justNonDef DontFilterByBlank filterBlank]
  in elementList
     (n_ "filters") attrs $ map filterCriterionToElement filterCriteria
fltColToElement (ColorFilter opts) = toElement (n_ "colorFilter") opts
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
fltColToElement (DynamicFilter opts) = toElement (n_ "dynamicFilter") opts
fltColToElement (IconFilter iconId iconSet) =
  leafElement (n_ "iconFilter") $
  ["iconSet" .= iconSet] ++ catMaybes ["iconId" .=? iconId]
fltColToElement (BottomNFilter opts) = edgeFilter False opts
fltColToElement (TopNFilter opts) = edgeFilter True opts

edgeFilter :: Bool -> EdgeFilterOptions -> Element
edgeFilter top EdgeFilterOptions {..} =
  leafElement (n_ "top10") $
  ["top" .= top, "percent" .= _efoUsePercents, "val" .= _efoVal] ++
  catMaybes ["filterVal" .=? _efoFilterVal]

filterCriterionToElement :: FilterCriterion -> Element
filterCriterionToElement (FilterValue v) =
  leafElement (n_ "filter") ["val" .= v]
filterCriterionToElement (FilterDateGroup (DateGroupByYear y)) =
  leafElement
    (n_ "dateGroupItem")
    ["dateTimeGrouping" .= ("year" :: Text), "year" .= y]
filterCriterionToElement (FilterDateGroup (DateGroupByMonth y m)) =
  leafElement
    (n_ "dateGroupItem")
    ["dateTimeGrouping" .= ("month" :: Text), "year" .= y, "month" .= m]
filterCriterionToElement (FilterDateGroup (DateGroupByDay y m d)) =
  leafElement
    (n_ "dateGroupItem")
    ["dateTimeGrouping" .= ("day" :: Text), "year" .= y, "month" .= m, "day" .= d]
filterCriterionToElement (FilterDateGroup (DateGroupByHour y m d h)) =
  leafElement
    (n_ "dateGroupItem")
    [ "dateTimeGrouping" .= ("hour" :: Text)
    , "year" .= y
    , "month" .= m
    , "day" .= d
    , "hour" .= h
    ]
filterCriterionToElement (FilterDateGroup (DateGroupByMinute y m d h mi)) =
  leafElement
    (n_ "dateGroupItem")
    [ "dateTimeGrouping" .= ("minute" :: Text)
    , "year" .= y
    , "month" .= m
    , "day" .= d
    , "hour" .= h
    , "minute" .= mi
    ]
filterCriterionToElement (FilterDateGroup (DateGroupBySecond y m d h mi s)) =
  leafElement
    (n_ "dateGroupItem")
    [ "dateTimeGrouping" .= ("second" :: Text)
    , "year" .= y
    , "month" .= m
    , "day" .= d
    , "hour" .= h
    , "minute" .= mi
    , "second" .= s
    ]

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

instance ToAttrVal FilterByBlank where
  toAttrVal FilterByBlank = toAttrVal True
  toAttrVal DontFilterByBlank = toAttrVal False

instance ToElement ColorFilterOptions where
  toElement nm ColorFilterOptions {..} =
    leafElement nm $
    catMaybes ["cellColor" .=? justFalse _cfoCellColor, "dxfId" .=? _cfoDxfId]

instance ToElement DynFilterOptions where
  toElement nm DynFilterOptions {..} =
    leafElement nm $
    ["type" .= _dfoType] ++
    catMaybes ["val" .=? _dfoVal, "maxVal" .=? _dfoMaxVal]

instance ToAttrVal DynFilterType where
  toAttrVal DynFilterAboveAverage = "aboveAverage"
  toAttrVal DynFilterBelowAverage = "belowAverage"
  toAttrVal DynFilterLastMonth = "lastMonth"
  toAttrVal DynFilterLastQuarter = "lastQuarter"
  toAttrVal DynFilterLastWeek = "lastWeek"
  toAttrVal DynFilterLastYear = "lastYear"
  toAttrVal DynFilterM1 = "M1"
  toAttrVal DynFilterM10 = "M10"
  toAttrVal DynFilterM11 = "M11"
  toAttrVal DynFilterM12 = "M12"
  toAttrVal DynFilterM2 = "M2"
  toAttrVal DynFilterM3 = "M3"
  toAttrVal DynFilterM4 = "M4"
  toAttrVal DynFilterM5 = "M5"
  toAttrVal DynFilterM6 = "M6"
  toAttrVal DynFilterM7 = "M7"
  toAttrVal DynFilterM8 = "M8"
  toAttrVal DynFilterM9 = "M9"
  toAttrVal DynFilterNextMonth = "nextMonth"
  toAttrVal DynFilterNextQuarter = "nextQuarter"
  toAttrVal DynFilterNextWeek = "nextWeek"
  toAttrVal DynFilterNextYear = "nextYear"
  toAttrVal DynFilterNull = "null"
  toAttrVal DynFilterQ1 = "Q1"
  toAttrVal DynFilterQ2 = "Q2"
  toAttrVal DynFilterQ3 = "Q3"
  toAttrVal DynFilterQ4 = "Q4"
  toAttrVal DynFilterThisMonth = "thisMonth"
  toAttrVal DynFilterThisQuarter = "thisQuarter"
  toAttrVal DynFilterThisWeek = "thisWeek"
  toAttrVal DynFilterThisYear = "thisYear"
  toAttrVal DynFilterToday = "today"
  toAttrVal DynFilterTomorrow = "tomorrow"
  toAttrVal DynFilterYearToDate = "yearToDate"
  toAttrVal DynFilterYesterday = "yesterday"
