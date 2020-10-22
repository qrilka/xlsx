{-# LANGUAGE CPP #-}
{-# LANGUAGE DeriveGeneric #-}
{-# LANGUAGE OverloadedStrings #-}
{-# LANGUAGE RecordWildCards #-}
{-# LANGUAGE ScopedTypeVariables #-}
{-# LANGUAGE TemplateHaskell #-}
module Codec.Xlsx.Types.ConditionalFormatting
  ( ConditionalFormatting
  , CfRule(..)
  , NStdDev(..)
  , Inclusion(..)
  , CfValue(..)
  , MinCfValue(..)
  , MaxCfValue(..)
  , Condition(..)
  , OperatorExpression(..)
  , TimePeriod(..)
  , IconSetOptions(..)
  , IconSetType(..)
  , DataBarOptions(..)
  , dataBarWithColor
    -- * Lenses
    -- ** CfRule
  , cfrCondition
  , cfrDxfId
  , cfrPriority
  , cfrStopIfTrue
    -- ** IconSetOptions
  , isoIconSet
  , isoValues
  , isoReverse
  , isoShowValue
    -- ** DataBarOptions
  , dboMaxLength
  , dboMinLength
  , dboShowValue
  , dboMinimum
  , dboMaximum
  , dboColor
    -- * Misc
  , topCfPriority
  ) where

import Control.Arrow (first, right)
import Control.DeepSeq (NFData)
#ifdef USE_MICROLENS
import Lens.Micro.TH (makeLenses)
#else
import Control.Lens (makeLenses)
#endif
import Data.Bool (bool)
import Data.ByteString (ByteString)
import Data.Default
import Data.Map (Map)
import qualified Data.Map as M
import Data.Maybe
import Data.Monoid ((<>))
import Data.Text (Text)
import qualified Data.Text as T
import GHC.Generics (Generic)
import Text.XML
import Text.XML.Cursor hiding (bool)
import qualified Xeno.DOM as Xeno

import Codec.Xlsx.Parser.Internal
import Codec.Xlsx.Types.Common
import Codec.Xlsx.Types.StyleSheet (Color)
import Codec.Xlsx.Writer.Internal

-- | Logical operation used in 'CellIs' condition
--
-- See 18.18.15 "ST_ConditionalFormattingOperator
-- (Conditional Format Operators)" (p. 2446)
data OperatorExpression
    = OpBeginsWith Formula         -- ^ 'Begins with' operator
    | OpBetween Formula Formula    -- ^ 'Between' operator
    | OpContainsText Formula       -- ^ 'Contains' operator
    | OpEndsWith Formula           -- ^ 'Ends with' operator
    | OpEqual Formula              -- ^ 'Equal to' operator
    | OpGreaterThan Formula        -- ^ 'Greater than' operator
    | OpGreaterThanOrEqual Formula -- ^ 'Greater than or equal to' operator
    | OpLessThan Formula           -- ^ 'Less than' operator
    | OpLessThanOrEqual Formula    -- ^ 'Less than or equal to' operator
    | OpNotBetween Formula Formula -- ^ 'Not between' operator
    | OpNotContains Formula        -- ^ 'Does not contain' operator
    | OpNotEqual Formula           -- ^ 'Not equal to' operator
    deriving (Eq, Ord, Show, Generic)
instance NFData OperatorExpression

-- | Used in a "contains dates" conditional formatting rule.
-- These are dynamic time periods, which change based on
-- the date the conditional formatting is refreshed / applied.
--
-- See 18.18.82 "ST_TimePeriod (Time Period Types)" (p. 2508)
data TimePeriod
    = PerLast7Days  -- ^ A date in the last seven days.
    | PerLastMonth  -- ^ A date occuring in the last calendar month.
    | PerLastWeek   -- ^ A date occuring last week.
    | PerNextMonth  -- ^ A date occuring in the next calendar month.
    | PerNextWeek   -- ^ A date occuring next week.
    | PerThisMonth  -- ^ A date occuring in this calendar month.
    | PerThisWeek   -- ^ A date occuring this week.
    | PerToday      -- ^ Today's date.
    | PerTomorrow   -- ^ Tomorrow's date.
    | PerYesterday  -- ^ Yesterday's date.
    deriving (Eq, Ord, Show, Generic)
instance NFData TimePeriod

-- | Flag indicating whether the 'aboveAverage' and 'belowAverage'
-- criteria is inclusive of the average itself, or exclusive of that
-- value.
data Inclusion
  = Inclusive
  | Exclusive
  deriving (Eq, Ord, Show, Generic)
instance NFData Inclusion

-- | The number of standard deviations to include above or below the
-- average in the conditional formatting rule.
newtype NStdDev =
  NStdDev Int
  deriving (Eq, Ord, Show, Generic)
instance NFData NStdDev

-- | Conditions which could be used for conditional formatting
--
-- See 18.18.12 "ST_CfType (Conditional Format Type)" (p. 2443)
data Condition
    -- | This conditional formatting rule highlights cells that are
    -- above (or maybe equal to) the average for all values in the range.
    = AboveAverage Inclusion (Maybe NStdDev)
    -- | This conditional formatting rule highlights cells in the
    -- range that begin with the given text. Equivalent to
    -- using the LEFT() sheet function and comparing values.
    | BeginsWith Text
    -- | This conditional formatting rule highlights cells that are
    -- below the average for all values in the range.
    | BelowAverage Inclusion (Maybe NStdDev)
    -- | This conditional formatting rule highlights cells whose
    -- values fall in the bottom N percent bracket.
    | BottomNPercent Int
    -- | This conditional formatting rule highlights cells whose
    -- values fall in the bottom N bracket.
    | BottomNValues Int
    -- | This conditional formatting rule compares a cell value
    -- to a formula calculated result, using an operator.
    | CellIs OperatorExpression
    -- | This conditional formatting rule creates a gradated color
    -- scale on the cells with specified colors for specified minimum
    -- and maximum.
    | ColorScale2 MinCfValue Color MaxCfValue Color
    -- | This conditional formatting rule creates a gradated color
    -- scale on the cells with specified colors for specified minimum,
    -- midpoint and maximum.
    | ColorScale3 MinCfValue Color CfValue Color MaxCfValue Color
    -- | This conditional formatting rule highlights cells that
    -- are completely blank. Equivalent of using LEN(TRIM()).
    -- This means that if the cell contains only characters
    -- that TRIM() would remove, then it is considered blank.
    -- An empty cell is also considered blank.
    | ContainsBlanks
    -- | This conditional formatting rule highlights cells with
    -- formula errors. Equivalent to using ISERROR() sheet
    -- function to determine if there is a formula error.
    | ContainsErrors
    -- | This conditional formatting rule highlights cells
    -- containing given text. Equivalent to using the SEARCH()
    -- sheet function to determine whether the cell contains
    -- the text.
    | ContainsText Text
    -- | This conditional formatting rule displays a gradated data bar
    -- in the range of cells.
    | DataBar DataBarOptions
    -- | This conditional formatting rule highlights cells
    -- without formula errors. Equivalent to using ISERROR()
    -- sheet function to determine if there is a formula error.
    | DoesNotContainErrors
    -- | This conditional formatting rule highlights cells that
    -- are not blank. Equivalent of using LEN(TRIM()). This
    -- means that if the cell contains only characters that
    -- TRIM() would remove, then it is considered blank. An
    -- empty cell is also considered blank.
    | DoesNotContainBlanks
    -- | This conditional formatting rule highlights cells that do
    -- not contain given text. Equivalent to using the
    -- SEARCH() sheet function.
    | DoesNotContainText Text
    -- | This conditional formatting rule highlights duplicated
    -- values.
    | DuplicateValues
    -- | This conditional formatting rule highlights cells ending
    -- with given text. Equivalent to using the RIGHT() sheet
    -- function and comparing values.
    | EndsWith Text
    -- | This conditional formatting rule contains a formula to
    -- evaluate. When the formula result is true, the cell is
    -- highlighted.
    | Expression Formula
    -- | This conditional formatting rule applies icons to cells
    -- according to their values.
    | IconSet IconSetOptions
    -- | This conditional formatting rule highlights cells
    -- containing dates in the specified time period. The
    -- underlying value of the cell is evaluated, therefore the
    -- cell does not need to be formatted as a date to be
    -- evaluated. For example, with a cell containing the
    -- value 38913 the conditional format shall be applied if
    -- the rule requires a value of 7/14/2006.
    | InTimePeriod TimePeriod
    -- | This conditional formatting rule highlights cells whose
    -- values fall in the top N percent bracket.
    | TopNPercent Int
    -- | This conditional formatting rule highlights cells whose
    -- values fall in the top N bracket.
    | TopNValues Int
    -- | This conditional formatting rule highlights unique values in the range.
    | UniqueValues
    deriving (Eq, Ord, Show, Generic)
instance NFData Condition

-- | Describes the values of the interpolation points in a color
-- scale, data bar or icon set conditional formatting rules.
--
-- See 18.3.1.11 "cfvo (Conditional Format Value Object)" (p. 1604)
data CfValue
  = CfValue Double
  | CfPercent Double
  | CfPercentile Double
  | CfFormula Formula
  deriving (Eq, Ord, Show, Generic)
instance NFData CfValue

data MinCfValue
  = CfvMin
  | MinCfValue CfValue
  deriving (Eq, Ord, Show, Generic)
instance NFData MinCfValue

data MaxCfValue
  = CfvMax
  | MaxCfValue CfValue
  deriving (Eq, Ord, Show, Generic)
instance NFData MaxCfValue

-- | internal type for (de)serialization
--
-- See 18.18.13 "ST_CfvoType (Conditional Format Value Object Type)" (p. 2445)
data CfvType =
  CfvtFormula
  -- ^ The minimum\/ midpoint \/ maximum value for the gradient is
  -- determined by a formula.
  | CfvtMax
  -- ^ Indicates that the maximum value in the range shall be used as
  -- the maximum value for the gradient.
  | CfvtMin
  -- ^ Indicates that the minimum value in the range shall be used as
  -- the minimum value for the gradient.
  | CfvtNum
  -- ^ Indicates that the minimum \/ midpoint \/ maximum value for the
  -- gradient is specified by a constant numeric value.
  | CfvtPercent
  -- ^ Value indicates a percentage between the minimum and maximum
  -- values in the range shall be used as the minimum \/ midpoint \/
  -- maximum value for the gradient.
  | CfvtPercentile
  -- ^ Value indicates a percentile ranking in the range shall be used
  -- as the minimum \/ midpoint \/ maximum value for the gradient.
  deriving (Eq, Ord, Show, Generic)
instance NFData CfvType

-- | Describes an icon set conditional formatting rule.
--
-- See 18.3.1.49 "iconSet (Icon Set)" (p. 1645)
data IconSetOptions = IconSetOptions
  { _isoIconSet :: IconSetType
  -- ^ icon set used, default value is 'IconSet3Trafficlights1'
  , _isoValues :: [CfValue]
  -- ^ values describing per icon ranges
  , _isoReverse :: Bool
  -- ^ reverses the default order of the icons in the specified icon set
  , _isoShowValue :: Bool
  -- ^ indicates whether to show the values of the cells on which this
  -- icon set is applied.
  } deriving (Eq, Ord, Show, Generic)
instance NFData IconSetOptions

-- | Icon set type for conditional formatting. 'CfValue' fields
-- determine lower range bounds. I.e. @IconSet3Signs (CfPercent 0)
-- (CfPercent 33) (CfPercent 67)@ say that 1st icon will be shown for
-- values ranging from 0 to 33 percents, 2nd for 33 to 67 percent and
-- the 3rd one for values from 67 to 100 percent.
--
-- 18.18.42 "ST_IconSetType (Icon Set Type)" (p. 2463)
data IconSetType =
  IconSet3Arrows -- CfValue CfValue CfValue
  | IconSet3ArrowsGray -- CfValue CfValue CfValue
  | IconSet3Flags -- CfValue CfValue CfValue
  | IconSet3Signs -- CfValue CfValue CfValue
  | IconSet3Symbols -- CfValue CfValue CfValue
  | IconSet3Symbols2 -- CfValue CfValue CfValue
  | IconSet3TrafficLights1 -- CfValue CfValue CfValue
  | IconSet3TrafficLights2 -- CfValue CfValue CfValue
  -- ^ 3 traffic lights icon set with thick black border.
  | IconSet4Arrows -- CfValue CfValue CfValue CfValue
  | IconSet4ArrowsGray -- CfValue CfValue CfValue CfValue
  | IconSet4Rating -- CfValue CfValue CfValue CfValue
  | IconSet4RedToBlack -- CfValue CfValue CfValue CfValue
  | IconSet4TrafficLights -- CfValue CfValue CfValue CfValue
  | IconSet5Arrows -- CfValue CfValue CfValue CfValue CfValue
  | IconSet5ArrowsGray -- CfValue CfValue CfValue CfValue CfValue
  | IconSet5Quarters -- CfValue CfValue CfValue CfValue CfValue
  | IconSet5Rating -- CfValue CfValue CfValue CfValue CfValue
  deriving  (Eq, Ord, Show, Generic)
instance NFData IconSetType

-- | Describes a data bar conditional formatting rule.
--
-- See 18.3.1.28 "dataBar (Data Bar)" (p. 1621)
data DataBarOptions = DataBarOptions
  { _dboMaxLength :: Int
  -- ^ The maximum length of the data bar, as a percentage of the cell
  -- width.
  , _dboMinLength :: Int
  -- ^ The minimum length of the data bar, as a percentage of the cell
  -- width.
  , _dboShowValue :: Bool
  -- ^ Indicates whether to show the values of the cells on which this
  -- data bar is applied.
  , _dboMinimum :: MinCfValue
  , _dboMaximum :: MaxCfValue
  , _dboColor :: Color
  } deriving (Eq, Ord, Show, Generic)
instance NFData DataBarOptions

defaultDboMaxLength :: Int
defaultDboMaxLength = 90

defaultDboMinLength :: Int
defaultDboMinLength = 10

dataBarWithColor :: Color -> Condition
dataBarWithColor c =
  DataBar
    DataBarOptions
    { _dboMaxLength = defaultDboMaxLength
    , _dboMinLength = defaultDboMinLength
    , _dboShowValue = True
    , _dboMinimum = CfvMin
    , _dboMaximum = CfvMax
    , _dboColor = c
    }

-- | This collection represents a description of a conditional formatting rule.
--
-- See 18.3.1.10 "cfRule (Conditional Formatting Rule)" (p. 1602)
data CfRule = CfRule
    { _cfrCondition  :: Condition
    -- | This is an index to a dxf element in the Styles Part
    -- indicating which cell formatting to
    -- apply when the conditional formatting rule criteria is met.
    , _cfrDxfId      :: Maybe Int
    -- | The priority of this conditional formatting rule. This value
    -- is used to determine which format should be evaluated and
    -- rendered. Lower numeric values are higher priority than
    -- higher numeric values, where 1 is the highest priority.
    , _cfrPriority   :: Int
    -- | If this flag is set, no rules with lower priority shall
    -- be applied over this rule, when this rule
    -- evaluates to true.
    , _cfrStopIfTrue :: Maybe Bool
    } deriving (Eq, Ord, Show, Generic)
instance NFData CfRule

instance Default IconSetOptions where
  def =
    IconSetOptions
    { _isoIconSet = IconSet3TrafficLights1
    , _isoValues = [CfPercent 0, CfPercent 33.33, CfPercent 66.67]
--        IconSet3TrafficLights1 (CfPercent 0) (CfPercent 33.33) (CfPercent 66.67)
    , _isoReverse = False
    , _isoShowValue = True
    }

makeLenses ''CfRule
makeLenses ''IconSetOptions
makeLenses ''DataBarOptions

type ConditionalFormatting = [CfRule]

topCfPriority :: Int
topCfPriority = 1

{-------------------------------------------------------------------------------
  Parsing
-------------------------------------------------------------------------------}

instance FromCursor CfRule where
    fromCursor cur = do
        _cfrDxfId      <- maybeAttribute "dxfId" cur
        _cfrPriority   <- fromAttribute "priority" cur
        _cfrStopIfTrue <- maybeAttribute "stopIfTrue" cur
        -- spec shows this attribute as optional but it's not clear why could
        -- conditional formatting record be needed with no condition type set
        cfType <- fromAttribute "type" cur
        _cfrCondition <- readCondition cfType cur
        return CfRule{..}

readCondition :: Text -> Cursor -> [Condition]
readCondition "aboveAverage" cur       = do
  above <- fromAttributeDef "aboveAverage" True cur
  inclusion <- fromAttributeDef "equalAverage" Exclusive cur
  nStdDev <- maybeAttribute "stdDev" cur
  if above
    then return $ AboveAverage inclusion nStdDev
    else return $ BelowAverage inclusion nStdDev
readCondition "beginsWith" cur = do
  txt <- fromAttribute "text" cur
  return $ BeginsWith txt
readCondition "colorScale" cur = do
  let cfvos = cur $/ element (n_ "colorScale") &/ element (n_ "cfvo") &| node
      colors = cur $/ element (n_ "colorScale") &/ element (n_ "color") &| node
  case (cfvos, colors) of
    ([n1, n2], [cn1, cn2]) -> do
      mincfv <- fromCursor $ fromNode n1
      minc <- fromCursor $ fromNode cn1
      maxcfv <- fromCursor $ fromNode n2
      maxc <- fromCursor $ fromNode cn2
      return $ ColorScale2 mincfv minc maxcfv maxc
    ([n1, n2, n3], [cn1, cn2, cn3]) -> do
      mincfv <- fromCursor $ fromNode n1
      minc <- fromCursor $ fromNode cn1
      midcfv <- fromCursor $ fromNode n2
      midc <- fromCursor $ fromNode cn2
      maxcfv <- fromCursor $ fromNode n3
      maxc <- fromCursor $ fromNode cn3
      return $ ColorScale3 mincfv minc midcfv midc maxcfv maxc
    _ ->
      error "Malformed colorScale condition"
readCondition "cellIs" cur           = do
    operator <- fromAttribute "operator" cur
    let formulas = cur $/ element (n_ "formula") >=> fromCursor
    expr <- readOpExpression operator formulas
    return $ CellIs expr
readCondition "containsBlanks" _     = return ContainsBlanks
readCondition "containsErrors" _     = return ContainsErrors
readCondition "containsText" cur     = do
    txt <- fromAttribute "text" cur
    return $ ContainsText txt
readCondition "dataBar" cur = fmap DataBar $ cur $/ element (n_ "dataBar") >=> fromCursor
readCondition "duplicateValues" _    = return DuplicateValues
readCondition "endsWith" cur         = do
    txt <- fromAttribute "text" cur
    return $ EndsWith txt
readCondition "expression" cur       = do
    formula <- cur $/ element (n_ "formula") >=> fromCursor
    return $ Expression formula
readCondition "iconSet" cur = fmap IconSet $ cur $/ element (n_ "iconSet") >=> fromCursor
readCondition "notContainsBlanks" _  = return DoesNotContainBlanks
readCondition "notContainsErrors" _  = return DoesNotContainErrors
readCondition "notContainsText" cur  = do
    txt <- fromAttribute "text" cur
    return $ DoesNotContainText txt
readCondition "timePeriod" cur  = do
    period <- fromAttribute "timePeriod" cur
    return $ InTimePeriod period
readCondition "top10" cur = do
  bottom <- fromAttributeDef "bottom" False cur
  percent <- fromAttributeDef "percent" False cur
  rank <- fromAttribute "rank" cur
  case (bottom, percent) of
    (True, True) -> return $ BottomNPercent rank
    (True, False) -> return $ BottomNValues rank
    (False, True) -> return $ TopNPercent rank
    (False, False) -> return $ TopNValues rank
readCondition "uniqueValues" _       = return UniqueValues
readCondition t _                    = error $ "Unexpected conditional formatting type " ++ show t

readOpExpression :: Text -> [Formula] -> [OperatorExpression]
readOpExpression "beginsWith" [f]         = [OpBeginsWith f ]
readOpExpression "between" [f1, f2]       = [OpBetween f1 f2]
readOpExpression "containsText" [f]       = [OpContainsText f]
readOpExpression "endsWith" [f]           = [OpEndsWith f]
readOpExpression "equal" [f]              = [OpEqual f]
readOpExpression "greaterThan" [f]        = [OpGreaterThan f]
readOpExpression "greaterThanOrEqual" [f] = [OpGreaterThanOrEqual f]
readOpExpression "lessThan" [f]           = [OpLessThan f]
readOpExpression "lessThanOrEqual" [f]    = [OpLessThanOrEqual f]
readOpExpression "notBetween" [f1, f2]    = [OpNotBetween f1 f2]
readOpExpression "notContains" [f]        = [OpNotContains f]
readOpExpression "notEqual" [f]           = [OpNotEqual f]
readOpExpression _ _                      = []

instance FromXenoNode CfRule where
  fromXenoNode root = parseAttributes root $ do
        _cfrDxfId <- maybeAttr "dxfId"
        _cfrPriority <- fromAttr "priority"
        _cfrStopIfTrue <- maybeAttr "stopIfTrue"
        -- spec shows this attribute as optional but it's not clear why could
        -- conditional formatting record be needed with no condition type set
        cfType <- fromAttr "type"
        _cfrCondition <- readConditionX cfType
        return CfRule {..}
    where
      readConditionX ("aboveAverage" :: ByteString) = do
        above <- fromAttrDef "aboveAverage" True
        inclusion <- fromAttrDef "equalAverage" Exclusive
        nStdDev <- maybeAttr "stdDev"
        if above
          then return $ AboveAverage inclusion nStdDev
          else return $ BelowAverage inclusion nStdDev
      readConditionX "beginsWith" = BeginsWith <$> fromAttr "text"
      readConditionX "colorScale" = toAttrParser $ do
        xs <- collectChildren root . maybeParse "colorScale" $ \node ->
          collectChildren node $ (,) <$> childList "cfvo"
                                     <*> childList "color"
        case xs of
          Just ([n1, n2], [cn1, cn2]) -> do
            mincfv <- fromXenoNode n1
            minc <- fromXenoNode cn1
            maxcfv <- fromXenoNode n2
            maxc <- fromXenoNode cn2
            return $ ColorScale2 mincfv minc maxcfv maxc
          Just ([n1, n2, n3], [cn1, cn2, cn3]) -> do
            mincfv <- fromXenoNode n1
            minc <- fromXenoNode cn1
            midcfv <- fromXenoNode n2
            midc <- fromXenoNode cn2
            maxcfv <- fromXenoNode n3
            maxc <- fromXenoNode cn3
            return $ ColorScale3 mincfv minc midcfv midc maxcfv maxc
          _ ->
            Left "Malformed colorScale condition"
      readConditionX "cellIs" = do
        operator <- fromAttr "operator"
        formulas <- toAttrParser . collectChildren root $ fromChildList "formula"
        case (operator, formulas) of
          ("beginsWith" :: ByteString, [f]) -> return . CellIs $ OpBeginsWith f
          ("between", [f1, f2]) -> return . CellIs $ OpBetween f1 f2
          ("containsText", [f]) -> return . CellIs $ OpContainsText f
          ("endsWith", [f]) -> return . CellIs $ OpEndsWith f
          ("equal", [f]) -> return . CellIs $ OpEqual f
          ("greaterThan", [f]) -> return . CellIs $ OpGreaterThan f
          ("greaterThanOrEqual", [f]) -> return . CellIs $ OpGreaterThanOrEqual f
          ("lessThan", [f]) -> return . CellIs $ OpLessThan f
          ("lessThanOrEqual", [f]) -> return . CellIs $ OpLessThanOrEqual f
          ("notBetween", [f1, f2]) -> return . CellIs $ OpNotBetween f1 f2
          ("notContains", [f]) -> return . CellIs $ OpNotContains f
          ("notEqual", [f]) -> return . CellIs $ OpNotEqual f
          _ -> toAttrParser $ Left "Bad cellIs rule"
      readConditionX "containsBlanks" = return ContainsBlanks
      readConditionX "containsErrors" = return ContainsErrors
      readConditionX "containsText" = ContainsText <$> fromAttr "text"
      readConditionX "dataBar" =
        fmap DataBar . toAttrParser . collectChildren root $ fromChild "dataBar"
      readConditionX "duplicateValues" = return DuplicateValues
      readConditionX "endsWith" = EndsWith <$> fromAttr "text"
      readConditionX "expression" =
        fmap Expression . toAttrParser . collectChildren root $ fromChild "formula"
      readConditionX "iconSet" =
        fmap IconSet . toAttrParser . collectChildren root $ fromChild "iconSet"
      readConditionX "notContainsBlanks" = return DoesNotContainBlanks
      readConditionX "notContainsErrors" = return DoesNotContainErrors
      readConditionX "notContainsText" =
        DoesNotContainText <$> fromAttr "text"
      readConditionX "timePeriod" = InTimePeriod <$> fromAttr "timePeriod"
      readConditionX "top10" = do
        bottom <- fromAttrDef "bottom" False
        percent <- fromAttrDef "percent" False
        rank <- fromAttr "rank"
        case (bottom, percent) of
          (True, True) -> return $ BottomNPercent rank
          (True, False) -> return $ BottomNValues rank
          (False, True) -> return $ TopNPercent rank
          (False, False) -> return $ TopNValues rank
      readConditionX "uniqueValues" = return UniqueValues
      readConditionX x =
        toAttrParser . Left $ "Unexpected conditional formatting type " <> T.pack (show x)

instance FromAttrVal TimePeriod where
    fromAttrVal "last7Days" = readSuccess PerLast7Days
    fromAttrVal "lastMonth" = readSuccess PerLastMonth
    fromAttrVal "lastWeek"  = readSuccess PerLastWeek
    fromAttrVal "nextMonth" = readSuccess PerNextMonth
    fromAttrVal "nextWeek"  = readSuccess PerNextWeek
    fromAttrVal "thisMonth" = readSuccess PerThisMonth
    fromAttrVal "thisWeek"  = readSuccess PerThisWeek
    fromAttrVal "today"     = readSuccess PerToday
    fromAttrVal "tomorrow"  = readSuccess PerTomorrow
    fromAttrVal "yesterday" = readSuccess PerYesterday
    fromAttrVal t           = invalidText "TimePeriod" t

instance FromAttrBs TimePeriod where
    fromAttrBs "last7Days" = return PerLast7Days
    fromAttrBs "lastMonth" = return PerLastMonth
    fromAttrBs "lastWeek"  = return PerLastWeek
    fromAttrBs "nextMonth" = return PerNextMonth
    fromAttrBs "nextWeek"  = return PerNextWeek
    fromAttrBs "thisMonth" = return PerThisMonth
    fromAttrBs "thisWeek"  = return PerThisWeek
    fromAttrBs "today"     = return PerToday
    fromAttrBs "tomorrow"  = return PerTomorrow
    fromAttrBs "yesterday" = return PerYesterday
    fromAttrBs x           = unexpectedAttrBs "TimePeriod" x

instance FromAttrVal CfvType where
  fromAttrVal "num"        = readSuccess CfvtNum
  fromAttrVal "percent"    = readSuccess CfvtPercent
  fromAttrVal "max"        = readSuccess CfvtMax
  fromAttrVal "min"        = readSuccess CfvtMin
  fromAttrVal "formula"    = readSuccess CfvtFormula
  fromAttrVal "percentile" = readSuccess CfvtPercentile
  fromAttrVal t            = invalidText "CfvType" t

instance FromAttrBs CfvType where
  fromAttrBs "num"        = return CfvtNum
  fromAttrBs "percent"    = return CfvtPercent
  fromAttrBs "max"        = return CfvtMax
  fromAttrBs "min"        = return CfvtMin
  fromAttrBs "formula"    = return CfvtFormula
  fromAttrBs "percentile" = return CfvtPercentile
  fromAttrBs x            = unexpectedAttrBs "CfvType" x

readCfValue :: (CfValue -> a) -> [a] -> [a] -> Cursor -> [a]
readCfValue f minVal maxVal c = do
  vType <- fromAttribute "type" c
  case vType of
    CfvtNum -> do
      v <- fromAttribute "val" c
      return . f $ CfValue v
    CfvtFormula -> do
      v <- fromAttribute "val" c
      return . f $ CfFormula v
    CfvtPercent -> do
      v <- fromAttribute "val" c
      return . f $ CfPercent v
    CfvtPercentile -> do
      v <- fromAttribute "val" c
      return . f $ CfPercentile v
    CfvtMin -> minVal
    CfvtMax -> maxVal

readCfValueX ::
     (CfValue -> a)
  -> Either Text a
  -> Either Text a
  -> Xeno.Node
  -> Either Text a
readCfValueX f minVal maxVal root =
  parseAttributes root $ do
    vType <- fromAttr "type"
    case vType of
      CfvtNum -> do
        v <- fromAttr "val"
        return . f $ CfValue v
      CfvtFormula -> do
        v <- fromAttr "val"
        return . f $ CfFormula v
      CfvtPercent -> do
        v <- fromAttr "val"
        return . f $ CfPercent v
      CfvtPercentile -> do
        v <- fromAttr "val"
        return . f $ CfPercentile v
      CfvtMin -> toAttrParser minVal
      CfvtMax -> toAttrParser maxVal

failMinCfvType :: [a]
failMinCfvType = fail "unexpected 'min' type"

failMinCfvTypeX :: Either Text a
failMinCfvTypeX = Left "unexpected 'min' type"

failMaxCfvType :: [a]
failMaxCfvType = fail "unexpected 'max' type"

failMaxCfvTypeX :: Either Text a
failMaxCfvTypeX = Left "unexpected 'max' type"

instance FromCursor CfValue where
  fromCursor = readCfValue id failMinCfvType failMaxCfvType

instance FromXenoNode CfValue where
  fromXenoNode root = readCfValueX id failMinCfvTypeX failMaxCfvTypeX root

instance FromCursor MinCfValue where
  fromCursor = readCfValue MinCfValue (return CfvMin) failMaxCfvType

instance FromXenoNode MinCfValue where
  fromXenoNode root =
    readCfValueX MinCfValue (return CfvMin) failMaxCfvTypeX root

instance FromCursor MaxCfValue where
  fromCursor = readCfValue MaxCfValue failMinCfvType (return CfvMax)

instance FromXenoNode MaxCfValue where
  fromXenoNode root =
    readCfValueX MaxCfValue failMinCfvTypeX (return CfvMax) root

defaultIconSet :: IconSetType
defaultIconSet =  IconSet3TrafficLights1

instance FromCursor IconSetOptions where
  fromCursor cur = do
    _isoIconSet <- fromAttributeDef "iconSet" defaultIconSet cur
    let _isoValues = cur $/ element (n_ "cfvo") >=> fromCursor
    _isoReverse <- fromAttributeDef "reverse" False cur
    _isoShowValue <- fromAttributeDef "showValue" True cur
    return IconSetOptions {..}

instance FromXenoNode IconSetOptions where
  fromXenoNode root = do
    (_isoIconSet, _isoReverse, _isoShowValue) <-
      parseAttributes root $ (,,) <$> fromAttrDef "iconSet" defaultIconSet
                                  <*> fromAttrDef "reverse" False
                                  <*> fromAttrDef "showValue" True
    _isoValues <- collectChildren root $ fromChildList "cfvo"
    return IconSetOptions {..}

instance FromAttrVal IconSetType where
  fromAttrVal "3Arrows" = readSuccess IconSet3Arrows
  fromAttrVal "3ArrowsGray" = readSuccess IconSet3ArrowsGray
  fromAttrVal "3Flags" = readSuccess IconSet3Flags
  fromAttrVal "3Signs" = readSuccess IconSet3Signs
  fromAttrVal "3Symbols" = readSuccess IconSet3Symbols
  fromAttrVal "3Symbols2" = readSuccess IconSet3Symbols2
  fromAttrVal "3TrafficLights1" = readSuccess IconSet3TrafficLights1
  fromAttrVal "3TrafficLights2" = readSuccess IconSet3TrafficLights2
  fromAttrVal "4Arrows" = readSuccess IconSet4Arrows
  fromAttrVal "4ArrowsGray" = readSuccess IconSet4ArrowsGray
  fromAttrVal "4Rating" = readSuccess IconSet4Rating
  fromAttrVal "4RedToBlack" = readSuccess IconSet4RedToBlack
  fromAttrVal "4TrafficLights" = readSuccess IconSet4TrafficLights
  fromAttrVal "5Arrows" = readSuccess IconSet5Arrows
  fromAttrVal "5ArrowsGray" = readSuccess IconSet5ArrowsGray
  fromAttrVal "5Quarters" = readSuccess IconSet5Quarters
  fromAttrVal "5Rating" = readSuccess IconSet5Rating
  fromAttrVal t = invalidText "IconSetType" t

instance FromAttrBs IconSetType where
  fromAttrBs "3Arrows" = return IconSet3Arrows
  fromAttrBs "3ArrowsGray" = return IconSet3ArrowsGray
  fromAttrBs "3Flags" = return IconSet3Flags
  fromAttrBs "3Signs" = return IconSet3Signs
  fromAttrBs "3Symbols" = return IconSet3Symbols
  fromAttrBs "3Symbols2" = return IconSet3Symbols2
  fromAttrBs "3TrafficLights1" = return IconSet3TrafficLights1
  fromAttrBs "3TrafficLights2" = return IconSet3TrafficLights2
  fromAttrBs "4Arrows" = return IconSet4Arrows
  fromAttrBs "4ArrowsGray" = return IconSet4ArrowsGray
  fromAttrBs "4Rating" = return IconSet4Rating
  fromAttrBs "4RedToBlack" = return IconSet4RedToBlack
  fromAttrBs "4TrafficLights" = return IconSet4TrafficLights
  fromAttrBs "5Arrows" = return IconSet5Arrows
  fromAttrBs "5ArrowsGray" = return IconSet5ArrowsGray
  fromAttrBs "5Quarters" = return IconSet5Quarters
  fromAttrBs "5Rating" = return IconSet5Rating
  fromAttrBs x = unexpectedAttrBs "IconSetType" x

instance FromCursor DataBarOptions where
  fromCursor cur = do
    _dboMaxLength <- fromAttributeDef "maxLength" defaultDboMaxLength cur
    _dboMinLength <- fromAttributeDef "minLength" defaultDboMinLength cur
    _dboShowValue <- fromAttributeDef "showValue" True cur
    let cfvos = cur $/ element (n_ "cfvo") &| node
    case cfvos of
      [nMin, nMax] -> do
        _dboMinimum <- fromCursor (fromNode nMin)
        _dboMaximum <- fromCursor (fromNode nMax)
        _dboColor <- cur $/ element (n_ "color") >=> fromCursor
        return DataBarOptions{..}
      ns -> do
        fail $ "expected minimum and maximum cfvo nodes but see instead " ++
          show (length ns) ++ " cfvo nodes"

instance FromXenoNode DataBarOptions where
  fromXenoNode root = do
    (_dboMaxLength, _dboMinLength, _dboShowValue) <-
      parseAttributes root $ (,,) <$> fromAttrDef "maxLength" defaultDboMaxLength
                                  <*> fromAttrDef "minLength" defaultDboMinLength
                                  <*> fromAttrDef "showValue" True
    (_dboMinimum, _dboMaximum, _dboColor) <-
      collectChildren root $ (,,) <$> fromChild "cfvo"
                                  <*> fromChild "cfvo"
                                  <*> fromChild "color"
    return DataBarOptions{..}

instance FromAttrVal Inclusion where
  fromAttrVal = right (first $ bool Exclusive Inclusive) . fromAttrVal

instance FromAttrBs Inclusion where
  fromAttrBs = fmap (bool Exclusive Inclusive) . fromAttrBs

instance FromAttrVal NStdDev where
  fromAttrVal = right (first NStdDev) . fromAttrVal

instance FromAttrBs NStdDev where
  fromAttrBs = fmap NStdDev . fromAttrBs

{-------------------------------------------------------------------------------
  Rendering
-------------------------------------------------------------------------------}

instance ToElement CfRule where
    toElement nm CfRule{..} =
        let (condType, condAttrs, condNodes) = conditionData _cfrCondition
            baseAttrs = M.fromList . catMaybes $
                [ Just $ "type"       .=  condType
                ,        "dxfId"      .=? _cfrDxfId
                , Just $ "priority"   .=  _cfrPriority
                ,        "stopIfTrue" .=? _cfrStopIfTrue
                ]
        in Element
           { elementName = nm
           , elementAttributes = M.union baseAttrs condAttrs
           , elementNodes = condNodes
           }

conditionData :: Condition -> (Text, Map Name Text, [Node])
conditionData (AboveAverage i sDevs) =
  ("aboveAverage", M.fromList $ ["aboveAverage" .= True] ++
                   catMaybes [ "equalAverage" .=? justNonDef Exclusive i
                             , "stdDev" .=? sDevs], [])
conditionData (BeginsWith t)         = ("beginsWith", M.fromList [ "text" .= t], [])
conditionData (BelowAverage i sDevs) =
  ("aboveAverage", M.fromList $ ["aboveAverage" .= False] ++
                   catMaybes [ "equalAverage" .=? justNonDef Exclusive i
                             , "stdDev" .=? sDevs], [])
conditionData (BottomNPercent n)     = ("top10", M.fromList [ "bottom" .= True, "rank" .= n, "percent" .= True ], [])
conditionData (BottomNValues n)      = ("top10", M.fromList [ "bottom" .= True, "rank" .= n ], [])
conditionData (CellIs opExpr)        = ("cellIs", M.fromList [ "operator" .= op], formulas)
    where (op, formulas) = operatorExpressionData opExpr
conditionData (ColorScale2 minv minc maxv maxc) =
  ( "colorScale"
  , M.empty
  , [ NodeElement $
      elementListSimple
        "colorScale"
        [ toElement "cfvo" minv
        , toElement "cfvo" maxv
        , toElement "color" minc
        , toElement "color" maxc
        ]
    ])
conditionData (ColorScale3 minv minc midv midc maxv maxc) =
  ( "colorScale"
  , M.empty
  , [ NodeElement $
      elementListSimple
        "colorScale"
        [ toElement "cfvo" minv
        , toElement "cfvo" midv
        , toElement "cfvo" maxv
        , toElement "color" minc
        , toElement "color" midc
        , toElement "color" maxc
        ]
    ])
conditionData ContainsBlanks         = ("containsBlanks", M.empty, [])
conditionData ContainsErrors         = ("containsErrors", M.empty, [])
conditionData (ContainsText t)       = ("containsText", M.fromList [ "text" .= t], [])
conditionData (DataBar dbOpts)       = ("dataBar", M.empty, [toNode "dataBar" dbOpts])
conditionData DoesNotContainBlanks   = ("notContainsBlanks", M.empty, [])
conditionData DoesNotContainErrors   = ("notContainsErrors", M.empty, [])
conditionData (DoesNotContainText t) = ("notContainsText", M.fromList [ "text" .= t], [])
conditionData DuplicateValues        = ("duplicateValues", M.empty, [])
conditionData (EndsWith t)           = ("endsWith", M.fromList [ "text" .= t], [])
conditionData (Expression formula)   = ("expression", M.empty, [formulaNode formula])
conditionData (InTimePeriod period)  = ("timePeriod", M.fromList [ "timePeriod" .= period ], [])
conditionData (IconSet isOptions)    = ("iconSet", M.empty, [toNode "iconSet" isOptions])
conditionData (TopNPercent n)        = ("top10", M.fromList [ "rank" .= n, "percent" .= True ], [])
conditionData (TopNValues n)         = ("top10", M.fromList [ "rank" .= n ], [])
conditionData UniqueValues           = ("uniqueValues", M.empty, [])

operatorExpressionData :: OperatorExpression -> (Text, [Node])
operatorExpressionData (OpBeginsWith f)          = ("beginsWith", [formulaNode f])
operatorExpressionData (OpBetween f1 f2)         = ("between", [formulaNode f1, formulaNode f2])
operatorExpressionData (OpContainsText f)        = ("containsText", [formulaNode f])
operatorExpressionData (OpEndsWith f)            = ("endsWith", [formulaNode f])
operatorExpressionData (OpEqual f)               = ("equal", [formulaNode f])
operatorExpressionData (OpGreaterThan f)         = ("greaterThan", [formulaNode f])
operatorExpressionData (OpGreaterThanOrEqual f)  = ("greaterThanOrEqual", [formulaNode f])
operatorExpressionData (OpLessThan f)            = ("lessThan", [formulaNode f])
operatorExpressionData (OpLessThanOrEqual f)     = ("lessThanOrEqual", [formulaNode f])
operatorExpressionData (OpNotBetween f1 f2)      = ("notBetween", [formulaNode f1, formulaNode f2])
operatorExpressionData (OpNotContains f)         = ("notContains", [formulaNode f])
operatorExpressionData (OpNotEqual f)            = ("notEqual", [formulaNode  f])

instance ToElement MinCfValue where
  toElement nm CfvMin = leafElement nm ["type" .= CfvtMin]
  toElement nm (MinCfValue cfv) = toElement nm cfv

instance ToElement MaxCfValue where
  toElement nm CfvMax = leafElement nm ["type" .= CfvtMax]
  toElement nm (MaxCfValue cfv) = toElement nm cfv

instance ToElement CfValue where
  toElement nm (CfValue v) = leafElement nm ["type" .= CfvtNum, "val" .= v]
  toElement nm (CfPercent v) =
    leafElement nm ["type" .= CfvtPercent, "val" .= v]
  toElement nm (CfPercentile v) =
    leafElement nm ["type" .= CfvtPercentile, "val" .= v]
  toElement nm (CfFormula f) =
    leafElement nm ["type" .= CfvtFormula, "val" .= unFormula f]

instance ToAttrVal CfvType where
  toAttrVal CfvtNum = "num"
  toAttrVal CfvtPercent = "percent"
  toAttrVal CfvtMax = "max"
  toAttrVal CfvtMin = "min"
  toAttrVal CfvtFormula = "formula"
  toAttrVal CfvtPercentile = "percentile"

instance ToElement IconSetOptions where
  toElement nm IconSetOptions {..} =
    elementList nm attrs $ map (toElement "cfvo") _isoValues
    where
      attrs = catMaybes
        [ "iconSet" .=? justNonDef defaultIconSet _isoIconSet
        , "reverse" .=? justTrue _isoReverse
        , "showValue" .=? justFalse _isoShowValue
        ]

instance ToAttrVal IconSetType where
  toAttrVal IconSet3Arrows = "3Arrows"
  toAttrVal IconSet3ArrowsGray = "3ArrowsGray"
  toAttrVal IconSet3Flags = "3Flags"
  toAttrVal IconSet3Signs = "3Signs"
  toAttrVal IconSet3Symbols = "3Symbols"
  toAttrVal IconSet3Symbols2 = "3Symbols2"
  toAttrVal IconSet3TrafficLights1 = "3TrafficLights1"
  toAttrVal IconSet3TrafficLights2 = "3TrafficLights2"
  toAttrVal IconSet4Arrows = "4Arrows"
  toAttrVal IconSet4ArrowsGray = "4ArrowsGray"
  toAttrVal IconSet4Rating = "4Rating"
  toAttrVal IconSet4RedToBlack = "4RedToBlack"
  toAttrVal IconSet4TrafficLights = "4TrafficLights"
  toAttrVal IconSet5Arrows = "5Arrows"
  toAttrVal IconSet5ArrowsGray = "5ArrowsGray"
  toAttrVal IconSet5Quarters = "5Quarters"
  toAttrVal IconSet5Rating = "5Rating"

instance ToElement DataBarOptions where
  toElement nm DataBarOptions {..} = elementList nm attrs elements
    where
      attrs = catMaybes
        [ "maxLength" .=? justNonDef defaultDboMaxLength _dboMaxLength
        , "minLength" .=? justNonDef defaultDboMinLength _dboMinLength
        , "showValue" .=? justFalse _dboShowValue
        ]
      elements =
        [ toElement "cfvo" _dboMinimum
        , toElement "cfvo" _dboMaximum
        , toElement "color" _dboColor
        ]

toNode :: ToElement a => Name -> a -> Node
toNode nm = NodeElement . toElement nm

formulaNode :: Formula -> Node
formulaNode = toNode "formula"

instance ToAttrVal TimePeriod where
    toAttrVal PerLast7Days = "last7Days"
    toAttrVal PerLastMonth = "lastMonth"
    toAttrVal PerLastWeek  = "lastWeek"
    toAttrVal PerNextMonth = "nextMonth"
    toAttrVal PerNextWeek  = "nextWeek"
    toAttrVal PerThisMonth = "thisMonth"
    toAttrVal PerThisWeek  = "thisWeek"
    toAttrVal PerToday     = "today"
    toAttrVal PerTomorrow  = "tomorrow"
    toAttrVal PerYesterday = "yesterday"

instance ToAttrVal Inclusion where
  toAttrVal = toAttrVal . (== Inclusive)

instance ToAttrVal NStdDev where
  toAttrVal (NStdDev n) = toAttrVal n
