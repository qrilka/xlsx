{-# LANGUAGE CPP #-}
{-# LANGUAGE OverloadedStrings #-}
{-# LANGUAGE RecordWildCards   #-}
{-# LANGUAGE TemplateHaskell   #-}
{-# LANGUAGE DeriveGeneric #-}
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

import Data.Bool (bool)
import Control.Arrow (first, right)
import Control.Lens (makeLenses)
import Data.Default
import Data.Map (Map)
import qualified Data.Map as M
import Data.Maybe
import Data.Text (Text)
import GHC.Generics (Generic)
import Text.XML
import Text.XML.Cursor hiding (bool)

#if !MIN_VERSION_base(4,8,0)
import Control.Applicative
#endif

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

-- | Flag indicating whether the 'aboveAverage' and 'belowAverage'
-- criteria is inclusive of the average itself, or exclusive of that
-- value.
data Inclusion
  = Inclusive
  | Exclusive
  deriving (Eq, Ord, Show, Generic)

-- | The number of standard deviations to include above or below the
-- average in the conditional formatting rule.
newtype NStdDev =
  NStdDev Int
  deriving (Eq, Ord, Show, Generic)

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

data MinCfValue
  = CfvMin
  | MinCfValue CfValue
  deriving (Eq, Ord, Show, Generic)

data MaxCfValue
  = CfvMax
  | MaxCfValue CfValue
  deriving (Eq, Ord, Show, Generic)

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

-- | Describes an icon set conditional formatting rule.
--
-- See 18.3.1.49 "iconSet (Icon Set)" (p. 1645)
data IconSetOptions = IconSetOptions
  { _isoIconSet :: IconSetType
  -- ^ icon set used, default value is 'IconSet3Trafficlights1'
  , _isoReverse :: Bool
  -- ^ reverses the default order of the icons in the specified icon set
  , _isoShowValue :: Bool
  -- ^ indicates whether to show the values of the cells on which this
  -- icon set is applied.
  } deriving (Eq, Ord, Show, Generic)

-- | Icon set type for conditional formatting. 'CfValue' fields
-- determine lower range bounds. I.e. @IconSet3Signs (CfPercent 0)
-- (CfPercent 33) (CfPercent 67)@ say that 1st icon will be shown for
-- values ranging from 0 to 33 percents, 2nd for 33 to 67 percent and
-- the 3rd one for values from 67 to 100 percent.
--
-- 18.18.42 "ST_IconSetType (Icon Set Type)" (p. 2463)
data IconSetType =
  IconSet3Arrows CfValue CfValue CfValue
  | IconSet3ArrowsGray CfValue CfValue CfValue
  | IconSet3Flags CfValue CfValue CfValue
  | IconSet3Signs CfValue CfValue CfValue
  | IconSet3Symbols CfValue CfValue CfValue
  | IconSet3Symbols2 CfValue CfValue CfValue
  | IconSet3TrafficLights1 CfValue CfValue CfValue
  | IconSet3TrafficLights2 CfValue CfValue CfValue
  -- ^ 3 traffic lights icon set with thick black border.
  | IconSet4Arrows CfValue CfValue CfValue CfValue
  | IconSet4ArrowsGray CfValue CfValue CfValue CfValue
  | IconSet4Rating CfValue CfValue CfValue CfValue
  | IconSet4RedToBlack CfValue CfValue CfValue CfValue
  | IconSet4TrafficLights CfValue CfValue CfValue CfValue
  | IconSet5Arrows CfValue CfValue CfValue CfValue CfValue
  | IconSet5ArrowsGray CfValue CfValue CfValue CfValue CfValue
  | IconSet5Quarters CfValue CfValue CfValue CfValue CfValue
  | IconSet5Rating CfValue CfValue CfValue CfValue CfValue
  deriving  (Eq, Ord, Show, Generic)

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

instance Default IconSetOptions where
  def =
    IconSetOptions
    { _isoIconSet =
        IconSet3TrafficLights1 (CfPercent 0) (CfPercent 33.33) (CfPercent 66.67)
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

instance FromAttrVal CfvType where
  fromAttrVal "num"        = readSuccess CfvtNum
  fromAttrVal "percent"    = readSuccess CfvtPercent
  fromAttrVal "max"        = readSuccess CfvtMax
  fromAttrVal "min"        = readSuccess CfvtMin
  fromAttrVal "formula"    = readSuccess CfvtFormula
  fromAttrVal "percentile" = readSuccess CfvtPercentile
  fromAttrVal t            = invalidText "CfvType" t

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

failMinCfvType :: [a]
failMinCfvType = fail "unexpected 'min' type"

failMaxCfvType :: [a]
failMaxCfvType = fail "unexpected 'max' type"

instance FromCursor CfValue where
  fromCursor = readCfValue id failMinCfvType failMaxCfvType

instance FromCursor MinCfValue where
  fromCursor = readCfValue MinCfValue (return CfvMin) failMaxCfvType

instance FromCursor MaxCfValue where
  fromCursor = readCfValue MaxCfValue failMinCfvType (return CfvMax)

defaultIconSet :: Text
defaultIconSet =  "3TrafficLights1"

instance FromCursor IconSetOptions where
  fromCursor cur = do
    isType <- fromAttributeDef "iconSet" defaultIconSet cur
    _isoReverse <- fromAttributeDef "reverse" False cur
    _isoShowValue <- fromAttributeDef "showValue" True cur
    let cfvoEls = cur $/ element (n_ "cfvo")
    _isoIconSet <- case (isType, cfvoEls) of
      ("3Arrows", [e1, e2, e3]) ->
        IconSet3Arrows <$> fromCursor e1 <*> fromCursor e2 <*> fromCursor e3
      ("3ArrowsGray", [e1, e2, e3]) ->
        IconSet3ArrowsGray <$> fromCursor e1 <*> fromCursor e2 <*> fromCursor e3
      ("3Flags", [e1, e2, e3]) ->
        IconSet3Flags <$> fromCursor e1 <*> fromCursor e2 <*> fromCursor e3
      ("3Signs", [e1, e2, e3]) ->
        IconSet3Signs <$> fromCursor e1 <*> fromCursor e2 <*> fromCursor e3
      ("3Symbols", [e1, e2, e3]) ->
        IconSet3Symbols <$> fromCursor e1 <*> fromCursor e2 <*> fromCursor e3
      ("3Symbols2", [e1, e2, e3]) ->
        IconSet3Symbols2 <$> fromCursor e1 <*> fromCursor e2 <*> fromCursor e3
      ("3TrafficLights1", [e1, e2, e3]) ->
        IconSet3TrafficLights1 <$> fromCursor e1 <*> fromCursor e2 <*> fromCursor e3
      ("3TrafficLights2", [e1, e2, e3]) ->
        IconSet3TrafficLights2 <$> fromCursor e1 <*> fromCursor e2 <*> fromCursor e3
      ("4Arrows", [e1, e2, e3, e4]) ->
        IconSet4Arrows <$> fromCursor e1 <*> fromCursor e2 <*> fromCursor e3 <*> fromCursor e4
      ("4ArrowsGray", [e1, e2, e3, e4]) ->
        IconSet4ArrowsGray <$> fromCursor e1 <*> fromCursor e2 <*> fromCursor e3 <*> fromCursor e4
      ("4Rating", [e1, e2, e3, e4]) ->
        IconSet4Rating <$> fromCursor e1 <*> fromCursor e2 <*> fromCursor e3 <*> fromCursor e4
      ("4RedToBlack", [e1, e2, e3, e4]) ->
        IconSet4RedToBlack <$> fromCursor e1 <*> fromCursor e2 <*> fromCursor e3 <*> fromCursor e4
      ("4TrafficLights", [e1, e2, e3, e4]) ->
        IconSet4TrafficLights <$> fromCursor e1 <*> fromCursor e2 <*> fromCursor e3 <*> fromCursor e4
      ("5Arrows", [e1, e2, e3, e4, e5]) ->
        IconSet5Arrows <$> fromCursor e1 <*> fromCursor e2 <*> fromCursor e3 <*> fromCursor e4 <*> fromCursor e5
      ("5ArrowsGray", [e1, e2, e3, e4, e5]) ->
        IconSet5ArrowsGray <$> fromCursor e1 <*> fromCursor e2 <*> fromCursor e3 <*> fromCursor e4 <*> fromCursor e5
      ("5Quarters", [e1, e2, e3, e4, e5]) ->
        IconSet5Quarters <$> fromCursor e1 <*> fromCursor e2 <*> fromCursor e3 <*> fromCursor e4 <*> fromCursor e5
      ("5Rating", [e1, e2, e3, e4, e5]) ->
        IconSet5Rating <$> fromCursor e1 <*> fromCursor e2 <*> fromCursor e3 <*> fromCursor e4 <*> fromCursor e5
      _ ->
        fail $ "unexpected combination of icon set type " ++ show isType ++ " and " ++
                   show (length cfvoEls) ++ " cfvo elements"
    return IconSetOptions{..}

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

instance FromAttrVal Inclusion where
  fromAttrVal = right (first $ bool Exclusive Inclusive) . fromAttrVal

instance FromAttrVal NStdDev where
  fromAttrVal = right (first NStdDev) . fromAttrVal

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
    elementList nm attrs $ map (toElement "cfvo") cfvs
    where
      attrs = catMaybes
        [ "iconSet" .=? justNonDef defaultIconSet iconSet
        , "reverse" .=? justTrue _isoReverse
        , "showValue" .=? justFalse _isoShowValue
        ]
      (iconSet, cfvs) =
        case _isoIconSet of
          IconSet3Arrows cfv1 cfv2 cfv3 ->
            ("3Arrows" :: Text, [cfv1, cfv2, cfv3])
          IconSet3ArrowsGray cfv1 cfv2 cfv3 ->
            ("3ArrowsGray", [cfv1, cfv2, cfv3])
          IconSet3Flags cfv1 cfv2 cfv3 -> ("3Flags", [cfv1, cfv2, cfv3])
          IconSet3Signs cfv1 cfv2 cfv3 -> ("3Signs", [cfv1, cfv2, cfv3])
          IconSet3Symbols cfv1 cfv2 cfv3 -> ("3Symbols", [cfv1, cfv2, cfv3])
          IconSet3Symbols2 cfv1 cfv2 cfv3 -> ("3Symbols2", [cfv1, cfv2, cfv3])
          IconSet3TrafficLights1 cfv1 cfv2 cfv3 ->
            ("3TrafficLights1", [cfv1, cfv2, cfv3])
          IconSet3TrafficLights2 cfv1 cfv2 cfv3 ->
            ("3TrafficLights2", [cfv1, cfv2, cfv3])
          IconSet4Arrows cfv1 cfv2 cfv3 cfv4 ->
            ("4Arrows", [cfv1, cfv2, cfv3, cfv4])
          IconSet4ArrowsGray cfv1 cfv2 cfv3 cfv4 ->
            ("4ArrowsGray", [cfv1, cfv2, cfv3, cfv4])
          IconSet4Rating cfv1 cfv2 cfv3 cfv4 ->
            ("4Rating", [cfv1, cfv2, cfv3, cfv4])
          IconSet4RedToBlack cfv1 cfv2 cfv3 cfv4 ->
            ("4RedToBlack", [cfv1, cfv2, cfv3, cfv4])
          IconSet4TrafficLights cfv1 cfv2 cfv3 cfv4 ->
            ("4TrafficLights", [cfv1, cfv2, cfv3, cfv4])
          IconSet5Arrows cfv1 cfv2 cfv3 cfv4 cfv5 ->
            ("5Arrows", [cfv1, cfv2, cfv3, cfv4, cfv5])
          IconSet5ArrowsGray cfv1 cfv2 cfv3 cfv4 cfv5 ->
            ("5ArrowsGray", [cfv1, cfv2, cfv3, cfv4, cfv5])
          IconSet5Quarters cfv1 cfv2 cfv3 cfv4 cfv5 ->
            ("5Quarters", [cfv1, cfv2, cfv3, cfv4, cfv5])
          IconSet5Rating cfv1 cfv2 cfv3 cfv4 cfv5 ->
            ("5Rating", [cfv1, cfv2, cfv3, cfv4, cfv5])

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
