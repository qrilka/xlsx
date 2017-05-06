{-# LANGUAGE OverloadedStrings #-}
{-# LANGUAGE RecordWildCards   #-}
{-# LANGUAGE TemplateHaskell   #-}
{-# LANGUAGE DeriveGeneric #-}
module Codec.Xlsx.Types.ConditionalFormatting
    ( ConditionalFormatting
    , CfRule(..)
    , Condition(..)
    , OperatorExpression (..)
    , TimePeriod (..)
    -- * Lenses
    -- ** CfRule
    , cfrCondition
    , cfrDxfId
    , cfrPriority
    , cfrStopIfTrue
    -- * Misc
    , topCfPriority
    ) where

import GHC.Generics (Generic)

import           Control.Lens               (makeLenses)
import           Data.Map                   (Map)
import qualified Data.Map                   as M
import           Data.Maybe
import           Data.Text                  (Text)
import           Text.XML
import           Text.XML.Cursor

import           Codec.Xlsx.Parser.Internal
import           Codec.Xlsx.Types.Common
import           Codec.Xlsx.Writer.Internal

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

-- | Conditions which could be used for conditional formatting
--
-- See 18.18.12 "ST_CfType (Conditional Format Type)" (p. 2443)
data Condition
    -- | This conditional formatting rule highlights cells in the
    -- range that begin with the given text. Equivalent to
    -- using the LEFT() sheet function and comparing values.
    = BeginsWith Text
    -- | This conditional formatting rule compares a cell value
    -- to a formula calculated result, using an operator.
    | CellIs OperatorExpression
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
    -- | This conditional formatting rule highlights cells ending
    -- with given text. Equivalent to using the RIGHT() sheet
    -- function and comparing values.
    | EndsWith Text
    -- | This conditional formatting rule contains a formula to
    -- evaluate. When the formula result is true, the cell is
    -- highlighted.
    | Expression Formula
    -- | This conditional formatting rule highlights cells
    -- containing dates in the specified time period. The
    -- underlying value of the cell is evaluated, therefore the
    -- cell does not need to be formatted as a date to be
    -- evaluated. For example, with a cell containing the
    -- value 38913 the conditional format shall be applied if
    -- the rule requires a value of 7/14/2006.
    | InTimePeriod TimePeriod
    -- TODO: aboveAverage
    -- TODO: colorScale
    -- TODO: dataBar
    -- TODO: iconSet
    -- TODO: timePeriod
    -- TODO: top10
    -- TODO: uniqueValues
    deriving (Eq, Ord, Show, Generic)

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

makeLenses ''CfRule

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
readCondition "beginsWith" cur       = do
    txt <- fromAttribute "text" cur
    return $ BeginsWith txt
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
readCondition "notContainsBlanks" _  = return DoesNotContainBlanks
readCondition "notContainsErrors" _  = return DoesNotContainErrors
readCondition "notContainsText" cur  = do
    txt <- fromAttribute "text" cur
    return $ DoesNotContainText txt
readCondition "endsWith" cur         = do
    txt <- fromAttribute "text" cur
    return $ EndsWith txt
readCondition "expression" cur       = do
    formula <- cur $/ element "formula" >=> fromCursor
    return $ Expression formula
readCondition "timePeriod" cur  = do
    period <- fromAttribute "timePeriod" cur
    return $ InTimePeriod period
readCondition _ _                    = error "Unexpected conditional formatting type"

readOpExpression :: Text -> [Formula] -> [OperatorExpression]
readOpExpression "beginsWith" [f ]        = [OpBeginsWith f ]
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
conditionData (BeginsWith t)         = ("beginsWith", M.fromList [ "text" .= t], [])
conditionData (CellIs opExpr)        = ("cellIs", M.fromList [ "operator" .= op], formulas)
    where (op, formulas) = operatorExpressionData opExpr
conditionData ContainsBlanks         = ("containsBlanks", M.empty, [])
conditionData ContainsErrors         = ("containsErrors", M.empty, [])
conditionData (ContainsText t)       = ("containsText", M.fromList [ "text" .= t], [])
conditionData DoesNotContainBlanks   = ("notContainsBlanks", M.empty, [])
conditionData DoesNotContainErrors   = ("notContainsErrors", M.empty, [])
conditionData (DoesNotContainText t) = ("notContainsText", M.fromList [ "text" .= t], [])
conditionData (EndsWith t)           = ("endsWith", M.fromList [ "text" .= t], [])
conditionData (Expression formula)   = ("expression", M.empty, [formulaNode formula])
conditionData (InTimePeriod period)  = ("timePeriod", M.fromList [ "timePeriod" .= period ], [])

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

formulaNode :: Formula -> Node
formulaNode = NodeElement . toElement "formula"

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
