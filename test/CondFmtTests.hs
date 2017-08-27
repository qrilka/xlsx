{-# LANGUAGE CPP #-}
{-# LANGUAGE FlexibleInstances #-}
{-# LANGUAGE MultiParamTypeClasses #-}
{-# LANGUAGE OverloadedStrings #-}
{-# LANGUAGE ScopedTypeVariables #-}
{-# OPTIONS_GHC -fno-warn-orphans #-}

module CondFmtTests
  ( tests
  ) where

import Data.Text (Text)
import qualified Data.Text as T
import Test.SmallCheck.Series
import Test.Tasty (testGroup, TestTree)
import Test.Tasty.SmallCheck (testProperty)
import Text.XML (Node(NodeElement))
import Text.XML.Cursor (fromNode)

#if !MIN_VERSION_base(4,8,0)
import Control.Applicative
#endif

import Codec.Xlsx
import Codec.Xlsx.Parser.Internal
import Codec.Xlsx.Writer.Internal

tests :: TestTree
tests =
  testGroup
    "Types.ConditionalFormatting tests"
    [ testProperty "fromCursor . toElement == id" $ \(cFmt :: CfRule) ->
        [cFmt] == fromCursor (cursorFromElement $ toElement (n_ "cfRule") cFmt)
    ]
  where
    cursorFromElement = fromNode . NodeElement . addNS mainNamespace Nothing

instance Monad m  => Serial m CfRule where
  series = cons4 CfRule

instance Monad m  => Serial m Condition where
  series = localDepth (const 2) $ cons0 AboveAverage
    \/ cons1 BeginsWith
    \/ cons0 BelowAverage
    \/ cons1 BottomNPercent
    \/ cons1 BottomNValues
    \/ cons1 CellIs
    \/ cons4 ColorScale2
    \/ cons6 ColorScale3
    \/ cons0 ContainsBlanks
    \/ cons0 ContainsErrors
    \/ cons1 ContainsText
    \/ cons1 DataBar
    \/ cons0 DoesNotContainErrors
    \/ cons0 DoesNotContainBlanks
    \/ cons1 DoesNotContainText
    \/ cons0 DuplicateValues
    \/ cons1 EndsWith
    \/ cons1 Expression
    \/ cons1 IconSet
    \/ cons1 InTimePeriod
    \/ cons1 TopNPercent
    \/ cons1 TopNValues
    \/ cons0 UniqueValues

instance Monad m => Serial m Text where
  series = T.pack <$> series

instance Monad m => Serial m OperatorExpression where
  series = cons1 OpBeginsWith
    \/ cons2 OpBetween
    \/ cons1 OpContainsText
    \/ cons1 OpEndsWith
    \/ cons1 OpEqual
    \/ cons1 OpGreaterThan
    \/ cons1 OpGreaterThanOrEqual
    \/ cons1 OpLessThan
    \/ cons1 OpLessThanOrEqual
    \/ cons2 OpNotBetween
    \/ cons1 OpNotContains
    \/ cons1 OpNotEqual

instance Monad m => Serial m DataBarOptions where
  series = cons6 DataBarOptions

instance Monad m => Serial m MaxCfValue where
  series = cons0 CfvMax \/ cons1 MaxCfValue

instance Monad m => Serial m MinCfValue where
  series = cons0 CfvMin \/ cons1 MinCfValue

instance Monad m => Serial m Color where
  series = cons4 Color

-- TODO: proper formula generator (?)
instance Monad m => Serial m Formula where
  series = Formula <$> series

instance Monad m => Serial m IconSetOptions where
  series = cons3 IconSetOptions

instance Monad m => Serial m IconSetType where
  series = cons3 IconSet3Arrows
    \/ cons3 IconSet3ArrowsGray
    \/ cons3 IconSet3Flags
    \/ cons3 IconSet3Signs
    \/ cons3 IconSet3Symbols
    \/ cons3 IconSet3Symbols
    \/ cons3 IconSet3TrafficLights1
    \/ cons3 IconSet3TrafficLights2
    \/ cons4 IconSet4Arrows
    \/ cons4 IconSet4ArrowsGray
    \/ cons4 IconSet4Rating
    \/ cons4 IconSet4RedToBlack
    \/ cons4 IconSet4TrafficLights
    \/ cons5 IconSet5Arrows
    \/ cons5 IconSet5ArrowsGray
    \/ cons5 IconSet5Quarters
    \/ cons5 IconSet5Rating

instance Monad m => Serial m CfValue where
  series = cons1 CfValue
    \/ cons1 CfPercent
    \/ cons1 CfPercentile
    \/ cons1 CfFormula

instance Monad m => Serial m TimePeriod where
  series =
    cons0 PerLast7Days
    \/ cons0 PerLastMonth
    \/ cons0 PerLastWeek
    \/ cons0 PerNextMonth
    \/ cons0 PerNextWeek
    \/ cons0 PerThisMonth
    \/ cons0 PerThisWeek
    \/ cons0 PerToday
    \/ cons0 PerTomorrow
    \/ cons0 PerYesterday

cons5 ::
     (Serial m a5, Serial m a4, Serial m a3, Serial m a2, Serial m a1)
  => (a5 -> a4 -> a3 -> a2 -> a1 -> a)
  -> Test.SmallCheck.Series.Series m a
cons5 f = decDepth $
  f <$> series
    <~> series
    <~> series
    <~> series
    <~> series

cons6 ::
     ( Serial m a6
     , Serial m a5
     , Serial m a4
     , Serial m a3
     , Serial m a2
     , Serial m a1
     )
  => (a6 -> a5 -> a4 -> a3 -> a2 -> a1 -> a)
  -> Test.SmallCheck.Series.Series m a
cons6 f = decDepth $
  f <$> series
    <~> series
    <~> series
    <~> series
    <~> series
    <~> series
