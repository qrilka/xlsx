{-# LANGUAGE FlexibleInstances #-}
{-# LANGUAGE MultiParamTypeClasses #-}
{-# OPTIONS_GHC -fno-warn-orphans #-}

module Test.SmallCheck.Series.Instances
  (
  ) where

import Data.Text (Text)
import qualified Data.Text as T
import Test.SmallCheck.Series

import Codec.Xlsx

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

instance Monad m => Serial m Text where
  series = T.pack <$> series


{-------------------------------------------------------------------------------
  Conditional formatting
-------------------------------------------------------------------------------}

instance Monad m  => Serial m CfRule

instance Monad m  => Serial m Condition where
  series = localDepth (const 2) $ cons2 AboveAverage
    \/ cons1 BeginsWith
    \/ cons2 BelowAverage
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

instance Monad m => Serial m NStdDev

instance Monad m => Serial m Inclusion

instance Monad m => Serial m OperatorExpression

instance Monad m => Serial m DataBarOptions

instance Monad m => Serial m MaxCfValue

instance Monad m => Serial m MinCfValue

instance Monad m => Serial m Color

-- TODO: proper formula generator (?)
instance Monad m => Serial m Formula

instance Monad m => Serial m IconSetOptions

instance Monad m => Serial m IconSetType

instance Monad m => Serial m CfValue

instance Monad m => Serial m TimePeriod
