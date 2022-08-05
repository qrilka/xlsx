{-# LANGUAGE FlexibleInstances #-}
{-# LANGUAGE MultiParamTypeClasses #-}
{-# LANGUAGE ScopedTypeVariables #-}
{-# LANGUAGE OverloadedStrings #-}
{-# OPTIONS_GHC -fno-warn-orphans #-}

module CommonTests
  ( tests
  ) where

import Data.Fixed (Pico, Fixed(..), E12)
import Data.Time.Calendar (fromGregorian)
import Data.Time.Clock (UTCTime(..))
import Test.Tasty.SmallCheck (testProperty)
import Test.SmallCheck.Series
       (Positive(..), Serial(..), newtypeCons, cons0, (\/))
import Test.Tasty (testGroup, TestTree)
import Test.Tasty.HUnit (testCase, (@?=))

import Codec.Xlsx.Types.Common

tests :: TestTree
tests =
  testGroup
    "Types.Common tests"
    [ testProperty "col2int . int2col == id" $
        \(Positive i) -> i == col2int (int2col i)
    , testCase "date conversions" $ do
        dateFromNumber DateBase1900 (- 2338.0) @?= read "1893-08-06 00:00:00 UTC"
        dateFromNumber DateBase1900 2.0 @?= read "1900-01-02 00:00:00 UTC"
        dateFromNumber DateBase1900 3687.0 @?= read "1910-02-03 00:00:00 UTC"
        dateFromNumber DateBase1900 38749.0 @?= read "2006-02-01 00:00:00 UTC"
        dateFromNumber DateBase1900 2958465.0 @?= read "9999-12-31 00:00:00 UTC"
        dateFromNumber DateBase1900 59.0 @?= read "1900-02-28 00:00:00 UTC"
        dateFromNumber DateBase1900 59.5 @?= read "1900-02-28 12:00:00 UTC"
        dateFromNumber DateBase1900 60.0 @?= read "1900-03-01 00:00:00 UTC"
        dateFromNumber DateBase1900 60.5 @?= read "1900-03-01 00:00:00 UTC"
        dateFromNumber DateBase1900 61 @?= read "1900-03-01 00:00:00 UTC"
        dateFromNumber DateBase1900 61.5 @?= read "1900-03-01 12:00:00 UTC"
        dateFromNumber DateBase1900 62 @?= read "1900-03-02 00:00:00 UTC"
        dateFromNumber DateBase1904 (-3800.0) @?= read "1893-08-05 00:00:00 UTC"
        dateFromNumber DateBase1904 0.0 @?= read "1904-01-01 00:00:00 UTC"
        dateFromNumber DateBase1904 2225.0 @?= read "1910-02-03 00:00:00 UTC"
        dateFromNumber DateBase1904 37287.0 @?= read "2006-02-01 00:00:00 UTC"
        dateFromNumber DateBase1904 2957003.0 @?= read "9999-12-31 00:00:00 UTC"
     , testCase "Converting dates in the vicinity of 1900-03-01 to numbers" $ do
        -- Note how the fact that 1900-02-29 exists for Excel forces us to skip 60
        dateToNumber DateBase1900 (UTCTime (fromGregorian 1900 2 28) 0) @?= (59 :: Double)
        dateToNumber DateBase1900 (UTCTime (fromGregorian 1900 3 1) 0) @?= (61 :: Double)
     , testProperty "dateToNumber . dateFromNumber == id" $
       -- Because excel treats 1900 as a leap year, dateToNumber and dateFromNumber
       -- aren't inverses of each other in the range n E [60, 61[ for DateBase1900
       \b (n :: Pico) -> (n >= 60 && n < 61 && b == DateBase1900) || n == dateToNumber b (dateFromNumber b $ n)
    , testCase "building single CellRefs" $ do
        singleCellRef' (both Rel (5, 25)) @?= CellRef "Y5"
        singleCellRef' (Rel 5, Abs 25) @?= CellRef "$Y5"
        singleCellRef' (Abs 5, Rel 25) @?= CellRef "Y$5"
        singleCellRef' (both Abs (5, 25)) @?= CellRef "$Y$5"
        singleCellRef (5, 25) @?= CellRef "Y5"
    , testCase "parsing single CellRefs as abstract coordinates" $ do
        fromSingleCellRef (CellRef "Y5") @?= Just (5, 25)
        fromSingleCellRef (CellRef "$Y5") @?= Just (5, 25)
        fromSingleCellRef (CellRef "Y$5") @?= Just (5, 25)
        fromSingleCellRef (CellRef "$Y$5") @?= Just (5, 25)
    , testCase "parsing single CellRefs as potentially absolute coordinates" $ do
        fromSingleCellRef' (CellRef "Y5") @?= Just (both Rel (5, 25))
        fromSingleCellRef' (CellRef "$Y5") @?= Just (Rel 5, Abs 25)
        fromSingleCellRef' (CellRef "Y$5") @?= Just (Abs 5, Rel 25)
        fromSingleCellRef' (CellRef "$Y$5") @?= Just (both Abs (5, 25))
        fromSingleCellRef' (CellRef "$Y$50") @?= Just (both Abs (50, 25))
        fromSingleCellRef' (CellRef "$Y$5$0") @?= Nothing
        fromSingleCellRef' (CellRef "Y5:Z10") @?= Nothing
    , testCase "building ranges" $ do
        mkRange (5, 25) (10, 26) @?= CellRef "Y5:Z10"
        mkRange' (both Rel (5, 25)) (both Rel (10, 26)) @?= CellRef "Y5:Z10"
        mkRange' (both Abs (5, 25)) (both Abs (10, 26)) @?= CellRef "$Y$5:$Z$10"
        mkRange' (Rel 5, Abs 25) (Abs 10, Rel 26) @?= CellRef "$Y5:Z$10"
        mkForeignRange "myWorksheet" (Rel 5, Abs 25) (Abs 10, Rel 26) @?= CellRef "'myWorksheet'!$Y5:Z$10"
        mkForeignRange "my sheet" (Rel 5, Abs 25) (Abs 10, Rel 26) @?= CellRef "'my sheet'!$Y5:Z$10"
    , testCase "parsing ranges CellRefs as abstract coordinates" $ do
        fromRange (CellRef "Y5:Z10") @?= Just ((5, 25), (10, 26))
        fromRange (CellRef "$Y$5:$Z$10") @?= Just ((5, 25), (10, 26))
        fromRange (CellRef "myWorksheet!$Y5:Z$10") @?= Just ((5, 25), (10, 26))
    , testCase "parsing ranges CellRefs as potentially absolute coordinates" $ do
        fromRange' (CellRef "Y5:Z10") @?= Just (both (both Rel) ((5, 25), (10, 26)))
        fromRange' (CellRef "$Y$5:$Z$10") @?= Just (both (both Abs) ((5, 25), (10, 26)))
        fromRange' (CellRef "myWorksheet!$Y5:Z$10") @?= Just ((Rel 5, Abs 25), (Abs 10, Rel 26))
        fromForeignRange (CellRef "myWorksheet!$Y5:Z$10") @?= Just ("myWorksheet", ((Rel 5, Abs 25), (Abs 10, Rel 26)))
        fromForeignRange (CellRef "'myWorksheet'!Y5:Z10") @?= Just ("myWorksheet", both (both Rel) ((5, 25), (10, 26)))
        fromForeignRange (CellRef "'my sheet'!Y5:Z10") @?= Just ("my sheet", both (both Rel) ((5, 25), (10, 26)))
        fromForeignRange (CellRef "$Y5:Z$10") @?= Nothing
    ]

instance Monad m => Serial m (Fixed E12) where
  series = newtypeCons MkFixed

instance Monad m  => Serial m DateBase where
  series = cons0 DateBase1900 \/ cons0 DateBase1904
