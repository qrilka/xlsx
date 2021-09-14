{-# LANGUAGE FlexibleInstances #-}
{-# LANGUAGE MultiParamTypeClasses #-}
{-# LANGUAGE ScopedTypeVariables #-}

module CommonTests
  ( tests
  ) where

import Data.Fixed (Pico, Fixed(..), E12)
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
        dateFromNumber DateBase1900 62 @?= read "1900-03-02 00:00:00 UTC"
        dateFromNumber DateBase1904 (-3800.0) @?= read "1893-08-05 00:00:00 UTC"
        dateFromNumber DateBase1904 0.0 @?= read "1904-01-01 00:00:00 UTC"
        dateFromNumber DateBase1904 2225.0 @?= read "1910-02-03 00:00:00 UTC"
        dateFromNumber DateBase1904 37287.0 @?= read "2006-02-01 00:00:00 UTC"
        dateFromNumber DateBase1904 2957003.0 @?= read "9999-12-31 00:00:00 UTC"
     , testProperty "dateToNumber . dateFromNumber == id" $
       \b (n :: Pico) -> n == dateToNumber b (dateFromNumber b $ n)
    ]

instance Monad m => Serial m (Fixed E12) where
  series = newtypeCons MkFixed

instance Monad m  => Serial m DateBase where
  series = cons0 DateBase1900 \/ cons0 DateBase1904
