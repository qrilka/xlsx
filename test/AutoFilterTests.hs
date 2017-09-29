{-# LANGUAGE OverloadedStrings #-}
{-# LANGUAGE ScopedTypeVariables #-}

module AutoFilterTests
  ( tests
  ) where

import Test.SmallCheck.Series
import Test.Tasty (testGroup, TestTree)
import Test.Tasty.SmallCheck (testProperty)

import Codec.Xlsx
import Codec.Xlsx.Parser.Internal
import Codec.Xlsx.Writer.Internal

import Common
import Test.SmallCheck.Series.Instances ()

tests :: TestTree
tests =
  testGroup
    "Types.AutFilter tests"
    [ testProperty "fromCursor . toElement == id" $ \(autoFilter :: AutoFilter) ->
        [autoFilter] == fromCursor (cursorFromElement $ toElement (n_ "autoFilter") autoFilter)
    ]
