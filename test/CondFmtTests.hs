{-# LANGUAGE CPP #-}
{-# LANGUAGE FlexibleInstances #-}
{-# LANGUAGE MultiParamTypeClasses #-}
{-# LANGUAGE OverloadedStrings #-}
{-# LANGUAGE ScopedTypeVariables #-}
{-# OPTIONS_GHC -fno-warn-orphans #-}

module CondFmtTests
  ( tests
  ) where

import Test.Tasty (testGroup, TestTree)
import Test.Tasty.SmallCheck (testProperty)

#if !MIN_VERSION_base(4,8,0)
import Control.Applicative
#endif

import Codec.Xlsx
import Codec.Xlsx.Parser.Internal
import Codec.Xlsx.Writer.Internal

import Common
import Test.SmallCheck.Series.Instances ()

tests :: TestTree
tests =
  testGroup
    "Types.ConditionalFormatting tests"
    [ testProperty "fromCursor . toElement == id" $ \(cFmt :: CfRule) ->
        [cFmt] == fromCursor (cursorFromElement $ toElement (n_ "cfRule") cFmt)
    ]
