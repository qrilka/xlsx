module Main (main) where

import Test.Tasty (defaultMain, testGroup)
import Test.Tasty.SmallCheck (testProperty)

import Test.SmallCheck.Series (Positive(..))

import Codec.Xlsx


main = defaultMain $
  testGroup "Tests"
    [ testProperty "col2int . int2col == id" $
        \(Positive i) -> i == col2int (int2col i)
    ]
