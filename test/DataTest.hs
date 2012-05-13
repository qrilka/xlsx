module Main (main) where

import Test.Framework (defaultMain, testGroup)
import Test.Framework.Providers.HUnit
import Test.Framework.Providers.QuickCheck2 (testProperty)

import Test.QuickCheck
import Test.HUnit
import Control.Monad

import Codec.Xlsx


main = defaultMain tests

tests =
    [ testGroup "Behaves to spec"
        [ testProperty "col to name" prop_col2name 
        ]
    ]

prop_col2name (Positive i) = i == (col2Int $ int2col i)
    where types = (i::Int)
