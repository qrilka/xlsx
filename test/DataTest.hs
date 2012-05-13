module Main (main) where

import Test.Framework (defaultMain, testGroup)
import Test.Framework.Providers.QuickCheck2 (testProperty)

import Test.QuickCheck

import Codec.Xlsx

main = defaultMain tests

tests =
    [ testGroup "Behaves to spec"
        [ testProperty "col to name" prop_col2name ]
    ]


prop_col2name (Positive i) = i == (xlsxCol2int $ int2xlsxCol i)
    where types = (i::Int)
