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
--        [ testCase "xlsxInt2Col works" test_xlsxInt2Col
        [ testProperty "col to name" prop_col2name 
        ]
    ]

test_xlsxInt2Col = forM_ ([0..maxBound]::[Int]) $ \i -> do
    let t = int2xlsxCol i
    assertBool "always ok" True

prop_col2name (Positive i) = i == (col2Int $ int2Col i)
    where types = (i::Int)
