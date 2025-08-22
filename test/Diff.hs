module Diff where

import           Data.Algorithm.Diff       (Diff (..), getGroupedDiff)
import           Data.Algorithm.DiffOutput (ppDiff)
import           Data.Monoid               ((<>))
import           Test.Tasty.HUnit          (Assertion, HasCallStack, assertBool)
import           Text.Groom                (groom)

-- | Like '@=?' but producing a diff on failure.
(@==?) :: HasCallStack => (Eq a, Show a) => a -> a -> Assertion
x @==? y =
    assertBool ("Expected:\n" <> groom x <> "\nDifference:\n" <> msg) (x == y)
  where
    msg = ppDiff $ getGroupedDiff (lines . groom $ x) (lines . groom $ y)
