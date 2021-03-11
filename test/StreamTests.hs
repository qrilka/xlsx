{-# LANGUAGE OverloadedStrings #-}
{-# LANGUAGE QuasiQuotes       #-}
{-# LANGUAGE TupleSections #-}

module StreamTests
  ( tests
  ) where

import           Conduit                         ( (.|))
import Codec.Xlsx.Types.Internal.SharedStringTable
import Codec.Xlsx
import           Codec.Xlsx.Parser.Stream
import           Codec.Xlsx.Types.Common
import qualified Conduit                  as C
import           Control.Lens             hiding (indexed)
import           Data.ByteString.Lazy     (ByteString)
import qualified Data.ByteString.Lazy     as LB
import           Data.Map                 (Map)
import qualified Data.Map                 as Map
import qualified Data.Map                 as M
import           Data.Maybe               (mapMaybe)
import           Data.Text                (Text)
import qualified Data.Text                as Text
import qualified Data.Text                as T
import           Data.Vector              (Vector, indexed, toList)
import           Diff
import           Test.Tasty               (TestTree, testGroup)
import           Test.Tasty.HUnit         (testCase)
import           TestXlsx
import           Text.RawString.QQ
import           Text.XML

tests :: TestTree
tests =
  testGroup
    "Stream tests"
    [ testCase "Get's out the shared strings" $ do
        let bs = fromXlsx testTime testXlsx
        LB.writeFile "data-test.xlsx" bs

        res <- C.runResourceT $ C.runConduit $  C.sourceFile "data-test.xlsx" .| parseStringsIntoMap
        let
          testSst :: Vector XlsxText
          testSst = sstTable $  sstConstruct (testXlsx ^.. xlSheets . traversed . _2)

          testTexts :: Vector (Maybe Text)
          testTexts = preview _XlsxText <$> testSst

          testMap :: Map Int Text
          testMap = Map.fromList $ do
            (x, my) <- toList $ indexed testTexts
            maybe [] (pure . (x,)) my

        testMap @==? res
    ]
