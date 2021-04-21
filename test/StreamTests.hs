{-# LANGUAGE BlockArguments    #-}
{-# LANGUAGE OverloadedStrings #-}
{-# LANGUAGE QuasiQuotes       #-}
{-# LANGUAGE TupleSections     #-}

module StreamTests
  ( tests
  ) where

import           Codec.Xlsx
import           Codec.Xlsx.Parser.Stream
import           Codec.Xlsx.Types.Common
import           Codec.Xlsx.Types.Internal.SharedStringTable
import           Conduit                                     ((.|))
import qualified Conduit                                     as C
import           Control.Exception                           (bracket)
import           Control.Lens                                hiding (indexed)
import           Data.ByteString.Lazy                        (ByteString)
import qualified Data.ByteString.Lazy                        as LB
import           Data.Map                                    (Map)
import qualified Data.Map                                    as M
import qualified Data.Map                                    as Map
import           Data.Maybe                                  (mapMaybe)
import           Data.Text                                   (Text)
import qualified Data.Text                                   as T
import qualified Data.Text                                   as Text
import           Data.Vector                                 (Vector, indexed,
                                                              toList)
import           Diff
import           System.FilePath.Posix
import           Test.Tasty                                  (TestName,
                                                              TestTree,
                                                              testGroup)
import           Test.Tasty.HUnit                            (testCase)
import           TestXlsx
import           Text.RawString.QQ
import           Text.XML

tempPath :: FilePath
tempPath = "test" </> "temp"

mkTestCase :: TestName -> FilePath -> Xlsx -> TestTree
mkTestCase testName filename xlsx = testCase testName $ do
  let filepath = tempPath </> filename
  let bs = fromXlsx testTime xlsx
  LB.writeFile filepath bs
  res <- C.runResourceT $ C.runConduit $  C.sourceFile filepath .| parseSharedStringsIntoMapC

  let
    testSst :: Vector XlsxText
    testSst = sstTable $ sstConstruct (xlsx ^.. xlSheets . traversed . _2)

    testTexts :: Vector (Maybe Text)
    testTexts = preview _XlsxText <$> testSst

    testMap :: Map Int Text
    testMap = Map.fromList $ do
      (x, my) <- toList $ indexed testTexts
      maybe [] (pure . (x,)) my

  testMap @==? res


tests :: TestTree
tests =
  testGroup
    "Stream tests"
    [ mkTestCase "Get's out the shared strings" "data-test.xlsx" testXlsx
    , mkTestCase "Workbook result is parsed correctly" "workbook-test.xlsx" testFormatWorkbookResult -- TODO: compare to testFormatWorkbook
    , mkTestCase "Workbook is parsed correctly" "format-workbook-test.xlsx" testFormatWorkbook
    ]
