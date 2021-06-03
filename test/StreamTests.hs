{-# LANGUAGE BlockArguments    #-}
{-# LANGUAGE OverloadedStrings #-}
{-# LANGUAGE QuasiQuotes       #-}
{-# LANGUAGE TupleSections     #-}
{-# LANGUAGE ScopedTypeVariables #-}
{-# LANGUAGE TypeApplications #-}

module StreamTests
  ( tests
  ) where
import Control.Exception
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
import qualified Data.ByteString  as BS
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
import qualified Codec.Xlsx.Writer.Stream as SW
import Control.Monad.State.Lazy
import Test.Tasty.SmallCheck
import qualified Data.Set as Set
import Data.Set (Set)
import Text.Printf
import Debug.Trace
import Data.Set.Lens
import Control.DeepSeq
import Data.Conduit

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
  testGroup "Stream tests"
    [
      testGroup "Reader"
      [ mkTestCase "Get's out the shared strings" "data-test.xlsx" testXlsx
      , mkTestCase "Workbook result is parsed correctly" "workbook-test.xlsx" testFormatWorkbookResult -- TODO: compare to testFormatWorkbook
      , mkTestCase "Workbook is parsed correctly" "format-workbook-test.xlsx" testFormatWorkbook
      ],
      testGroup "Writer"
      [ testProperty "Input same as the output" testInputSameAsOutput
      , testProperty "Set of input texts is same as map length" testSetOfInputTextsIsSameAsMapLength
      , testProperty "Set of input texts is as value set length" testSetOfInputTextsIsSameAsValueSetLength
      ],
      testGroup "Reader/Writer"
      [ testCase "Write as stream, see if memory based implementation can read it" readWriteSimple
      ]
    ]

readWriteSimple :: IO ()
readWriteSimple = do
  eadStr  <- C.runResourceT $ readXlsxC (C.sourceFile "data/simple.xlsx")
  bs <- runConduitRes $ void (SW.writeXlsx readStr) .| C.foldC
  case (toXlsxEither $ LB.fromStrict bs) of
    Right _ -> putStrLn "success"
    Left x -> do
      putStrLn "writing failed file failed.xlsx"
      BS.writeFile "failed.xlsx" bs
      throwIO x


-- test if the input text is also the result (a property we use for convenience)
testInputSameAsOutput :: Text -> Either String String
testInputSameAsOutput someText =
  if someText  == out then Right msg  else Left msg

  where
    out = fst $ evalState (SW.getSetNumber someText) SW.initialSharedString
    msg = printf "'%s' = '%s'" (Text.unpack out) (Text.unpack someText)

-- test if unique strings actually get set in the map as keys
testSetOfInputTextsIsSameAsMapLength :: [Text] -> Bool
testSetOfInputTextsIsSameAsMapLength someTexts =
    length result == length unqTexts
  where
   result  :: Map Text Int
   result = view SW.string_map $ traverse SW.getSetNumber someTexts `execState` SW.initialSharedString
   unqTexts :: Set Text
   unqTexts = Set.fromList someTexts

-- test for every unique string we get a unique number
testSetOfInputTextsIsSameAsValueSetLength :: [Text] -> Bool
testSetOfInputTextsIsSameAsValueSetLength someTexts =
    length result == length unqTexts
  where
   result  :: Set Int
   result = setOf (SW.string_map . traversed) $ traverse SW.getSetNumber someTexts `execState` SW.initialSharedString
   unqTexts :: Set Text
   unqTexts = Set.fromList someTexts
