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
import qualified Codec.Xlsx.Writer.Internal.Stream as SW
import Control.Monad.State.Lazy
import Test.Tasty.SmallCheck
import qualified Data.Set as Set
import Data.Set (Set)
import Text.Printf
import Debug.Trace
import Data.Set.Lens
import Control.DeepSeq
import Data.Conduit
import Codec.Xlsx.Formatted

tempPath :: FilePath
tempPath = "test" </> "temp"

toBs = LB.toStrict . fromXlsx testTime

mkTestCase :: TestName -> Xlsx -> TestTree
mkTestCase testName xlsx = testCase testName $ do
  res <- C.runResourceT $ C.runConduit $  yield (toBs xlsx) .| parseSharedStringsIntoMapC

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
      [ mkTestCase "Get's out the shared strings" testXlsx
      , mkTestCase "Workbook result is parsed correctly" testFormatWorkbookResult -- TODO: compare to testFormatWorkbook
      , mkTestCase "Workbook is parsed correctly" testFormatWorkbook
      ],

      testGroup "Writer"
      [ testProperty "Input same as the output" testInputSameAsOutput
      , testProperty "Set of input texts is same as map length" testSetOfInputTextsIsSameAsMapLength
      , testProperty "Set of input texts is as value set length" testSetOfInputTextsIsSameAsValueSetLength
      ],

      testGroup "Reader/Writer"
      [ testCase "Write as stream, see if memory based implementation can read it" $ readWrite simpleWorkbook
      , testCase "Test a big workbook which caused issues with zipstream" $ readWrite bigWorkbook
      -- , testCase "Write as stream, see if memory based implementation can read it" $ readWrite testXlsx
      -- TODO forall SheetItem write that can be read
      ]
    ]


readWrite :: Xlsx -> IO ()
readWrite input = do
  readStr  <- C.runResourceT $ readXlsxC $ yield (toBs input)
  bs <- runConduitRes $ void (SW.writeXlsx readStr) .| C.foldC
  case toXlsxEither $ LB.fromStrict bs of
    Right result  ->
      simplified  @==?  result
    Left x -> do
      throwIO x


  where
    -- we remove some properties because the streaming is incomplete
    simplified = (xlStyles .~ Styles mempty $ input )


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

simpleWorkbook :: Xlsx
simpleWorkbook = formatWorkbook sheets minimalStyleSheet
  where
    sheets = [("Sheet1" , M.fromList [((1,1), (def & formattedCell . cellValue ?~ CellText "text at A1 Sheet1"))])]

bigWorkbook :: Xlsx
bigWorkbook = formatWorkbook sheets minimalStyleSheet
  where
    sheets = [("Sheet1" , M.fromList $ [0..512] >>= \row ->
                  [((row,1), (def & formattedCell . cellValue ?~ CellText "text at A1 Sheet1")), ((row,2), (def & formattedCell . cellValue ?~ CellText "text at B1 Sheet1"))])]
