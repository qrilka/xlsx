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
import System.Directory (getTemporaryDirectory)
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

toBs = LB.toStrict . fromXlsx testTime

tests :: TestTree
tests =
  testGroup "Stream tests"
    [
      testGroup "Writer"
      [ testProperty "Input same as the output" testInputSameAsOutput
      , testProperty "Set of input texts is same as map length" testSetOfInputTextsIsSameAsMapLength
      , testProperty "Set of input texts is as value set length" testSetOfInputTextsIsSameAsValueSetLength
      ],

      testGroup "Reader/Writer"
      [ testCase "Write as stream, see if memory based implementation can read it" $ readWrite simpleWorkbook
      , testCase "Write as stream, see if memory based implementation can read it" $ readWrite simpleWorkbookRow
      , testCase "Test a big workbook which caused issues with zipstream" $ readWrite bigWorkbook
      -- , testCase "Write as stream, see if memory based implementation can read it" $ readWrite testXlsx
      -- TODO forall SheetItem write that can be read
      ]
    ]

readWrite :: Xlsx -> IO ()
readWrite input = do
  BS.writeFile "testinput.xlsx" (toBs input)
  items <- runXlsxM "testinput.xlsx" $ collectItems 1
  bs <- runConduitRes $ void (SW.writeXlsx SW.defaultSettings $ C.yieldMany items) .| C.foldC
  case toXlsxEither $ LB.fromStrict bs of
    Right result  ->
      input @==?  result
    Left x -> do
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

-- can we do xx
simpleWorkbook :: Xlsx
simpleWorkbook = set xlSheets sheets def
  where
    sheets = [("Sheet1" , toWs [((1,1), a1), ((1,2), cellValue ?~ CellText "text at B1 Sheet1" $ def)])]

a1 :: Cell
a1 = cellValue ?~ CellText "text at A1 Sheet1" $ cellStyle ?~ 1 $ def

-- can we do x
--           x
simpleWorkbookRow :: Xlsx
simpleWorkbookRow = set xlSheets sheets def
  where
    sheets = [("Sheet1" , toWs [((1,1), a1), ((2,1), cellValue ?~ CellText "text at A2 Sheet1" $ def)])]


tshow :: Show a => a -> Text
tshow = Text.pack . show

toWs :: [((Int,Int), Cell)] -> Worksheet
toWs x = set wsCells (M.fromList x) def

-- can we do xxx
--           xxx
--           .
--           .
bigWorkbook :: Xlsx
bigWorkbook = set xlSheets sheets def
  where
    sheets = [("Sheet1" , toWs $ [0..512] >>= \row ->
                  [((row,1), a1)
                  ,((row,2), def & cellValue ?~ CellText ("text at B"<> tshow row <> " Sheet1"))
                  ,((row,3), def & cellValue ?~ CellText "text at C1 Sheet1")
                  ]
              )]
