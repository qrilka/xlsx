{-# LANGUAGE OverloadedStrings #-}
{-# LANGUAGE QuasiQuotes       #-}
{-# LANGUAGE TupleSections     #-}
{-# LANGUAGE ScopedTypeVariables #-}
{-# LANGUAGE TypeApplications #-}
{-# LANGUAGE CPP #-}

module StreamTests
  ( tests
  ) where

#ifdef USE_MICROLENS

import Test.Tasty (TestName, TestTree, testGroup)
tests :: TestTree
tests = testGroup
  "I stubbed out the tests module for microlens \
   because it doesn't understand setOf. \
   Volunteers are welcome to fix this!"
    []
#else

import Control.Exception
import Codec.Xlsx
import Codec.Xlsx.Parser.Stream
import Conduit ((.|))
import qualified Conduit as C
import Control.Lens hiding (indexed)
import Data.Set.Lens
import qualified Data.ByteString.Lazy as LB
import qualified Data.ByteString as BS
import Data.Map (Map)
import qualified Data.Map as M
import qualified Data.IntMap.Strict as IM
import Data.Text (Text)
import qualified Data.Text as Text
import Diff
import Test.Tasty (TestTree, testGroup)
import Test.Tasty.HUnit (testCase)
import TestXlsx
import qualified Codec.Xlsx.Writer.Stream as SW
import qualified Codec.Xlsx.Writer.Internal.Stream as SW
import Control.Monad.State.Lazy
import Test.Tasty.SmallCheck
import Test.SmallCheck.Series.Instances ()
import qualified Data.Set as Set
import Data.Set (Set)
import Text.Printf
import Data.Conduit

toBs :: Xlsx -> BS.ByteString
toBs = LB.toStrict . fromXlsx testTime

tests :: TestTree
tests =
  testGroup "Stream tests"
    [
      testGroup "Writer/shared strings"
      [ testProperty "Input same as the output" sharedStringInputSameAsOutput
      , testProperty "Set of input texts is same as map length" sharedStringInputTextsIsSameAsMapLength
      , testProperty "Set of input texts is as value set length" sharedStringInputTextsIsSameAsValueSetLength
      ],

      testGroup "Reader/Writer"
      [ testCase "Write as stream, see if memory based implementation can read it" $ readWrite simpleWorkbook
      , testCase "Write as stream, see if memory based implementation can read it" $ readWrite simpleWorkbookRow
      , testCase "Test a small workbook which has a fullblown sqaure" $ readWrite smallWorkbook
      , testCase "Test a big workbook as a full square which caused issues with zipstream \
                 The buffer of zipstream maybe 1kb, this workbook is big enough \
                 to be more than that. \
                 So if this encodes/decodes we know we can handle those sizes. \
                 In some older version the bytestring got cut off resulting in a corrupt xlsx file"
                  $ readWrite bigWorkbook
      -- , testCase "Write as stream, see if memory based implementation can read it" $ readWrite testXlsx
      -- TODO forall SheetItem write that can be read
      ],

      testGroup "Reader/inline strings"
      [ testCase "Can parse row with inline strings" inlineStringsAreParsed
      ],

      testGroup "Reader/floats parsing"
      [ testCase "Can parse untyped values as floats" untypedCellsAreParsedAsFloats
      ]
    ]

readWrite :: Xlsx -> IO ()
readWrite input = do
  BS.writeFile "testinput.xlsx" (toBs input)
  items <- fmap (toListOf (traversed . si_row)) $ runXlsxM "testinput.xlsx" $ collectItems $ makeIndex 1
  bs <- runConduitRes $ void (SW.writeXlsx SW.defaultSettings $ C.yieldMany items) .| C.foldC
  case toXlsxEither $ LB.fromStrict bs of
    Right result  ->
      input @==?  result
    Left x -> do
      throwIO x

-- test if the input text is also the result (a property we use for convenience)
sharedStringInputSameAsOutput :: Text -> Either String String
sharedStringInputSameAsOutput someText =
  if someText  == out then Right msg  else Left msg
  where
    out = fst $ evalState (SW.upsertSharedString someText) SW.initialSharedString
    msg = printf "'%s' = '%s'" (Text.unpack out) (Text.unpack someText)

-- test if unique strings actually get set in the map as keys
sharedStringInputTextsIsSameAsMapLength :: [Text] -> Bool
sharedStringInputTextsIsSameAsMapLength someTexts =
    length result == length unqTexts
  where
   result  :: Map Text Int
   result = view SW.string_map $ traverse SW.upsertSharedString someTexts `execState` SW.initialSharedString
   unqTexts :: Set Text
   unqTexts = Set.fromList someTexts

-- test for every unique string we get a unique number
sharedStringInputTextsIsSameAsValueSetLength :: [Text] -> Bool
sharedStringInputTextsIsSameAsValueSetLength someTexts =
    length result == length unqTexts
  where
   result  :: Set Int
   result = setOf (SW.string_map . traversed) $ traverse SW.upsertSharedString someTexts `execState` SW.initialSharedString
   unqTexts :: Set Text
   unqTexts = Set.fromList someTexts

-- can we do xx
simpleWorkbook :: Xlsx
simpleWorkbook = def & atSheet "Sheet1" ?~ sheet
  where
    sheet = toWs [((1,1), a1), ((1,2), cellValue ?~ CellText "text at B1 Sheet1" $ def)]

a1 :: Cell
a1 = cellValue ?~ CellText "text at A1 Sheet1" $ cellStyle ?~ 1 $ def

-- can we do x
--           x
simpleWorkbookRow :: Xlsx
simpleWorkbookRow = def & atSheet "Sheet1" ?~ sheet
  where
    sheet = toWs [((1,1), a1), ((2,1), cellValue ?~ CellText "text at A2 Sheet1" $ def)]


tshow :: Show a => a -> Text
tshow = Text.pack . show

toWs :: [((Int,Int), Cell)] -> Worksheet
toWs x = set wsCells (M.fromList x) def

-- can we do xxx
--           xxx
--           .
--           .
smallWorkbook :: Xlsx
smallWorkbook = def & atSheet "Sheet1" ?~ sheet
  where
    sheet = toWs $ [1..2] >>= \row ->
                  [((row,1), a1)
                  , ((row,2), def & cellValue ?~ CellText ("text at B"<> tshow row <> " Sheet1"))
                  , ((row,3), def & cellValue ?~ CellText "text at C1 Sheet1")
                  , ((row,4), def & cellValue ?~ CellDouble (0.2 + 0.1))
                  , ((row,5), def & cellValue ?~ CellBool False)
                  ]

bigWorkbook :: Xlsx
bigWorkbook = def & atSheet "Sheet1" ?~ sheet
  where
    sheet = toWs $ [1..512] >>= \row ->
                  [((row,1), a1)
                  ,((row,2), def & cellValue ?~ CellText ("text at B"<> tshow row <> " Sheet1"))
                  ,((row,3), def & cellValue ?~ CellText "text at C1 Sheet1")
                  ]


inlineStringsAreParsed :: IO ()
inlineStringsAreParsed = do
  items <- runXlsxM "data/inline-strings.xlsx" $ collectItems $ makeIndex 1
  let expected =
        [ IM.fromList
            [ ( 1,
                Cell
                  { _cellStyle = Nothing,
                    _cellValue = Just (CellText "My Inline String"),
                    _cellComment = Nothing,
                    _cellFormula = Nothing
                  }
              ),
              ( 2,
                Cell
                  { _cellStyle = Nothing,
                    _cellValue = Just (CellText "two"),
                    _cellComment = Nothing,
                    _cellFormula = Nothing
                  }
              )
            ]
        ]
  expected @==? (items ^.. traversed . si_row . ri_cell_row)

untypedCellsAreParsedAsFloats :: IO ()
untypedCellsAreParsedAsFloats = do
  -- values in that file are under `General` cell-type and are not marked
  -- as numbers explicitly in `t` attribute.
  items <- runXlsxM "data/floats.xlsx" $ collectItems $ makeIndex 1
  let expected =
        [ IM.fromList [ (1, def & cellValue ?~ CellDouble 12.0) ]
        , IM.fromList [ (1, def & cellValue ?~ CellDouble 13.0) ]
        -- cell below has explicit `Numeric` type, while others are all `General`,
        -- but sometimes excel does not add a `t="n"` attr even to numeric cells
        -- but it should be default as number in any cases if `t` is missing
        , IM.fromList [ (1, def & cellValue ?~ CellDouble 14.0 & cellStyle ?~ 1 ) ]
        , IM.fromList [ (1, def & cellValue ?~ CellDouble 15.0) ]
        ]
  expected @==? (_ri_cell_row . _si_row <$> items)

#endif
