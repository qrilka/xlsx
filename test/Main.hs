{-# LANGUAGE CPP               #-}
{-# LANGUAGE OverloadedLists   #-}
{-# LANGUAGE OverloadedStrings #-}
module Main
  ( main
  ) where

#ifdef USE_MICROLENS
import Lens.Micro
#else
import Control.Lens
#endif
import Data.ByteString.Lazy (ByteString)
import qualified Data.ByteString.Lazy as LB
import qualified Data.Map as M
import Data.Maybe (listToMaybe)
import Data.Text (Text)
import qualified StreamTests
import Text.XML

import Test.Tasty (defaultMain, testGroup)
import Test.Tasty.HUnit (testCase, (@=?))
import Test.Tasty.SmallCheck (testProperty)

import TestXlsx

import Codec.Xlsx
import Codec.Xlsx.Formatted

import AutoFilterTests
import Common
import CommonTests
import CondFmtTests
import Diff
import DrawingTests
import PivotTableTests

main :: IO ()
main = defaultMain $
  testGroup "Tests"
    [
       testCase "write . read == id" $ do
        let bs = fromXlsx testTime testXlsx
        LB.writeFile "data-test.xlsx" bs
        testXlsx @==? toXlsx (fromXlsx testTime testXlsx)
    ,  testCase "write . fast-read == id" $ do
        let bs = fromXlsx testTime testXlsx
        LB.writeFile "data-test.xlsx" bs
        testXlsx @==? toXlsxFast (fromXlsx testTime testXlsx)
    , testCase "fromRows . toRows == id" $
        testCellMap1 @=? fromRows (toRows testCellMap1)
    , testCase "fromRight . parseStyleSheet . renderStyleSheet == id" $
        testStyleSheet @==? fromRight (parseStyleSheet (renderStyleSheet  testStyleSheet))
    , testCase "correct shared strings parsing" $
        [testSharedStringTable] @=? parseBS testStrings
    , testCase "correct shared strings parsing: single underline" $
        [withSingleUnderline testSharedStringTable] @=? parseBS testStringsWithSingleUnderline
    , testCase "correct shared strings parsing: double underline" $
        [withDoubleUnderline testSharedStringTable] @=? parseBS testStringsWithDoubleUnderline
    , testCase "correct shared strings parsing even when one of the shared strings entry is just <t/>" $
        [testSharedStringTableWithEmpty] @=? parseBS testStringsWithEmpty
    , testCase "correct comments parsing" $
        [testCommentTable] @=? parseBS testComments
    , testCase "correct custom properties parsing" $
        [testCustomProperties] @==? parseBS testCustomPropertiesXml
    , testCase "proper results from `formatted`" $
        testFormattedResult @==? testRunFormatted
    , testCase "proper results from `formatWorkbook`" $
        testFormatWorkbookResult @==? testFormatWorkbook
    , testCase "formatted . toFormattedCells = id" $ do
        let fmtd = formatted testFormattedCells minimalStyleSheet
        testFormattedCells @==? toFormattedCells (formattedCellMap fmtd) (formattedMerges fmtd)
                                                 (formattedStyleSheet fmtd)
    , testCase "proper results from `conditionallyFormatted`" $
        testCondFormattedResult @==? testRunCondFormatted
    , testCase "toXlsxEither: properly formatted" $
        Right testXlsx @==? toXlsxEither (fromXlsx testTime testXlsx)
    , testCase "toXlsxEither: invalid format" $
        Left (InvalidZipArchive "Did not find end of central directory signature") @==? toXlsxEither "this is not a valid XLSX file"
    , testCase "toXlsx: correct floats parsing (typed and untyped cells are floats by default)"
        $ floatsParsingTests toXlsx
    , testCase "toXlsxFast: correct floats parsing (typed and untyped cells are floats by default)"
        $ floatsParsingTests toXlsxFast
    , testGroup "Codec: sheet state visibility"
        [ testGroup "toXlsxEitherFast"
            [ testProperty "pure state == toXlsxEitherFast (fromXlsx (defXlsxWithState state))" $
                \state ->
                    (Right (Just state) ==) $
                    fmap sheetStateOfDefXlsx $
                    toXlsxEitherFast . fromXlsx testTime $
                    defXlsxWithState state
            , testCase "should otherwise infer visible state by default" $
                Right (Just Visible) @=? (fmap sheetStateOfDefXlsx . toXlsxEitherFast) (fromXlsx testTime defXlsx)
            ]
        , testGroup "toXlsxEither"
            [ testProperty "pure state == toXlsxEither (fromXlsx (defXlsxWithState state))" $
                \state ->
                    (Right (Just state) ==) $
                    fmap sheetStateOfDefXlsx $
                    toXlsxEither . fromXlsx testTime $
                    defXlsxWithState state
            , testCase "should otherwise infer visible state by default" $
                Right (Just Visible) @=? (fmap sheetStateOfDefXlsx . toXlsxEither) (fromXlsx testTime defXlsx)
            ]
        ]
    , CommonTests.tests
    , CondFmtTests.tests
    , PivotTableTests.tests
    , DrawingTests.tests
    , AutoFilterTests.tests
    , StreamTests.tests
    ]

floatsParsingTests :: (ByteString -> Xlsx) -> IO ()
floatsParsingTests parser = do
  bs <- LB.readFile "data/floats.xlsx"
  let xlsx = parser bs
      parsedCells = maybe mempty (_wsCells . snd) $ listToMaybe $ xlsx ^. xlSheets
      expectedCells = M.fromList
        [ ((1,1), def & cellValue ?~ CellDouble 12.0)
        , ((2,1), def & cellValue ?~ CellDouble 13.0)
        , ((3,1), def & cellValue ?~ CellDouble 14.0 & cellStyle ?~ 1)
        , ((4,1), def & cellValue ?~ CellDouble 15.0)
        ]
  expectedCells @==? parsedCells

constSheetName :: Text
constSheetName = "sheet1"

defXlsx :: Xlsx
defXlsx = def & atSheet constSheetName ?~ def

defXlsxWithState :: SheetState -> Xlsx
defXlsxWithState state =
    def & atSheet constSheetName ?~ (wsState .~ state $ def)

sheetStateOfDefXlsx :: Xlsx -> Maybe SheetState
sheetStateOfDefXlsx xlsx =
  xlsx ^. atSheet constSheetName & mapped %~ _wsState
