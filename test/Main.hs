{-# LANGUAGE OverloadedLists   #-}
{-# LANGUAGE OverloadedStrings #-}
{-# LANGUAGE CPP #-}
module Main
  ( main
  ) where

#ifdef USE_MICROLENS
import Lens.Micro
#else
import Control.Lens
#endif
import Control.Monad.State.Lazy
import Data.ByteString.Lazy (ByteString)
import qualified Data.ByteString.Lazy as LB
import Data.Map (Map)
import Data.Maybe (listToMaybe)
import qualified Data.Map as M
import Data.Time.Clock.POSIX (POSIXTime)
import qualified Data.Vector as V
import qualified StreamTests
import Text.RawString.QQ
import Text.XML

import Test.Tasty (defaultMain, testGroup)
import Test.Tasty.HUnit (testCase)

import Test.Tasty.HUnit ((@=?))
import TestXlsx

import Codec.Xlsx
import Codec.Xlsx.Formatted
import Codec.Xlsx.Types.Internal
import Codec.Xlsx.Types.Internal.CommentTable
import Codec.Xlsx.Types.Internal.CustomProperties as CustomProperties
import Codec.Xlsx.Types.Internal.SharedStringTable

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
      parsedCells = maybe mempty ((^. wsCells) . snd) $ listToMaybe $ xlsx ^. xlSheets
      expectedCells = M.fromList
        [ ((1,1), def & cellValue ?~ CellDouble 12.0)
        , ((2,1), def & cellValue ?~ CellDouble 13.0)
        , ((3,1), def & cellValue ?~ CellDouble 14.0 & cellStyle ?~ 1)
        , ((4,1), def & cellValue ?~ CellDouble 15.0)
        ]
  expectedCells @==? parsedCells