-- {-# OPTIONS_GHC -F -pgmF tasty-discover #-}
{-# LANGUAGE OverloadedLists   #-}
{-# LANGUAGE OverloadedStrings #-}
{-# LANGUAGE QuasiQuotes       #-}
{-# LANGUAGE RecordWildCards   #-}
module Main
  ( main
  ) where

import           Control.Lens
import           Control.Monad.State.Lazy
import           Data.ByteString.Lazy                        (ByteString)
import qualified Data.ByteString.Lazy                        as LB
import           Data.Map                                    (Map)
import qualified Data.Map                                    as M
import           Data.Time.Clock.POSIX                       (POSIXTime)
import qualified Data.Vector                                 as V
import qualified StreamTests
import           Text.RawString.QQ
import           Text.XML

import           Test.Tasty                                  (defaultMain,
                                                              testGroup)
import           Test.Tasty.HUnit                            (testCase)

import           Test.Tasty.HUnit                            ((@=?))
import           TestXlsx

import           Codec.Xlsx
import           Codec.Xlsx.Formatted
import           Codec.Xlsx.Types.Internal
import           Codec.Xlsx.Types.Internal.CommentTable
import           Codec.Xlsx.Types.Internal.CustomProperties  as CustomProperties
import           Codec.Xlsx.Types.Internal.SharedStringTable

import           AutoFilterTests
import           Common
import           CommonTests
import           CondFmtTests
import           Diff
import           DrawingTests
import           PivotTableTests

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
    , testCase "proper results from `conditionalltyFormatted`" $
        testCondFormattedResult @==? testRunCondFormatted
    , testCase "toXlsxEither: properly formatted" $
        Right testXlsx @==? toXlsxEither (fromXlsx testTime testXlsx)
    , testCase "toXlsxEither: invalid format" $
        Left (InvalidZipArchive "Did not find end of central directory signature") @==? toXlsxEither "this is not a valid XLSX file"
    , CommonTests.tests
    , CondFmtTests.tests
    , PivotTableTests.tests
    , DrawingTests.tests
    , AutoFilterTests.tests
    , StreamTests.tests
    ]
