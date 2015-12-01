{-# LANGUAGE OverloadedStrings #-}
module Main (main) where

import qualified Data.IntMap as IM
import           Data.Map (Map)
import qualified Data.Map as M
import           Data.Time.Calendar
import           Data.Time.LocalTime
import           System.Time
import           Text.XML
import           Text.XML.Cursor

import           Test.Tasty (defaultMain, testGroup)
import           Test.Tasty.SmallCheck (testProperty)
import           Test.Tasty.HUnit (testCase)

import           Test.HUnit ((@=?))
import           Test.SmallCheck.Series (Positive(..))

import           Codec.Xlsx
import           Codec.Xlsx.Parser.Internal
import           Codec.Xlsx.StyleSheet


main = defaultMain $
  testGroup "Tests"
    [ testProperty "col2int . int2col == id" $
        \(Positive i) -> i == col2int (int2col i)
    , testCase "write . read == id" $
         testXlsx @=? toXlsx (fromXlsx testTime testXlsx)
    , testCase "fromRows . toRows == id" $
         testCellMap @=? fromRows (toRows testCellMap)
    , testCase "correct shared strings parsing" $
         testSharedStrings @=? testParseSharedStrings
    ]

testXlsx :: Xlsx
testXlsx = Xlsx sheets minimalStyles definedNames
  where
    sheets = M.fromList [( "List1", sheet )]
    sheet = Worksheet cols rowProps testCellMap ranges Nothing Nothing
    rowProps = M.fromList [(1, RowProps (Just 50) (Just 3))]
    cols = [ColumnsWidth 1 10 15 1]
    ranges = [mkRange (1,1) (1,2), mkRange (2,2) (10, 5)]
    minimalStyles = renderStyleSheet minimalStyleSheet
    definedNames = DefinedNames [("SampleName", Nothing, "A10:A20")]

testCellMap :: CellMap
testCellMap = M.fromList [ ((1, 2), cd1), ((1, 5), cd2)
                         , ((3, 1), cd3), ((3, 2), cd4), ((3, 7), cd5)
                         ]
  where
    cd v = Cell{_cellValue=Just v, _cellStyle=Nothing}
    cd1 = cd (CellText "just a text")
    cd2 = cd (CellDouble 42.4567)
    cd3 = cd (CellText "another text")
    cd4 = Cell{_cellValue=Nothing, _cellStyle=Nothing} -- shouldn't it be skipped?
    cd5 = cd $(CellBool True)

testTime :: ClockTime
testTime = TOD 123 567

testSharedStrings = IM.fromAscList $ zip [0..] ["plain text", "Just example"]

testParseSharedStrings = parseSharedStrings $ fromDocument $ parseLBS_ def strings
    where
      strings = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\
                \<sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" count=\"2\" uniqueCount=\"2\">\
                \<si><t>plain text</t></si>\
                \<si><r><t>Just </t></r><r><rPr><b val=\"true\"/><u val=\"single\"/>\
                \<sz val=\"10\"/><rFont val=\"Arial\"/><family val=\"2\"/></rPr><t>example</t></r></si>\
                \</sst>"
