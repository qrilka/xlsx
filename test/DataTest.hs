{-# LANGUAGE OverloadedStrings #-}
module Main (main) where

import           Control.Lens
import           Data.ByteString.Lazy (ByteString)
import           Data.Map (Map)
import qualified Data.Map as M
import           Data.Time.Clock.POSIX (POSIXTime)
import qualified Data.Vector as V
import           Text.XML
import           Text.XML.Cursor

import           Test.Tasty (defaultMain, testGroup)
import           Test.Tasty.SmallCheck (testProperty)
import           Test.Tasty.HUnit (testCase)

import           Test.HUnit ((@=?))
import           Test.SmallCheck.Series (Positive(..))

import           Codec.Xlsx
import           Codec.Xlsx.Types.SharedStringTable
import           Codec.Xlsx.Parser.Internal


main = defaultMain $
  testGroup "Tests"
    [ testProperty "col2int . int2col == id" $
        \(Positive i) -> i == col2int (int2col i)
    , testCase "write . read == id" $
         testXlsx @=? toXlsx (fromXlsx testTime testXlsx)
    , testCase "fromRows . toRows == id" $
         testCellMap @=? fromRows (toRows testCellMap)
    , testCase "fromRight . parseStyleSheet . renderStyleSheet == id" $
         testStyleSheet @=? fromRight (parseStyleSheet (renderStyleSheet  testStyleSheet))
    , testCase "correct shared strings parsing" $
         [testSharedStringTable] @=? testParsedSharedStringTables
    , testCase "correct shared strings parsing even when one of the shared strings entry is just <t/>" $
         [testSharedStringTableWithEmpty] @=? testParsedSharedStringTablesWithEmpty
    ]

testXlsx :: Xlsx
testXlsx = Xlsx sheets minimalStyles definedNames
  where
    sheets = M.fromList [( "List1", sheet )]
    sheet = Worksheet cols rowProps testCellMap ranges sheetViews pageSetup
    rowProps = M.fromList [(1, RowProps (Just 50) (Just 3))]
    cols = [ColumnsWidth 1 10 15 1]
    ranges = [mkRange (1,1) (1,2), mkRange (2,2) (10, 5)]
    minimalStyles = renderStyleSheet minimalStyleSheet
    definedNames = DefinedNames [("SampleName", Nothing, "A10:A20")]
    sheetViews = Just [sheetView1, sheetView2]
    sheetView1 = def & sheetViewRightToLeft .~ Just True
                     & sheetViewTopLeftCell .~ Just "B5"
    sheetView2 = def & sheetViewType .~ Just SheetViewTypePageBreakPreview
                     & sheetViewWorkbookViewId .~ 5
                     & sheetViewSelection .~ [ def & selectionActiveCell .~ Just "C2"
                                                   & selectionPane .~ Just PaneTypeBottomRight
                                             , def & selectionActiveCellId .~ Just 1
                                                   & selectionSqref .~ Just ["A3:A10","B1:G3"]
                                             ]
    pageSetup = Just $ def & pageSetupBlackAndWhite .~ Just True
                           & pageSetupCopies .~ Just 2
                           & pageSetupErrors .~ Just PrintErrorsDash
                           & pageSetupPaperSize .~ Just PaperA4

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

testTime :: POSIXTime
testTime = 123

fromRight :: Either a b -> b
fromRight (Right b) = b

testStyleSheet :: StyleSheet
testStyleSheet = minimalStyleSheet

testSharedStringTable :: SharedStringTable
testSharedStringTable = SharedStringTable $ V.fromList items
  where
    items = [text, rich]
    text = XlsxText "plain text"
    empty = XlsxText ""
    rich = XlsxRichText [ RichTextRun Nothing "Just "
                        , RichTextRun (Just props) "example" ]
    props = def & runPropertiesBold .~ Just True
                & runPropertiesUnderline .~ Just FontUnderlineSingle
                & runPropertiesSize .~ Just 10
                & runPropertiesFont .~ Just "Arial"
                & runPropertiesFontFamily .~ Just FontFamilySwiss

testSharedStringTableWithEmpty :: SharedStringTable
testSharedStringTableWithEmpty =
  SharedStringTable $ V.fromList [XlsxText ""]

testParsedSharedStringTables ::[SharedStringTable]
testParsedSharedStringTables = fromCursor . fromDocument $ parseLBS_ def testStrings

testParsedSharedStringTablesWithEmpty :: [SharedStringTable]
testParsedSharedStringTablesWithEmpty = fromCursor . fromDocument $ parseLBS_ def testStringsWithEmpty

testStrings :: ByteString
testStrings = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\
  \<sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" count=\"2\" uniqueCount=\"2\">\
  \<si><t>plain text</t></si>\
  \<si><r><t>Just </t></r><r><rPr><b val=\"true\"/><u val=\"single\"/>\
  \<sz val=\"10\"/><rFont val=\"Arial\"/><family val=\"2\"/></rPr><t>example</t></r></si>\
  \</sst>"

testStringsWithEmpty :: ByteString
testStringsWithEmpty = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\
  \<sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" count=\"2\" uniqueCount=\"2\">\
  \<si><t/></si>\
  \</sst>"
