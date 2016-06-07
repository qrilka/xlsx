{-# LANGUAGE OverloadedStrings #-}
{-# LANGUAGE QuasiQuotes       #-}
module Main (main) where

import           Control.Lens
import           Data.ByteString.Lazy               (ByteString)
import qualified Data.Map                           as M
import qualified Data.HashMap.Strict                as HM
import           Data.Time.Clock.POSIX              (POSIXTime)
import qualified Data.Vector                        as V
import           Text.RawString.QQ
import           Text.XML
import           Text.XML.Cursor

import           Test.Tasty                         (defaultMain, testGroup)
import           Test.Tasty.HUnit                   (testCase)
import           Test.Tasty.SmallCheck              (testProperty)

import           Test.HUnit                         ((@=?))
import           Test.SmallCheck.Series             (Positive (..))

import           Codec.Xlsx
import           Codec.Xlsx.Parser.Internal
import           Codec.Xlsx.Types.Internal.CustomProperties as CustomProperties
import           Codec.Xlsx.Types.Internal.SharedStringTable


main :: IO ()
main = defaultMain $
  testGroup "Tests"
    [ testProperty "col2int . int2col == id" $
        \(Positive i) -> i == col2int (int2col i)
    , testCase "write . read == id" $
        testXlsx @=? toXlsx (fromXlsx testTime testXlsx)
    , testCase "fromRows . toRows == id" $
        testCellMap1 @=? fromRows (toRows testCellMap1)
    , testCase "fromRight . parseStyleSheet . renderStyleSheet == id" $
        testStyleSheet @=? fromRight (parseStyleSheet (renderStyleSheet  testStyleSheet))
    , testCase "correct shared strings parsing" $
        [testSharedStringTable] @=? testParsedSharedStringTables
    , testCase "correct shared strings parsing even when one of the shared strings entry is just <t/>" $
        [testSharedStringTableWithEmpty] @=? testParsedSharedStringTablesWithEmpty
    , testCase "correct comments parsing" $
        [testCommentsTable] @=? testParsedComments
    , testCase "correct custom properties parsing" $
        [testCustomProperties] @=? testParsedCustomProperties
    ]

testXlsx :: Xlsx
testXlsx = Xlsx sheets minimalStyles definedNames customProperties
  where
    sheets = M.fromList [("List1", sheet1), ("Another sheet", sheet2)]
    sheet1 = Worksheet cols rowProps testCellMap1 ranges sheetViews pageSetup
    sheet2 = def & wsCells .~ testCellMap2
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
    customProperties = M.fromList [("some_prop", VtInt 42)]

testCellMap1 :: CellMap
testCellMap1 = M.fromList [ ((1, 2), cd1), ((1, 5), cd2)
                         , ((3, 1), cd3), ((3, 2), cd4), ((3, 7), cd5)
                         ]
  where
    cd v = def {_cellValue=Just v}
    cd1 = cd (CellText "just a text")
    cd2 = cd (CellDouble 42.4567)
    cd3 = cd (CellText "another text")
    cd4 = def -- shouldn't it be skipped?
    cd5 = cd (CellBool True)

testCellMap2 :: CellMap
testCellMap2 = M.fromList [ ((1, 2), def & cellValue ?~ CellText "something here")
                          , ((3, 5), def & cellValue ?~ CellDouble 123.456)
                          , ((2, 4),
                             def & cellValue ?~ CellText "value"
                                 & cellComment ?~ comment1
                            )
                          , ((10, 7),
                             def & cellValue ?~ CellText "value"
                                 & cellComment ?~ comment2
                            )
                          ]
  where
    comment1 = Comment (XlsxText "simple comment") "bob"
    comment2 = Comment (XlsxRichText [rich1, rich2]) "alice"
    rich1 = def & richTextRunText.~ "Look ma!"
                & richTextRunProperties ?~ (
                   def & runPropertiesBold ?~ True
                       & runPropertiesFont ?~ "Tahoma")
    rich2 = def & richTextRunText .~ "It's blue!"
                & richTextRunProperties ?~ (
                   def & runPropertiesItalic ?~ True
                       & runPropertiesColor ?~ (def & colorARGB ?~ "FF000080"))

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

testCommentsTable = CommentsTable $ HM.fromList
    [ ("D4", Comment (XlsxRichText rich) "Bob")
    , ("A2", Comment (XlsxText "Some comment here") "CBR") ]
  where
    rich = [ RichTextRun
             { _richTextRunProperties =
               Just $ def & runPropertiesCharset ?~ 1
                          & runPropertiesColor ?~ def -- TODO: why not Nothing here?
                          & runPropertiesFont ?~ "Calibri"
                          & runPropertiesScheme ?~ FontSchemeMinor
                          & runPropertiesSize ?~ 8.0
             , _richTextRunText = "Bob:"}
           , RichTextRun
             { _richTextRunProperties =
               Just $ def & runPropertiesCharset ?~ 1
                          & runPropertiesColor ?~ def
                          & runPropertiesFont ?~ "Calibri"
                          & runPropertiesScheme ?~ FontSchemeMinor
                          & runPropertiesSize ?~ 8.0
             , _richTextRunText = "Why such high expense?"}]

testParsedComments ::[CommentsTable]
testParsedComments = fromCursor . fromDocument $ parseLBS_ def testComments

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

testComments :: ByteString
testComments = [r|
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <authors>
    <author>Bob</author>
    <author>CBR</author>
  </authors>
  <commentList>
    <comment ref="D4" authorId="0">
      <text>
        <r>
          <rPr>
            <b/><sz val="8"/><color indexed="81"/><rFont val="Calibri"/>
            <charset val="1"/><scheme val="minor"/>
          </rPr>
          <t>Bob:</t>
        </r>
        <r>
          <rPr>
            <sz val="8"/><color indexed="81"/><rFont val="Calibri"/>
            <charset val="1"/> <scheme val="minor"/>
          </rPr>
          <t xml:space="preserve">Why such high expense?</t>
        </r>
      </text>
    </comment>
    <comment ref="A2" authorId="1">
      <text><t>Some comment here</t></text>
    </comment>
  </commentList>
</comments>
|]

testCustomProperties :: CustomProperties
testCustomProperties = CustomProperties.fromList
    [ ("testTextProp", VtLpwstr "test text property value")
    , ("prop2", VtLpwstr "222")
    , ("bool", VtBool False)
    , ("prop333", VtInt 1)
    , ("decimal", VtDecimal 1.234) ]

testParsedCustomProperties ::[CustomProperties]
testParsedCustomProperties = fromCursor . fromDocument $ parseLBS_ def testCustomPropertiesXml

testCustomPropertiesXml :: ByteString
testCustomPropertiesXml = [r|
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="2" name="prop2">
    <vt:lpwstr>222</vt:lpwstr>
  </property>
  <property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="3" name="prop333">
    <vt:int>1</vt:int>
  </property>
  <property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="4" name="testTextProp">
    <vt:lpwstr>test text property value</vt:lpwstr>
  </property>
  <property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="5" name="decimal">
    <vt:decimal>1.234</vt:decimal>
  </property>
  <property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="6" name="bool">
    <vt:bool>false</vt:bool>
  </property>
  <property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="7" name="blob">
    <vt:blob>
    ZXhhbXBs
    ZSBibG9i
    IGNvbnRl
    bnRz
    </vt:blob>
  </property>
</Properties>
|]
