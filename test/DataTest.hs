{-# LANGUAGE OverloadedStrings #-}
{-# LANGUAGE QuasiQuotes       #-}
module Main (main) where

import           Control.Lens
import           Control.Monad.State.Lazy
import           Data.ByteString.Lazy                        (ByteString)
import qualified Data.Map                                    as M
import           Data.Time.Clock.POSIX                       (POSIXTime)
import qualified Data.Vector                                 as V
import           Text.RawString.QQ
import           Text.XML
import           Text.XML.Cursor

import           Test.Tasty                                  (defaultMain,
                                                              testGroup)
import           Test.Tasty.HUnit                            (testCase)
import           Test.Tasty.SmallCheck                       (testProperty)

import           Test.SmallCheck.Series                      (Positive (..))
import           Test.Tasty.HUnit                            (HUnitFailure (..),
                                                              (@=?))

import           Codec.Xlsx
import           Codec.Xlsx.Formatted
import           Codec.Xlsx.Parser.Internal
import           Codec.Xlsx.Types.Internal.CommentTable
import           Codec.Xlsx.Types.Internal.CustomProperties  as CustomProperties
import           Codec.Xlsx.Types.Internal.SharedStringTable

import           Diff

main :: IO ()
main = defaultMain $
  testGroup "Tests"
    [ testProperty "col2int . int2col == id" $
        \(Positive i) -> i == col2int (int2col i)
    , testCase "write . read == id" $
        testXlsx @==? toXlsx (fromXlsx testTime testXlsx)
    , testCase "fromRows . toRows == id" $
        testCellMap1 @=? fromRows (toRows testCellMap1)
    , testCase "fromRight . parseStyleSheet . renderStyleSheet == id" $
        testStyleSheet @=? fromRight (parseStyleSheet (renderStyleSheet  testStyleSheet))
    , testCase "correct shared strings parsing" $
        [testSharedStringTable] @=? testParsedSharedStringTables
    , testCase "correct shared strings parsing even when one of the shared strings entry is just <t/>" $
        [testSharedStringTableWithEmpty] @=? testParsedSharedStringTablesWithEmpty
    , testCase "correct comments parsing" $
        [testCommentTable] @=? testParsedComments
    , testCase "correct custom properties parsing" $
        [testCustomProperties] @==? testParsedCustomProperties
    , testCase "proper results from `formatted`" $
        testFormattedResult @==? testRunFormatted
    , testCase "proper results from `conditionalltyFormatted`" $
        testCondFormattedResult @==? testRunCondFormatted
    ]

testXlsx :: Xlsx
testXlsx = Xlsx sheets minimalStyles definedNames customProperties
  where
    sheets = M.fromList [("List1", sheet1), ("Another sheet", sheet2)]
    sheet1 = Worksheet cols rowProps testCellMap1 ranges sheetViews pageSetup cFormatting
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
                                                   & selectionSqref ?~ SqRef ["A3:A10","B1:G3"]
                                             ]
    pageSetup = Just $ def & pageSetupBlackAndWhite .~ Just True
                           & pageSetupCopies .~ Just 2
                           & pageSetupErrors .~ Just PrintErrorsDash
                           & pageSetupPaperSize .~ Just PaperA4
    customProperties = M.fromList [("some_prop", VtInt 42)]
    cFormatting = M.fromList [(SqRef ["A1:B3"], rules1), (SqRef ["C1:C10"], rules2)]
    cfRule c d = CfRule { _cfrCondition  = c
                        , _cfrDxfId      = Just d
                        , _cfrPriority   = topCfPriority
                        , _cfrStopIfTrue = Nothing
                        }
    rules1 = [ cfRule ContainsBlanks 1
             , cfRule (ContainsText "foo") 2
             , cfRule (CellIs (OpBetween (Formula "A1") (Formula "B10"))) 3
             ]
    rules2 = [ cfRule ContainsErrors 3 ]

testCellMap1 :: CellMap
testCellMap1 = M.fromList [ ((1, 2), cd1_2), ((1, 5), cd1_5)
                          , ((3, 1), cd3_1), ((3, 2), cd3_2), ((3, 7), cd3_7)
                          , ((4, 1), cd4_1), ((4, 2), cd4_2), ((4, 3), cd4_3)
                          ]
  where
    cd v = def {_cellValue=Just v}
    cd1_2 = cd (CellText "just a text")
    cd1_5 = cd (CellDouble 42.4567)
    cd3_1 = cd (CellText "another text")
    cd3_2 = def -- shouldn't it be skipped?
    cd3_7 = cd (CellBool True)
    cd4_1 = cd (CellDouble 1)
    cd4_2 = cd (CellDouble 2)
    cd4_3 = (cd (CellDouble (1+2))) { _cellFormula =
                                            Just $ simpleCellFormula "A4+B4"
                                    }

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
testStyleSheet = minimalStyleSheet & styleSheetDxfs .~ [dxf1, dxf2]
  where
    dxf1 = def & dxfFont ?~ (def & fontBold ?~ True
                                 & fontSize ?~ 12)
    dxf2 = def & dxfFill ?~ (def & fillPattern ?~ (def & fillPatternBgColor ?~ red))
    red = def & colorARGB ?~ "FFFF0000"

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

testCommentTable = CommentTable $ M.fromList
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

testParsedComments ::[CommentTable]
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

testFormattedResult :: Formatted
testFormattedResult = Formatted cm styleSheet merges
  where
    cm = M.fromList [((1, 1), cell11),((1, 2), cell2)]
    cell11 = Cell
        { _cellStyle   = Just 1
        , _cellValue   = Just (CellText "text at A1")
        , _cellComment = Nothing
        , _cellFormula = Nothing }
    cell2 = Cell
        { _cellStyle   = Just 2
        , _cellValue   = Just (CellDouble 1.23)
        , _cellComment = Nothing
        , _cellFormula = Nothing }
    merges = []
    styleSheet =
        minimalStyleSheet & styleSheetCellXfs %~ (++ [cellXf1, cellXf2])
                          & styleSheetFonts   %~ (++ [font1, font2])
    nextFontId = length (minimalStyleSheet ^. styleSheetFonts)
    cellXf1 = def
        { _cellXfApplyFont = Just True
        , _cellXfFontId    = Just nextFontId }
    font1 = def
        { _fontName = Just "Calibri"
        , _fontBold = Just True }
    cellXf2 = def
        { _cellXfApplyFont = Just True
        , _cellXfFontId    = Just (nextFontId + 1) }
    font2 = def
        { _fontItalic = Just True }

testRunFormatted :: Formatted
testRunFormatted = formatted formattedCellMap minimalStyleSheet
  where
    formattedCellMap = flip execState def $ do
        let font1 = def & fontBold ?~ True
                        & fontName ?~ "Calibri"
        at (1, 1) ?= (def & formattedValue ?~ CellText "text at A1"
                          & formattedFont  ?~ font1)
        at (1, 2) ?= (def & formattedValue ?~ CellDouble 1.23
                          & formattedFont . non def . fontItalic ?~ True)

testCondFormattedResult :: CondFormatted
testCondFormattedResult = CondFormatted styleSheet formattings
  where
    styleSheet =
        minimalStyleSheet & styleSheetDxfs .~ dxfs
    dxfs = [ def & dxfFont ?~ (def & fontUnderline ?~ FontUnderlineSingle)
           , def & dxfFont ?~ (def & fontStrikeThrough ?~ True)
           , def & dxfFont ?~ (def & fontBold ?~ True) ]
    formattings = M.fromList [ (SqRef ["A1:A2", "B2:B3"], [cfRule1, cfRule2])
                             , (SqRef ["C3:E10"], [cfRule1])
                             , (SqRef ["F1:G10"], [cfRule3]) ]
    cfRule1 = CfRule
        { _cfrCondition  = ContainsBlanks
        , _cfrDxfId      = Just 0
        , _cfrPriority   = 1
        , _cfrStopIfTrue = Nothing }
    cfRule2 = CfRule
        { _cfrCondition  = BeginsWith "foo"
        , _cfrDxfId      = Just 1
        , _cfrPriority   = 1
        , _cfrStopIfTrue = Nothing }
    cfRule3 = CfRule
        { _cfrCondition  = CellIs (OpGreaterThan (Formula "A1"))
        , _cfrDxfId      = Just 2
        , _cfrPriority   = 1
        , _cfrStopIfTrue = Nothing }

testRunCondFormatted :: CondFormatted
testRunCondFormatted = conditionallyFormatted condFmts minimalStyleSheet
  where
    condFmts = flip execState def $ do
        let cfRule1 = def & condfmtCondition .~ ContainsBlanks
                          & condfmtDxf . dxfFont . non def . fontUnderline ?~ FontUnderlineSingle
            cfRule2 = def & condfmtCondition .~ BeginsWith "foo"
                          & condfmtDxf . dxfFont . non def . fontStrikeThrough ?~ True
            cfRule3 = def & condfmtCondition .~ CellIs (OpGreaterThan (Formula "A1"))
                          & condfmtDxf . dxfFont . non def . fontBold ?~ True
        at "A1:A2"  ?= [cfRule1, cfRule2]
        at "B2:B3"  ?= [cfRule1, cfRule2]
        at "C3:E10" ?= [cfRule1]
        at "F1:G10" ?= [cfRule3]
