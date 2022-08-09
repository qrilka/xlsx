{-# LANGUAGE OverloadedStrings #-}
{-# LANGUAGE RecordWildCards   #-}
{-# LANGUAGE QuasiQuotes       #-}
module Main
  ( main
  ) where

import Control.Lens
import Control.Monad.State.Lazy
import Data.ByteString.Lazy (ByteString)
import qualified Data.ByteString.Lazy as LB
import Data.Map (Map)
import qualified Data.Map as M
import Data.Maybe (isNothing, isJust)
import Data.Time.Clock.POSIX (POSIXTime)
import qualified Data.Vector as V
import Text.RawString.QQ
import Text.XML

import Test.Tasty (defaultMain, testGroup)
import Test.Tasty.HUnit (testCase)

import Test.Tasty.HUnit ((@=?),(@?=),(@?))

import Codec.Xlsx
import Codec.Xlsx.Formatted
import Codec.Xlsx.Types.Internal
import Codec.Xlsx.Types.Internal.CommentTable
import Codec.Xlsx.Types.Internal.CustomProperties
       as CustomProperties
import Codec.Xlsx.Types.Internal.SharedStringTable
import qualified Codec.Xlsx.Types.SheetState as SheetState

import AutoFilterTests
import Common
import CommonTests
import CondFmtTests
import Diff
import PivotTableTests
import DrawingTests

main :: IO ()
main = defaultMain $
  testGroup "Tests"
    [ testCase "write . read == id" $ do
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
    , testCase "correct shared strings parsing even when one of the shared strings entry is just <t/>" $
        [testSharedStringTableWithEmpty] @=? parseBS testStringsWithEmpty
    , testCase "correct comments parsing" $
        [testCommentTable] @=? parseBS testComments
    , testCase "correct custom properties parsing" $
        [testCustomProperties] @==? parseBS testCustomPropertiesXml
    , testCase "proper results from `formatted`" $
        testFormattedResult @==? testRunFormatted
    , testCase "formatted . toFormattedCells = id" $ do
        let fmtd = formatted testFormattedCells minimalStyleSheet
        testFormattedCells @==? toFormattedCells (formattedCellMap fmtd) (formattedMerges fmtd)
                                                 (formattedStyleSheet fmtd)
    , testCase "proper results from `conditionalltyFormatted`" $
        testCondFormattedResult @==? testRunCondFormatted
    , testCase "toXlsxEither: properly formatted" $
        Right testXlsx @==? toXlsxEither (fromXlsx testTime testXlsx)
    , testCase "toXlsxEither: invalid format" $
        Left InvalidZipArchive @==? toXlsxEither "this is not a valid XLSX file"
    , testGroup "atSheet lens setter" $
      let getSheet n xlsx = xlsx ^. atSheet n
          getVisibility n xlsx = xlsx ^. atSheet' n & fmap fst
          xlsx' = testXlsx & atSheet "Abc" ?~ def
        in
          [ testCase "control: sheet not created yet" $
              (testXlsx ^. atSheet "Abc" & isNothing) @? "should be Nothing"
          , testCase "given Just, should create a new visible sheet" $
              (testXlsx & atSheet "Abc" ?~ def & getVisibility "Abc") @?= Just SheetState.Visible
          , testCase "control: sheet created" $
              (xlsx' & getSheet "Abc" & isJust) @? "should be Just"
          , testCase "given Nothing, should delete a sheet" $
              (xlsx' & atSheet "Abc" .~ Nothing & getSheet "Abc" & isNothing) @? "should be Nothing"
          ]
    , CommonTests.tests
    , CondFmtTests.tests
    , PivotTableTests.tests
    , DrawingTests.tests
    , AutoFilterTests.tests
    ]

testXlsx :: Xlsx
testXlsx = Xlsx sheets minimalStyles definedNames customProperties DateBase1904
  where
    sheets =
      map (\(n, ws) -> (n, SheetState.Visible, ws))
        [("List1", sheet1), ("Another sheet", sheet2), ("with pivot table", pvSheet), ("cellrange DV source", cellRangeDvSheet)]
    sheet1 = Worksheet cols rowProps testCellMap1 drawing ranges
      sheetViews pageSetup cFormatting validations [] (Just autoFilter)
      tables (Just protection) sharedFormulas
    sharedFormulas =
      M.fromList
        [ (SharedFormulaIndex 0, SharedFormulaOptions (CellRef "A5:C5") (Formula "A4"))
        , (SharedFormulaIndex 1, SharedFormulaOptions (CellRef "B6:C6") (Formula "B3+12"))
        ]
    autoFilter = def & afRef ?~ CellRef "A1:E10"
                     & afFilterColumns .~ fCols
    fCols = M.fromList [ (1, Filters DontFilterByBlank
                             [FilterValue "a", FilterValue "b",FilterValue "ZZZ"])
                       , (2, CustomFiltersAnd (CustomFilter FltrGreaterThanOrEqual "0")
                           (CustomFilter FltrLessThan "42"))]
    tables =
      [ Table
        { tblName = Just "Table1"
        , tblDisplayName = "Table1"
        , tblRef = CellRef "A3"
        , tblColumns = [TableColumn "another text"]
        , tblAutoFilter = Just (def & afRef ?~ CellRef "A3")
        }
      ]
    protection =
      fullSheetProtection
      { _sprScenarios = False
      , _sprLegacyPassword = Just $ legacyPassword "hard password"
      }
    sheet2 = def & wsCells .~ testCellMap2
    cellRangeDvSheet = def & wsCells .~ cellRangeTestCellMap
    pvSheet = sheetWithPvCells & wsPivotTables .~ [testPivotTable]
    sheetWithPvCells = def & wsCells .~ testPivotSrcCells
    rowProps = M.fromList [(1, RowProps { rowHeight       = Just (CustomHeight 50)
                                        , rowStyle        = Just 3
                                        , rowHidden       = False
                                        })]
    cols = [ColumnsProperties 1 10 (Just 15) (Just 1) False False False]
    drawing = Just $ testDrawing { _xdrAnchors = map resolve $ _xdrAnchors testDrawing }
    resolve :: Anchor RefId RefId -> Anchor FileInfo ChartSpace
    resolve Anchor {..} =
      let obj =
            case _anchObject of
              Picture {..} ->
                let blipFill = (_picBlipFill & bfpImageInfo ?~ fileInfo)
                in Picture
                   { _picMacro = _picMacro
                   , _picPublished = _picPublished
                   , _picNonVisual = _picNonVisual
                   , _picBlipFill = blipFill
                   , _picShapeProperties = _picShapeProperties
                   }
              Graphic nv _ tr ->
                Graphic nv testLineChartSpace tr
      in Anchor
         { _anchAnchoring = _anchAnchoring
         , _anchObject = obj
         , _anchClientData = _anchClientData
         }
    fileInfo = FileInfo "dummy.png" "image/png" "fake contents"
    ranges = [mkRange (1,1) (1,2), mkRange (2,2) (10, 5)]
    minimalStyles = renderStyleSheet minimalStyleSheet
    definedNames = DefinedNames [("SampleName", Nothing, "A10:A20")]
    sheetViews = Just [sheetView1, sheetView2]
    sheetView1 = def & sheetViewRightToLeft ?~ True
                     & sheetViewTopLeftCell ?~ CellRef "B5"
    sheetView2 = def & sheetViewType ?~ SheetViewTypePageBreakPreview
                     & sheetViewWorkbookViewId .~ 5
                     & sheetViewSelection .~ [ def & selectionActiveCell ?~ CellRef "C2"
                                                   & selectionPane ?~ PaneTypeBottomRight
                                             , def & selectionActiveCellId ?~ 1
                                                   & selectionSqref ?~ SqRef [ CellRef "A3:A10"
                                                                             , CellRef "B1:G3"]
                                             ]
    pageSetup = Just $ def & pageSetupBlackAndWhite ?~  True
                           & pageSetupCopies ?~ 2
                           & pageSetupErrors ?~ PrintErrorsDash
                           & pageSetupPaperSize ?~ PaperA4
    customProperties = M.fromList [("some_prop", VtInt 42)]
    cFormatting = M.fromList [(SqRef [CellRef "A1:B3"], rules1), (SqRef [CellRef "C1:C10"], rules2)]
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
testCellMap1 = M.fromList [ ((1, 2), cd1_2), ((1, 5), cd1_5), ((1, 10), cd1_10)
                          , ((3, 1), cd3_1), ((3, 2), cd3_2), ((3, 3), cd3_3), ((3, 7), cd3_7)
                          , ((4, 1), cd4_1), ((4, 2), cd4_2), ((4, 3), cd4_3)
                          , ((5, 1), cd5_1), ((5, 2), cd5_2), ((5, 3), cd5_3)
                          , ((6, 2), cd6_2), ((6, 3), cd6_3)
                          ]
  where
    cd v = def {_cellValue=Just v}
    cd1_2 = cd (CellText "just a text, fließen, русский <> и & \"in quotes\"")
    cd1_5 = cd (CellDouble 42.4567)
    cd1_10 = cd (CellText "")
    cd3_1 = cd (CellText "another text")
    cd3_2 = def -- shouldn't it be skipped?
    cd3_3 = def & cellValue ?~ CellError ErrorDiv0
                & cellFormula ?~ simpleCellFormula "1/0"
    cd3_7 = cd (CellBool True)
    cd4_1 = cd (CellDouble 1)
    cd4_2 = cd (CellDouble 123456789012345)
    cd4_3 = (cd (CellDouble (1+2))) { _cellFormula =
                                            Just $ simpleCellFormula "A4+B4<>11"
                                    }
    cd5_1 = def & cellFormula ?~ sharedFormulaByIndex (SharedFormulaIndex 0)
    cd5_2 = def & cellFormula ?~ sharedFormulaByIndex (SharedFormulaIndex 0)
    cd5_3 = def & cellFormula ?~ sharedFormulaByIndex (SharedFormulaIndex 0)
    cd6_2 = def & cellFormula ?~ sharedFormulaByIndex (SharedFormulaIndex 1)
    cd6_3 = def & cellFormula ?~ sharedFormulaByIndex (SharedFormulaIndex 1)

cellRangeTestCellMap :: CellMap
cellRangeTestCellMap = M.fromList [ ((1, 1), def & cellValue ?~ CellText "A-A-A")
                                    , ((2, 1), def & cellValue ?~ CellText "B-B-B")
                                    , ((1, 2), def & cellValue ?~ CellText "C-C-C")
                                    , ((2, 2), def & cellValue ?~ CellText "D-D-D")
                                    , ((1, 3), def & cellValue ?~ CellDouble 6)
                                    , ((2, 3), def & cellValue ?~ CellDouble 7)
                                    , ((3, 1), def & cellValue ?~ CellDouble 5)
                                    , ((3, 2), def & cellValue ?~ CellText "numbers!")
                                    , ((3, 3), def & cellValue ?~ CellDouble 5)
                                  ]

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
                          , ((11, 4), def & cellComment ?~ comment3)
                          ]
  where
    comment1 = Comment (XlsxText "simple comment") "bob" True
    comment2 = Comment (XlsxRichText [rich1, rich2]) "alice" False
    comment3 = Comment (XlsxText "comment for an empty cell") "bob" True
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

fromRight :: Show a => Either a b -> b
fromRight (Right b) = b
fromRight (Left x) = error $ "Right _ was expected but Left " ++ show x ++ " found"

testStyleSheet :: StyleSheet
testStyleSheet = minimalStyleSheet & styleSheetDxfs .~ [dxf1, dxf2, dxf3]
                                   & styleSheetNumFmts .~ M.fromList [(164, "0.000")]
                                   & styleSheetCellXfs %~ (++ [cellXf1, cellXf2])
  where
    dxf1 = def & dxfFont ?~ (def & fontBold ?~ True
                                 & fontSize ?~ 12)
    dxf2 = def & dxfFill ?~ (def & fillPattern ?~ (def & fillPatternBgColor ?~ red))
    dxf3 = def & dxfNumFmt ?~ NumFmt 164 "0.000"
    red = def & colorARGB ?~ "FFFF0000"
    cellXf1 = def
        { _cellXfApplyNumberFormat = Just True
        , _cellXfNumFmtId          = Just 2 }
    cellXf2 = def
        { _cellXfApplyNumberFormat = Just True
        , _cellXfNumFmtId          = Just 164 }

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

testCommentTable :: CommentTable
testCommentTable = CommentTable $ M.fromList
    [ (CellRef "D4", Comment (XlsxRichText rich) "Bob" True)
    , (CellRef "A2", Comment (XlsxText "Some comment here") "CBR" True) ]
  where
    rich = [ RichTextRun
             { _richTextRunProperties =
               Just $ def & runPropertiesBold ?~ True
                          & runPropertiesCharset ?~ 1
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
    cm = M.fromList [ ((1, 1), cell11)
                    , ((1, 2), cell12)
                    , ((2, 5), cell25) ]
    cell11 = Cell
        { _cellStyle   = Just 1
        , _cellValue   = Just (CellText "text at A1")
        , _cellComment = Nothing
        , _cellFormula = Nothing }
    cell12 = Cell
        { _cellStyle   = Just 2
        , _cellValue   = Just (CellDouble 1.23)
        , _cellComment = Nothing
        , _cellFormula = Nothing }
    cell25 = Cell
        { _cellStyle   = Just 3
        , _cellValue   = Just (CellDouble 1.23456)
        , _cellComment = Nothing
        , _cellFormula = Nothing }
    merges = []
    styleSheet =
        minimalStyleSheet & styleSheetCellXfs %~ (++ [cellXf1, cellXf2, cellXf3])
                          & styleSheetFonts   %~ (++ [font1, font2])
                          & styleSheetNumFmts .~ numFmts
    nextFontId = length (minimalStyleSheet ^. styleSheetFonts)
    cellXf1 = def
        { _cellXfApplyFont = Just True
        , _cellXfFontId    = Just nextFontId }
    font1 = def
        { _fontName = Just "Calibri"
        , _fontBold = Just True }
    cellXf2 = def
        { _cellXfApplyFont         = Just True
        , _cellXfFontId            = Just (nextFontId + 1)
        , _cellXfApplyNumberFormat = Just True
        , _cellXfNumFmtId          = Just 164 }
    font2 = def
        { _fontItalic = Just True }
    cellXf3 = def
        { _cellXfApplyNumberFormat = Just True
        , _cellXfNumFmtId          = Just 2 }
    numFmts = M.fromList [(164, "0.0000")]

testRunFormatted :: Formatted
testRunFormatted = formatted formattedCellMap minimalStyleSheet
  where
    formattedCellMap = flip execState def $ do
        let font1 = def & fontBold ?~ True
                        & fontName ?~ "Calibri"
        at (1, 1) ?= (def & formattedCell . cellValue ?~ CellText "text at A1"
                          & formattedFormat . formatFont  ?~ font1)
        at (1, 2) ?= (def & formattedCell . cellValue ?~ CellDouble 1.23
                          & formattedFormat . formatFont . non def . fontItalic ?~ True
                          & formattedFormat . formatNumberFormat ?~ fmtDecimalsZeroes 4)
        at (2, 5) ?= (def & formattedCell . cellValue ?~ CellDouble 1.23456
                          & formattedFormat . formatNumberFormat ?~ StdNumberFormat Nf2Decimal)

testCondFormattedResult :: CondFormatted
testCondFormattedResult = CondFormatted styleSheet formattings
  where
    styleSheet =
        minimalStyleSheet & styleSheetDxfs .~ dxfs
    dxfs = [ def & dxfFont ?~ (def & fontUnderline ?~ FontUnderlineSingle)
           , def & dxfFont ?~ (def & fontStrikeThrough ?~ True)
           , def & dxfFont ?~ (def & fontBold ?~ True) ]
    formattings = M.fromList [ (SqRef [CellRef "A1:A2", CellRef "B2:B3"], [cfRule1, cfRule2])
                             , (SqRef [CellRef "C3:E10"], [cfRule1])
                             , (SqRef [CellRef "F1:G10"], [cfRule3]) ]
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

testFormattedCells :: Map (Int, Int) FormattedCell
testFormattedCells = flip execState def $ do
    at (1,1) ?= (def & formattedRowSpan .~ 5
                     & formattedColSpan .~ 5
                     & formattedFormat . formatBorder . non def . borderTop .
                                                        non def . borderStyleLine ?~ LineStyleDashed
                     & formattedFormat . formatBorder . non def . borderBottom .
                                                        non def . borderStyleLine ?~ LineStyleDashed)
    at (10,2) ?= (def & formattedFormat . formatFont . non def . fontBold ?~ True)

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
        at (CellRef "A1:A2")  ?= [cfRule1, cfRule2]
        at (CellRef "B2:B3")  ?= [cfRule1, cfRule2]
        at (CellRef "C3:E10") ?= [cfRule1]
        at (CellRef "F1:G10") ?= [cfRule3]

validations :: Map SqRef DataValidation
validations = M.fromList
    [ ( SqRef [CellRef "A1"], def
      )
      , ( SqRef [CellRef "A1", CellRef "B2:C3"], def
        { _dvAllowBlank       = True
        , _dvError            = Just "incorrect data"
        , _dvErrorStyle       = ErrorStyleInformation
        , _dvErrorTitle       = Just "error title"
        , _dvPrompt           = Just "enter data"
        , _dvPromptTitle      = Just "prompt title"
        , _dvShowDropDown     = True
        , _dvShowErrorMessage = True
        , _dvShowInputMessage = True
        , _dvValidationType   = ValidationTypeList $ ListExpression ["aaaa","bbbb","cccc"]
        }
      )
      , ( SqRef [CellRef "E50:E55"], def
        { _dvAllowBlank       = True
        , _dvError            = Just "incorrect data"
        , _dvErrorStyle       = ErrorStyleInformation
        , _dvErrorTitle       = Just "error title"
        , _dvPrompt           = Just "Input kebab string"
        , _dvPromptTitle      = Just "I love kebab-case"
        , _dvShowDropDown     = True
        , _dvShowErrorMessage = True
        , _dvShowInputMessage = True
        , _dvValidationType   = ValidationTypeList $ RangeExpression $ CellRef "'cellrange DV source'!$A$1:$B$2"
        }
      )
    , ( SqRef [CellRef "A6", CellRef "I2"], def
        { _dvAllowBlank       = False
        , _dvError            = Just "aaa"
        , _dvErrorStyle       = ErrorStyleWarning
        , _dvErrorTitle       = Just "bbb"
        , _dvPrompt           = Just "ccc"
        , _dvPromptTitle      = Just "ddd"
        , _dvShowDropDown     = False
        , _dvShowErrorMessage = False
        , _dvShowInputMessage = False
        , _dvValidationType   = ValidationTypeDecimal $ ValGreaterThan $ Formula "10"
        }
      )
    , ( SqRef [CellRef "A7"], def
        { _dvAllowBlank       = False
        , _dvError            = Just "aaa"
        , _dvErrorStyle       = ErrorStyleStop
        , _dvErrorTitle       = Just "bbb"
        , _dvPrompt           = Just "ccc"
        , _dvPromptTitle      = Just "ddd"
        , _dvShowDropDown     = False
        , _dvShowErrorMessage = False
        , _dvShowInputMessage = False
        , _dvValidationType   = ValidationTypeWhole $ ValNotBetween (Formula "10") (Formula "12")
        }
      )
    ]
