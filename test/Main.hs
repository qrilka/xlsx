{-# LANGUAGE OverloadedStrings #-}
{-# LANGUAGE RecordWildCards   #-}
{-# LANGUAGE QuasiQuotes       #-}
module Main (main) where

import Control.Lens
import Control.Monad.State.Lazy
import Data.ByteString.Lazy (ByteString)
import qualified Data.ByteString.Lazy as LB
import Data.Map (Map)
import qualified Data.Map as M
import Data.Maybe (mapMaybe)
import qualified Data.Text as T
import Data.Time.Clock.POSIX (POSIXTime)
import qualified Data.Vector as V
import Text.RawString.QQ
import Text.XML

import Test.Tasty (defaultMain, testGroup)
import Test.Tasty.HUnit (testCase)
import Test.Tasty.SmallCheck (testProperty)

import Test.SmallCheck.Series (Positive(..))
import Test.Tasty.HUnit ((@=?))

import Codec.Xlsx
import Codec.Xlsx.Formatted
import Codec.Xlsx.Parser.Internal.PivotTable
import Codec.Xlsx.Types.Internal
import Codec.Xlsx.Types.Internal.CommentTable
import Codec.Xlsx.Types.Internal.CustomProperties
       as CustomProperties
import Codec.Xlsx.Types.Internal.SharedStringTable
import Codec.Xlsx.Types.PivotTable.Internal
import Codec.Xlsx.Writer.Internal.PivotTable

import Common
import Diff
import DrawingTests

main :: IO ()
main = defaultMain $
  testGroup "Tests"
    [ testProperty "col2int . int2col == id" $
        \(Positive i) -> i == col2int (int2col i)
    , testCase "write . read == id" $ do
        let bs = fromXlsx testTime testXlsx
        LB.writeFile "data-test.xlsx" bs
        testXlsx @==? toXlsx (fromXlsx testTime testXlsx)
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
    , testCase "proper pivot table rendering" $ do
      let ptFiles = renderPivotTableFiles testPivotSrcCells 3 testPivotTable
      parseLBS_ def (pvtfTable ptFiles) @==?
        stripContentSpaces (parseLBS_ def testPivotTableDefinition)
      parseLBS_ def (pvtfCacheDefinition ptFiles) @==?
        stripContentSpaces (parseLBS_ def testPivotCacheDefinition)
    , testCase "proper pivot table parsing" $ do
      let sheetName = "Sheet1"
          ref = CellRef "A1:D5"
          forCacheId (CacheId 3) = Just (sheetName, ref, testPivotCacheFields)
          forCacheId _ = Nothing
      Just (sheetName, ref, testPivotCacheFields) @==? parseCache testPivotCacheDefinition
      Just testPivotTable @==? parsePivotTable forCacheId testPivotTableDefinition
    , DrawingTests.tests
    ]

testXlsx :: Xlsx
testXlsx = Xlsx sheets minimalStyles definedNames customProperties
  where
    sheets =
      [("List1", sheet1), ("Another sheet", sheet2), ("with pivot table", pvSheet)]
    sheet1 = Worksheet cols rowProps testCellMap1 drawing ranges
      sheetViews pageSetup cFormatting validations [] (Just autoFilter) tables (Just protection)
    autoFilter = def & afRef ?~ CellRef "A1:E10"
                     & afFilterColumns .~ fCols
    fCols = M.fromList [ (1, Filters ["a","b","ZZZ"])
                       , (2, CustomFiltersAnd (CustomFilter FltrGreaterThanOrEqual "0")
                           (CustomFilter FltrLessThan "42"))]
    tables =
      [ Table
        { tblName = "Table1"
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
    pvSheet = sheetWithPvCells & wsPivotTables .~ [testPivotTable]
    sheetWithPvCells = def & wsCells .~ testPivotSrcCells
    rowProps = M.fromList [(1, RowProps (Just 50) (Just 3))]
    cols = [ColumnsWidth 1 10 15 (Just 1)]
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
testStyleSheet = minimalStyleSheet & styleSheetDxfs .~ [dxf1, dxf2]
                                   & styleSheetNumFmts .~ M.fromList [(164, "0.000")]
                                   & styleSheetCellXfs %~ (++ [cellXf1, cellXf2])
  where
    dxf1 = def & dxfFont ?~ (def & fontBold ?~ True
                                 & fontSize ?~ 12)
    dxf2 = def & dxfFill ?~ (def & fillPattern ?~ (def & fillPatternBgColor ?~ red))
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
        , _dvValidationType   = ValidationTypeList ["aaaa","bbbb","cccc"]
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

testPivotTable :: PivotTable
testPivotTable =
  PivotTable
  { _pvtName = "PivotTable1"
  , _pvtDataCaption = "Values"
  , _pvtLocation = CellRef "A3:D12"
  , _pvtSrcRef = CellRef "A1:D5"
  , _pvtSrcSheet = "Sheet1"
  , _pvtRowFields = [FieldPosition colorField, DataPosition]
  , _pvtColumnFields = [FieldPosition yearField]
  , _pvtDataFields =
      [ DataField
        { _dfName = "Sum of field Price"
        , _dfField = priceField
        , _dfFunction = ConsolidateSum
        }
      , DataField
        { _dfName = "Sum of field Count"
        , _dfField = countField
        , _dfFunction = ConsolidateSum
        }
      ]
  , _pvtFields =
      [ PivotFieldInfo colorField False [CellText "green"]
      , PivotFieldInfo yearField True []
      , PivotFieldInfo priceField False []
      , PivotFieldInfo countField False []
      ]
  , _pvtRowGrandTotals = True
  , _pvtColumnGrandTotals = False
  , _pvtOutline = False
  , _pvtOutlineData = False
  }
  where
    colorField = PivotFieldName "Color"
    yearField = PivotFieldName "Year"
    priceField = PivotFieldName "Price"
    countField = PivotFieldName "Count"

testPivotSrcCells :: CellMap
testPivotSrcCells =
  M.fromList $
  concat
    [ [((row, col), def & cellValue ?~ v) | (col, v) <- zip [1 ..] cells]
    | (row, cells) <- zip [1 ..] cellMap
    ]
  where
    cellMap =
      [ [CellText "Color", CellText "Year", CellText "Price", CellText "Count"]
      , [CellText "green", CellDouble 2012, CellDouble 12.23, CellDouble 17]
      , [CellText "white", CellDouble 2011, CellDouble 73.99, CellDouble 21]
      , [CellText "red", CellDouble 2012, CellDouble 10.19, CellDouble 172]
      , [CellText "white", CellDouble 2012, CellDouble 34.99, CellDouble 49]
      ]

testPivotCacheFields :: [CacheField]
testPivotCacheFields =
  [ CacheField
      (PivotFieldName "Color")
      [CellText "green", CellText "white", CellText "red"]
  , CacheField (PivotFieldName "Year") [CellDouble 2012, CellDouble 2011]
  , CacheField
      (PivotFieldName "Price")
      [CellDouble 12.23, CellDouble 73.99, CellDouble 10.19, CellDouble 34.99]
  , CacheField
      (PivotFieldName "Count")
      [CellDouble 17, CellDouble 21, CellDouble 172, CellDouble 49]
  ]

testPivotTableDefinition :: ByteString
testPivotTableDefinition = [r|
<?xml version="1.0" encoding="UTF-8" standalone="yes"?><!--Pivot table generated by xlsx-->
<pivotTableDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" name="PivotTable1" cacheId="3" dataOnRows="1" colGrandTotals="0" dataCaption="Values">
  <location ref="A3:D12" firstHeaderRow="1" firstDataRow="2" firstDataCol="1"/>
  <pivotFields>
    <pivotField name="Color" axis="axisRow" showAll="0" outline="0">
      <items>
        <item h="1" x="0"/><item x="1"/><item x="2"/><item t="default"/>
      </items>
    </pivotField>
    <pivotField name="Year" axis="axisCol" showAll="0" outline="1">
      <items>
        <item x="0"/><item x="1"/><item t="default"/>
      </items>
    </pivotField>
    <pivotField name="Price" dataField="1" showAll="0" outline="0"/>
    <pivotField name="Count" dataField="1" showAll="0" outline="0"/>
 </pivotFields>
  <rowFields><field x="0"/><field x="-2"/></rowFields>
  <colFields><field x="1"/></colFields>
  <dataFields>
    <dataField name="Sum of field Price" fld="2"/>
    <dataField name="Sum of field Count" fld="3"/>
  </dataFields>
</pivotTableDefinition>
|]

testPivotCacheDefinition :: ByteString
testPivotCacheDefinition = [r|
<?xml version="1.0" encoding="UTF-8" standalone="yes"?><!--Pivot cache definition generated by xlsx-->
<pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    invalid="1" refreshOnLoad="1"
    xmlns:ns="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <cacheSource type="worksheet">
    <worksheetSource ref="A1:D5" sheet="Sheet1"/>
  </cacheSource>
  <cacheFields>
    <cacheField name="Color">
     <sharedItems>
       <s v="green"/><s v="white"/><s v="red"/>
     </sharedItems>
    </cacheField>
    <cacheField name="Year">
     <sharedItems containsNumber="1" containsString="0" containsSemiMixedTypes="0">
       <n v="2012"/><n v="2011"/>
     </sharedItems>
    </cacheField>
    <cacheField name="Price">
     <sharedItems containsNumber="1" containsString="0" containsSemiMixedTypes="0">
       <n v="12.23"/><n v="73.99"/><n v="10.19"/><n v="34.99"/>
     </sharedItems>
    </cacheField>
    <cacheField name="Count">
     <sharedItems containsNumber="1" containsString="0" containsSemiMixedTypes="0">
       <n v="17"/><n v="21"/><n v="172"/><n v="49"/>
     </sharedItems>
    </cacheField>
</cacheFields>
</pivotCacheDefinition>
|]

stripContentSpaces :: Document -> Document
stripContentSpaces doc@Document {documentRoot = root} =
  doc {documentRoot = go root}
  where
    go e@Element {elementNodes = nodes} =
      e {elementNodes = mapMaybe goNode nodes}
    goNode (NodeElement el) = Just $ NodeElement (go el)
    goNode t@(NodeContent txt) =
      if T.strip txt == T.empty
        then Nothing
        else Just t
    goNode other = Just $ other
