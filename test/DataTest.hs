{-# LANGUAGE OverloadedStrings #-}
{-# LANGUAGE RecordWildCards   #-}
{-# LANGUAGE QuasiQuotes       #-}
module Main (main) where

import           Control.Lens
import           Control.Monad.State.Lazy
import           Data.ByteString.Lazy                        (ByteString)
import qualified Data.ByteString.Lazy as LB
import           Data.Map                                    (Map)
import qualified Data.Map                                    as M
import           Data.Maybe                                  (mapMaybe)
import qualified Data.Text                                   as T
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
import           Codec.Xlsx.Parser.Internal.PivotTable
import           Codec.Xlsx.Types.Internal
import           Codec.Xlsx.Types.Internal.CommentTable
import           Codec.Xlsx.Types.Internal.CustomProperties  as CustomProperties
import           Codec.Xlsx.Types.Internal.SharedStringTable
import           Codec.Xlsx.Types.PivotTable.Internal
import           Codec.Xlsx.Types.StyleSheet
import           Codec.Xlsx.Writer.Internal
import           Codec.Xlsx.Writer.Internal.PivotTable

import           Diff

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
    , testCase "correct drawing parsing" $
        [testDrawing] @==? parseBS testDrawingFile
    , testCase "write . read == id for Drawings" $
        [testDrawing] @==? parseBS testWrittenDrawing
    , testCase "correct chart parsing" $
        [testChartSpace] @==? parseBS testChartFile
    , testCase "write . read == id for Charts" $
        [testChartSpace] @==? parseBS testWrittenChartSpace
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
      let ptFiles = renderPivotTableFiles 3 testPivotTable
      parseLBS_ def (pvtfTable ptFiles) @==?
        stripContentSpaces (parseLBS_ def testPivotTableDefinition)
      parseLBS_ def (pvtfCacheDefinition ptFiles) @==?
        stripContentSpaces (parseLBS_ def testPivotCacheDefinition)
    , testCase "proper pivot table parsing" $ do
      let sheetName = "Sheet1"
          ref = CellRef "A1:D5"
          fields =
            [ PivotFieldName "Color"
            , PivotFieldName "Year"
            , PivotFieldName "Price"
            , PivotFieldName "Count"
            ]
          forCacheId (CacheId 3) = Just (sheetName, ref, fields)
          forCacheId _ = Nothing
      Just (sheetName, ref, fields) @==? parseCache testPivotCacheDefinition
      Just testPivotTable @==? parsePivotTable forCacheId testPivotTableDefinition
    ]

parseBS :: FromCursor a => ByteString -> [a]
parseBS = fromCursor . fromDocument . parseLBS_ def

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
    sheetWithPvCells = flip execState def $ do
      forM_ (zip [1..] ["Color", "Year", "Price", "Count"]) $ \(c, n) ->
        cellValueAt (1, c) ?= CellText n
      cellValueAt (2, 1) ?= CellText "brown"
      cellValueAt (2, 2) ?= CellDouble 2016
      cellValueAt (2, 3) ?= CellDouble 12.34
      cellValueAt (2, 4) ?= CellDouble 42
    rowProps = M.fromList [(1, RowProps (Just 50) (Just 3))]
    cols = [ColumnsWidth 1 10 15 (Just 1)]
    drawing = Just $ testDrawing { _xdrAnchors = map resolve $ _xdrAnchors testDrawing }
    resolve :: Anchor RefId RefId -> Anchor FileInfo ChartSpace
    resolve Anchor {..} =
      let obj =
            case _anchObject of
              Picture {..} ->
                let picBlipFill = (_picBlipFill & bfpImageInfo ?~ fileInfo)
                in Picture
                   { _picMacro = _picMacro
                   , _picPublished = _picPublished
                   , _picNonVisual = _picNonVisual
                   , _picBlipFill = picBlipFill
                   , _picShapeProperties = _picShapeProperties
                   }
              Graphic nv _ tr ->
                Graphic nv testChartSpace tr
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
                          ]
  where
    comment1 = Comment (XlsxText "simple comment") "bob" True
    comment2 = Comment (XlsxRichText [rich1, rich2]) "alice" False
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

testDrawing :: UnresolvedDrawing
testDrawing = Drawing [anchor1, anchor2]
  where
    anchor1 =
      Anchor
      {_anchAnchoring = anchoring1, _anchObject = pic, _anchClientData = def}
    anchoring1 =
      TwoCellAnchor
      { tcaFrom = unqMarker (0, 0) (0, 0)
      , tcaTo = unqMarker (12, 320760) (33, 38160)
      , tcaEditAs = EditAsAbsolute
      }
    pic =
      Picture
      { _picMacro = Nothing
      , _picPublished = False
      , _picNonVisual = nonVis1
      , _picBlipFill = bfProps
      , _picShapeProperties = shProps
      }
    nonVis1 =
      PicNonVisual $
      NonVisualDrawingProperties
      { _nvdpId = 0
      , _nvdpName = "Picture 1"
      , _nvdpDescription = Just ""
      , _nvdpHidden = False
      , _nvdpTitle = Nothing
      }
    bfProps =
      BlipFillProperties
      {_bfpImageInfo = Just (RefId "rId1"), _bfpFillMode = Just FillStretch}
    shProps =
      ShapeProperties
      { _spXfrm = Just trnsfrm
      , _spGeometry = Just PresetGeometry
      , _spFill = Nothing
      , _spOutline = Just $ LineProperties (Just NoFill)
      }
    trnsfrm =
      Transform2D
      { _trRot = Angle 0
      , _trFlipH = False
      , _trFlipV = False
      , _trOffset = Just (unqPoint2D 0 0)
      , _trExtents =
          Just
            (PositiveSize2D
               (PositiveCoordinate 10074240)
               (PositiveCoordinate 5402520))
      }
    anchor2 =
      Anchor
      { _anchAnchoring = anchoring2
      , _anchObject = graphic
      , _anchClientData = def
      }
    anchoring2 =
      TwoCellAnchor
      { tcaFrom = unqMarker (0, 87840) (21, 131040)
      , tcaTo = unqMarker (7, 580320) (38, 132480)
      , tcaEditAs = EditAsOneCell
      }
    graphic =
      Graphic
      { _grNonVisual = nonVis2
      , _grChartSpace = RefId "rId2"
      , _grTransform = transform
      }
    nonVis2 = GraphNonVisual $
      NonVisualDrawingProperties
      { _nvdpId = 1
      , _nvdpName = ""
      , _nvdpDescription = Nothing
      , _nvdpHidden = False
      , _nvdpTitle = Nothing
      }
    transform =
      Transform2D
      { _trRot = Angle 0
      , _trFlipH = False
      , _trFlipV = False
      , _trOffset = Just (unqPoint2D 0 0)
      , _trExtents =
          Just
            (PositiveSize2D
               (PositiveCoordinate 10074240)
               (PositiveCoordinate 5402520))
      }


testDrawingFile :: ByteString
testDrawingFile = [r|
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <xdr:twoCellAnchor editAs="absolute">
    <xdr:from>
      <xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff>
      <xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff>
    </xdr:from>
    <xdr:to>
      <xdr:col>12</xdr:col><xdr:colOff>320760</xdr:colOff>
      <xdr:row>33</xdr:row><xdr:rowOff>38160</xdr:rowOff>
    </xdr:to>
    <xdr:pic>
      <xdr:nvPicPr><xdr:cNvPr id="0" name="Picture 1" descr=""/><xdr:cNvPicPr/></xdr:nvPicPr>
      <xdr:blipFill><a:blip r:embed="rId1"></a:blip><a:stretch/></xdr:blipFill>
      <xdr:spPr>
        <a:xfrm><a:off x="0" y="0"/><a:ext cx="10074240" cy="5402520"/></a:xfrm>
        <a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:ln><a:noFill/></a:ln>
      </xdr:spPr>
    </xdr:pic>
    <xdr:clientData/>
  </xdr:twoCellAnchor>
  <xdr:twoCellAnchor editAs="oneCell">
    <xdr:from>
      <xdr:col>0</xdr:col><xdr:colOff>87840</xdr:colOff>
      <xdr:row>21</xdr:row><xdr:rowOff>131040</xdr:rowOff>
    </xdr:from>
    <xdr:to>
      <xdr:col>7</xdr:col><xdr:colOff>580320</xdr:colOff>
      <xdr:row>38</xdr:row><xdr:rowOff>132480</xdr:rowOff>
    </xdr:to>
    <xdr:graphicFrame>
      <xdr:nvGraphicFramePr><xdr:cNvPr id="1" name=""/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr>
      <xdr:xfrm><a:off x="0" y="0"/><a:ext cx="10074240" cy="5402520"/></xdr:xfrm>
      <a:graphic>
        <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">
          <c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" r:id="rId2"/>
        </a:graphicData>
      </a:graphic>
    </xdr:graphicFrame>
    <xdr:clientData/>
  </xdr:twoCellAnchor>
</xdr:wsDr>
|]

testWrittenDrawing :: ByteString
testWrittenDrawing = renderLBS def $ toDocument testDrawing

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

testChartFile :: ByteString
testChartFile = [r|
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<c:chart>
<c:title>
 <c:tx><c:rich><a:bodyPr rot="0" anchor="b"/><a:p><a:r><a:t>Line chart title</a:t></a:r></a:p></c:rich></c:tx>
</c:title>
<c:plotArea>
<c:lineChart>
<c:grouping val="standard"/>
<c:ser>
  <c:idx val="0"/><c:order val="0"/>
  <c:tx><c:strRef><c:f>Sheet1!$A$1</c:f></c:strRef></c:tx>
  <c:marker><c:symbol val="none"/></c:marker>
  <c:val><c:numRef><c:f>Sheet1!$B$1:$D$1</c:f></c:numRef></c:val>
  <c:smooth val="0"/>
</c:ser>
<c:ser>
  <c:idx val="1"/><c:order val="1"/>
  <c:tx><c:strRef><c:f>Sheet1!$A$2</c:f></c:strRef></c:tx>
  <c:marker><c:symbol val="none"/></c:marker>
  <c:val><c:numRef><c:f>Sheet1!$B$2:$D$2</c:f></c:numRef></c:val>
  <c:smooth val="0"/>
</c:ser>
<c:marker val="0"/>
<c:smooth val="0"/>
</c:lineChart>
</c:plotArea>
<c:plotVisOnly val="1"/>
<c:dispBlanksAs val="gap"/>
</c:chart>
</c:chartSpace>
|]

testChartSpace :: ChartSpace
testChartSpace =
  ChartSpace
  { _chspTitle = Just $ ChartTitle titleBody
  , _chspCharts = charts
  , _chspLegend = Nothing
  , _chspPlotVisOnly = Just True
  , _chspDispBlanksAs = Just DispBlanksAsGap
  }
  where
    titleBody =
      TextBody
      { _txbdRotation = Angle 0
      , _txbdSpcFirstLastPara = False
      , _txbdVertOverflow = TextVertOverflow
      , _txbdVertical = TextVerticalHorz
      , _txbdWrap = TextWrapSquare
      , _txbdAnchor = TextAnchoringBottom
      , _txbdAnchorCenter = False
      , _txbdParagraphs =
          [TextParagraph Nothing [RegularRun Nothing "Line chart title"]]
      }
    charts =
      [ LineChart
        { _lnchGrouping = StandardGrouping
        , _lnchSeries = series
        , _lnchMarker = Just False
        , _lnchSmooth = Just False
        }
      ]
    series =
      [ LineSeries
        { _lnserShared = Series . Just $ Formula "Sheet1!$A$1"
        , _lnserMarker = Just markerNone
        , _lnserDataLblProps = Nothing
        , _lnserVal = Just $ Formula "Sheet1!$B$1:$D$1"
        , _lnserSmooth = Just False
        }
      , LineSeries
        { _lnserShared = Series . Just $ Formula "Sheet1!$A$2"
        , _lnserMarker = Just markerNone
        , _lnserDataLblProps = Nothing
        , _lnserVal = Just $ Formula "Sheet1!$B$2:$D$2"
        , _lnserSmooth = Just False
        }
      ]
    markerNone =
      DataMarker {_dmrkSymbol = Just DataMarkerNone, _dmrkSize = Nothing}

testWrittenChartSpace :: ByteString
testWrittenChartSpace = renderLBS def{rsNamespaces=nss} $ toDocument testChartSpace
  where
    nss = [ ("c", "http://schemas.openxmlformats.org/drawingml/2006/chart")
          , ("a", "http://schemas.openxmlformats.org/drawingml/2006/main") ]


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
      [ PivotFieldInfo colorField False
      , PivotFieldInfo yearField True
      , PivotFieldInfo priceField False
      , PivotFieldInfo countField False
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

testPivotTableDefinition :: ByteString
testPivotTableDefinition = [r|
<?xml version="1.0" encoding="UTF-8" standalone="yes"?><!--Pivot table generated by xlsx-->
<pivotTableDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" name="PivotTable1" cacheId="3" dataOnRows="1" colGrandTotals="0" dataCaption="Values">
  <location ref="A3:D12" firstHeaderRow="1" firstDataRow="2" firstDataCol="1"/>
  <pivotFields>
    <pivotField name="Color" axis="axisRow" showAll="0" outline="0">
      <items>
        <item t="default"/>
      </items>
    </pivotField>
    <pivotField name="Year" axis="axisCol" showAll="0" outline="1">
      <items>
        <item t="default"/>
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
    <cacheField name="Color"/>
    <cacheField name="Year"/>
    <cacheField name="Price"/>
    <cacheField name="Count"/>
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
