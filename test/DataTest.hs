{-# LANGUAGE OverloadedStrings #-}
{-# LANGUAGE QuasiQuotes       #-}
module Main (main) where

import           Control.Lens
import           Control.Monad.State.Lazy
import           Data.ByteString.Lazy                        (ByteString)
import           Data.Map                                    (Map)
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
import           Codec.Xlsx.Types.Internal
import           Codec.Xlsx.Types.Internal.CommentTable
import           Codec.Xlsx.Types.Internal.CustomProperties  as CustomProperties
import           Codec.Xlsx.Types.Internal.SharedStringTable
import           Codec.Xlsx.Types.StyleSheet
import           Codec.Xlsx.Writer.Internal

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
        testStyleSheet @==? fromRight (parseStyleSheet (renderStyleSheet  testStyleSheet))
    , testCase "correct shared strings parsing" $
        [testSharedStringTable] @=? testParsedSharedStringTables
    , testCase "correct shared strings parsing even when one of the shared strings entry is just <t/>" $
        [testSharedStringTableWithEmpty] @=? testParsedSharedStringTablesWithEmpty
    , testCase "correct comments parsing" $
        [testCommentTable] @=? testParsedComments
    , testCase "correct drawing parsing" $
        [testDrawing] @==? parseDrawing testDrawingFile
    , testCase "write . read == id for Drawings" $
        [testDrawing] @==? parseDrawing testWrittenDrawing
    , testCase "correct custom properties parsing" $
        [testCustomProperties] @==? testParsedCustomProperties
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
    ]

testXlsx :: Xlsx
testXlsx = Xlsx sheets minimalStyles definedNames customProperties
  where
    sheets = M.fromList [("List1", sheet1), ("Another sheet", sheet2)]
    sheet1 = Worksheet cols rowProps testCellMap1 drawing ranges sheetViews pageSetup cFormatting
    sheet2 = def & wsCells .~ testCellMap2
    rowProps = M.fromList [(1, RowProps (Just 50) (Just 3))]
    cols = [ColumnsWidth 1 10 15 1]
    drawing = Just $ testDrawing { _xdrAnchors = [pic] }
    pic = head (testDrawing ^. xdrAnchors) & anchObject . picBlipFill . bfpImageInfo ?~ fileInfo
    fileInfo = FileInfo "dummy.png" "image/png" "fake contents"
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

testDrawing = Drawing [ anchor ]
  where
    anchor = Anchor
        { _anchAnchoring  = anchoring
        , _anchObject     = pic
        , _anchClientData = def }
    anchoring = TwoCellAnchor
        { tcaFrom   = unqMarker ( 0,      0) ( 0,     0)
        , tcaTo     = unqMarker (12, 320760) (33, 38160)
        , tcaEditAs = EditAsAbsolute }
    pic = Picture
        { _picMacro           = Nothing
        , _picPublished       = False
        , _picNonVisual       = nonVis
        , _picBlipFill        = bfProps
        , _picShapeProperties = shProps }
    nonVis = PicNonVisual $ PicDrawingNonVisual
        { _pdnvId          = 0
        , _pdnvName        = "Picture 1"
        , _pdnvDescription = Just ""
        , _pdnvHidden      = False
        , _pdnvTitle       = Nothing }
    bfProps = BlipFillProperties
        { _bfpImageInfo = Just (RefId "rId1")
        , _bfpFillMode  = Just FillStretch }
    shProps = ShapeProperties
        { _spXfrm      = Just trnsfrm
        , _spGeometry  = Just PresetGeometry
        , _spOutline   = Just $ LineProperties (Just LineNoFill) }
    trnsfrm = Transform2D
        {  _trRot    = Angle 0
        , _trFlipH   = False
        , _trFlipV   = False
        , _trOffset  = Just (unqPoint2D 0 0)
        , _trExtents = Just (PositiveSize2D (PositiveCoordinate 10074240)
                                            (PositiveCoordinate 5402520)) }

parseDrawing :: ByteString -> [UnresolvedDrawing]
parseDrawing bs = fromCursor . fromDocument $ parseLBS_ def bs

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
        at "A1:A2"  ?= [cfRule1, cfRule2]
        at "B2:B3"  ?= [cfRule1, cfRule2]
        at "C3:E10" ?= [cfRule1]
        at "F1:G10" ?= [cfRule3]
