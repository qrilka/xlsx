{-# LANGUAGE CPP #-}
{-# LANGUAGE OverloadedStrings #-}
{-# LANGUAGE QuasiQuotes #-}
module DrawingTests
  ( tests
  , testDrawing
  , testLineChartSpace
  ) where

#ifdef USE_MICROLENS
import Lens.Micro
#else
import Control.Lens
#endif
import Data.ByteString.Lazy (ByteString)
import Test.Tasty (testGroup, TestTree)
import Test.Tasty.HUnit (testCase)
import Text.RawString.QQ
import Text.XML

import Codec.Xlsx
import Codec.Xlsx.Types.Internal
import Codec.Xlsx.Writer.Internal

import Common
import Diff

tests :: TestTree
tests =
  testGroup
    "Drawing tests"
    [ testCase "correct drawing parsing" $
      [testDrawing] @==? parseBS testDrawingFile
    , testCase "write . read == id for Drawings" $
      [testDrawing] @==? parseBS testWrittenDrawing
    , testCase "correct chart parsing" $
      [testLineChartSpace] @==? parseBS testLineChartFile
    , testCase "parse . render == id for line Charts" $
      [testLineChartSpace] @==? parseBS (renderChartSpace testLineChartSpace)
    , testCase "parse . render == id for area Charts" $
      [testAreaChartSpace] @==? parseBS (renderChartSpace testAreaChartSpace)
    , testCase "parse . render == id for bar Charts" $
      [testBarChartSpace] @==? parseBS (renderChartSpace testBarChartSpace)
    , testCase "parse . render == id for pie Charts" $
      [testPieChartSpace] @==? parseBS (renderChartSpace testPieChartSpace)
    , testCase "parse . render == id for scatter Charts" $
      [testScatterChartSpace] @==? parseBS (renderChartSpace testScatterChartSpace)
    ]

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
      { _nvdpId = DrawingElementId 0
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
      , _spOutline = Just $ def {_lnFill = Just NoFill}
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
    nonVis2 =
      GraphNonVisual $
      NonVisualDrawingProperties
      { _nvdpId = DrawingElementId 1
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

testLineChartFile :: ByteString
testLineChartFile = [r|
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
  <c:spPr>
    <a:solidFill><a:srgbClr val="0000FF"/></a:solidFill>
    <a:ln w="28800"><a:solidFill><a:srgbClr val="0000FF"/></a:solidFill></a:ln>
  </c:spPr>
  <c:marker><c:symbol val="none"/></c:marker>
  <c:val><c:numRef><c:f>Sheet1!$B$1:$D$1</c:f></c:numRef></c:val>
  <c:smooth val="0"/>
</c:ser>
<c:ser>
  <c:idx val="1"/><c:order val="1"/>
  <c:tx><c:strRef><c:f>Sheet1!$A$2</c:f></c:strRef></c:tx>
  <c:spPr>
    <a:solidFill><a:srgbClr val="FF0000"/></a:solidFill>
    <a:ln w="28800"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:ln>
  </c:spPr>
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

oneChartChartSpace :: Chart -> ChartSpace
oneChartChartSpace chart =
  ChartSpace
  { _chspTitle = Just $ ChartTitle (Just titleBody)
  , _chspCharts = [chart]
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

renderChartSpace :: ChartSpace -> ByteString
renderChartSpace = renderLBS def {rsNamespaces = nss} . toDocument
  where
    nss =
      [ ("c", "http://schemas.openxmlformats.org/drawingml/2006/chart")
      , ("a", "http://schemas.openxmlformats.org/drawingml/2006/main")
      ]

testLineChartSpace :: ChartSpace
testLineChartSpace = oneChartChartSpace lineChart
  where
    lineChart =
      LineChart
      { _lnchGrouping = StandardGrouping
      , _lnchSeries = series
      , _lnchMarker = Just False
      , _lnchSmooth = Just False
      }
    series =
      [ LineSeries
        { _lnserShared =
            Series
            { _serTx = Just $ Formula "Sheet1!$A$1"
            , _serShapeProperties = Just $ rgbShape "0000FF"
            }
        , _lnserMarker = Just markerNone
        , _lnserDataLblProps = Nothing
        , _lnserVal = Just $ Formula "Sheet1!$B$1:$D$1"
        , _lnserSmooth = Just False
        }
      , LineSeries
        { _lnserShared =
            Series
            { _serTx = Just $ Formula "Sheet1!$A$2"
            , _serShapeProperties = Just $ rgbShape "FF0000"
            }
        , _lnserMarker = Just markerNone
        , _lnserDataLblProps = Nothing
        , _lnserVal = Just $ Formula "Sheet1!$B$2:$D$2"
        , _lnserSmooth = Just False
        }
      ]
    rgbShape color =
      def
      { _spFill = Just $ solidRgb color
      , _spOutline =
          Just $
          LineProperties {_lnFill = Just $ solidRgb color, _lnWidth = 28800}
      }
    markerNone =
      DataMarker {_dmrkSymbol = Just DataMarkerNone, _dmrkSize = Nothing}

testAreaChartSpace :: ChartSpace
testAreaChartSpace = oneChartChartSpace areaChart
  where
    areaChart =
      AreaChart {_archGrouping = Just StandardGrouping, _archSeries = series}
    series =
      [ AreaSeries
        { _arserShared =
            Series
            { _serTx = Just $ Formula "Sheet1!$A$1"
            , _serShapeProperties =
                Just $
                def
                { _spFill = Just $ solidRgb "000088"
                , _spOutline = Just $ def {_lnFill = Just NoFill}
                }
            }
        , _arserDataLblProps = Nothing
        , _arserVal = Just $ Formula "Sheet1!$B$1:$D$1"
        }
      ]

testBarChartSpace :: ChartSpace
testBarChartSpace =
  oneChartChartSpace
    BarChart
    { _brchDirection = DirectionColumn
    , _brchGrouping = Just BarStandardGrouping
    , _brchSeries =
        [ BarSeries
          { _brserShared =
              Series
              { _serTx = Just $ Formula "Sheet1!$A$1"
              , _serShapeProperties =
                  Just $
                  def
                  { _spFill = Just $ solidRgb "000088"
                  , _spOutline = Just $ def {_lnFill = Just NoFill}
                  }
              }
          , _brserDataLblProps = Nothing
          , _brserVal = Just $ Formula "Sheet1!$B$1:$D$1"
          }
        ]
    }

testPieChartSpace :: ChartSpace
testPieChartSpace =
  oneChartChartSpace
    PieChart
    { _pichSeries =
        [ PieSeries
          { _piserShared =
              Series
              { _serTx = Just $ Formula "Sheet1!$A$1"
              , _serShapeProperties = Nothing
              }
          , _piserDataPoints =
              [ def & dpShapeProperties ?~ solidFill "000088"
              , def & dpShapeProperties ?~ solidFill "008800"
              , def & dpShapeProperties ?~ solidFill "880000"
              ]
          , _piserDataLblProps = Nothing
          , _piserVal = Just $ Formula "Sheet1!$B$1:$D$1"
          }
        ]
    }
  where
    solidFill color = def & spFill ?~ solidRgb color

testScatterChartSpace :: ChartSpace
testScatterChartSpace =
  oneChartChartSpace
    ScatterChart
    { _scchStyle = ScatterMarker
    , _scchSeries =
        [ ScatterSeries
          { _scserShared =
              Series
              { _serTx = Just $ Formula "Sheet1!$A$2"
              , _serShapeProperties =
                  Just $ def {_spOutline = Just $ def {_lnFill = Just NoFill}}
              }
          , _scserMarker = Just $ DataMarker (Just DataMarkerSquare) Nothing
          , _scserDataLblProps = Nothing
          , _scserXVal = Just $ Formula "Sheet1!$B$1:$D$1"
          , _scserYVal = Just $ Formula "Sheet1!$B$2:$D$2"
          , _scserSmooth = Nothing
          }
        ]
    }
