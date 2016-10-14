{-# LANGUAGE OverloadedStrings #-}
{-# LANGUAGE RecordWildCards   #-}
module Codec.Xlsx.Types.Drawing.Chart where

import           Data.Default
import           Data.Maybe                        (catMaybes, maybeToList)
import           Data.Text                         (Text)
import           Text.XML
import           Text.XML.Cursor

import           Codec.Xlsx.Parser.Internal hiding (n)
import           Codec.Xlsx.Types.Common
import           Codec.Xlsx.Types.Drawing.Common
import           Codec.Xlsx.Writer.Internal

-- | Main Chart holder, combines
-- TODO: title, autoTitleDeleted, pivotFmts
--  view3D, floor, sideWall, backWall, showDLblsOverMax, extLst
data ChartSpace = ChartSpace
  { _chspTitle :: Maybe ChartTitle
  , _chspCharts :: [Chart]
  , _chspLegend :: Maybe Legend
  , _chspPlotVisOnly :: Maybe Bool
  , _chspDispBlanksAs :: Maybe DispBlanksAs
  } deriving (Eq, Show)

-- | Chart title
--
-- TODO: layout, overlay, spPr, txPr, extLst
newtype ChartTitle =
  ChartTitle TextBody
  deriving (Eq, Show)

-- | This simple type specifies the possible ways to display blanks.
--
-- See 21.2.3.10 "ST_DispBlanksAs (Display Blanks As)" (p. 3444)
data DispBlanksAs
  = DispBlanksAsGap
    -- ^ Specifies that blank values shall be left as a gap.
  | DispBlanksAsSpan
    -- ^ Specifies that blank values shall be spanned with a line.
  | DispBlanksAsZero
    -- ^ Specifies that blank values shall be treated as zero.
  deriving (Eq, Show)

-- TODO: legendEntry, layout, overlay, spPr, txPr, extLst
data Legend = Legend
    { _legendPos     :: Maybe LegendPos
    , _legendOverlay :: Maybe Bool
    } deriving (Eq, Show)

-- See 21.2.3.24 "ST_LegendPos (Legend Position)" (p. 3449)
data LegendPos
  = LegendBottom
    -- ^ b (Bottom) Specifies that the legend shall be drawn at the
    -- bottom of the chart.
  | LegendLeft
    -- ^ l (Left) Specifies that the legend shall be drawn at the left
    -- of the chart.
  | LegendRight
    -- ^ r (Right) Specifies that the legend shall be drawn at the
    -- right of the chart.
  | LegendTop
    -- ^ t (Top) Specifies that the legend shall be drawn at the top
    -- of the chart.
  | LegendTopRight
    -- ^ tr (Top Right) Specifies that the legend shall be drawn at
    -- the top right of the chart.
  deriving (Eq, Show)

-- | Specific Chart
-- TODO:
--   areaChart, area3DChart, line3DChart, stockChart, radarChart, scatterChart,
--   pieChart, pie3DChart, doughnutChart, barChart, bar3DChart, ofPieChart,
--   surfaceChart, surface3DChart, bubbleChart
data Chart = LineChart
  { _lnchGrouping :: ChartGrouping
  , _lnchSeries :: [LineSeries]
  , _lnchMarker :: Maybe Bool
    -- ^ specifies that the marker shall be shown
  , _lnchSmooth :: Maybe Bool
    -- ^ specifies the line connecting the points on the chart shall be
    -- smoothed using Catmull-Rom splines
  } deriving (Eq, Show)

-- | Possible groupings for a bar chart
--
-- See 21.2.3.17 "ST_Grouping (Grouping)" (p. 3446)
data ChartGrouping
  = PercentStackedGrouping
    -- ^ (100% Stacked) Specifies that the chart series are drawn next to each
    -- other along the value axis and scaled to total 100%.
  | StackedGrouping
    -- ^ (Stacked) Specifies that the chart series are drawn next to each
    -- other on the value axis.
  | StandardGrouping
    -- ^(Standard) Specifies that the chart series are drawn on the value
    -- axis.
  deriving (Eq, Show)

-- | Specifies common series options
-- TODO: spPr
--
-- See @EG_SerShared@ (p. 4063)
data Series = Series
  { _serTx :: Maybe Formula
    -- ^ specifies text for a series name, without rich text formatting
    -- currently only reference formula is supported
  } deriving (Eq, Show)

-- | A series on a line chart
--
-- TODO: dPt, trendline, errBars, cat, extLst
--
-- See @CT_LineSer@ (p. 4064)
data LineSeries = LineSeries
  { _lnserShared :: Series
  , _lnserMarker :: Maybe DataMarker
  , _lnserDataLblProps :: Maybe DataLblProps
  , _lnserVal :: Maybe Formula
    -- ^ currently only reference formula is supported
  , _lnserSmooth :: Maybe Bool
  } deriving (Eq, Show)

-- See @CT_Marker@ (p. 4061)
data DataMarker = DataMarker
  { _dmrkSymbol :: Maybe DataMarkerSymbol
  , _dmrkSize :: Maybe Int
    -- ^ integer between 2 and 72, specifying a size in points
  } deriving (Eq, Show)

data DataMarkerSymbol
  = DataMarkerCircle
  | DataMarkerDash
  | DataMarkerDiamond
  | DataMarkerDot
  | DataMarkerNone
  | DataMarkerPicture
  | DataMarkerPlus
  | DataMarkerSquare
  | DataMarkerStar
  | DataMarkerTriangle
  | DataMarkerX
  | DataMarkerAuto
  deriving (Eq, Show)

-- | Settings for the data labels for an entire series or the
-- entire chart
--
-- TODO: numFmt, spPr, txPr, dLblPos, showBubbleSize,
-- separator, showLeaderLines, leaderLines
-- See 21.2.2.49 "dLbls (Data Labels)" (p. 3384)
data DataLblProps = DataLblProps
  { _dlblShowLegendKey :: Maybe Bool
  , _dlblShowVal :: Maybe Bool
  , _dlblShowCatName :: Maybe Bool
  , _dlblShowSerName :: Maybe Bool
  , _dlblShowPercent :: Maybe Bool
  } deriving (Eq, Show)

-- | Specifies the possible positions for tick marks.

-- See 21.2.3.48 "ST_TickMark (Tick Mark)" (p. 3467)
data TickMark
  = TickMarkCross
    -- ^ (Cross) Specifies the tick marks shall cross the axis.
  | TickMarkIn
    -- ^ (Inside) Specifies the tick marks shall be inside the plot area.
  | TickMarkNone
    -- ^ (None) Specifies there shall be no tick marks.
  | TickMarkOut
    -- ^ (Outside) Specifies the tick marks shall be outside the plot area.
  deriving (Eq, Show)

{-------------------------------------------------------------------------------
  Parsing
-------------------------------------------------------------------------------}

instance FromCursor ChartSpace where
  fromCursor cur = do
    cur' <- cur $/ element (c_ "chart")
    _chspTitle <- maybeFromElement (c_ "title") cur'
    let _chspCharts =
          cur' $/ element (c_ "plotArea") &/ anyElement >=> chartFromNode . node
    _chspLegend <- maybeFromElement (c_ "legend") cur'
    _chspPlotVisOnly <- maybeBoolElementValue (c_ "plotVisOnly") cur'
    _chspDispBlanksAs <- maybeElementValue (c_ "dispBlanksAs") cur'
    return ChartSpace {..}

chartFromNode :: Node -> [Chart]
chartFromNode n
  | n `nodeElNameIs` (c_ "lineChart") = do
    _lnchGrouping <- fromElementValue (c_ "grouping") cur
    let _lnchSeries = cur $/ element (c_ "ser") >=> fromCursor
    _lnchMarker <- maybeBoolElementValue (c_ "marker") cur
    _lnchSmooth <- maybeBoolElementValue (c_ "smooth") cur
    return LineChart {..}
  | otherwise = fail "no matching chart node"
  where
    cur = fromNode n

instance FromCursor LineSeries where
  fromCursor cur = do
    _lnserShared <- fromCursor cur
    _lnserMarker <- maybeFromElement (c_ "marker") cur
    _lnserDataLblProps <- maybeFromElement (c_ "dLbls") cur
    _lnserVal <-
      cur $/ element (c_ "val") &/ element (c_ "numRef") >=>
      maybeFromElement (c_ "f")
    _lnserSmooth <- maybeElementValueDef (c_ "smooth") True cur
    return LineSeries {..}

-- should we respect idx and order?
instance FromCursor Series where
  fromCursor cur = do
    _serTx <-
      cur $/ element (c_ "tx") &/ element (c_ "strRef") >=>
      maybeFromElement (c_ "f")
    return Series {..}

instance FromCursor DataMarker where
  fromCursor cur = do
    _dmrkSymbol <- maybeElementValue (c_ "symbol") cur
    _dmrkSize <- maybeElementValue (c_ "size") cur
    return DataMarker {..}

instance FromAttrVal DataMarkerSymbol where
  fromAttrVal "circle" = readSuccess DataMarkerCircle
  fromAttrVal "dash" = readSuccess DataMarkerDash
  fromAttrVal "diamond" = readSuccess DataMarkerDiamond
  fromAttrVal "dot" = readSuccess DataMarkerDot
  fromAttrVal "none" = readSuccess DataMarkerNone
  fromAttrVal "picture" = readSuccess DataMarkerPicture
  fromAttrVal "plus" = readSuccess DataMarkerPlus
  fromAttrVal "square" = readSuccess DataMarkerSquare
  fromAttrVal "star" = readSuccess DataMarkerStar
  fromAttrVal "triangle" = readSuccess DataMarkerTriangle
  fromAttrVal "x" = readSuccess DataMarkerX
  fromAttrVal "auto" = readSuccess DataMarkerAuto
  fromAttrVal t = invalidText "DataMarkerSymbol" t

instance FromCursor DataLblProps where
  fromCursor cur = do
    _dlblShowLegendKey <- maybeBoolElementValue (c_ "showLegendKey") cur
    _dlblShowVal <- maybeBoolElementValue (c_ "showVal") cur
    _dlblShowCatName <- maybeBoolElementValue (c_ "showCatName") cur
    _dlblShowSerName <- maybeBoolElementValue (c_ "showSerName") cur
    _dlblShowPercent <- maybeBoolElementValue (c_ "showPercent") cur
    return DataLblProps {..}

instance FromAttrVal ChartGrouping where
  fromAttrVal "percentStacked" = readSuccess PercentStackedGrouping
  fromAttrVal "standard" = readSuccess StandardGrouping
  fromAttrVal "stacked" = readSuccess StackedGrouping
  fromAttrVal t = invalidText "ChartGrouping" t

instance FromCursor ChartTitle where
  fromCursor cur =
    cur $/ element (c_ "tx") &/ element (c_ "rich") >=> fmap ChartTitle . fromCursor

instance FromCursor Legend where
  fromCursor cur = do
    _legendPos <- maybeElementValue (c_ "legendPos") cur
    _legendOverlay <- maybeElementValueDef (c_ "overlay") True cur
    return Legend {..}

instance FromAttrVal LegendPos where
  fromAttrVal "b" = readSuccess LegendBottom
  fromAttrVal "l" = readSuccess LegendLeft
  fromAttrVal "r" = readSuccess LegendRight
  fromAttrVal "t" = readSuccess LegendTop
  fromAttrVal "tr" = readSuccess LegendTopRight
  fromAttrVal t = invalidText "LegendPos" t

instance FromAttrVal DispBlanksAs where
  fromAttrVal "gap" = readSuccess DispBlanksAsGap
  fromAttrVal "span" = readSuccess DispBlanksAsSpan
  fromAttrVal "zero" = readSuccess DispBlanksAsZero
  fromAttrVal t = invalidText "DispBlanksAs" t

{-------------------------------------------------------------------------------
  Default instances
-------------------------------------------------------------------------------}

instance Default Legend where
  def = Legend {_legendPos = Just LegendBottom, _legendOverlay = Just False}

{-------------------------------------------------------------------------------
  Rendering
-------------------------------------------------------------------------------}

instance ToDocument ChartSpace where
  toDocument =
    documentFromNsPrefElement "Charts generated by xlsx" chartNs (Just "c") .
    toElement "chartSpace"

instance ToElement ChartSpace where
  toElement nm ChartSpace {..} =
    elementListSimple nm [nonRounded, chartEl, chSpPr]
    where
      -- no such element gives a chart space with rounded corners
      nonRounded = elementValue "roundedCorners" False
      chSpPr = toElement "spPr" $ def {_spFill = Just SolidFill}
      chartEl = elementListSimple "chart" elements
      elements =
        catMaybes
          [ toElement "title" <$> _chspTitle
          -- LO?
          , Just $ elementValue "autoTitleDeleted" False
          , Just $ elementListSimple "plotArea" areaEls
          , toElement "legend" <$> _chspLegend
          , elementValue "plotVisOnly" <$> _chspPlotVisOnly
          , elementValue "dispBlanksAs" <$> _chspDispBlanksAs
          ]
      -- we reserve 2 axes - X and Y for line charts
      -- this needs to be reworked when other chart types will be added
      enumCharts = zip [1,3 ..] _chspCharts
      charts = [chartToElement ch i (i + 1) | (i, ch) <- enumCharts]
      areaEls = charts ++ valAxes ++ catAxes
      catAxes = [catAxEl i (i + 1) | (i, _) <- enumCharts]
      valAxes = [valAxEl (i + 1) i | (i, _) <- enumCharts]
      catAxEl :: Int -> Int -> Element
      catAxEl i cr =
        elementListSimple "catAx" $
        [ elementValue "axId" i
        , emptyElement "scaling"
        , elementValue "delete" False
        , elementValue "axPos" ("b" :: Text)
        , elementValue "majorTickMark" TickMarkNone
        , elementValue "minorTickMark" TickMarkNone
        , toElement "spPr" noFill
        , elementValue "crossAx" cr
        , elementValue "auto" True
        ]
      valAxEl :: Int -> Int -> Element
      valAxEl i cr =
        elementListSimple "valAx" $
        [ elementValue "axId" i
        , emptyElement "scaling"
        , elementValue "delete" False
        , elementValue "axPos" ("l" :: Text)
        , gridLinesEl
        , elementValue "majorTickMark" TickMarkNone
        , elementValue "minorTickMark" TickMarkNone
        , toElement "spPr" noFill
        , elementValue "crossAx" cr
        ]
      noFill =
        def
        { _spFill = Just NoFill
        , _spOutline = Just . LineProperties $ Just NoFill
        }
      gridLinesEl =
        elementListSimple "majorGridlines" [toElement "spPr" lineFill]
      lineFill = def { _spOutline = Just . LineProperties $ Just SolidFill }

chartToElement :: Chart -> Int -> Int -> Element
chartToElement LineChart {..} cId vId = elementListSimple "lineChart" elements
  where
    elements =
      (grouping : varyColors : series) ++
      catMaybes
        [ elementValue "marker" <$> _lnchMarker
        , elementValue "smooth" <$> _lnchSmooth
        ] ++
      map (elementValue "axId") [cId, vId]
    grouping = elementValue "grouping" _lnchGrouping
    -- no element seems to be equal to varyColors=true in Excel Online
    varyColors = elementValue "varyColors" False
    series = [indexedSeriesEl i s | (i, s) <- zip [0 ..] _lnchSeries]
    indexedSeriesEl :: Int -> LineSeries -> Element
    indexedSeriesEl i s = prependI i $ toElement "ser" s
    prependI i e@Element {..} = e {elementNodes = iNodes i ++ elementNodes}
    iNodes i = map NodeElement [elementValue n i | n <- ["idx", "order"]]

instance ToAttrVal ChartGrouping where
  toAttrVal PercentStackedGrouping = "percentStacked"
  toAttrVal StandardGrouping = "standard"
  toAttrVal StackedGrouping = "stacked"

instance ToElement LineSeries where
  toElement nm LineSeries {..} = serEl
      { elementNodes = elementNodes serEl ++
                     map NodeElement elements }
    where
      serEl = toElement nm _lnserShared
      elements =
        catMaybes
          [ toElement "marker" <$> _lnserMarker
          , toElement "dLbls" <$> _lnserDataLblProps
          , Just $ valEl _lnserVal
          , elementValue "smooth" <$> _lnserSmooth
          ]
      valEl v =
        elementListSimple
          "val"
          [elementListSimple "numRef" $ maybeToList (toElement "f" <$> v)]

instance ToElement DataMarker where
  toElement nm DataMarker {..} = elementListSimple nm elements
    where
      elements =
        catMaybes
          [ elementValue "symbol" <$> _dmrkSymbol
          , elementValue "size" <$> _dmrkSize
          ]

instance ToAttrVal DataMarkerSymbol where
  toAttrVal DataMarkerCircle = "circle"
  toAttrVal DataMarkerDash = "dash"
  toAttrVal DataMarkerDiamond = "diamond"
  toAttrVal DataMarkerDot = "dot"
  toAttrVal DataMarkerNone = "none"
  toAttrVal DataMarkerPicture = "picture"
  toAttrVal DataMarkerPlus = "plus"
  toAttrVal DataMarkerSquare = "square"
  toAttrVal DataMarkerStar = "star"
  toAttrVal DataMarkerTriangle = "triangle"
  toAttrVal DataMarkerX = "x"
  toAttrVal DataMarkerAuto = "auto"

instance ToElement DataLblProps where
  toElement nm DataLblProps {..} = elementListSimple nm elements
    where
      elements =
        catMaybes
          [ elementValue "showLegendKey" <$> _dlblShowLegendKey
          , elementValue "showVal" <$> _dlblShowVal
          , elementValue "showCatName" <$> _dlblShowCatName
          , elementValue "showSerName" <$> _dlblShowSerName
          , elementValue "showPercent" <$> _dlblShowPercent
          ]

-- should we respect idx and order?
instance ToElement Series where
  toElement nm Series {..} =
    elementListSimple
      nm
      [ elementListSimple
          "tx"
          [elementListSimple "strRef" $ maybeToList (toElement "f" <$> _serTx)]
      ]

instance ToElement ChartTitle where
  toElement nm (ChartTitle body) =
    elementListSimple nm [txEl, elementValue "overlay" False]
    where
      txEl = elementListSimple "tx" [toElement (c_ "rich") body]

instance ToElement Legend where
  toElement nm Legend{..} = elementListSimple nm elements
    where
       elements = catMaybes [ elementValue "legendPos" <$> _legendPos
                            , elementValue "overlay" <$>_legendOverlay]

instance ToAttrVal LegendPos where
  toAttrVal LegendBottom   = "b" 
  toAttrVal LegendLeft     = "l" 
  toAttrVal LegendRight    = "r" 
  toAttrVal LegendTop      = "t" 
  toAttrVal LegendTopRight = "tr"

instance ToAttrVal DispBlanksAs where
  toAttrVal DispBlanksAsGap  = "gap"
  toAttrVal DispBlanksAsSpan = "span"
  toAttrVal DispBlanksAsZero = "zero"

instance ToAttrVal TickMark where
  toAttrVal TickMarkCross = "cross"
  toAttrVal TickMarkIn = "in"
  toAttrVal TickMarkNone = "none"
  toAttrVal TickMarkOut = "out"

-- | Add chart namespace to name
c_ :: Text -> Name
c_ x =
  Name {nameLocalName = x, nameNamespace = Just chartNs, namePrefix = Just "c"}

chartNs :: Text
chartNs = "http://schemas.openxmlformats.org/drawingml/2006/chart"
