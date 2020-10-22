{-# LANGUAGE CPP #-}
{-# LANGUAGE OverloadedStrings #-}
{-# LANGUAGE RecordWildCards #-}
{-# LANGUAGE TemplateHaskell #-}
{-# LANGUAGE DeriveGeneric #-}
module Codec.Xlsx.Types.Drawing.Chart where

import GHC.Generics (Generic)

#ifdef USE_MICROLENS
import Lens.Micro.TH (makeLenses)
#else
import Control.Lens.TH
#endif
import Control.DeepSeq (NFData)
import Data.Default
import Data.Maybe (catMaybes, listToMaybe, maybeToList)
import Data.Text (Text)
import Text.XML
import Text.XML.Cursor

import Codec.Xlsx.Parser.Internal
import Codec.Xlsx.Types.Common
import Codec.Xlsx.Types.Drawing.Common
import Codec.Xlsx.Writer.Internal

-- | Main Chart holder, combines
-- TODO: title, autoTitleDeleted, pivotFmts
--  view3D, floor, sideWall, backWall, showDLblsOverMax, extLst
data ChartSpace = ChartSpace
  { _chspTitle :: Maybe ChartTitle
  , _chspCharts :: [Chart]
  , _chspLegend :: Maybe Legend
  , _chspPlotVisOnly :: Maybe Bool
  , _chspDispBlanksAs :: Maybe DispBlanksAs
  } deriving (Eq, Show, Generic)
instance NFData ChartSpace

-- | Chart title
--
-- TODO: layout, overlay, spPr, txPr, extLst
newtype ChartTitle =
  ChartTitle (Maybe TextBody)
  deriving (Eq, Show, Generic)
instance NFData ChartTitle

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
  deriving (Eq, Show, Generic)
instance NFData DispBlanksAs

-- TODO: legendEntry, layout, overlay, spPr, txPr, extLst
data Legend = Legend
    { _legendPos     :: Maybe LegendPos
    , _legendOverlay :: Maybe Bool
    } deriving (Eq, Show, Generic)
instance NFData Legend

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
  deriving (Eq, Show, Generic)
instance NFData LegendPos

-- | Specific Chart
-- TODO:
--   area3DChart, line3DChart, stockChart, radarChart,
--   pie3DChart, doughnutChart, bar3DChart, ofPieChart,
--   surfaceChart, surface3DChart, bubbleChart
data Chart
  = LineChart { _lnchGrouping :: ChartGrouping
              , _lnchSeries :: [LineSeries]
              , _lnchMarker :: Maybe Bool
                -- ^ specifies that the marker shall be shown
              , _lnchSmooth :: Maybe Bool
                -- ^ specifies the line connecting the points on the chart shall be
                -- smoothed using Catmull-Rom splines
              }
  | AreaChart { _archGrouping :: Maybe ChartGrouping
              , _archSeries :: [AreaSeries]
              }
  | BarChart { _brchDirection :: BarDirection
             , _brchGrouping :: Maybe BarChartGrouping
             , _brchSeries :: [BarSeries]
             }
  | PieChart { _pichSeries :: [PieSeries]
             }
  | ScatterChart { _scchStyle :: ScatterStyle
                 , _scchSeries :: [ScatterSeries]
                 }
  deriving (Eq, Show, Generic)
instance NFData Chart

-- | Possible groupings for a chart
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
  deriving (Eq, Show, Generic)
instance NFData ChartGrouping

-- | Possible groupings for a bar chart
--
-- See 21.2.3.4 "ST_BarGrouping (Bar Grouping)" (p. 3441)
data BarChartGrouping
  = BarClusteredGrouping
    -- ^ Specifies that the chart series are drawn next to each other
    -- along the category axis.
  | BarPercentStackedGrouping
    -- ^ (100% Stacked) Specifies that the chart series are drawn next to each
    -- other along the value axis and scaled to total 100%.
  | BarStackedGrouping
    -- ^ (Stacked) Specifies that the chart series are drawn next to each
    -- other on the value axis.
  | BarStandardGrouping
    -- ^(Standard) Specifies that the chart series are drawn on the value
    -- axis.
  deriving (Eq, Show, Generic)
instance NFData BarChartGrouping

-- | Possible directions for a bar chart
--
-- See 21.2.3.3 "ST_BarDir (Bar Direction)" (p. 3441)
data BarDirection
  = DirectionBar
  | DirectionColumn
  deriving (Eq, Show, Generic)
instance NFData BarDirection

-- | Possible styles of scatter chart
--
-- /Note:/ It appears that even for 'ScatterMarker' style Exel draws a
-- line between chart points if otline fill for '_scserShared' isn't
-- set to so it's not quite clear how could Excel use this property
--
-- See 21.2.3.40 "ST_ScatterStyle (Scatter Style)" (p. 3455)
data ScatterStyle
  = ScatterNone
  | ScatterLine
  | ScatterLineMarker
  | ScatterMarker
  | ScatterSmooth
  | ScatterSmoothMarker
  deriving (Eq, Show, Generic)
instance NFData ScatterStyle

-- | Single data point options
--
-- TODO:  invertIfNegative,  bubble3D, explosion, pictureOptions, extLst
--
-- See 21.2.2.52 "dPt (Data Point)" (p. 3384)
data DataPoint = DataPoint
  { _dpMarker :: Maybe DataMarker
  , _dpShapeProperties :: Maybe ShapeProperties
  } deriving (Eq, Show, Generic)
instance NFData DataPoint

-- | Specifies common series options
-- TODO: spPr
--
-- See @EG_SerShared@ (p. 4063)
data Series = Series
  { _serTx :: Maybe Formula
    -- ^ specifies text for a series name, without rich text formatting
    -- currently only reference formula is supported
  , _serShapeProperties :: Maybe ShapeProperties
  } deriving (Eq, Show, Generic)
instance NFData Series

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
  } deriving (Eq, Show, Generic)
instance NFData LineSeries

-- | A series on an area chart
--
-- TODO: pictureOptions, dPt, trendline, errBars, cat, extLst
--
-- See @CT_AreaSer@ (p. 4065)
data AreaSeries = AreaSeries
  { _arserShared :: Series
  , _arserDataLblProps :: Maybe DataLblProps
  , _arserVal :: Maybe Formula
  } deriving (Eq, Show, Generic)
instance NFData AreaSeries

-- | A series on a bar chart
--
-- TODO: invertIfNegative, pictureOptions, dPt, trendline, errBars,
-- cat, shape, extLst
--
-- See @CT_BarSer@ (p. 4064)
data BarSeries = BarSeries
  { _brserShared :: Series
  , _brserDataLblProps :: Maybe DataLblProps
  , _brserVal :: Maybe Formula
  } deriving (Eq, Show, Generic)
instance NFData BarSeries

-- | A series on a pie chart
--
-- TODO: explosion, cat, extLst
--
-- See @CT_PieSer@ (p. 4065)
data PieSeries = PieSeries
  { _piserShared :: Series
  , _piserDataPoints :: [DataPoint]
  -- ^ normally you should set fill for chart datapoints to make them
  -- properly colored
  , _piserDataLblProps :: Maybe DataLblProps
  , _piserVal :: Maybe Formula
  } deriving (Eq, Show, Generic)
instance NFData PieSeries

-- | A series on a scatter chart
--
-- TODO: dPt, trendline, errBars, smooth, extLst
--
-- See @CT_ScatterSer@ (p. 4064)
data ScatterSeries = ScatterSeries
  { _scserShared :: Series
  , _scserMarker :: Maybe DataMarker
  , _scserDataLblProps :: Maybe DataLblProps
  , _scserXVal :: Maybe Formula
  , _scserYVal :: Maybe Formula
  , _scserSmooth :: Maybe Bool
  } deriving (Eq, Show, Generic)
instance NFData ScatterSeries

-- See @CT_Marker@ (p. 4061)
data DataMarker = DataMarker
  { _dmrkSymbol :: Maybe DataMarkerSymbol
  , _dmrkSize :: Maybe Int
    -- ^ integer between 2 and 72, specifying a size in points
  } deriving (Eq, Show, Generic)
instance NFData DataMarker

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
  deriving (Eq, Show, Generic)
instance NFData DataMarkerSymbol

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
  } deriving (Eq, Show, Generic)
instance NFData DataLblProps

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
  deriving (Eq, Show, Generic)
instance NFData TickMark

makeLenses ''DataPoint

{-------------------------------------------------------------------------------
  Default instances
-------------------------------------------------------------------------------}

instance Default DataPoint where
    def = DataPoint Nothing Nothing

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
  | n `nodeElNameIs` (c_ "areaChart") = do
    _archGrouping <- maybeElementValue (c_ "grouping") cur
    let _archSeries = cur $/ element (c_ "ser") >=> fromCursor
    return AreaChart {..}
  | n `nodeElNameIs` (c_ "barChart") = do
    _brchDirection <- fromElementValue (c_ "barDir") cur
    _brchGrouping <-
      maybeElementValueDef (c_ "grouping") BarClusteredGrouping cur
    let _brchSeries = cur $/ element (c_ "ser") >=> fromCursor
    return BarChart {..}
  | n `nodeElNameIs` (c_ "pieChart") = do
    let _pichSeries = cur $/ element (c_ "ser") >=> fromCursor
    return PieChart {..}
  | n `nodeElNameIs` (c_ "scatterChart") = do
    _scchStyle <- fromElementValue (c_ "scatterStyle") cur
    let _scchSeries = cur $/ element (c_ "ser") >=> fromCursor
    return ScatterChart {..}
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

instance FromCursor AreaSeries where
  fromCursor cur = do
    _arserShared <- fromCursor cur
    _arserDataLblProps <- maybeFromElement (c_ "dLbls") cur
    _arserVal <-
      cur $/ element (c_ "val") &/ element (c_ "numRef") >=>
      maybeFromElement (c_ "f")
    return AreaSeries {..}

instance FromCursor BarSeries where
  fromCursor cur = do
    _brserShared <- fromCursor cur
    _brserDataLblProps <- maybeFromElement (c_ "dLbls") cur
    _brserVal <-
      cur $/ element (c_ "val") &/ element (c_ "numRef") >=>
      maybeFromElement (c_ "f")
    return BarSeries {..}

instance FromCursor PieSeries where
  fromCursor cur = do
    _piserShared <- fromCursor cur
    let _piserDataPoints = cur $/ element (c_ "dPt") >=> fromCursor
    _piserDataLblProps <- maybeFromElement (c_ "dLbls") cur
    _piserVal <-
      cur $/ element (c_ "val") &/ element (c_ "numRef") >=>
      maybeFromElement (c_ "f")
    return PieSeries {..}

instance FromCursor ScatterSeries where
  fromCursor cur = do
    _scserShared <- fromCursor cur
    _scserMarker <- maybeFromElement (c_ "marker") cur
    _scserDataLblProps <- maybeFromElement (c_ "dLbls") cur
    _scserXVal <-
      cur $/ element (c_ "xVal") &/ element (c_ "numRef") >=>
      maybeFromElement (c_ "f")
    _scserYVal <-
      cur $/ element (c_ "yVal") &/ element (c_ "numRef") >=>
      maybeFromElement (c_ "f")
    _scserSmooth <- maybeElementValueDef (c_ "smooth") True cur
    return ScatterSeries {..}

-- should we respect idx and order?
instance FromCursor Series where
  fromCursor cur = do
    _serTx <-
      cur $/ element (c_ "tx") &/ element (c_ "strRef") >=>
      maybeFromElement (c_ "f")
    _serShapeProperties <- maybeFromElement (c_ "spPr") cur
    return Series {..}

instance FromCursor DataMarker where
  fromCursor cur = do
    _dmrkSymbol <- maybeElementValue (c_ "symbol") cur
    _dmrkSize <- maybeElementValue (c_ "size") cur
    return DataMarker {..}

instance FromCursor DataPoint where
  fromCursor cur = do
    _dpMarker <- maybeFromElement (c_ "marker") cur
    _dpShapeProperties <- maybeFromElement (c_ "spPr") cur
    return DataPoint {..}

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

instance FromAttrVal BarDirection where
  fromAttrVal "bar" = readSuccess DirectionBar
  fromAttrVal "col" = readSuccess DirectionColumn
  fromAttrVal t = invalidText "BarDirection" t

instance FromAttrVal ScatterStyle where
  fromAttrVal "none" = readSuccess ScatterNone
  fromAttrVal "line" = readSuccess ScatterLine
  fromAttrVal "lineMarker" = readSuccess ScatterLineMarker
  fromAttrVal "marker" = readSuccess ScatterMarker
  fromAttrVal "smooth" = readSuccess ScatterSmooth
  fromAttrVal "smoothMarker" = readSuccess ScatterSmoothMarker
  fromAttrVal t = invalidText "ScatterStyle" t

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

instance FromAttrVal BarChartGrouping where
  fromAttrVal "clustered" = readSuccess BarClusteredGrouping
  fromAttrVal "percentStacked" = readSuccess BarPercentStackedGrouping
  fromAttrVal "standard" = readSuccess BarStandardGrouping
  fromAttrVal "stacked" = readSuccess BarStackedGrouping
  fromAttrVal t = invalidText "BarChartGrouping" t

instance FromCursor ChartTitle where
  fromCursor cur = do
    let mTitle = listToMaybe $
          cur $/ element (c_ "tx") &/ element (c_ "rich") >=> fromCursor
    return $ ChartTitle mTitle

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
      chSpPr = toElement "spPr" $ def {_spFill = Just $ solidRgb "ffffff"}
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
      areaEls = charts ++ axes
      (_, charts, axes) = foldr addChart (1, [], []) _chspCharts
      addChart ch (i, cs, as) =
        let (c, as') = chartToElements ch i
        in (i + length as', c : cs, as' ++ as)

chartToElements :: Chart -> Int -> (Element, [Element])
chartToElements chart axId =
  case chart of
    LineChart {..} ->
      chartElement
        "lineChart"
        stdAxes
        (Just _lnchGrouping)
        _lnchSeries
        []
        (catMaybes
           [ elementValue "marker" <$> _lnchMarker
           , elementValue "smooth" <$> _lnchSmooth
           ])
    AreaChart {..} ->
      chartElement "areaChart" stdAxes _archGrouping _archSeries [] []
    BarChart {..} ->
      chartElement
        "barChart"
        stdAxes
        _brchGrouping
        _brchSeries
        [elementValue "barDir" _brchDirection]
        []
    PieChart {..} -> chartElement "pieChart" [] noGrouping _pichSeries [] []
    ScatterChart {..} ->
      chartElement
        "scatterChart"
        xyAxes
        noGrouping
        _scchSeries
        [elementValue "scatterStyle" _scchStyle]
        []
  where
    noGrouping :: Maybe ChartGrouping
    noGrouping = Nothing
    chartElement
      :: (ToElement s, ToAttrVal gr)
      => Name
      -> [Element]
      -> Maybe gr
      -> [s]
      -> [Element]
      -> [Element]
      -> (Element, [Element])
    chartElement nm axes mGrouping series prepended appended =
      ( elementListSimple nm $
        prepended ++
        (maybeToList $ elementValue "grouping" <$> mGrouping) ++
        (varyColors : seriesEls series) ++
        appended ++ zipWith (\n _ -> elementValue "axId" n) [axId ..] axes
      , axes)
    -- no element seems to be equal to varyColors=true in Excel Online
    varyColors = elementValue "varyColors" False
    seriesEls series = [indexedSeriesEl i s | (i, s) <- zip [0 ..] series]
    indexedSeriesEl
      :: ToElement a
      => Int -> a -> Element
    indexedSeriesEl i s = prependI i $ toElement "ser" s
    prependI i e@Element {..} = e {elementNodes = iNodes i ++ elementNodes}
    iNodes i = map NodeElement [elementValue n i | n <- ["idx", "order"]]

    stdAxes = [catAx axId (axId + 1), valAx "l" (axId + 1) axId]
    xyAxes = [valAx "b" axId (axId + 1), valAx "l" (axId + 1) axId]
    catAx :: Int -> Int -> Element
    catAx i cr =
      elementListSimple "catAx" $
      [ elementValue "axId" i
      , emptyElement "scaling"
      , elementValue "delete" False
      , elementValue "axPos" ("b" :: Text)
      , elementValue "majorTickMark" TickMarkNone
      , elementValue "minorTickMark" TickMarkNone
      , toElement "spPr" grayLines
      , elementValue "crossAx" cr
      , elementValue "auto" True
      ]
    valAx :: Text -> Int -> Int -> Element
    valAx pos i cr =
      elementListSimple "valAx" $
      [ elementValue "axId" i
      , emptyElement "scaling"
      , elementValue "delete" False
      , elementValue "axPos" pos
      , gridLinesEl
      , elementValue "majorTickMark" TickMarkNone
      , elementValue "minorTickMark" TickMarkNone
      , toElement "spPr" grayLines
      , elementValue "crossAx" cr
      ]
    grayLines = def {_spOutline = Just def {_lnFill = Just $ solidRgb "b3b3b3"}}
    gridLinesEl =
      elementListSimple "majorGridlines" [toElement "spPr" grayLines]

instance ToAttrVal ChartGrouping where
  toAttrVal PercentStackedGrouping = "percentStacked"
  toAttrVal StandardGrouping = "standard"
  toAttrVal StackedGrouping = "stacked"

instance ToAttrVal BarChartGrouping where
  toAttrVal BarClusteredGrouping = "clustered"
  toAttrVal BarPercentStackedGrouping = "percentStacked"
  toAttrVal BarStandardGrouping = "standard"
  toAttrVal BarStackedGrouping = "stacked"

instance ToAttrVal BarDirection where
  toAttrVal DirectionBar = "bar"
  toAttrVal DirectionColumn = "col"

instance ToAttrVal ScatterStyle where
  toAttrVal ScatterNone = "none"
  toAttrVal ScatterLine = "line"
  toAttrVal ScatterLineMarker = "lineMarker"
  toAttrVal ScatterMarker = "marker"
  toAttrVal ScatterSmooth = "smooth"
  toAttrVal ScatterSmoothMarker = "smoothMarker"

instance ToElement LineSeries where
  toElement nm LineSeries {..} = simpleSeries nm _lnserShared _lnserVal pr ap
    where
      pr =
        catMaybes
          [ toElement "marker" <$> _lnserMarker
          , toElement "dLbls" <$> _lnserDataLblProps
          ]
      ap = maybeToList $ elementValue "smooth" <$> _lnserSmooth

simpleSeries :: Name
             -> Series
             -> Maybe Formula
             -> [Element]
             -> [Element]
             -> Element
simpleSeries nm shared val prepended appended =
  serEl {elementNodes = elementNodes serEl ++ map NodeElement elements}
  where
    serEl = toElement nm shared
    elements = prepended ++ (valEl val : appended)
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

instance ToElement AreaSeries where
  toElement nm AreaSeries {..} = simpleSeries nm _arserShared _arserVal pr []
    where
      pr = maybeToList $ fmap (toElement "dLbls") _arserDataLblProps

instance ToElement BarSeries where
  toElement nm BarSeries {..} = simpleSeries nm _brserShared _brserVal pr []
    where
      pr = maybeToList $ fmap (toElement "dLbls") _brserDataLblProps

instance ToElement PieSeries where
  toElement nm PieSeries {..} = simpleSeries nm _piserShared _piserVal pr []
    where
      pr = dPts ++ maybeToList (fmap (toElement "dLbls") _piserDataLblProps)
      dPts = zipWith dPtEl [(0 :: Int) ..] _piserDataPoints
      dPtEl i DataPoint {..} =
        elementListSimple
          "dPt"
          (elementValue "idx" i :
           catMaybes
             [ toElement "marker" <$> _dpMarker
             , toElement "spPr" <$> _dpShapeProperties
             ])

instance ToElement ScatterSeries where
  toElement nm ScatterSeries {..} =
    serEl {elementNodes = elementNodes serEl ++ map NodeElement elements}
    where
      serEl = toElement nm _scserShared
      elements =
        catMaybes
          [ toElement "marker" <$> _scserMarker
          , toElement "dLbls" <$> _scserDataLblProps
          ] ++
        [valEl "xVal" _scserXVal, valEl "yVal" _scserYVal] ++
        (maybeToList $ fmap (elementValue "smooth") _scserSmooth)
      valEl vnm v =
        elementListSimple
          vnm
          [elementListSimple "numRef" $ maybeToList (toElement "f" <$> v)]

-- should we respect idx and order?
instance ToElement Series where
  toElement nm Series {..} =
    elementListSimple nm $
    [ elementListSimple
        "tx"
        [elementListSimple "strRef" $ maybeToList (toElement "f" <$> _serTx)]
    ] ++
    maybeToList (toElement "spPr" <$> _serShapeProperties)

instance ToElement ChartTitle where
  toElement nm (ChartTitle body) =
    elementListSimple nm [txEl, elementValue "overlay" False]
    where
      txEl = elementListSimple "tx" $ catMaybes [toElement (c_ "rich") <$> body]

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
