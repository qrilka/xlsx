{-# LANGUAGE CPP                  #-}
{-# LANGUAGE FlexibleInstances    #-}
{-# LANGUAGE OverloadedStrings    #-}
{-# LANGUAGE RecordWildCards      #-}
{-# LANGUAGE TemplateHaskell      #-}
{-# LANGUAGE TypeSynonymInstances #-}
module Codec.Xlsx.Types.Drawing where

import           Control.Arrow                           (first)
import           Control.Lens.TH
import           Data.ByteString.Lazy                    (ByteString)
import           Data.Default
import qualified Data.Map                                as M
import           Data.Maybe                              (catMaybes,
                                                          listToMaybe)
import           Data.Monoid                             ((<>))
import           Data.Text                               (Text)
import qualified Data.Text                               as T
import qualified Data.Text.Read                          as T
import           Text.XML
import           Text.XML.Cursor

#if !MIN_VERSION_base(4,8,0)
import           Control.Applicative
#endif

import           Codec.Xlsx.Parser.Internal              hiding (n)
import           Codec.Xlsx.Types.Internal
import           Codec.Xlsx.Types.Internal.Relationships
import           Codec.Xlsx.Writer.Internal

-- | information about image file as a par of a drawing
data FileInfo = FileInfo
    { _fiFilename    :: FilePath
    -- ^ image filename, images are assumed to be stored under path "xl/media/"
    , _fiContentType :: Text
    -- ^ image content type, ECMA-376 advises to use "image/png" or "image/jpeg"
    -- if interoperability is wanted
    , _fiContents    :: ByteString
    -- ^ image file contents
    } deriving (Eq, Show)

-- | This simple type represents a one dimensional position or length
--
-- See 20.1.10.16 "ST_Coordinate (Coordinate)" (p. 2921)
data Coordinate
    = UnqCoordinate Int
    -- ^ see 20.1.10.19 "ST_CoordinateUnqualified (Coordinate)" (p. 2922)
    | UniversalMeasure UnitIdentifier Double
    -- ^ see 22.9.2.15 "ST_UniversalMeasure (Universal Measurement)" (p. 3793)
    deriving (Eq, Show)

-- | Units used in "Universal measure" coordinates
-- see 22.9.2.15 "ST_UniversalMeasure (Universal Measurement)" (p. 3793)
data UnitIdentifier
    = UnitCm -- "cm" As defined in ISO 31.
    | UnitMm -- "mm" As defined in ISO 31.
    | UnitIn -- "in" 1 in = 2.54 cm (informative)
    | UnitPt -- "pt" 1 pt = 1/72 in (informative)
    | UnitPc -- "pc" 1 pc = 12 pt (informative)
    | UnitPi -- "pi" 1 pi = 12 pt (informative)
    deriving (Eq, Show)

-- See @CT_Point2D@ (p. 3989)
data Point2D = Point2D
    { _pt2dX :: Coordinate
    , _pt2dY :: Coordinate
    } deriving (Eq, Show)

unqPoint2D :: Int -> Int -> Point2D
unqPoint2D x y = Point2D (UnqCoordinate x) (UnqCoordinate y)

-- | Positive position or length in EMUs, maximu allowed value is 27273042316900.
-- see 20.1.10.41 "ST_PositiveCoordinate (Positive Coordinate)" (p. 2942)
newtype PositiveCoordinate = PositiveCoordinate Integer
    deriving (Eq, Ord, Show)

-- | This simple type represents an angle in 60,000ths of a degree.
-- Positive angles are clockwise (i.e., towards the positive y axis); negative
-- angles are counter-clockwise (i.e., towards the negative y axis).
newtype Angle = Angle Int
    deriving (Eq, Show)

data PositiveSize2D = PositiveSize2D
    { _ps2dX :: PositiveCoordinate
    , _ps2dY :: PositiveCoordinate
    } deriving (Eq, Show)

data Marker = Marker
    { _mrkCol    :: Int
    , _mrkColOff :: Coordinate
    , _mrkRow    :: Int
    , _mrkRowOff :: Coordinate
    } deriving (Eq, Show)

unqMarker :: (Int, Int) -> (Int, Int) -> Marker
unqMarker (col, colOff) (row, rowOff) =
    Marker col (UnqCoordinate colOff) row (UnqCoordinate rowOff)

data EditAs
    = EditAsTwoCell
    | EditAsOneCell
    | EditAsAbsolute
    deriving (Eq,Show)

data Anchoring
    = AbsoluteAnchor
      { absaPos :: Point2D
      , absaExt :: PositiveSize2D
      }
    | OneCellAnchor
      { onecaFrom :: Marker
      , onecaExt  :: PositiveSize2D
      }
    | TwoCellAnchor
      { tcaFrom   :: Marker
      , tcaTo     :: Marker
      , tcaEditAs :: EditAs
      }
    deriving (Eq, Show)

data DrawingObject a
    = Picture
      { _picMacro           :: Maybe Text
      , _picPublished       :: Bool
      , _picNonVisual       :: PicNonVisual
      , _picBlipFill        :: BlipFillProperties a
      , _picShapeProperties :: ShapeProperties
      -- TODO: style
      }
    -- TODO: sp, grpSp, graphicFrame, cxnSp, contentPart
    deriving (Eq, Show)

-- | This element is used to set certain properties related to a drawing
-- element on the client spreadsheet application.
--
-- see 20.5.2.3 "clientData (Client Data)" (p. 3156)
data ClientData = ClientData
    { _cldLcksWithSheet   :: Bool
    -- ^ This attribute indicates whether to disable selection on
    -- drawing elements when the sheet is protected.
    , _cldPrintsWithSheet :: Bool
    -- ^ This attribute indicates whether to print drawing elements
    -- when printing the sheet.
    } deriving (Eq, Show)

data PicNonVisual = PicNonVisual
    { _pnvDrawingProps :: PicDrawingNonVisual
    -- TODO: cNvPicPr
    }
    deriving (Eq, Show)

-- see 20.1.2.2.8 "cNvPr (Non-Visual Drawing Properties)" (p. 2731)
data PicDrawingNonVisual = PicDrawingNonVisual
    { _pdnvId          :: Int
    -- ^ Specifies a unique identifier for the current
    -- DrawingML object within the current
    --
    -- TODO: make ids internal and consistent by construction
    , _pdnvName        :: Text
    -- ^ Specifies the name of the object.
    -- Typically, this is used to store the original file
    -- name of a picture object.
    , _pdnvDescription :: Maybe Text
    -- ^ Alternative Text for Object
    , _pdnvHidden      :: Bool
    , _pdnvTitle       :: Maybe Text
    } deriving (Eq, Show)

data BlipFillProperties a = BlipFillProperties
    { _bfpImageInfo :: Maybe a
    , _bfpFillMode  :: Maybe FillMode
    -- TODO: dpi, rotWithShape, srcRect
    } deriving (Eq, Show)

-- see @a_EG_FillModeProperties@ (p. 4319)
data FillMode
    -- See 20.1.8.58 "tile (Tile)" (p. 2880)
    = FillTile    -- TODO: tx, ty, sx, sy, flip, algn
    -- See 20.1.8.56 "stretch (Stretch)" (p. 2879)
    | FillStretch -- TODO: srcRect
    deriving (Eq, Show)

-- See 20.1.2.2.35 "spPr (Shape Properties)" (p. 2751)
data ShapeProperties = ShapeProperties
    { _spXfrm     :: Maybe Transform2D
    , _spGeometry :: Maybe Geometry
    , _spOutline  :: Maybe LineProperties
    -- TODO: bwMode, a_EG_FillProperties, a_EG_EffectProperties, scene3d, sp3d, extLst
    } deriving (Eq, Show)

-- See 20.1.7.6 "xfrm (2D Transform for Individual Objects)" (p. 2849)
data Transform2D = Transform2D
    { _trRot     :: Angle
    -- ^ Specifies the rotation of the Graphic Frame.
    , _trFlipH   :: Bool
    -- ^ Specifies a horizontal flip. When true, this attribute defines
    -- that the shape is flipped horizontally about the center of its bounding box.
    , _trFlipV   :: Bool
    -- ^ Specifies a vertical flip. When true, this attribute defines
    -- that the shape is flipped vetically about the center of its bounding box.
    , _trOffset  :: Maybe Point2D
    -- ^ See 20.1.7.4 "off (Offset)" (p. 2847)
    , _trExtents :: Maybe PositiveSize2D
    -- ^ See 20.1.7.3 "ext (Extents)" (p. 2846) or
    -- 20.5.2.14 "ext (Shape Extent)" (p. 3165)
    }
    deriving (Eq, Show)

data Geometry
    -- TODO: custGeom
    = PresetGeometry
      -- TODO: prst, avList
      -- currently uses "rect" with empty avList
    deriving (Eq, Show)

-- See 20.1.2.2.24 "ln (Outline)" (p. 2744)
data LineProperties = LineProperties
    { _lnFill :: Maybe LineFill
    -- TODO: w, cap, cmpd, algn, a_EG_LineDashProperties,
    --   a_EG_LineJoinProperties, headEnd, tailEnd, extLst
    }
    deriving (Eq, Show)

data LineFill
    -- See 20.1.8.44 "noFill (No Fill)" (p. 2872)
    = LineNoFill
    -- TODO: solidFill, gradFill, pattFill
    deriving (Eq, Show)

-- See @EG_Anchor@ (p. 4052)
data Anchor a = Anchor
    { _anchAnchoring  :: Anchoring
    , _anchObject     :: DrawingObject a
    , _anchClientData :: ClientData
    } deriving (Eq, Show)

data GenericDrawing a = Drawing
    { _xdrAnchors :: [Anchor a]
    } deriving (Eq, Show)

-- See 20.5.2.35 "wsDr (Worksheet Drawing)" (p. 3176)
type Drawing = GenericDrawing FileInfo

type UnresolvedDrawing = GenericDrawing RefId

makeLenses ''Anchor
makeLenses ''DrawingObject
makeLenses ''BlipFillProperties
makeLenses ''GenericDrawing

{-------------------------------------------------------------------------------
  Default instances
-------------------------------------------------------------------------------}

instance Default ClientData where
    def = ClientData True True

{-------------------------------------------------------------------------------
  Parsing
-------------------------------------------------------------------------------}

instance FromCursor UnresolvedDrawing where
    fromCursor cur = [Drawing $ cur $/ anyElement >=> fromCursor]

instance FromCursor (Anchor RefId) where
    fromCursor cur = do
        _anchAnchoring  <- fromCursor cur
        _anchObject     <- cur $/ anyElement >=> fromCursor
        _anchClientData <- cur $/ element (xdr"clientData") >=> fromCursor
        return Anchor{..}

instance FromCursor Anchoring where
    fromCursor = anchoringFromNode . node

anchoringFromNode :: Node -> [Anchoring]
anchoringFromNode n | n `nodeElNameIs` xdr "twoCellAnchor" = do
                          tcaEditAs <- fromAttributeDef "editAs" EditAsTwoCell cur
                          tcaFrom <- cur $/ element (xdr"from") >=> fromCursor
                          tcaTo <- cur $/ element (xdr"to") >=> fromCursor
                          return TwoCellAnchor{..}
                    | n `nodeElNameIs` xdr "oneCellAnchor" = do
                          onecaFrom <- cur $/ element (xdr"from") >=> fromCursor
                          onecaExt <- cur $/ element (xdr"ext") >=> fromCursor
                          return OneCellAnchor{..}
                    | n `nodeElNameIs` xdr "absolueAnchor" = do
                          absaPos <- cur $/ element (xdr"pos") >=> fromCursor
                          absaExt <- cur $/ element (xdr"ext") >=> fromCursor
                          return AbsoluteAnchor{..}
                    | otherwise = fail "no matching anchoring node"
  where
    cur = fromNode n

nodeElNameIs :: Node -> Name -> Bool
nodeElNameIs (NodeElement el) name = elementName el == name
nodeElNameIs _ _                   = False

instance FromCursor Marker where
    fromCursor cur = do
        _mrkCol <- cur $/ element (xdr"col") &/ content >=> decimal
        _mrkColOff <- cur $/ element (xdr"colOff") &/ content >=> coordinate
        _mrkRow <- cur $/ element (xdr"row") &/ content >=> decimal
        _mrkRowOff <- cur $/ element (xdr"rowOff") &/ content >=> coordinate
        return Marker{..}

instance FromCursor (DrawingObject RefId) where
    fromCursor = drawingObjectFromNode . node

drawingObjectFromNode :: Node -> [DrawingObject RefId]
drawingObjectFromNode n | n `nodeElNameIs` xdr "pic" = do
                              _picMacro <- maybeAttribute "macro" cur
                              _picPublished <- fromAttributeDef "fPublished" False cur
                              _picNonVisual <- cur $/ element (xdr"nvPicPr") >=> fromCursor
                              _picBlipFill  <- cur $/ element (xdr"blipFill") >=> fromCursor
                              _picShapeProperties <- cur $/ element (xdr"spPr") >=> fromCursor
                              return Picture{..}
                        | otherwise = fail "no matching drawing object node"
    where
      cur = fromNode n

instance FromCursor PicNonVisual where
    fromCursor cur = do
        _pnvDrawingProps <- cur $/ element (xdr"cNvPr") >=> fromCursor
        return PicNonVisual{..}

instance FromCursor PicDrawingNonVisual where
    fromCursor cur = do
        _pdnvId <- fromAttribute "id" cur
        _pdnvName <- fromAttribute "name" cur
        _pdnvDescription <- maybeAttribute "descr" cur
        _pdnvHidden <- fromAttributeDef "hidden" False cur
        _pdnvTitle <- maybeAttribute "title" cur
        return PicDrawingNonVisual{..}

instance FromCursor (BlipFillProperties RefId) where
    fromCursor cur = do
        let _bfpImageInfo = listToMaybe $ cur $/ element (a"blip") >=>
                            fmap RefId . attribute (odr"embed")
            _bfpFillMode  = listToMaybe $ cur $/ anyElement >=> fromCursor
        return BlipFillProperties{..}

instance FromCursor ShapeProperties where
    fromCursor cur = do
        _spXfrm <- maybeFromElement (a"xfrm") cur
        let _spGeometry = listToMaybe $ cur $/ anyElement >=> fromCursor
        _spOutline <- maybeFromElement (a"ln") cur
        return ShapeProperties{..}

instance FromCursor FillMode where
    fromCursor = fillModeFromNode . node

fillModeFromNode :: Node -> [FillMode]
fillModeFromNode n | n `nodeElNameIs` a "stretch" = return FillStretch
                   | n `nodeElNameIs` a "stretch" = return FillTile
                   | otherwise = fail "no matching fill mode node"

instance FromCursor Transform2D where
    fromCursor cur = do
        _trRot     <- fromAttributeDef "rot" (Angle 0) cur
        _trFlipH   <- fromAttributeDef "flipH" False cur
        _trFlipV   <- fromAttributeDef "flipV" False cur
        _trOffset  <- maybeFromElement (a"off") cur
        _trExtents <- maybeFromElement (a"ext") cur
        return Transform2D{..}

instance FromCursor Geometry where
    fromCursor = geometryFromNode . node

geometryFromNode :: Node -> [Geometry]
geometryFromNode n | n `nodeElNameIs` a "prstGeom" =
                         return PresetGeometry
                   | otherwise = fail "no matching geometry node"

instance FromCursor LineProperties where
    fromCursor cur = do
        let _lnFill = listToMaybe $ cur $/ anyElement >=> fromCursor
        return LineProperties{..}

instance FromCursor ClientData where
    fromCursor cur = do
        _cldLcksWithSheet   <- fromAttributeDef "fLocksWithSheet"  True cur
        _cldPrintsWithSheet <- fromAttributeDef "fPrintsWithSheet" True cur
        return ClientData{..}

instance FromCursor Point2D where
    fromCursor cur = do
        x <- coordinate =<< fromAttribute "x" cur
        y <- coordinate =<< fromAttribute "y" cur
        return $ Point2D x y

instance FromCursor PositiveSize2D where
    fromCursor cur = do
        cx <- PositiveCoordinate <$> fromAttribute "cx" cur
        cy <- PositiveCoordinate <$> fromAttribute "cy" cur
        return $ PositiveSize2D cx cy

-- see 20.1.10.3 "ST_Angle (Angle)" (p. 2912)
instance FromAttrVal Angle where
    fromAttrVal t = first Angle <$> fromAttrVal t

-- see 20.5.3.2 "ST_EditAs (Resizing Behaviors)" (p. 3177)
instance FromAttrVal EditAs where
    fromAttrVal "absolute" = readSuccess EditAsAbsolute
    fromAttrVal "oneCell"  = readSuccess EditAsOneCell
    fromAttrVal "twoCell"  = readSuccess EditAsTwoCell
    fromAttrVal t          = invalidText "EditAs" t

instance FromCursor LineFill where
    fromCursor = lineFillFromNode . node

lineFillFromNode :: Node -> [LineFill]
lineFillFromNode n | n `nodeElNameIs` a "noFill" = return LineNoFill
                   | otherwise = fail "no matching line fill node"

coordinate :: Monad m => Text -> m Coordinate
coordinate t =  case T.decimal t of
  Right (d, leftover) | leftover == T.empty ->
      return $ UnqCoordinate d
  _ ->
      case T.rational t of
          Right (r, "cm") -> return $ UniversalMeasure UnitCm r
          Right (r, "mm") -> return $ UniversalMeasure UnitMm r
          Right (r, "in") -> return $ UniversalMeasure UnitIn r
          Right (r, "pt") -> return $ UniversalMeasure UnitPt r
          Right (r, "pc") -> return $ UniversalMeasure UnitPc r
          Right (r, "pi") -> return $ UniversalMeasure UnitPi r
          _               -> fail $ "invalid coordinate: " ++ show t

{-------------------------------------------------------------------------------
  Rendering
-------------------------------------------------------------------------------}

instance ToDocument UnresolvedDrawing where
    toDocument = documentFromNsElement "Drawing generated by xlsx" xlDrawingNs
                 . toElement "wsDr"

instance ToElement UnresolvedDrawing where
    toElement nm (Drawing anchors) = Element
        { elementName       = nm
        , elementAttributes = M.empty
        , elementNodes      = map NodeElement $
                              map anchorToElement anchors
        }

anchorToElement :: Anchor RefId -> Element
anchorToElement Anchor{..} = el
    { elementNodes = elementNodes el ++
                     map NodeElement [ drawingObjEl, cdEl ] }
  where
    el = anchoringToElement _anchAnchoring
    drawingObjEl = drawingObjToElement _anchObject
    cdEl = toElement "clientData" _anchClientData

anchoringToElement :: Anchoring -> Element
anchoringToElement anchoring = elementList nm attrs elements
  where
    (nm, attrs, elements) = case anchoring of
        AbsoluteAnchor{..} ->
            ("absoluteAnchor", [],
             [ toElement "pos" absaPos, toElement "ext" absaExt ])
        OneCellAnchor{..}  ->
            ("oneCellAnchor", [],
             [ toElement "from" onecaFrom, toElement "ext" onecaExt ])
        TwoCellAnchor{..}  ->
            ("twoCellAnchor", [ "editAs" .= tcaEditAs ],
             [ toElement "from" tcaFrom, toElement "to" tcaTo ])

instance ToElement Marker where
    toElement nm Marker{..} = elementListSimple nm elements
      where
        elements = [ elementContent "col"    (toAttrVal _mrkCol)
                   , elementContent "colOff" (toAttrVal _mrkColOff)
                   , elementContent "row"    (toAttrVal _mrkRow)
                   , elementContent "rowOff" (toAttrVal _mrkRowOff)]

drawingObjToElement :: DrawingObject RefId -> Element
drawingObjToElement Picture{..} =
    elementList "pic" attrs elements
  where
    attrs = catMaybes [ "macro"      .=? _picMacro
                      , "fPublished" .=? justTrue _picPublished]
    elements = [ toElement "nvPicPr"  _picNonVisual
               , toElement "blipFill" _picBlipFill
               , toElement "spPr"     _picShapeProperties ]

instance ToElement PicNonVisual where
    toElement nm PicNonVisual{..} =
        elementListSimple nm [toElement "cNvPr" _pnvDrawingProps]

instance ToElement PicDrawingNonVisual where
    toElement nm PicDrawingNonVisual{..} =
        leafElement nm attrs
      where
        attrs = [ "id"    .= _pdnvId
                , "name"  .= _pdnvName ] ++
                catMaybes [ "descr"  .=? _pdnvDescription
                          , "hidden" .=? justTrue _pdnvHidden
                          , "title"  .=? _pdnvTitle ]

instance ToElement (BlipFillProperties RefId) where
    toElement nm BlipFillProperties{..} =
        elementListSimple nm elements
      where
        elements = catMaybes [ (\rId -> leafElement (a"blip") [ odr "embed" .= rId ]) <$> _bfpImageInfo
                             , fillModeToElement <$> _bfpFillMode ]

fillModeToElement :: FillMode -> Element
fillModeToElement FillStretch = emptyElement (a"stretch")
fillModeToElement FillTile    = emptyElement (a"stretch")

instance ToElement ShapeProperties where
    toElement nm ShapeProperties{..} = elementListSimple nm elements
      where
        elements = catMaybes [ toElement (a"xfrm") <$> _spXfrm
                             , geometryToElement <$> _spGeometry
                             , toElement (a"ln")  <$> _spOutline ]

instance ToElement Transform2D where
    toElement nm Transform2D{..} = elementList nm attrs elements
      where
        attrs = catMaybes [ "rot"   .=? justNonDef (Angle 0) _trRot
                          , "flipH" .=? justTrue _trFlipH
                          , "flipV" .=? justTrue _trFlipV ]
        elements = catMaybes [ toElement (a"off") <$> _trOffset
                             , toElement (a"ext") <$> _trExtents ]

geometryToElement :: Geometry -> Element
geometryToElement PresetGeometry = emptyElement (a"prstGeom")

instance ToElement LineProperties where
    toElement nm LineProperties{..} =
      elementListSimple nm $ catMaybes [ lineFillToElement <$> _lnFill ]

lineFillToElement :: LineFill -> Element
lineFillToElement LineNoFill = emptyElement (a"noFill")

instance ToElement ClientData where
    toElement nm ClientData{..} = leafElement nm attrs
      where
        attrs = catMaybes [ "fLocksWithSheet"  .=? justFalse _cldLcksWithSheet
                          , "fPrintsWithSheet" .=? justFalse _cldPrintsWithSheet
                          ]

instance ToElement Point2D where
    toElement nm Point2D{..} = leafElement nm [ "x" .= _pt2dX
                                              , "y" .= _pt2dY
                                              ]

instance ToElement PositiveSize2D where
    toElement nm PositiveSize2D{..} = leafElement nm [ "cx" .= _ps2dX
                                                     , "cy" .= _ps2dY ]

instance ToAttrVal Coordinate where
    toAttrVal (UnqCoordinate x) = toAttrVal x
    toAttrVal (UniversalMeasure unit x) = toAttrVal x <> unitToText unit
      where
        unitToText UnitCm = "cm"
        unitToText UnitMm = "mm"
        unitToText UnitIn = "in"
        unitToText UnitPt = "pt"
        unitToText UnitPc = "pc"
        unitToText UnitPi = "pi"

instance ToAttrVal PositiveCoordinate where
    toAttrVal (PositiveCoordinate x) = toAttrVal x

instance ToAttrVal EditAs where
    toAttrVal EditAsAbsolute = "absolute"
    toAttrVal EditAsOneCell  = "oneCell"
    toAttrVal EditAsTwoCell  = "twoCell"

instance ToAttrVal Angle where
    toAttrVal (Angle x) = toAttrVal x

-- | Add DrawingML namespace to name
a :: Text -> Name
a x = Name
  { nameLocalName = x
  , nameNamespace = Just drawingNs
  , namePrefix = Nothing
  }

drawingNs :: Text
drawingNs = "http://schemas.openxmlformats.org/drawingml/2006/main"

-- | Add Spreadsheet DrawingML namespace to name
xdr :: Text -> Name
xdr x = Name
  { nameLocalName = x
  , nameNamespace = Just xlDrawingNs
  , namePrefix = Nothing
  }

xlDrawingNs :: Text
xlDrawingNs = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
