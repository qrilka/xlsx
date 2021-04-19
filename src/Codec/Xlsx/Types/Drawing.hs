{-# LANGUAGE CPP #-}
{-# LANGUAGE DataKinds #-}
{-# LANGUAGE FlexibleInstances #-}
{-# LANGUAGE OverloadedStrings #-}
{-# LANGUAGE RecordWildCards #-}
{-# LANGUAGE StandaloneDeriving #-}
{-# LANGUAGE TemplateHaskell #-}
{-# LANGUAGE TupleSections #-}
{-# LANGUAGE TypeFamilies #-}
{-# LANGUAGE TypeSynonymInstances #-}
{-# LANGUAGE DeriveGeneric #-}
module Codec.Xlsx.Types.Drawing where

import Control.Arrow (first)
import Control.DeepSeq (NFData)
#ifdef USE_MICROLENS
import Lens.Micro.TH (makeLenses)
#else
import Control.Lens.TH
#endif
import Data.ByteString.Lazy (ByteString)
import Data.Default
import qualified Data.Map as M
import Data.Maybe (catMaybes, listToMaybe, mapMaybe)
import Data.Text (Text)
import qualified Data.Text as T
import GHC.Generics (Generic)
import Text.XML
import Text.XML.Cursor

import Codec.Xlsx.Parser.Internal
import Codec.Xlsx.Types.Drawing.Chart
import Codec.Xlsx.Types.Drawing.Common
import Codec.Xlsx.Types.Internal
import Codec.Xlsx.Types.Internal.Relationships
import Codec.Xlsx.Writer.Internal

-- | information about image file as a par of a drawing
data FileInfo = FileInfo
    { _fiFilename    :: FilePath
    -- ^ image filename, images are assumed to be stored under path "xl\/media\/"
    , _fiContentType :: Text
    -- ^ image content type, ECMA-376 advises to use "image\/png" or "image\/jpeg"
    -- if interoperability is wanted
    , _fiContents    :: ByteString
    -- ^ image file contents
    } deriving (Eq, Show, Generic)
instance NFData FileInfo

data Marker = Marker
    { _mrkCol    :: Int
    , _mrkColOff :: Coordinate
    , _mrkRow    :: Int
    , _mrkRowOff :: Coordinate
    } deriving (Eq, Show, Generic)
instance NFData Marker

unqMarker :: (Int, Int) -> (Int, Int) -> Marker
unqMarker (col, colOff) (row, rowOff) =
    Marker col (UnqCoordinate colOff) row (UnqCoordinate rowOff)

data EditAs
    = EditAsTwoCell
    | EditAsOneCell
    | EditAsAbsolute
    deriving (Eq, Show, Generic)
instance NFData EditAs

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
    deriving (Eq, Show, Generic)
instance NFData Anchoring

data DrawingObject p g
  = Picture { _picMacro :: Maybe Text
           ,  _picPublished :: Bool
           ,  _picNonVisual :: PicNonVisual
           ,  _picBlipFill :: BlipFillProperties p
           ,  _picShapeProperties :: ShapeProperties
              -- TODO: style
            }
  | Graphic { _grNonVisual :: GraphNonVisual
           ,  _grChartSpace :: g
           ,  _grTransform :: Transform2D}
    -- TODO: sp, grpSp, graphicFrame, cxnSp, contentPart
  deriving (Eq, Show, Generic)
instance (NFData p, NFData g) => NFData (DrawingObject p g)

-- | basic function to create picture drawing object
--
-- /Note:/ specification says that drawing element ids need to be
-- unique within 1 document, otherwise /...document shall be
-- considered non-conformant/.
picture :: DrawingElementId -> FileInfo -> DrawingObject FileInfo c
picture dId fi =
  Picture
  { _picMacro = Nothing
  , _picPublished = False
  , _picNonVisual = nonVis
  , _picBlipFill = bfProps
  , _picShapeProperties = shProps
  }
  where
    nonVis =
      PicNonVisual $
      NonVisualDrawingProperties
      { _nvdpId = dId
      , _nvdpName = T.pack $ _fiFilename fi
      , _nvdpDescription = Nothing
      , _nvdpHidden = False
      , _nvdpTitle = Nothing
      }
    bfProps =
      BlipFillProperties
      {_bfpImageInfo = Just fi, _bfpFillMode = Just FillStretch}
    shProps =
      ShapeProperties
      { _spXfrm = Nothing
      , _spGeometry = Nothing
      , _spFill = Just NoFill
      , _spOutline = Just $ def {_lnFill = Just NoFill}
      }

-- | helper to retrive information about all picture files in
-- particular drawing alongside with their anchorings (i.e. sizes and
-- positions)
extractPictures :: Drawing -> [(Anchoring, FileInfo)]
extractPictures dr = mapMaybe maybePictureInfo $ _xdrAnchors dr
  where
    maybePictureInfo Anchor {..} =
      case _anchObject of
        Picture {..} -> (_anchAnchoring,) <$> _bfpImageInfo _picBlipFill
        _ -> Nothing

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
    } deriving (Eq, Show, Generic)
instance NFData ClientData

data PicNonVisual = PicNonVisual
  { _pnvDrawingProps :: NonVisualDrawingProperties
    -- TODO: cNvPicPr
  } deriving (Eq, Show, Generic)
instance NFData PicNonVisual

data GraphNonVisual = GraphNonVisual
  { _gnvDrawingProps :: NonVisualDrawingProperties
    -- TODO cNvGraphicFramePr
  } deriving (Eq, Show, Generic)
instance NFData GraphNonVisual

newtype DrawingElementId = DrawingElementId
  { unDrawingElementId :: Int
  } deriving (Eq, Show, Generic)
instance NFData DrawingElementId

-- see 20.1.2.2.8 "cNvPr (Non-Visual Drawing Properties)" (p. 2731)
data NonVisualDrawingProperties = NonVisualDrawingProperties
    { _nvdpId          :: DrawingElementId
    -- ^ Specifies a unique identifier for the current
    -- DrawingML object within the current
    --
    -- TODO: make ids internal and consistent by construction
    , _nvdpName        :: Text
    -- ^ Specifies the name of the object.
    -- Typically, this is used to store the original file
    -- name of a picture object.
    , _nvdpDescription :: Maybe Text
    -- ^ Alternative Text for Object
    , _nvdpHidden      :: Bool
    , _nvdpTitle       :: Maybe Text
    } deriving (Eq, Show, Generic)
instance NFData NonVisualDrawingProperties

data BlipFillProperties a = BlipFillProperties
    { _bfpImageInfo :: Maybe a
    , _bfpFillMode  :: Maybe FillMode
    -- TODO: dpi, rotWithShape, srcRect
    } deriving (Eq, Show, Generic)
instance NFData a => NFData (BlipFillProperties a)

-- see @a_EG_FillModeProperties@ (p. 4319)
data FillMode
    -- See 20.1.8.58 "tile (Tile)" (p. 2880)
    = FillTile    -- TODO: tx, ty, sx, sy, flip, algn
    -- See 20.1.8.56 "stretch (Stretch)" (p. 2879)
    | FillStretch -- TODO: srcRect
    deriving (Eq, Show, Generic)
instance NFData FillMode

-- See @EG_Anchor@ (p. 4052)
data Anchor p g = Anchor
    { _anchAnchoring  :: Anchoring
    , _anchObject     :: DrawingObject p g
    , _anchClientData :: ClientData
    } deriving (Eq, Show, Generic)
instance (NFData p, NFData g) => NFData (Anchor p g)

data GenericDrawing p g = Drawing
    { _xdrAnchors :: [Anchor p g]
    } deriving (Eq, Show, Generic)
instance (NFData p, NFData g) => NFData (GenericDrawing p g)

-- See 20.5.2.35 "wsDr (Worksheet Drawing)" (p. 3176)
type Drawing = GenericDrawing FileInfo ChartSpace

type UnresolvedDrawing = GenericDrawing RefId RefId

makeLenses ''Anchor
makeLenses ''DrawingObject
makeLenses ''BlipFillProperties
makeLenses ''GenericDrawing

-- | simple drawing object anchoring using one cell as a top lelft
-- corner and dimensions of that object
simpleAnchorXY :: (Int, Int) -- ^ x+y coordinates of a cell used as
                             -- top left anchoring corner
               -> PositiveSize2D -- ^ size of drawing object to be
                                 -- anchored
               -> DrawingObject p g
               -> Anchor p g
simpleAnchorXY (x, y) sz obj =
  Anchor
  { _anchAnchoring =
      OneCellAnchor {onecaFrom = unqMarker (x, 0) (y, 0), onecaExt = sz}
  , _anchObject = obj
  , _anchClientData = def
  }

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

instance FromCursor (Anchor RefId RefId) where
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

instance FromCursor Marker where
    fromCursor cur = do
        _mrkCol <- cur $/ element (xdr"col") &/ content >=> decimal
        _mrkColOff <- cur $/ element (xdr"colOff") &/ content >=> coordinate
        _mrkRow <- cur $/ element (xdr"row") &/ content >=> decimal
        _mrkRowOff <- cur $/ element (xdr"rowOff") &/ content >=> coordinate
        return Marker{..}

instance FromCursor (DrawingObject RefId RefId) where
    fromCursor = drawingObjectFromNode . node

drawingObjectFromNode :: Node -> [DrawingObject RefId RefId]
drawingObjectFromNode n
  | n `nodeElNameIs` xdr "pic" = do
    _picMacro <- maybeAttribute "macro" cur
    _picPublished <- fromAttributeDef "fPublished" False cur
    _picNonVisual <- cur $/ element (xdr "nvPicPr") >=> fromCursor
    _picBlipFill <- cur $/ element (xdr "blipFill") >=> fromCursor
    _picShapeProperties <- cur $/ element (xdr "spPr") >=> fromCursor
    return Picture {..}
  | n `nodeElNameIs` xdr "graphicFrame" = do
    _grNonVisual <-
      cur $/ element (xdr "nvGraphicFramePr") >=> fromCursor
    _grTransform <- cur $/ element (xdr "xfrm") >=> fromCursor
    _grChartSpace <-
      cur $/ element (a_ "graphic") &/ element (a_ "graphicData") &/
      element (c_ "chart") >=> fmap RefId .  attribute (odr "id")
    return Graphic {..}
  | otherwise = fail "no matching drawing object node"
  where
    cur = fromNode n

instance FromCursor PicNonVisual where
    fromCursor cur = do
        _pnvDrawingProps <- cur $/ element (xdr"cNvPr") >=> fromCursor
        return PicNonVisual{..}

instance FromCursor GraphNonVisual where
  fromCursor cur = do
    _gnvDrawingProps <- cur $/ element (xdr "cNvPr") >=> fromCursor
    return GraphNonVisual {..}

instance FromCursor NonVisualDrawingProperties where
    fromCursor cur = do
        _nvdpId <- fromAttribute "id" cur
        _nvdpName <- fromAttribute "name" cur
        _nvdpDescription <- maybeAttribute "descr" cur
        _nvdpHidden <- fromAttributeDef "hidden" False cur
        _nvdpTitle <- maybeAttribute "title" cur
        return NonVisualDrawingProperties{..}

instance FromAttrVal DrawingElementId where
  fromAttrVal = fmap (first DrawingElementId) . fromAttrVal

instance FromCursor (BlipFillProperties RefId) where
    fromCursor cur = do
        let _bfpImageInfo = listToMaybe $ cur $/ element (a_ "blip") >=>
                            fmap RefId . attribute (odr"embed")
            _bfpFillMode  = listToMaybe $ cur $/ anyElement >=> fromCursor
        return BlipFillProperties{..}

instance FromCursor FillMode where
    fromCursor = fillModeFromNode . node

fillModeFromNode :: Node -> [FillMode]
fillModeFromNode n | n `nodeElNameIs` a_ "stretch" = return FillStretch
                   | n `nodeElNameIs` a_ "stretch" = return FillTile
                   | otherwise = fail "no matching fill mode node"

-- see 20.5.3.2 "ST_EditAs (Resizing Behaviors)" (p. 3177)
instance FromAttrVal EditAs where
    fromAttrVal "absolute" = readSuccess EditAsAbsolute
    fromAttrVal "oneCell"  = readSuccess EditAsOneCell
    fromAttrVal "twoCell"  = readSuccess EditAsTwoCell
    fromAttrVal t          = invalidText "EditAs" t

instance FromCursor ClientData where
    fromCursor cur = do
        _cldLcksWithSheet   <- fromAttributeDef "fLocksWithSheet"  True cur
        _cldPrintsWithSheet <- fromAttributeDef "fPrintsWithSheet" True cur
        return ClientData{..}

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

anchorToElement :: Anchor RefId RefId -> Element
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

drawingObjToElement :: DrawingObject RefId RefId -> Element
drawingObjToElement Picture {..} = elementList "pic" attrs elements
  where
    attrs =
      catMaybes ["macro" .=? _picMacro, "fPublished" .=? justTrue _picPublished]
    elements =
      [ toElement "nvPicPr" _picNonVisual
      , toElement "blipFill" _picBlipFill
      , toElement "spPr" _picShapeProperties
      ]
drawingObjToElement Graphic {..} = elementListSimple "graphicFrame" elements
  where
    elements =
      [ toElement "nvGraphicFramePr" _grNonVisual
      , toElement "xfrm" _grTransform
      , graphicEl
      ]
    graphicEl =
      elementListSimple
        (a_ "graphic")
        [ elementList
            (a_ "graphicData")
            ["uri" .= chartNs]
            [leafElement (c_ "chart") [odr "id" .= _grChartSpace]]
        ]

instance ToElement PicNonVisual where
  toElement nm PicNonVisual {..} =
    elementListSimple
      nm
      [toElement "cNvPr" _pnvDrawingProps, emptyElement "cNvPicPr"]

instance ToElement GraphNonVisual where
  toElement nm GraphNonVisual {..} =
    elementListSimple
      nm
      [toElement "cNvPr" _gnvDrawingProps, emptyElement "cNvGraphicFramePr"]

instance ToElement NonVisualDrawingProperties where
    toElement nm NonVisualDrawingProperties{..} =
        leafElement nm attrs
      where
        attrs = [ "id"    .= _nvdpId
                , "name"  .= _nvdpName ] ++
                catMaybes [ "descr"  .=? _nvdpDescription
                          , "hidden" .=? justTrue _nvdpHidden
                          , "title"  .=? _nvdpTitle ]

instance ToAttrVal DrawingElementId where
  toAttrVal = toAttrVal . unDrawingElementId

instance ToElement (BlipFillProperties RefId) where
    toElement nm BlipFillProperties{..} =
        elementListSimple nm elements
      where
        elements = catMaybes [ (\rId -> leafElement (a_ "blip") [ odr "embed" .= rId ]) <$> _bfpImageInfo
                             , fillModeToElement <$> _bfpFillMode ]

fillModeToElement :: FillMode -> Element
fillModeToElement FillStretch = emptyElement (a_ "stretch")
fillModeToElement FillTile    = emptyElement (a_ "stretch")

instance ToElement ClientData where
    toElement nm ClientData{..} = leafElement nm attrs
      where
        attrs = catMaybes [ "fLocksWithSheet"  .=? justFalse _cldLcksWithSheet
                          , "fPrintsWithSheet" .=? justFalse _cldPrintsWithSheet
                          ]

instance ToAttrVal EditAs where
    toAttrVal EditAsAbsolute = "absolute"
    toAttrVal EditAsOneCell  = "oneCell"
    toAttrVal EditAsTwoCell  = "twoCell"

-- | Add Spreadsheet DrawingML namespace to name
xdr :: Text -> Name
xdr x = Name
  { nameLocalName = x
  , nameNamespace = Just xlDrawingNs
  , namePrefix = Nothing
  }

xlDrawingNs :: Text
xlDrawingNs = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
