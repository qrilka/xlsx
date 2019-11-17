{-# LANGUAGE OverloadedStrings #-}
{-# LANGUAGE RecordWildCards #-}
{-# LANGUAGE TemplateHaskell #-}
{-# LANGUAGE DeriveGeneric #-}
module Codec.Xlsx.Types.Drawing.Common where

import GHC.Generics (Generic)

import Control.Arrow (first)
import Control.Lens.TH
import Control.Monad (join)
import Control.Monad.Fail (MonadFail)
import Control.DeepSeq (NFData)
import Data.Default
import Data.Maybe (catMaybes, listToMaybe)
import Data.Monoid ((<>))
import Data.Text (Text)
import qualified Data.Text as T
import qualified Data.Text.Read as T
import Text.XML
import Text.XML.Cursor

import Codec.Xlsx.Parser.Internal
import Codec.Xlsx.Writer.Internal

-- | This simple type represents an angle in 60,000ths of a degree.
-- Positive angles are clockwise (i.e., towards the positive y axis); negative
-- angles are counter-clockwise (i.e., towards the negative y axis).
newtype Angle =
  Angle Int
  deriving (Eq, Show, Generic)
instance NFData Angle

-- | A string with rich text formatting
--
-- TODO: horzOverflow, lIns, tIns, rIns, bIns, numCol, spcCol, rtlCol,
--   fromWordArt, forceAA, upright, compatLnSpc, prstTxWarp,
--   a_EG_TextAutofit, scene3d, a_EG_Text3D, extLst
--
-- See @CT_TextBody@ (p. 4034)
data TextBody = TextBody
  { _txbdRotation :: Angle
    -- ^ Specifies the rotation that is being applied to the text within the bounding box.
  , _txbdSpcFirstLastPara :: Bool
    -- ^ Specifies whether the before and after paragraph spacing defined by the user is
    -- to be respected.
  , _txbdVertOverflow :: TextVertOverflow
    -- ^ Determines whether the text can flow out of the bounding box vertically.
  , _txbdVertical :: TextVertical
    -- ^ Determines if the text within the given text body should be displayed vertically.
  , _txbdWrap :: TextWrap
    -- ^ Specifies the wrapping options to be used for this text body.
  , _txbdAnchor :: TextAnchoring
    -- ^ Specifies the anchoring position of the txBody within the shape.
  , _txbdAnchorCenter :: Bool
    -- ^ Specifies the centering of the text box. The way it works fundamentally is
    -- to determine the smallest possible "bounds box" for the text and then to center
    -- that "bounds box" accordingly. This is different than paragraph alignment, which
    -- aligns the text within the "bounds box" for the text.
  , _txbdParagraphs :: [TextParagraph]
    -- ^ Paragraphs of text within the containing text body
  } deriving (Eq, Show, Generic)
instance NFData TextBody


-- | Text vertical overflow
-- See 20.1.10.83 "ST_TextVertOverflowType (Text Vertical Overflow)" (p. 3083)
data TextVertOverflow
  = TextVertOverflowClip
    -- ^ Pay attention to top and bottom barriers. Provide no indication that there is
    -- text which is not visible.
  | TextVertOverflowEllipsis
    -- ^ Pay attention to top and bottom barriers. Use an ellipsis to denote that
    -- there is text which is not visible.
  | TextVertOverflow
    -- ^ Overflow the text and pay no attention to top and bottom barriers.
  deriving (Eq, Show, Generic)
instance NFData TextVertOverflow

-- | If there is vertical text, determines what kind of vertical text is going to be used.
--
--  See 20.1.10.82 "ST_TextVerticalType (Vertical Text Types)" (p. 3083)
data TextVertical
  = TextVerticalEA
    -- ^ A special version of vertical text, where some fonts are displayed as if rotated
    -- by 90 degrees while some fonts (mostly East Asian) are displayed vertical.
  | TextVerticalHorz
    -- ^ Horizontal text. This should be default.
  | TextVerticalMongolian
    -- ^ A special version of vertical text, where some fonts are displayed as if rotated
    -- by 90 degrees while some fonts (mostly East Asian) are displayed vertical. The
    -- difference between this and the 'TextVerticalEA' is the text flows top down then
    -- LEFT RIGHT, instead of RIGHT LEFT
  | TextVertical
    -- ^ Determines if all of the text is vertical orientation (each line is 90 degrees
    -- rotated clockwise, so it goes from top to bottom; each next line is to the left
    -- from the previous one).
  | TextVertical270
    -- ^ Determines if all of the text is vertical orientation (each line is 270 degrees
    -- rotated clockwise, so it goes from bottom to top; each next line is to the right
    -- from the previous one).
  | TextVerticalWordArt
    -- ^ Determines if all of the text is vertical ("one letter on top of another").
  | TextVerticalWordArtRtl
    -- ^  Specifies that vertical WordArt should be shown from right to left rather than
    -- left to right.
  deriving (Eq, Show, Generic)
instance NFData TextVertical

-- | Text wrapping types
--
-- See 20.1.10.84 "ST_TextWrappingType (Text Wrapping Types)" (p. 3084)
data TextWrap
    = TextWrapNone
    -- ^ No wrapping occurs on this text body. Words spill out without
    -- paying attention to the bounding rectangle boundaries.
    | TextWrapSquare
    -- ^ Determines whether we wrap words within the bounding rectangle.
    deriving (Eq, Show, Generic)
instance NFData TextWrap

-- | This type specifies a list of available anchoring types for text.
--
-- See 20.1.10.59 "ST_TextAnchoringType (Text Anchoring Types)" (p. 3058)
data TextAnchoring
  = TextAnchoringBottom
    -- ^ Anchor the text at the bottom of the bounding rectangle.
  | TextAnchoringCenter
    -- ^ Anchor the text at the middle of the bounding rectangle.
  | TextAnchoringDistributed
    -- ^  Anchor the text so that it is distributed vertically. When text is horizontal,
    -- this spaces out the actual lines of text and is almost always identical in
    -- behavior to 'TextAnchoringJustified' (special case: if only 1 line, then anchored
    -- in middle). When text is vertical, then it distributes the letters vertically.
    -- This is different than 'TextAnchoringJustified', because it always forces distribution
    -- of the words, even if there are only one or two words in a line.
  | TextAnchoringJustified
    -- ^ Anchor the text so that it is justified vertically. When text is horizontal,
    -- this spaces out the actual lines of text and is almost always identical in
    -- behavior to 'TextAnchoringDistributed' (special case: if only 1 line, then anchored at
    -- top). When text is vertical, then it justifies the letters vertically. This is
    -- different than 'TextAnchoringDistributed' because in some cases such as very little
    -- text in a line, it does not justify.
  | TextAnchoringTop
    -- ^ Anchor the text at the top of the bounding rectangle.
  deriving (Eq, Show, Generic)
instance NFData TextAnchoring

-- See 21.1.2.2.6 "p (Text Paragraphs)" (p. 3211)
data TextParagraph = TextParagraph
  { _txpaDefCharProps :: Maybe TextCharacterProperties
  , _txpaRuns :: [TextRun]
  } deriving (Eq, Show, Generic)
instance NFData TextParagraph

-- | Text character properties
--
-- TODO: kumimoji, lang, altLang, sz, strike, kern, cap, spc,
--   normalizeH, baseline, noProof, dirty, err, smtClean, smtId,
--   bmk, ln, a_EG_FillProperties, a_EG_EffectProperties, highlight,
--   a_EG_TextUnderlineLine, a_EG_TextUnderlineFill, latin, ea, cs,
--   sym, hlinkClick, hlinkMouseOver, rtl, extLst
--
-- See @CT_TextCharacterProperties@ (p. 4039)
data TextCharacterProperties = TextCharacterProperties
  { _txchBold :: Bool
  , _txchItalic :: Bool
  , _txchUnderline :: Bool
  } deriving (Eq, Show, Generic)
instance NFData TextCharacterProperties

-- | Text run
--
-- TODO: br, fld
data TextRun = RegularRun
  { _txrCharProps :: Maybe TextCharacterProperties
  , _txrText :: Text
  } deriving (Eq, Show, Generic)
instance NFData TextRun

-- | This simple type represents a one dimensional position or length
--
-- See 20.1.10.16 "ST_Coordinate (Coordinate)" (p. 2921)
data Coordinate
  = UnqCoordinate Int
    -- ^ see 20.1.10.19 "ST_CoordinateUnqualified (Coordinate)" (p. 2922)
  | UniversalMeasure UnitIdentifier
                     Double
    -- ^ see 22.9.2.15 "ST_UniversalMeasure (Universal Measurement)" (p. 3793)
  deriving (Eq, Show, Generic)
instance NFData Coordinate

-- | Units used in "Universal measure" coordinates
-- see 22.9.2.15 "ST_UniversalMeasure (Universal Measurement)" (p. 3793)
data UnitIdentifier
  = UnitCm -- "cm" As defined in ISO 31.
  | UnitMm -- "mm" As defined in ISO 31.
  | UnitIn -- "in" 1 in = 2.54 cm (informative)
  | UnitPt -- "pt" 1 pt = 1/72 in (informative)
  | UnitPc -- "pc" 1 pc = 12 pt (informative)
  | UnitPi -- "pi" 1 pi = 12 pt (informative)
  deriving (Eq, Show, Generic)
instance NFData UnitIdentifier

-- See @CT_Point2D@ (p. 3989)
data Point2D = Point2D
  { _pt2dX :: Coordinate
  , _pt2dY :: Coordinate
  } deriving (Eq, Show, Generic)
instance NFData Point2D

unqPoint2D :: Int -> Int -> Point2D
unqPoint2D x y = Point2D (UnqCoordinate x) (UnqCoordinate y)

-- | Positive position or length in EMUs, maximu allowed value is 27273042316900.
-- see 20.1.10.41 "ST_PositiveCoordinate (Positive Coordinate)" (p. 2942)
newtype PositiveCoordinate =
  PositiveCoordinate Integer
  deriving (Eq, Ord, Show, Generic)
instance NFData PositiveCoordinate

data PositiveSize2D = PositiveSize2D
  { _ps2dX :: PositiveCoordinate
  , _ps2dY :: PositiveCoordinate
  } deriving (Eq, Show, Generic)
instance NFData PositiveSize2D

positiveSize2D :: Integer -> Integer -> PositiveSize2D
positiveSize2D x y =
  PositiveSize2D (PositiveCoordinate x) (PositiveCoordinate y)

cmSize2D :: Integer -> Integer -> PositiveSize2D
cmSize2D x y = positiveSize2D (cm2emu x) (cm2emu y)

cm2emu :: Integer -> Integer
cm2emu cm = 360000 * cm

-- See 20.1.7.6 "xfrm (2D Transform for Individual Objects)" (p. 2849)
data Transform2D = Transform2D
  { _trRot :: Angle
    -- ^ Specifies the rotation of the Graphic Frame.
  , _trFlipH :: Bool
    -- ^ Specifies a horizontal flip. When true, this attribute defines
    -- that the shape is flipped horizontally about the center of its bounding box.
  , _trFlipV :: Bool
    -- ^ Specifies a vertical flip. When true, this attribute defines
    -- that the shape is flipped vetically about the center of its bounding box.
  , _trOffset :: Maybe Point2D
    -- ^ See 20.1.7.4 "off (Offset)" (p. 2847)
  , _trExtents :: Maybe PositiveSize2D
    -- ^ See 20.1.7.3 "ext (Extents)" (p. 2846) or
    -- 20.5.2.14 "ext (Shape Extent)" (p. 3165)
  } deriving (Eq, Show, Generic)
instance NFData Transform2D

-- TODO: custGeom
data Geometry =
  PresetGeometry
  -- TODO: prst, avList
  -- currently uses "rect" with empty avList
  deriving (Eq, Show, Generic)
instance NFData Geometry

-- See 20.1.2.2.35 "spPr (Shape Properties)" (p. 2751)
data ShapeProperties = ShapeProperties
  { _spXfrm :: Maybe Transform2D
  , _spGeometry :: Maybe Geometry
  , _spFill :: Maybe FillProperties
  , _spOutline :: Maybe LineProperties
    -- TODO: bwMode, a_EG_EffectProperties, scene3d, sp3d, extLst
  } deriving (Eq, Show, Generic)
instance NFData ShapeProperties

-- | Specifies an outline style that can be applied to a number of
-- different objects such as shapes and text.
--
-- TODO: cap, cmpd, algn, a_EG_LineDashProperties,
--   a_EG_LineJoinProperties, headEnd, tailEnd, extLst
--
-- See 20.1.2.2.24 "ln (Outline)" (p. 2744)
data LineProperties = LineProperties
  { _lnFill :: Maybe FillProperties
  , _lnWidth :: Int
  -- ^ Specifies the width to be used for the underline stroke.  The
  -- value is in EMU, is greater of equal to 0 and maximum value is
  -- 20116800.
  } deriving (Eq, Show, Generic)
instance NFData LineProperties

-- | Color choice for some drawing element
--
-- TODO: scrgbClr, hslClr, sysClr, schemeClr, prstClr
--
-- See @EG_ColorChoice@ (p. 3996)
data ColorChoice =
  RgbColor Text
  -- ^ Specifies a color using the red, green, blue RGB color
  -- model. Red, green, and blue is expressed as sequence of hex
  -- digits, RRGGBB. A perceptual gamma of 2.2 is used.
  --
  -- See 20.1.2.3.32 "srgbClr (RGB Color Model - Hex Variant)" (p. 2773)
  deriving (Eq, Show, Generic)
instance NFData ColorChoice

-- TODO: gradFill, pattFill
data FillProperties =
  NoFill
  -- ^ See 20.1.8.44 "noFill (No Fill)" (p. 2872)
  | SolidFill (Maybe ColorChoice)
  -- ^ Solid fill
  -- See 20.1.8.54 "solidFill (Solid Fill)" (p. 2879)
  deriving (Eq, Show, Generic)
instance NFData FillProperties

-- | solid fill with color specified by hexadecimal RGB color
solidRgb :: Text -> FillProperties
solidRgb t = SolidFill . Just $ RgbColor t

makeLenses ''ShapeProperties

{-------------------------------------------------------------------------------
  Default instances
-------------------------------------------------------------------------------}

instance Default ShapeProperties where
    def = ShapeProperties Nothing Nothing Nothing Nothing

instance Default LineProperties where
  def = LineProperties Nothing 0

{-------------------------------------------------------------------------------
  Parsing
-------------------------------------------------------------------------------}

instance FromCursor TextBody where
  fromCursor cur = do
    cur' <- cur $/ element (a_ "bodyPr")
    _txbdRotation <- fromAttributeDef "rot" (Angle 0) cur'
    _txbdSpcFirstLastPara <- fromAttributeDef "spcFirstLastPara" False cur'
    _txbdVertOverflow <- fromAttributeDef "vertOverflow" TextVertOverflow cur'
    _txbdVertical <- fromAttributeDef "vert" TextVerticalHorz cur'
    _txbdWrap <- fromAttributeDef "wrap" TextWrapSquare cur'
    _txbdAnchor <- fromAttributeDef "anchor" TextAnchoringTop cur'
    _txbdAnchorCenter <- fromAttributeDef "anchorCtr" False cur'
    let _txbdParagraphs = cur $/ element (a_ "p") >=> fromCursor
    return TextBody {..}

instance FromCursor TextParagraph where
  fromCursor cur = do
    let _txpaDefCharProps =
          join . listToMaybe $
          cur $/ element (a_ "pPr") >=> maybeFromElement (a_ "defRPr")
        _txpaRuns = cur $/ element (a_ "r") >=> fromCursor
    return TextParagraph {..}

instance FromCursor TextCharacterProperties where
  fromCursor cur = do
    _txchBold <- fromAttributeDef "b" False cur
    _txchItalic <- fromAttributeDef "i" False cur
    _txchUnderline <- fromAttributeDef "u" False cur
    return TextCharacterProperties {..}

instance FromCursor TextRun where
  fromCursor cur = do
    _txrCharProps <- maybeFromElement (a_ "rPr") cur
    _txrText <- cur $/ element (a_ "t") &/ content
    return RegularRun {..}

-- See 20.1.10.3 "ST_Angle (Angle)" (p. 2912)
instance FromAttrVal Angle where
  fromAttrVal t = first Angle <$> fromAttrVal t

-- See 20.1.10.83 "ST_TextVertOverflowType (Text Vertical Overflow)" (p. 3083)
instance FromAttrVal TextVertOverflow where
  fromAttrVal "overflow" = readSuccess TextVertOverflow
  fromAttrVal "ellipsis" = readSuccess TextVertOverflowEllipsis
  fromAttrVal "clip" = readSuccess TextVertOverflowClip
  fromAttrVal t = invalidText "TextVertOverflow" t

instance FromAttrVal TextVertical where
  fromAttrVal "horz" = readSuccess TextVerticalHorz
  fromAttrVal "vert" = readSuccess TextVertical
  fromAttrVal "vert270" = readSuccess TextVertical270
  fromAttrVal "wordArtVert" = readSuccess TextVerticalWordArt
  fromAttrVal "eaVert" = readSuccess TextVerticalEA
  fromAttrVal "mongolianVert" = readSuccess TextVerticalMongolian
  fromAttrVal "wordArtVertRtl" = readSuccess TextVerticalWordArtRtl
  fromAttrVal t = invalidText "TextVertical" t

instance FromAttrVal TextWrap where
  fromAttrVal "none" = readSuccess TextWrapNone
  fromAttrVal "square" = readSuccess TextWrapSquare
  fromAttrVal t = invalidText "TextWrap" t

-- See 20.1.10.59 "ST_TextAnchoringType (Text Anchoring Types)" (p. 3058)
instance FromAttrVal TextAnchoring where
  fromAttrVal "t" = readSuccess TextAnchoringTop
  fromAttrVal "ctr" = readSuccess TextAnchoringCenter
  fromAttrVal "b" = readSuccess TextAnchoringBottom
  fromAttrVal "just" = readSuccess TextAnchoringJustified
  fromAttrVal "dist" = readSuccess TextAnchoringDistributed
  fromAttrVal t = invalidText "TextAnchoring" t

instance FromCursor ShapeProperties where
  fromCursor cur = do
    _spXfrm <- maybeFromElement (a_ "xfrm") cur
    let _spGeometry = listToMaybe $ cur $/ anyElement >=> fromCursor
        _spFill = listToMaybe $ cur $/ anyElement >=> fillPropsFromNode . node
    _spOutline <- maybeFromElement (a_ "ln") cur
    return ShapeProperties {..}

instance FromCursor Transform2D where
    fromCursor cur = do
        _trRot     <- fromAttributeDef "rot" (Angle 0) cur
        _trFlipH   <- fromAttributeDef "flipH" False cur
        _trFlipV   <- fromAttributeDef "flipV" False cur
        _trOffset  <- maybeFromElement (a_ "off") cur
        _trExtents <- maybeFromElement (a_ "ext") cur
        return Transform2D{..}

instance FromCursor Geometry where
    fromCursor = geometryFromNode . node

geometryFromNode :: Node -> [Geometry]
geometryFromNode n | n `nodeElNameIs` a_ "prstGeom" =
                         return PresetGeometry
                   | otherwise = fail "no matching geometry node"

instance FromCursor LineProperties where
    fromCursor cur = do
        let _lnFill = listToMaybe $ cur $/ anyElement >=> fromCursor
        _lnWidth <- fromAttributeDef "w" 0 cur
        return LineProperties{..}

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

instance FromCursor FillProperties where
    fromCursor = fillPropsFromNode . node

fillPropsFromNode :: Node -> [FillProperties]
fillPropsFromNode n
  | n `nodeElNameIs` a_ "noFill" = return NoFill
  | n `nodeElNameIs` a_ "solidFill" = do
    let color =
          listToMaybe $ fromNode n $/ anyElement >=> colorChoiceFromNode . node
    return $ SolidFill color
  | otherwise = fail "no matching line fill node"

colorChoiceFromNode :: Node -> [ColorChoice]
colorChoiceFromNode n
  | n `nodeElNameIs` a_ "srgbClr" = do
    val <- fromAttribute "val" $ fromNode n
    return $ RgbColor val
  | otherwise = fail "no matching color choice node"

coordinate :: MonadFail m => Text -> m Coordinate
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

instance ToElement TextBody where
  toElement nm TextBody {..} = elementListSimple nm (bodyPr : paragraphs)
    where
      bodyPr = leafElement (a_ "bodyPr") bodyPrAttrs
      bodyPrAttrs =
        catMaybes
          [ "rot" .=? justNonDef (Angle 0) _txbdRotation
          , "spcFirstLastPara" .=? justTrue _txbdSpcFirstLastPara
          , "vertOverflow" .=? justNonDef TextVertOverflow _txbdVertOverflow
          , "vert" .=? justNonDef TextVerticalHorz _txbdVertical
          , "wrap" .=? justNonDef TextWrapSquare _txbdWrap
          , "anchor" .=? justNonDef TextAnchoringTop _txbdAnchor
          , "anchorCtr" .=? justTrue _txbdAnchorCenter
          ]
      paragraphs = map (toElement (a_ "p")) _txbdParagraphs

instance ToElement TextParagraph where
  toElement nm TextParagraph {..} = elementListSimple nm elements
    where
      elements =
        case _txpaDefCharProps of
          Just props -> (defRPr props) : runs
          Nothing -> runs
      defRPr props =
        elementListSimple (a_ "pPr") [toElement (a_ "defRPr") props]
      runs = map (toElement (a_ "r")) _txpaRuns

instance ToElement TextCharacterProperties where
  toElement nm TextCharacterProperties {..} = leafElement nm attrs
    where
      attrs = ["b" .= _txchBold, "i" .= _txchItalic, "u" .= _txchUnderline]

instance ToElement TextRun where
  toElement nm RegularRun {..} = elementListSimple nm elements
    where
      elements =
        catMaybes
          [ toElement (a_ "rPr") <$> _txrCharProps
          , Just $ elementContent (a_ "t") _txrText
          ]

instance ToAttrVal TextVertOverflow where
  toAttrVal TextVertOverflow = "overflow"
  toAttrVal TextVertOverflowEllipsis = "ellipsis"
  toAttrVal TextVertOverflowClip = "clip"

instance ToAttrVal TextVertical where
  toAttrVal TextVerticalHorz = "horz"
  toAttrVal TextVertical = "vert"
  toAttrVal TextVertical270 = "vert270"
  toAttrVal TextVerticalWordArt = "wordArtVert"
  toAttrVal TextVerticalEA = "eaVert"
  toAttrVal TextVerticalMongolian = "mongolianVert"
  toAttrVal TextVerticalWordArtRtl = "wordArtVertRtl"

instance ToAttrVal TextWrap where
  toAttrVal TextWrapNone = "none"
  toAttrVal TextWrapSquare = "square"

-- See 20.1.10.59 "ST_TextAnchoringType (Text Anchoring Types)" (p. 3058)
instance ToAttrVal TextAnchoring where
  toAttrVal TextAnchoringTop = "t"
  toAttrVal TextAnchoringCenter = "ctr"
  toAttrVal TextAnchoringBottom = "b"
  toAttrVal TextAnchoringJustified = "just"
  toAttrVal TextAnchoringDistributed = "dist"

instance ToAttrVal Angle where
  toAttrVal (Angle x) = toAttrVal x

instance ToElement ShapeProperties where
    toElement nm ShapeProperties{..} = elementListSimple nm elements
      where
        elements = catMaybes [ toElement (a_ "xfrm") <$> _spXfrm
                             , geometryToElement <$> _spGeometry
                             , fillPropsToElement <$> _spFill
                             , toElement (a_ "ln")  <$> _spOutline ]

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

instance ToElement Transform2D where
    toElement nm Transform2D{..} = elementList nm attrs elements
      where
        attrs = catMaybes [ "rot"   .=? justNonDef (Angle 0) _trRot
                          , "flipH" .=? justTrue _trFlipH
                          , "flipV" .=? justTrue _trFlipV ]
        elements = catMaybes [ toElement (a_ "off") <$> _trOffset
                             , toElement (a_ "ext") <$> _trExtents ]

geometryToElement :: Geometry -> Element
geometryToElement PresetGeometry =
  leafElement (a_ "prstGeom") ["prst" .= ("rect" :: Text)]

instance ToElement LineProperties where
  toElement nm LineProperties {..} = elementList nm attrs elements
    where
      attrs = catMaybes ["w" .=? justNonDef 0 _lnWidth]
      elements = catMaybes [fillPropsToElement <$> _lnFill]

fillPropsToElement :: FillProperties -> Element
fillPropsToElement NoFill = emptyElement (a_ "noFill")
fillPropsToElement (SolidFill color) =
  elementListSimple (a_ "solidFill") $ catMaybes [colorChoiceToElement <$> color]

colorChoiceToElement :: ColorChoice -> Element
colorChoiceToElement (RgbColor color) =
  leafElement (a_ "srgbClr") ["val" .= color]

-- | Add DrawingML namespace to name
a_ :: Text -> Name
a_ x =
  Name {nameLocalName = x, nameNamespace = Just drawingNs, namePrefix = Just "a"}

drawingNs :: Text
drawingNs = "http://schemas.openxmlformats.org/drawingml/2006/main"
