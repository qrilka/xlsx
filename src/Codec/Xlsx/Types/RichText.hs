{-# LANGUAGE CPP #-}
{-# LANGUAGE OverloadedStrings  #-}
{-# LANGUAGE RecordWildCards    #-}
{-# LANGUAGE TemplateHaskell    #-}
{-# LANGUAGE DeriveGeneric #-}
module Codec.Xlsx.Types.RichText (
    -- * Main types
    RichTextRun(..)
  , RunProperties(..)
  , applyRunProperties
    -- * Lenses
    -- ** RichTextRun
  , richTextRunProperties
  , richTextRunText
    -- ** RunProperties
  , runPropertiesBold
  , runPropertiesCharset
  , runPropertiesColor
  , runPropertiesCondense
  , runPropertiesExtend
  , runPropertiesFontFamily
  , runPropertiesItalic
  , runPropertiesOutline
  , runPropertiesFont
  , runPropertiesScheme
  , runPropertiesShadow
  , runPropertiesStrikeThrough
  , runPropertiesSize
  , runPropertiesUnderline
  , runPropertiesVertAlign
  ) where

import GHC.Generics (Generic)

#ifdef USE_MICROLENS
import Lens.Micro.TH (makeLenses)
#else
import Control.Lens hiding (element)
#endif
import Control.Monad
import Control.DeepSeq (NFData)
import Data.Default
import Data.Maybe (catMaybes)
import Data.Text (Text)
import Text.XML
import Text.XML.Cursor
import qualified Data.Map as Map

import Codec.Xlsx.Parser.Internal
import Codec.Xlsx.Types.StyleSheet
import Codec.Xlsx.Writer.Internal

-- | Rich Text Run
--
-- This element represents a run of rich text. A rich text run is a region of
-- text that share a common set of properties, such as formatting properties.
--
-- Section 18.4.4, "r (Rich Text Run)" (p. 1724)
data RichTextRun = RichTextRun {
    -- | This element represents a set of properties to apply to the contents of
    -- this rich text run.
    _richTextRunProperties :: Maybe RunProperties

    -- | This element represents the text content shown as part of a string.
    --
    -- NOTE: 'RichTextRun' elements with an empty text field will result in
    -- an error when opening the file in Excel.
    --
    -- Section 18.4.12, "t (Text)" (p. 1727)
  , _richTextRunText :: Text
  }
  deriving (Eq, Ord, Show, Generic)

instance NFData RichTextRun

-- | Run properties
--
-- Section 18.4.7, "rPr (Run Properties)" (p. 1725)
data RunProperties = RunProperties {
    -- | Displays characters in bold face font style.
    --
    -- Section 18.8.2, "b (Bold)" (p. 1757)
    _runPropertiesBold :: Maybe Bool

    -- | This element defines the font character set of this font.
    --
    -- Section 18.4.1, "charset (Character Set)" (p. 1721)
  , _runPropertiesCharset :: Maybe Int

    -- | One of the colors associated with the data bar or color scale.
    --
    -- Section 18.3.1.15, "color (Data Bar Color)" (p. 1608)
  , _runPropertiesColor :: Maybe Color

    -- | Macintosh compatibility setting. Represents special word/character
    -- rendering on Macintosh, when this flag is set. The effect is to condense
    -- the text (squeeze it together).
    --
    -- Section 18.8.12, "condense (Condense)" (p. 1764)
  , _runPropertiesCondense :: Maybe Bool

    -- | This element specifies a compatibility setting used for previous
    -- spreadsheet applications, resulting in special word/character rendering
    -- on those legacy applications, when this flag is set. The effect extends
    -- or stretches out the text.
    --
    -- Section 18.8.17, "extend (Extend)" (p. 1766)
  , _runPropertiesExtend :: Maybe Bool

    -- | The font family this font belongs to. A font family is a set of fonts
    -- having common stroke width and serif characteristics. This is system
    -- level font information. The font name overrides when there are
    -- conflicting values.
    --
    -- Section 18.8.18, "family (Font Family)" (p. 1766)
  , _runPropertiesFontFamily :: Maybe FontFamily

    -- | Displays characters in italic font style. The italic style is defined
    -- by the font at a system level and is not specified by ECMA-376.
    --
    -- Section 18.8.26, "i (Italic)" (p. 1773)
  , _runPropertiesItalic :: Maybe Bool

    -- | This element displays only the inner and outer borders of each
    -- character. This is very similar to Bold in behavior.
    --
    -- Section 18.4.2, "outline (Outline)" (p. 1722)
  , _runPropertiesOutline :: Maybe Bool

    -- | This element is a string representing the name of the font assigned to
    -- display this run.
    --
    -- Section 18.4.5, "rFont (Font)" (p. 1724)
  , _runPropertiesFont :: Maybe Text

    -- | Defines the font scheme, if any, to which this font belongs. When a
    -- font definition is part of a theme definition, then the font is
    -- categorized as either a major or minor font scheme component. When a new
    -- theme is chosen, every font that is part of a theme definition is updated
    -- to use the new major or minor font definition for that theme. Usually
    -- major fonts are used for styles like headings, and minor fonts are used
    -- for body and paragraph text.
    --
    -- Section 18.8.35, "scheme (Scheme)" (p. 1794)
  , _runPropertiesScheme :: Maybe FontScheme

    -- | Macintosh compatibility setting. Represents special word/character
    -- rendering on Macintosh, when this flag is set. The effect is to render a
    -- shadow behind, beneath and to the right of the text.
    --
    -- Section 18.8.36, "shadow (Shadow)" (p. 1795)
  , _runPropertiesShadow :: Maybe Bool

    -- | This element draws a strikethrough line through the horizontal middle
    -- of the text.
    --
    -- Section 18.4.10, "strike (Strike Through)" (p. 1726)
  , _runPropertiesStrikeThrough :: Maybe Bool

    -- | This element represents the point size (1/72 of an inch) of the Latin
    -- and East Asian text.
    --
    -- Section 18.4.11, "sz (Font Size)" (p. 1727)
  , _runPropertiesSize :: Maybe Double

    -- | This element represents the underline formatting style.
    --
    -- Section 18.4.13, "u (Underline)" (p. 1728)
  , _runPropertiesUnderline :: Maybe FontUnderline

    -- | This element adjusts the vertical position of the text relative to the
    -- text's default appearance for this run. It is used to get 'superscript'
    -- or 'subscript' texts, and shall reduce the font size (if a smaller size
    -- is available) accordingly.
    --
    -- Section 18.4.14, "vertAlign (Vertical Alignment)" (p. 1728)
  , _runPropertiesVertAlign :: Maybe FontVerticalAlignment
  }
  deriving (Eq, Ord, Show, Generic)

instance NFData RunProperties

{-------------------------------------------------------------------------------
  Lenses
-------------------------------------------------------------------------------}

makeLenses ''RichTextRun
makeLenses ''RunProperties

{-------------------------------------------------------------------------------
  Default instances
-------------------------------------------------------------------------------}

instance Default RichTextRun where
  def = RichTextRun {
      _richTextRunProperties = Nothing
    , _richTextRunText       = ""
    }

instance Default RunProperties where
   def = RunProperties {
      _runPropertiesBold          = Nothing
    , _runPropertiesCharset       = Nothing
    , _runPropertiesColor         = Nothing
    , _runPropertiesCondense      = Nothing
    , _runPropertiesExtend        = Nothing
    , _runPropertiesFontFamily    = Nothing
    , _runPropertiesItalic        = Nothing
    , _runPropertiesOutline       = Nothing
    , _runPropertiesFont          = Nothing
    , _runPropertiesScheme        = Nothing
    , _runPropertiesShadow        = Nothing
    , _runPropertiesStrikeThrough = Nothing
    , _runPropertiesSize          = Nothing
    , _runPropertiesUnderline     = Nothing
    , _runPropertiesVertAlign     = Nothing
    }

{-------------------------------------------------------------------------------
  Rendering
-------------------------------------------------------------------------------}

-- | See @CT_RElt@, p. 3903
instance ToElement RichTextRun where
  toElement nm RichTextRun{..} = Element {
      elementName       = nm
    , elementAttributes = Map.empty
    , elementNodes      = map NodeElement . catMaybes $ [
          toElement "rPr" <$> _richTextRunProperties
        , Just $ elementContentPreserved "t" _richTextRunText
        ]
    }

-- | See @CT_RPrElt@, p. 3903
instance ToElement RunProperties where
  toElement nm RunProperties{..} = Element {
      elementName       = nm
    , elementAttributes = Map.empty
    , elementNodes      = map NodeElement . catMaybes $ [
          elementValue "rFont"     <$> _runPropertiesFont
        , elementValue "charset"   <$> _runPropertiesCharset
        , elementValue "family"    <$> _runPropertiesFontFamily
        , elementValue "b"         <$> _runPropertiesBold
        , elementValue "i"         <$> _runPropertiesItalic
        , elementValue "strike"    <$> _runPropertiesStrikeThrough
        , elementValue "outline"   <$> _runPropertiesOutline
        , elementValue "shadow"    <$> _runPropertiesShadow
        , elementValue "condense"  <$> _runPropertiesCondense
        , elementValue "extend"    <$> _runPropertiesExtend
        , toElement    "color"     <$> _runPropertiesColor
        , elementValue "sz"        <$> _runPropertiesSize
        , elementValueDef "u" FontUnderlineSingle
                                   <$> _runPropertiesUnderline
        , elementValue "vertAlign" <$> _runPropertiesVertAlign
        , elementValue "scheme"    <$> _runPropertiesScheme
        ]
    }

{-------------------------------------------------------------------------------
  Parsing
-------------------------------------------------------------------------------}

-- | See @CT_RElt@, p. 3903
instance FromCursor RichTextRun where
  fromCursor cur = do
    _richTextRunText <- cur $/ element (n_ "t") &/ content
    _richTextRunProperties <- maybeFromElement (n_ "rPr") cur
    return RichTextRun{..}

instance FromXenoNode RichTextRun where
  fromXenoNode root = do
    (prNode, tNode) <- collectChildren root $ (,) <$> maybeChild "rPr" <*> requireChild "t"
    _richTextRunProperties <- mapM fromXenoNode prNode
    _richTextRunText <- contentX tNode
    return RichTextRun {..}

-- | See @CT_RPrElt@, p. 3903
instance FromCursor RunProperties where
  fromCursor cur = do
    _runPropertiesFont          <- maybeElementValue (n_ "rFont") cur
    _runPropertiesCharset       <- maybeElementValue (n_ "charset") cur
    _runPropertiesFontFamily    <- maybeElementValue (n_ "family") cur
    _runPropertiesBold          <- maybeBoolElementValue (n_ "b") cur
    _runPropertiesItalic        <- maybeBoolElementValue (n_ "i") cur
    _runPropertiesStrikeThrough <- maybeBoolElementValue (n_ "strike") cur
    _runPropertiesOutline       <- maybeBoolElementValue (n_ "outline") cur
    _runPropertiesShadow        <- maybeBoolElementValue (n_ "shadow") cur
    _runPropertiesCondense      <- maybeBoolElementValue (n_ "condense") cur
    _runPropertiesExtend        <- maybeBoolElementValue (n_ "extend") cur
    _runPropertiesColor         <- maybeFromElement  (n_ "color") cur
    _runPropertiesSize          <- maybeElementValue (n_ "sz") cur
    _runPropertiesUnderline     <- maybeElementValueDef (n_ "u") FontUnderlineSingle cur
    _runPropertiesVertAlign     <- maybeElementValue (n_ "vertAlign") cur
    _runPropertiesScheme        <- maybeElementValue (n_ "scheme") cur
    return RunProperties{..}

instance FromXenoNode RunProperties where
  fromXenoNode root = collectChildren root $ do
    _runPropertiesFont          <- maybeElementVal "rFont"
    _runPropertiesCharset       <- maybeElementVal "charset"
    _runPropertiesFontFamily    <- maybeElementVal "family"
    _runPropertiesBold          <- maybeElementVal "b"
    _runPropertiesItalic        <- maybeElementVal "i"
    _runPropertiesStrikeThrough <- maybeElementVal "strike"
    _runPropertiesOutline       <- maybeElementVal "outline"
    _runPropertiesShadow        <- maybeElementVal "shadow"
    _runPropertiesCondense      <- maybeElementVal "condense"
    _runPropertiesExtend        <- maybeElementVal "extend"
    _runPropertiesColor         <- maybeFromChild "color"
    _runPropertiesSize          <- maybeElementVal "sz"
    _runPropertiesUnderline     <- maybeElementVal "u"
    _runPropertiesVertAlign     <- maybeElementVal "vertAlign"
    _runPropertiesScheme        <- maybeElementVal "scheme"
    return RunProperties{..}

{-------------------------------------------------------------------------------
  Applying formatting
-------------------------------------------------------------------------------}

#if (MIN_VERSION_base(4,11,0))
instance Semigroup RunProperties where
  a <> b = RunProperties {
      _runPropertiesBold          = override _runPropertiesBold
    , _runPropertiesCharset       = override _runPropertiesCharset
    , _runPropertiesColor         = override _runPropertiesColor
    , _runPropertiesCondense      = override _runPropertiesCondense
    , _runPropertiesExtend        = override _runPropertiesExtend
    , _runPropertiesFontFamily    = override _runPropertiesFontFamily
    , _runPropertiesItalic        = override _runPropertiesItalic
    , _runPropertiesOutline       = override _runPropertiesOutline
    , _runPropertiesFont          = override _runPropertiesFont
    , _runPropertiesScheme        = override _runPropertiesScheme
    , _runPropertiesShadow        = override _runPropertiesShadow
    , _runPropertiesStrikeThrough = override _runPropertiesStrikeThrough
    , _runPropertiesSize          = override _runPropertiesSize
    , _runPropertiesUnderline     = override _runPropertiesUnderline
    , _runPropertiesVertAlign     = override _runPropertiesVertAlign
    }
    where
      override :: (RunProperties -> Maybe x) -> Maybe x
      override f = f b `mplus` f a

#endif

-- | The 'Monoid' instance for 'RunProperties' is biased: later properties
-- override earlier ones.
instance Monoid RunProperties where
  mempty = def
  a `mappend` b = RunProperties {
      _runPropertiesBold          = override _runPropertiesBold
    , _runPropertiesCharset       = override _runPropertiesCharset
    , _runPropertiesColor         = override _runPropertiesColor
    , _runPropertiesCondense      = override _runPropertiesCondense
    , _runPropertiesExtend        = override _runPropertiesExtend
    , _runPropertiesFontFamily    = override _runPropertiesFontFamily
    , _runPropertiesItalic        = override _runPropertiesItalic
    , _runPropertiesOutline       = override _runPropertiesOutline
    , _runPropertiesFont          = override _runPropertiesFont
    , _runPropertiesScheme        = override _runPropertiesScheme
    , _runPropertiesShadow        = override _runPropertiesShadow
    , _runPropertiesStrikeThrough = override _runPropertiesStrikeThrough
    , _runPropertiesSize          = override _runPropertiesSize
    , _runPropertiesUnderline     = override _runPropertiesUnderline
    , _runPropertiesVertAlign     = override _runPropertiesVertAlign
    }
    where
      override :: (RunProperties -> Maybe x) -> Maybe x
      override f = f b `mplus` f a

-- | Apply properties to a 'RichTextRun'
--
-- If the 'RichTextRun' specifies its own properties, then these overrule the
-- properties specified here. For example, adding @bold@ to a 'RichTextRun'
-- which is already @italic@ will make the 'RichTextRun' both @bold and @italic@
-- but adding it to one that that is explicitly _not_ bold will leave the
-- 'RichTextRun' unchanged.
applyRunProperties :: RunProperties -> RichTextRun -> RichTextRun
applyRunProperties p (RichTextRun Nothing   t) = RichTextRun (Just p) t
applyRunProperties p (RichTextRun (Just p') t) = RichTextRun (Just (p `mappend` p')) t
