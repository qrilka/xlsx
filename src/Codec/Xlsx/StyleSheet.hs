{-# LANGUAGE CPP                #-}
{-# LANGUAGE OverloadedStrings  #-}
{-# LANGUAGE RecordWildCards    #-}
{-# LANGUAGE TemplateHaskell    #-}
{-# OPTIONS_GHC -Wall #-}
-- | Support for writing (but not reading) style sheets
module Codec.Xlsx.StyleSheet (
    -- * The main two types
    StyleSheet(..)
  , CellXf(..)
  , minimalStyleSheet
    -- * Supporting record types
  , Alignment(..)
  , Border(..)
  , BorderStyle(..)
  , Color(..)
  , Fill(..)
  , FillPattern(..)
  , Font(..)
  , Protection(..)
    -- * Supporting enumerations
  , CellHorizontalAlignment(..)
  , CellVerticalAlignment(..)
  , FontFamily(..)
  , FontScheme(..)
  , FontVerticalAlignment(..)
  , LineStyle(..)
  , PatternType(..)
  , ReadingOrder(..)
    -- * Lenses
    -- ** StyleSheet
  , styleSheetBorders
  , styleSheetFonts
  , styleSheetFills
  , styleSheetCellXfs
    -- ** CellXf
  , cellXfApplyAlignment
  , cellXfApplyBorder
  , cellXfApplyFill
  , cellXfApplyFont
  , cellXfApplyNumberFormat
  , cellXfApplyProtection
  , cellXfBorderId
  , cellXfFillId
  , cellXfFontId
  , cellXfNumFmtId
  , cellXfPivotButton
  , cellXfQuotePrefix
  , cellXfId
  , cellXfAlignment
  , cellXfProtection
    -- ** Alignment
  , alignmentHorizontal
  , alignmentIndent
  , alignmentJustifyLastLine
  , alignmentReadingOrder
  , alignmentRelativeIndent
  , alignmentShrinkToFit
  , alignmentTextRotation
  , alignmentVertical
  , alignmentWrapText
    -- ** Border
  , borderDiagonalDown
  , borderDiagonalUp
  , borderOutline
  , borderBottom
  , borderDiagonal
  , borderEnd
  , borderHorizontal
  , borderStart
  , borderTop
  , borderVertical
  , borderLeft
  , borderRight
    -- ** BorderStyle
  , borderStyleColor
  , borderStyleLine
    -- ** Color
  , colorAutomatic
  , colorARGB
  , colorTheme
  , colorTint
    -- ** Fill
  , fillPattern
    -- ** FillPattern
  , fillPatternBgColor
  , fillPatternFgColor
  , fillPatternType
    -- ** Font
  , fontBold
  , fontCharset
  , fontColor
  , fontCondense
  , fontExtend
  , fontFamily
  , fontItalic
  , fontName
  , fontOutline
  , fontScheme
  , fontShadow
  , fontStrikeThrough
  , fontSize
  , fontUnderline
  , fontVertAlign
    -- ** Protection
  , protectionHidden
  , protectionLocked
  ) where

import Control.Lens hiding ((.=))
import Data.Default
import Data.Maybe (catMaybes)
import Text.XML
import qualified Data.Map as Map

#if !MIN_VERSION_base(4,8,0)
import Control.Applicative
#endif

import Codec.Xlsx.Writer.Internal

{-------------------------------------------------------------------------------
  The main types
-------------------------------------------------------------------------------}

-- | StyleSheet for an XML document
--
-- Relevant parts of the EMCA standard (2nd edition, part 1,
-- <http://www.ecma-international.org/publications/standards/Ecma-376.htm>),
-- page numbers refer to the page in the PDF rather than the page number as
-- printed on the page):
--
-- * Chapter 12, "SpreadsheetML" (p. 74)
--   In particular Section 12.3.20, "Styles Part" (p. 104)
-- * Chapter 18, "SpreadsheetML Reference Material" (p. 1692)
--   In particular Section 18.8, "Styles" (p. 1944) and Section 18.8.39
--   "styleSheet" (Style Sheet)" (p. 1993); it is the latter section that
--   specifies the top-level style sheet format.
--
-- TODO: the following child elements:
-- * cellStyles
-- * cellStyleXfs
-- * colors
-- * dxfs
-- * extLst
-- * numFmts
-- * tableStyles
--
-- NOTE: You will probably want to base your style sheet on 'minimalStyleSheet'.
-- See also 'renderStyleSheet'.
data StyleSheet = StyleSheet {
    -- | This element contains borders formatting information, specifying all
    -- border definitions for all cells in the workbook.
    --
    -- Section 18.8.5, "borders (Borders)" (p. 1951)
    _styleSheetBorders :: [Border]

    -- | Cell formats
    --
     -- This element contains the master formatting records (xf) which define the
    -- formatting applied to cells in this workbook. These records are the
    -- starting point for determining the formatting for a cell. Cells in the
    -- Sheet Part reference the xf records by zero-based index.
    --
    -- Section 18.8.10, "cellXfs (Cell Formats)" (p. 1956)
  , _styleSheetCellXfs :: [CellXf]

    -- | This element defines the cell fills portion of the Styles part,
    -- consisting of a sequence of fill records. A cell fill consists of a
    -- background color, foreground color, and pattern to be applied across the
    -- cell.
    --
    -- Section 18.8.21, "fills (Fills)" (p. 1963)
  , _styleSheetFills :: [Fill]

    -- | This element contains all font definitions for this workbook.
    --
    -- Section 18.8.23 "fonts (Fonts)" (p. 1965)
  , _styleSheetFonts :: [Font]
  }
  deriving (Show, Eq, Ord)

-- | Cell formatting
--
-- TODO: The @extLst@ field is currently unsupported.
--
-- Section 18.8.45 "xf (Format)" (p. 1999)
data CellXf = CellXf {
    -- | A boolean value indicating whether the alignment formatting specified
    -- for this xf should be applied.
    _cellXfApplyAlignment :: Maybe Bool

    -- | A boolean value indicating whether the border formatting specified for
    -- this xf should be applied.
  , _cellXfApplyBorder :: Maybe Bool

    -- | A boolean value indicating whether the fill formatting specified for
    -- this xf should be applied.
  , _cellXfApplyFill :: Maybe Bool

    -- | A boolean value indicating whether the font formatting specified for
    -- this xf should be applied.
  , _cellXfApplyFont :: Maybe Bool

    -- | A boolean value indicating whether the number formatting specified for
    -- this xf should be applied.
  , _cellXfApplyNumberFormat :: Maybe Bool

    -- | A boolean value indicating whether the protection formatting specified
    -- for this xf should be applied.
  , _cellXfApplyProtection :: Maybe Bool

    -- | Zero-based index of the border record used by this cell format.
    --
    -- (18.18.2, p. 2673).
  , _cellXfBorderId :: Maybe Int

    -- | Zero-based index of the fill record used by this cell format.
    --
    -- (18.18.30, p. 2693)
  , _cellXfFillId :: Maybe Int

    -- | Zero-based index of the font record used by this cell format.
    --
    -- An integer that represents a zero based index into the `styleSheetFonts`
    -- collection in the style sheet (18.18.32, p. 2694).
  , _cellXfFontId :: Maybe Int

    -- | Id of the number format (numFmt) record used by this cell format.
    --
    -- This simple type defines the identifier to a style sheet number format
    -- entry in CT_NumFmts. Number formats are written to the styles part
    -- (18.18.47, p. 2708). See also 18.8.31 (p. 1981) for more information on
    -- number formats.
    --
    -- TODO: The numFmts part of the style sheet is currently not implemented.
  , _cellXfNumFmtId :: Maybe Int

    -- | A boolean value indicating whether the cell rendering includes a pivot
    -- table dropdown button.
  , _cellXfPivotButton :: Maybe Bool

    -- | A boolean value indicating whether the text string in a cell should be
    -- prefixed by a single quote mark (e.g., 'text). In these cases, the quote
    -- is not stored in the Shared Strings Part.
  , _cellXfQuotePrefix :: Maybe Bool

    -- | For xf records contained in cellXfs this is the zero-based index of an
    -- xf record contained in cellStyleXfs corresponding to the cell style
    -- applied to the cell.
    --
    -- Not present for xf records contained in cellStyleXfs.
    --
    -- Used by xf records and cellStyle records to reference xf records defined
    -- in the cellStyleXfs collection. (18.18.10, p. 2678)
    -- TODO: the cellStyleXfs field of a style sheet not currently implemented.
  , _cellXfId :: Maybe Int

    -- | Formatting information pertaining to text alignment in cells. There are
    -- a variety of choices for how text is aligned both horizontally and
    -- vertically, as well as indentation settings, and so on.
  , _cellXfAlignment :: Maybe Alignment

    -- | Contains protection properties associated with the cell. Each cell has
    -- protection properties that can be set. The cell protection properties do
    -- not take effect unless the sheet has been protected.
  , _cellXfProtection :: Maybe Protection
  }
  deriving (Show, Eq, Ord)

{-------------------------------------------------------------------------------
  Supporting record types
-------------------------------------------------------------------------------}

-- | Alignment
--
-- See 18.8.1 "alignment (Alignment)" (p. 1944)
data Alignment = Alignment {
    -- | Specifies the type of horizontal alignment in cells.
    _alignmentHorizontal :: Maybe CellHorizontalAlignment

    -- | An integer value, where an increment of 1 represents 3 spaces.
    -- Indicates the number of spaces (of the normal style font) of indentation
    -- for text in a cell.
  , _alignmentIndent :: Maybe Int

    -- | A boolean value indicating if the cells justified or distributed
    -- alignment should be used on the last line of text. (This is typical for
    -- East Asian alignments but not typical in other contexts.)
  , _alignmentJustifyLastLine :: Maybe Bool

    -- | An integer value indicating whether the reading order
    -- (bidirectionality) of the cell is leftto- right, right-to-left, or
    -- context dependent.
  , _alignmentReadingOrder :: Maybe ReadingOrder

    -- | An integer value (used only in a dxf element) to indicate the
    -- additional number of spaces of indentation to adjust for text in a cell.
  , _alignmentRelativeIndent :: Maybe Int

    -- | A boolean value indicating if the displayed text in the cell should be
    -- shrunk to fit the cell width. Not applicable when a cell contains
    -- multiple lines of text.
  , _alignmentShrinkToFit :: Maybe Bool

    -- | Text rotation in cells. Expressed in degrees. Values range from 0 to
    -- 180. The first letter of the text is considered the center-point of the
    -- arc.
  , _alignmentTextRotation :: Maybe Int

    -- | Vertical alignment in cells.
  , _alignmentVertical :: Maybe CellVerticalAlignment

    -- | A boolean value indicating if the text in a cell should be line-wrapped
    -- within the cell.
  , _alignmentWrapText :: Maybe Bool
  }
  deriving (Show, Eq, Ord)

-- | Expresses a single set of cell border formats (left, right, top, bottom,
-- diagonal). Color is optional. When missing, 'automatic' is implied.
--
-- See 18.8.4 "border (Border)" (p. 1949)
data Border = Border {
    -- | A boolean value indicating if the cell's diagonal border includes a
    -- diagonal line, starting at the top left corner of the cell and moving
    -- down to the bottom right corner of the cell.
    _borderDiagonalDown :: Maybe Bool

    -- | A boolean value indicating if the cell's diagonal border includes a
    -- diagonal line, starting at the bottom left corner of the cell and moving
    -- up to the top right corner of the cell.
  , _borderDiagonalUp :: Maybe Bool

    -- | A boolean value indicating if left, right, top, and bottom borders
    -- should be applied only to outside borders of a cell range.
  , _borderOutline :: Maybe Bool

    -- | Bottom border
  , _borderBottom :: Maybe BorderStyle

    -- | Diagonal
  , _borderDiagonal :: Maybe BorderStyle

    -- | Trailing edge border
    --
    -- See also 'borderRight'
  , _borderEnd :: Maybe BorderStyle

    -- | Horizontal inner borders
  , _borderHorizontal :: Maybe BorderStyle

    -- | Leading edge border
    --
    -- See also 'borderLeft'
  , _borderStart :: Maybe BorderStyle

    -- | Top border
  , _borderTop :: Maybe BorderStyle

    -- | Vertical inner border
  , _borderVertical :: Maybe BorderStyle

    -- | Left border
    --
    -- NOTE: The spec does not formally support a 'left' border, but the
    -- examples do mention 'left'. See also 'borderStart'.
  , _borderLeft :: Maybe BorderStyle

    -- | Right border
    --
    -- NOTE: The spec does not formally support a 'right' border, but the
    -- examples do mention 'right'. See also 'borderEnd'.
  , _borderRight :: Maybe BorderStyle
  }
  deriving (Show, Eq, Ord)

-- | Border style
data BorderStyle = BorderStyle {
    _borderStyleColor :: Maybe Color
  , _borderStyleLine  :: Maybe LineStyle
  }
  deriving (Show, Eq, Ord)

-- | One of the colors associated with the data bar or color scale.
--
-- The 'indexed' attribute (used for backwards compatibility only) is not
-- modelled here.
--
-- See 18.3.1.15 "color (Data Bar Color)" (p. 1780)
data Color = Color {
    -- | A boolean value indicating the color is automatic and system color
    -- dependent.
    _colorAutomatic :: Maybe Bool

    -- | Standard Alpha Red Green Blue color value (ARGB).
    --
    -- This simple type's contents have a length of exactly 8 hexadecimal
    -- digit(s); see "18.18.86 ST_UnsignedIntHex (Hex Unsigned Integer)" (p.
    -- 2754).
  , _colorARGB :: Maybe String

    -- | A zero-based index into the <clrScheme> collection (20.1.6.2),
    -- referencing a particular <sysClr> or <srgbClr> value expressed in the
    -- Theme part.
  , _colorTheme :: Maybe Int

    -- | Specifies the tint value applied to the color.
    --
    -- If tint is supplied, then it is applied to the RGB value of the color to
    -- determine the final color applied.
    --
    -- The tint value is stored as a double from -1.0 .. 1.0, where -1.0 means
    -- 100% darken and 1.0 means 100% lighten. Also, 0.0 means no change.
  , _colorTint :: Maybe Double
  }
  deriving (Show, Eq, Ord)

-- | This element specifies fill formatting.
--
-- TODO: Gradient fills (18.8.4) are currently unsupported. If we add them,
-- then the spec says (@CT_Fill@, p. 4484), _eithre_ a gradient _or_ a solid
-- fill pattern should be specified.
--
-- Section 18.8.20, "fill (Fill)" (p. 1962)
data Fill = Fill {
    _fillPattern :: Maybe FillPattern
  }
  deriving (Show, Eq, Ord)

-- | This element is used to specify cell fill information for pattern and solid
-- color cell fills. For solid cell fills (no pattern), fgColor is used. For
-- cell fills with patterns specified, then the cell fill color is specified by
-- the bgColor element.
--
-- Section 18.8.32 "patternFill (Pattern)" (p. 1990)
data FillPattern = FillPattern {
    _fillPatternBgColor :: Maybe Color
  , _fillPatternFgColor :: Maybe Color
  , _fillPatternType    :: Maybe PatternType
  }
  deriving (Show, Eq, Ord)

-- | This element defines the properties for one of the fonts used in this
-- workbook.
--
-- Section 18.2.22 "font (Font)" (p. 1964)
data Font = Font {
    -- | Displays characters in bold face font style.
    _fontBold :: Maybe Bool

    -- | This element defines the font character set of this font.
    --
    -- This field is used in font creation and selection if a font of the given
    -- facename is not available on the system. Although it is not required to
    -- have around when resolving font facename, the information can be stored
    -- for when needed to help resolve which font face to use of all available
    -- fonts on a system.
    --
    -- Charset represents the basic set of characters associated with a font
    -- (that it can display), and roughly corresponds to the ANSI codepage
    -- (8-bit or DBCS) of that character set used by a given language. Given
    -- more common use of Unicode where many fonts support more than one of the
    -- traditional charset categories, and the use of font linking, using
    -- charset to resolve font name is less and less common, but still can be
    -- useful.
    --
    -- These are operating-system-dependent values.
    --
    -- Section 18.4.1 "charset (Character Set)" provides some example values.
  , _fontCharset :: Maybe Int

    -- | Color
  , _fontColor :: Maybe Color

    -- | Macintosh compatibility setting. Represents special word/character
    -- rendering on Macintosh, when this flag is set. The effect is to condense
    -- the text (squeeze it together). SpreadsheetML applications are not
    -- required to render according to this flag.
  , _fontCondense :: Maybe Bool

    -- | This element specifies a compatibility setting used for previous
    -- spreadsheet applications, resulting in special word/character rendering
    -- on those legacy applications, when this flag is set. The effect extends
    -- or stretches out the text. SpreadsheetML applications are not required to
    -- render according to this flag.
  , _fontExtend :: Maybe Bool

    -- | The font family this font belongs to. A font family is a set of fonts
    -- having common stroke width and serif characteristics. This is system
    -- level font information. The font name overrides when there are
    -- conflicting values.
  , _fontFamily :: Maybe FontFamily

    -- | Displays characters in italic font style. The italic style is defined
    -- by the font at a system level and is not specified by ECMA-376.
  , _fontItalic :: Maybe Bool

    -- | This element specifies the face name of this font.
    --
    -- A string representing the name of the font. If the font doesn't exist
    -- (because it isn't installed on the system), or the charset not supported
    -- by that font, then another font should be substituted.
    --
    -- The string length for this attribute shall be 0 to 31 characters.
  , _fontName :: Maybe String

    -- | This element displays only the inner and outer borders of each
    -- character. This is very similar to Bold in behavior.
  , _fontOutline :: Maybe Bool

    -- | Defines the font scheme, if any, to which this font belongs. When a
    -- font definition is part of a theme definition, then the font is
    -- categorized as either a major or minor font scheme component. When a new
    -- theme is chosen, every font that is part of a theme definition is updated
    -- to use the new major or minor font definition for that theme. Usually
    -- major fonts are used for styles like headings, and minor fonts are used
    -- for body and paragraph text.
  , _fontScheme :: Maybe FontScheme

    -- | Macintosh compatibility setting. Represents special word/character
    -- rendering on Macintosh, when this flag is set. The effect is to render a
    -- shadow behind, beneath and to the right of the text. SpreadsheetML
    -- applications are not required to render according to this flag.
  , _fontShadow :: Maybe Bool

    -- | This element draws a strikethrough line through the horizontal middle
    -- of the text.
  , _fontStrikeThrough :: Maybe Bool

    -- | This element represents the point size (1/72 of an inch) of the Latin
    -- and East Asian text.
  , _fontSize :: Maybe Double

    -- | This element represents the underline formatting style.
  , _fontUnderline :: Maybe Bool

    -- | This element adjusts the vertical position of the text relative to the
    -- text's default appearance for this run. It is used to get 'superscript'
    -- or 'subscript' texts, and shall reduce the font size (if a smaller size
    -- is available) accordingly.
  , _fontVertAlign :: Maybe FontVerticalAlignment
  }
  deriving (Show, Eq, Ord)

-- | Protection properties
--
-- Contains protection properties associated with the cell. Each cell has
-- protection properties that can be set. The cell protection properties do not
-- take effect unless the sheet has been protected.
--
-- Section 18.8.3, "protection (Protection Properties)", p. 1990
data Protection = Protection {
    _protectionHidden :: Maybe Bool
  , _protectionLocked :: Maybe Bool
  }
  deriving (Show, Eq, Ord)

{-------------------------------------------------------------------------------
  Enumerations
-------------------------------------------------------------------------------}

-- | Horizontal alignment in cells
--
-- See 18.18.40 "ST_HorizontalAlignment (Horizontal Alignment Type)" (p. 2698)
data CellHorizontalAlignment =
    CellHorizontalAlignmentCenter
  | CellHorizontalAlignmentCenterContinuous
  | CellHorizontalAlignmentDistributed
  | CellHorizontalAlignmentFill
  | CellHorizontalAlignmentGeneral
  | CellHorizontalAlignmentJustify
  | CellHorizontalAlignmentLeft
  | CellHorizontalAlignmentRight
  deriving (Show, Eq, Ord)

-- | Vertical alignment in cells
--
-- See 18.18.88 "ST_VerticalAlignment (Vertical Alignment Types)" (p. 2755)
data CellVerticalAlignment =
    CellVerticalAlignmentBottom
  | CellVerticalAlignmentCenter
  | CellVerticalAlignmentDistributed
  | CellVerticalAlignmentJustify
  | CellVerticalAlignmentTop
  deriving (Show, Eq, Ord)

-- | Font family
--
-- See 18.8.18 "family (Font Family)" (p. 1960)
-- and 17.18.30 "ST_FontFamily (Font Family Value)" (p. 1546)
data FontFamily =
    -- | Proportional font with serifs
    FontFamilyRoman

    -- | Proportional font without serifs
  | FontFamilySwiss

    -- | Monospace font with or without serifs
  | FontFamilyModern

    -- | Script font designed to mimic the appearance of handwriting
  | FontFamilyScript

    -- | Novelty font
  | FontFamilyDecorative
  deriving (Show, Eq, Ord)

-- | Font scheme
--
-- See 18.18.33 "ST_FontScheme (Font scheme Styles)" (p. 2694)
data FontScheme =
    -- | This font is the major font for this theme.
    FontSchemeMajor

    -- | This font is the minor font for this theme.
  | FontSchemeMinor

    -- | This font is not a theme font.
  | FontSchemeNone
  deriving (Show, Eq, Ord)

-- | Vertical alignment
--
-- See 22.9.2.17 "ST_VerticalAlignRun (Vertical Positioning Location)" (p. 4341)
data FontVerticalAlignment =
    FontVerticalAlignmentBaseline
  | FontVerticalAlignmentSubscript
  | FontVerticalAlignmentSuperscript
  deriving (Show, Eq, Ord)

data LineStyle =
    LineStyleDashDot
  | LineStyleDashDotDot
  | LineStyleDashed
  | LineStyleDotted
  | LineStyleDouble
  | LineStyleHair
  | LineStyleMedium
  | LineStyleMediumDashDot
  | LineStyleMediumDashDotDot
  | LineStyleMediumDashed
  | LineStyleNone
  | LineStyleSlantDashDot
  | LineStyleThick
  | LineStyleThin
  deriving (Show, Eq, Ord)

-- | Indicates the style of fill pattern being used for a cell format.
--
-- Section 18.18.55 "ST_PatternType (Pattern Type)" (p. 2713)
data PatternType =
    PatternTypeDarkDown
  | PatternTypeDarkGray
  | PatternTypeDarkGrid
  | PatternTypeDarkHorizontal
  | PatternTypeDarkTrellis
  | PatternTypeDarkUp
  | PatternTypeDarkVertical
  | PatternTypeGray0625
  | PatternTypeGray125
  | PatternTypeLightDown
  | PatternTypeLightGray
  | PatternTypeLightGrid
  | PatternTypeLightHorizontal
  | PatternTypeLightTrellis
  | PatternTypeLightUp
  | PatternTypeLightVertical
  | PatternTypeMediumGray
  | PatternTypeNone
  | PatternTypeSolid
  deriving (Show, Eq, Ord)

-- | Reading order
--
-- See 18.8.1 "alignment (Alignment)" (p. 1944, esp. p. 1945)
data ReadingOrder =
    ReadingOrderContextDependent
  | ReadingOrderLeftToRight
  | ReadingOrderRightToLeft
  deriving (Show, Eq, Ord)

{-------------------------------------------------------------------------------
  Lenses
-------------------------------------------------------------------------------}

makeLenses ''StyleSheet
makeLenses ''CellXf

makeLenses ''Alignment
makeLenses ''Border
makeLenses ''BorderStyle
makeLenses ''Color
makeLenses ''Fill
makeLenses ''FillPattern
makeLenses ''Font
makeLenses ''Protection

{-------------------------------------------------------------------------------
  Minimal stylesheet
-------------------------------------------------------------------------------}

-- | Minimal style sheet
--
-- Excel expects some minimal definitions in the stylesheet; you probably want
-- to define your own stylesheets based on this one.
--
-- This more-or-less follows the recommendations at
-- <http://stackoverflow.com/questions/26050708/minimal-style-sheet-for-excel-open-xml-with-dates>,
-- but with some additions based on experimental evidence.
minimalStyleSheet :: StyleSheet
minimalStyleSheet = def
    & styleSheetBorders .~ [defaultBorder]
    & styleSheetFonts   .~ [defaultFont]
    & styleSheetFills   .~ [fillNone, fillGray125]
    & styleSheetCellXfs .~ [defaultCellXf]
  where
    -- The 'Default' instance for 'Border' uses 'left' and 'right' rather than
    -- 'start' and 'end', because this is what Excel does (even though the spec
    -- says different)
    defaultBorder :: Border
    defaultBorder = def
      & borderBottom .~ Just def
      & borderTop    .~ Just def
      & borderLeft   .~ Just def
      & borderRight  .~ Just def

    defaultFont :: Font
    defaultFont = def
      & fontFamily .~ Just FontFamilySwiss
      & fontSize   .~ Just 11

    fillNone, fillGray125 :: Fill
    fillNone = def
      & fillPattern .~ Just (def & fillPatternType .~ Just PatternTypeNone)
    fillGray125 = def
      & fillPattern .~ Just (def & fillPatternType .~ Just PatternTypeGray125)

    defaultCellXf :: CellXf
    defaultCellXf = def
      & cellXfBorderId .~ Just 0
      & cellXfFillId   .~ Just 0
      & cellXfFontId   .~ Just 0

{-------------------------------------------------------------------------------
  Default instances
-------------------------------------------------------------------------------}

instance Default StyleSheet where
  def = StyleSheet {
      _styleSheetBorders = []
    , _styleSheetFonts   = []
    , _styleSheetFills   = []
    , _styleSheetCellXfs = []
    }

instance Default CellXf where
  def = CellXf {
      _cellXfApplyAlignment    = Nothing
    , _cellXfApplyBorder       = Nothing
    , _cellXfApplyFill         = Nothing
    , _cellXfApplyFont         = Nothing
    , _cellXfApplyNumberFormat = Nothing
    , _cellXfApplyProtection   = Nothing
    , _cellXfBorderId          = Nothing
    , _cellXfFillId            = Nothing
    , _cellXfFontId            = Nothing
    , _cellXfNumFmtId          = Nothing
    , _cellXfPivotButton       = Nothing
    , _cellXfQuotePrefix       = Nothing
    , _cellXfId                = Nothing
    , _cellXfAlignment         = Nothing
    , _cellXfProtection        = Nothing
    }

instance Default Alignment where
  def = Alignment {
     _alignmentHorizontal      = Nothing
   , _alignmentIndent          = Nothing
   , _alignmentJustifyLastLine = Nothing
   , _alignmentReadingOrder    = Nothing
   , _alignmentRelativeIndent  = Nothing
   , _alignmentShrinkToFit     = Nothing
   , _alignmentTextRotation    = Nothing
   , _alignmentVertical        = Nothing
   , _alignmentWrapText        = Nothing
   }

instance Default Border where
  def = Border {
      _borderDiagonalDown = Nothing
    , _borderDiagonalUp   = Nothing
    , _borderOutline      = Nothing
    , _borderBottom       = Nothing
    , _borderDiagonal     = Nothing
    , _borderEnd          = Nothing
    , _borderHorizontal   = Nothing
    , _borderStart        = Nothing
    , _borderTop          = Nothing
    , _borderVertical     = Nothing
    , _borderLeft         = Nothing
    , _borderRight        = Nothing
    }

instance Default BorderStyle where
  def = BorderStyle {
      _borderStyleColor = Nothing
    , _borderStyleLine  = Nothing
    }

instance Default Color where
  def = Color {
    _colorAutomatic = Nothing
  , _colorARGB      = Nothing
  , _colorTheme     = Nothing
  , _colorTint      = Nothing
  }

instance Default Fill where
  def = Fill {
      _fillPattern = Nothing
    }

instance Default FillPattern where
  def = FillPattern {
      _fillPatternBgColor = Nothing
    , _fillPatternFgColor = Nothing
    , _fillPatternType    = Nothing
    }

instance Default Font where
  def = Font {
      _fontBold          = Nothing
    , _fontCharset       = Nothing
    , _fontColor         = Nothing
    , _fontCondense      = Nothing
    , _fontExtend        = Nothing
    , _fontFamily        = Nothing
    , _fontItalic        = Nothing
    , _fontName          = Nothing
    , _fontOutline       = Nothing
    , _fontScheme        = Nothing
    , _fontShadow        = Nothing
    , _fontStrikeThrough = Nothing
    , _fontSize          = Nothing
    , _fontUnderline     = Nothing
    , _fontVertAlign     = Nothing
    }

instance Default Protection where
  def = Protection {
      _protectionHidden = Nothing
    , _protectionLocked = Nothing
    }

{-------------------------------------------------------------------------------
  Rendering record types

  NOTE: Excel is sensitive to the order of the child nodes, so we are careful
  to follow the XML schema here. We are also careful to follow the ordering
  for attributes, although this is actually pointless, as xml-conduit stores
  these as a Map, so we lose the ordering. But if we change representation,
  at least they are in the right order (hopefully) in the source code.
-------------------------------------------------------------------------------}

instance ToDocument StyleSheet where
  toDocument = documentFromElement "Stylesheet generated by xlsx"
             . toElement "styleSheet"

-- | See @CT_Stylesheet@, p. 4482
instance ToElement StyleSheet where
  toElement nm StyleSheet{..} = Element {
      elementName       = nm
    , elementAttributes = Map.empty
    , elementNodes      = map NodeElement $ [
         -- TODO: numFmts
         elementList "fonts"   $ map (toElement "font")   _styleSheetFonts
       , elementList "fills"   $ map (toElement "fill")   _styleSheetFills
       , elementList "borders" $ map (toElement "border") _styleSheetBorders
         -- TODO: cellStyleXfs
       , elementList "cellXfs" $ map (toElement "xf")     _styleSheetCellXfs
         -- TODO: cellStyles
         -- TODO: dxfs
         -- TODO: tableStyles
         -- TODO: colors
         -- TODO: extLst
       ]
    }

-- | See @CT_Xf@, p. 4486
instance ToElement CellXf where
  toElement nm CellXf{..} = Element {
      elementName       = nm
    , elementNodes      = map NodeElement . catMaybes $ [
          toElement "alignment"  <$> _cellXfAlignment
        , toElement "protection" <$> _cellXfProtection
          -- TODO: extLst
        ]
    , elementAttributes = Map.fromList . catMaybes $ [
          "numFmtId"          .=? _cellXfNumFmtId
        , "fontId"            .=? _cellXfFontId
        , "fillId"            .=? _cellXfFillId
        , "borderId"          .=? _cellXfBorderId
        , "xfId"              .=? _cellXfId
        , "quotePrefix"       .=? _cellXfQuotePrefix
        , "pivotButton"       .=? _cellXfPivotButton
        , "applyNumberFormat" .=? _cellXfApplyNumberFormat
        , "applyFont"         .=? _cellXfApplyFont
        , "applyFill"         .=? _cellXfApplyFill
        , "applyBorder"       .=? _cellXfApplyBorder
        , "applyAlignment"    .=? _cellXfApplyAlignment
        , "applyProtection"   .=? _cellXfApplyProtection
        ]
    }

-- | See @CT_CellAlignment@, p. 4482
instance ToElement Alignment where
  toElement nm Alignment{..} = Element {
      elementName       = nm
    , elementNodes      = []
    , elementAttributes = Map.fromList . catMaybes $ [
          "horizontal"      .=? _alignmentHorizontal
        , "vertical"        .=? _alignmentVertical
        , "textRotation"    .=? _alignmentTextRotation
        , "wrapText"        .=? _alignmentWrapText
        , "relativeIndent"  .=? _alignmentRelativeIndent
        , "indent"          .=? _alignmentIndent
        , "justifyLastLine" .=? _alignmentJustifyLastLine
        , "shrinkToFit"     .=? _alignmentShrinkToFit
        , "readingOrder"    .=? _alignmentReadingOrder
        ]
    }

-- | See @CT_Border@, p. 4483
instance ToElement Border where
  toElement nm Border{..} = Element {
      elementName       = nm
    , elementAttributes = Map.fromList . catMaybes $ [
          "diagonalUp"   .=? _borderDiagonalUp
        , "diagonalDown" .=? _borderDiagonalDown
        , "outline"      .=? _borderOutline
        ]
    , elementNodes      = map NodeElement . catMaybes $ [
          toElement "start"      <$> _borderStart
          -- Using position of @start@ for the not-per-spec @left@ element
        , toElement "left"       <$> _borderLeft
        , toElement "end"        <$> _borderEnd
          -- Using position of @end@ for the not-per-spec @right@ element
        , toElement "right"      <$> _borderRight
        , toElement "top"        <$> _borderTop
        , toElement "bottom"     <$> _borderBottom
        , toElement "diagonal"   <$> _borderDiagonal
        , toElement "vertical"   <$> _borderVertical
        , toElement "horizontal" <$> _borderHorizontal
        ]
    }

-- | See @CT_BorderPr@, p. 4483
instance ToElement BorderStyle where
  toElement nm BorderStyle{..} = Element {
      elementName       = nm
    , elementAttributes = Map.fromList . catMaybes $ [
          "style" .=? _borderStyleLine
        ]
    , elementNodes      = map NodeElement . catMaybes $ [
          toElement "color" <$> _borderStyleColor
        ]
    }

-- | See @CT_Color@, p. 4484
instance ToElement Color where
  toElement nm Color{..} = Element {
      elementName       = nm
    , elementNodes      = []
    , elementAttributes = Map.fromList . catMaybes $ [
          "auto"  .=? _colorAutomatic
        , "rgb"   .=? _colorARGB
        , "theme" .=? _colorTheme
        , "tint"  .=? _colorTint
        ]
    }

-- | See @CT_Fill@, p. 4484
instance ToElement Fill where
  toElement nm Fill{..} = Element {
      elementName       = nm
    , elementAttributes = Map.empty
    , elementNodes      = map NodeElement . catMaybes $ [
          toElement "patternFill" <$> _fillPattern
        ]
    }

-- | See @CT_PatternFill@, p. 4484
instance ToElement FillPattern where
  toElement nm FillPattern{..} = Element {
      elementName       = nm
    , elementAttributes = Map.fromList . catMaybes $ [
          "patternType" .=? _fillPatternType
        ]
    , elementNodes      = map NodeElement . catMaybes $ [
          toElement "fgColor" <$> _fillPatternFgColor
        , toElement "bgColor" <$> _fillPatternBgColor
        ]
    }

-- | See @CT_Font@, p. 4489
instance ToElement Font where
  toElement nm Font{..} = Element {
      elementName       = nm
    , elementAttributes = Map.empty -- all properties specified as child nodes
    , elementNodes      = map NodeElement . catMaybes $ [
          elementValue "name"      <$> _fontName
        , elementValue "charset"   <$> _fontCharset
        , elementValue "family"    <$> _fontFamily
        , elementValue "b"         <$> _fontBold
        , elementValue "i"         <$> _fontItalic
        , elementValue "strike"    <$> _fontStrikeThrough
        , elementValue "outline"   <$> _fontOutline
        , elementValue "shadow"    <$> _fontShadow
        , elementValue "condense"  <$> _fontCondense
        , elementValue "extend"    <$> _fontExtend
        , toElement    "color"     <$> _fontColor
        , elementValue "sz"        <$> _fontSize
        , elementValue "u"         <$> _fontUnderline
        , elementValue "vertAlign" <$> _fontVertAlign
        , elementValue "scheme"    <$> _fontScheme
        ]
    }

-- | See @CT_CellProtection@, p. 4484
instance ToElement Protection where
  toElement nm Protection{..} = Element {
      elementName       = nm
    , elementNodes      = []
    , elementAttributes = Map.fromList . catMaybes $ [
          "locked" .=? _protectionLocked
        , "hidden" .=? _protectionHidden
        ]
    }

{-------------------------------------------------------------------------------
  Rendering attribute values
-------------------------------------------------------------------------------}

instance ToAttrVal CellHorizontalAlignment where
  toAttrVal CellHorizontalAlignmentCenter           = "center"
  toAttrVal CellHorizontalAlignmentCenterContinuous = "centerContinuous"
  toAttrVal CellHorizontalAlignmentDistributed      = "distributed"
  toAttrVal CellHorizontalAlignmentFill             = "fill"
  toAttrVal CellHorizontalAlignmentGeneral          = "general"
  toAttrVal CellHorizontalAlignmentJustify          = "justify"
  toAttrVal CellHorizontalAlignmentLeft             = "left"
  toAttrVal CellHorizontalAlignmentRight            = "right"

instance ToAttrVal CellVerticalAlignment where
  toAttrVal CellVerticalAlignmentBottom      = "bottom"
  toAttrVal CellVerticalAlignmentCenter      = "center"
  toAttrVal CellVerticalAlignmentDistributed = "distributed"
  toAttrVal CellVerticalAlignmentJustify     = "justify"
  toAttrVal CellVerticalAlignmentTop         = "top"

instance ToAttrVal FontFamily where
  toAttrVal FontFamilyRoman      = "1"
  toAttrVal FontFamilySwiss      = "2"
  toAttrVal FontFamilyModern     = "3"
  toAttrVal FontFamilyScript     = "4"
  toAttrVal FontFamilyDecorative = "5"

instance ToAttrVal FontScheme where
  toAttrVal FontSchemeMajor = "major"
  toAttrVal FontSchemeMinor = "minor"
  toAttrVal FontSchemeNone  = "none"

instance ToAttrVal FontVerticalAlignment where
  toAttrVal FontVerticalAlignmentBaseline    = "baseline"
  toAttrVal FontVerticalAlignmentSubscript   = "subscript"
  toAttrVal FontVerticalAlignmentSuperscript = "superscript"

instance ToAttrVal LineStyle where
  toAttrVal LineStyleDashDot          = "dashDot"
  toAttrVal LineStyleDashDotDot       = "dashDotDot"
  toAttrVal LineStyleDashed           = "dashed"
  toAttrVal LineStyleDotted           = "dotted"
  toAttrVal LineStyleDouble           = "double"
  toAttrVal LineStyleHair             = "hair"
  toAttrVal LineStyleMedium           = "medium"
  toAttrVal LineStyleMediumDashDot    = "mediumDashDot"
  toAttrVal LineStyleMediumDashDotDot = "mediumDashDotDot"
  toAttrVal LineStyleMediumDashed     = "mediumDashed"
  toAttrVal LineStyleNone             = "none"
  toAttrVal LineStyleSlantDashDot     = "slantDashDot"
  toAttrVal LineStyleThick            = "thick"
  toAttrVal LineStyleThin             = "thin"

instance ToAttrVal PatternType where
  toAttrVal PatternTypeDarkDown        = "darkDown"
  toAttrVal PatternTypeDarkGray        = "darkGray"
  toAttrVal PatternTypeDarkGrid        = "darkGrid"
  toAttrVal PatternTypeDarkHorizontal  = "darkHorizontal"
  toAttrVal PatternTypeDarkTrellis     = "darkTrellis"
  toAttrVal PatternTypeDarkUp          = "darkUp"
  toAttrVal PatternTypeDarkVertical    = "darkVertical"
  toAttrVal PatternTypeGray0625        = "gray0625"
  toAttrVal PatternTypeGray125         = "gray125"
  toAttrVal PatternTypeLightDown       = "lightDown"
  toAttrVal PatternTypeLightGray       = "lightGray"
  toAttrVal PatternTypeLightGrid       = "lightGrid"
  toAttrVal PatternTypeLightHorizontal = "lightHorizontal"
  toAttrVal PatternTypeLightTrellis    = "lightTrellis"
  toAttrVal PatternTypeLightUp         = "lightUp"
  toAttrVal PatternTypeLightVertical   = "lightVertical"
  toAttrVal PatternTypeMediumGray      = "mediumGray"
  toAttrVal PatternTypeNone            = "none"
  toAttrVal PatternTypeSolid           = "solid"

instance ToAttrVal ReadingOrder where
  toAttrVal ReadingOrderContextDependent = "0"
  toAttrVal ReadingOrderLeftToRight      = "1"
  toAttrVal ReadingOrderRightToLeft      = "2"
