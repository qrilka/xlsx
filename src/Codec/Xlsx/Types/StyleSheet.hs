{-# LANGUAGE CPP               #-}
{-# LANGUAGE OverloadedStrings #-}
{-# LANGUAGE RecordWildCards   #-}
{-# LANGUAGE TemplateHaskell   #-}
{-# LANGUAGE DeriveGeneric #-}
-- | Support for writing (but not reading) style sheets
module Codec.Xlsx.Types.StyleSheet (
    -- * The main two types
    StyleSheet(..)
  , CellXf(..)
  , minimalStyleSheet
    -- * Supporting record types
  , Alignment(..)
  , Border(..)
  , BorderStyle(..)
  , Color(..)
  , Dxf(..)
  , Fill(..)
  , FillPattern(..)
  , Font(..)
  , NumberFormat(..)
  , ImpliedNumberFormat (..)
  , NumFmt
  , Protection(..)
    -- * Supporting enumerations
  , CellHorizontalAlignment(..)
  , CellVerticalAlignment(..)
  , FontFamily(..)
  , FontScheme(..)
  , FontUnderline(..)
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
  , styleSheetDxfs
  , styleSheetNumFmts
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
    -- ** Dxf
  , dxfAlignment
  , dxfBorder
  , dxfFill
  , dxfFont
  , dxfProtection
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
    -- * Helpers
    -- ** Number formats
  , fmtDecimals
  , fmtDecimalsZeroes
  , stdNumberFormatId
  , idToStdNumberFormat
  , firstUserNumFmtId
    -- ** Colors
  , colorToARGB
  ) where

import Control.Lens hiding (element, elements, (.=))
import Data.Default
import Data.Map.Strict (Map)
import qualified Data.Map.Strict as M
import Data.Maybe (catMaybes)
import Data.Monoid
import Data.Text (Text)
import qualified Data.Text as T
import GHC.Generics (Generic)
import Text.XML
import Text.XML.Cursor

#if !MIN_VERSION_base(4,8,0)
import Control.Applicative
#endif

import Codec.Xlsx.Parser.Internal
import Codec.Xlsx.Types.Internal.NumFmtPair
import Codec.Xlsx.Writer.Internal

{-------------------------------------------------------------------------------
  The main types
-------------------------------------------------------------------------------}

-- | StyleSheet for an XML document
--
-- Relevant parts of the EMCA standard (4th edition, part 1,
-- <http://www.ecma-international.org/publications/standards/Ecma-376.htm>),
-- page numbers refer to the page in the PDF rather than the page number as
-- printed on the page):
--
-- * Chapter 12, "SpreadsheetML" (p. 74)
--   In particular Section 12.3.20, "Styles Part" (p. 104)
-- * Chapter 18, "SpreadsheetML Reference Material" (p. 1528)
--   In particular Section 18.8, "Styles" (p. 1754) and Section 18.8.39
--   "styleSheet" (Style Sheet)" (p. 1796); it is the latter section that
--   specifies the top-level style sheet format.
--
-- TODO: the following child elements:
--
-- * cellStyles
-- * cellStyleXfs
-- * colors
-- * extLst
-- * tableStyles
--
-- NOTE: You will probably want to base your style sheet on 'minimalStyleSheet'.
-- See also:
--
-- * 'Codec.Xlsx.Types.renderStyleSheet' to translate a 'StyleSheet' to 'Styles'
-- * 'Codec.Xlsx.Formatted.formatted' for a higher level interface.
-- * 'Codec.Xlsx.Types.parseStyleSheet' to translate a raw 'StyleSheet' into 'Styles'
data StyleSheet = StyleSheet
    { _styleSheetBorders :: [Border]
    -- ^ This element contains borders formatting information, specifying all
    -- border definitions for all cells in the workbook.
    --
    -- Section 18.8.5, "borders (Borders)" (p. 1760)

    , _styleSheetCellXfs :: [CellXf]
    -- ^ Cell formats
    --
    -- This element contains the master formatting records (xf) which define the
    -- formatting applied to cells in this workbook. These records are the
    -- starting point for determining the formatting for a cell. Cells in the
    -- Sheet Part reference the xf records by zero-based index.
    --
    -- Section 18.8.10, "cellXfs (Cell Formats)" (p. 1764)

    , _styleSheetFills   :: [Fill]
    -- ^ This element defines the cell fills portion of the Styles part,
    -- consisting of a sequence of fill records. A cell fill consists of a
    -- background color, foreground color, and pattern to be applied across the
    -- cell.
    --
    -- Section 18.8.21, "fills (Fills)" (p. 1768)

    , _styleSheetFonts   :: [Font]
    -- ^ This element contains all font definitions for this workbook.
    --
    -- Section 18.8.23 "fonts (Fonts)" (p. 1769)

    , _styleSheetDxfs    :: [Dxf]
    -- ^ Differential formatting
    --
    -- This element contains the master differential formatting records (dxf's)
    -- which define formatting for all non-cell formatting in this workbook.
    -- Whereas xf records fully specify a particular aspect of formatting (e.g.,
    -- cell borders) by referencing those formatting definitions elsewhere in
    -- the Styles part, dxf records specify incremental (or differential) aspects
    -- of formatting directly inline within the dxf element. The dxf formatting
    -- is to be applied on top of or in addition to any formatting already
    -- present on the object using the dxf record.
    --
    -- Section 18.8.15, "dxfs (Formats)" (p. 1765)

    , _styleSheetNumFmts :: Map Int NumFmt
    -- ^ Number formats
    --
    -- This element contains custom number formats defined in this style sheet
    --
    -- Section 18.8.31, "numFmts (Number Formats)" (p. 1784)
    } deriving (Eq, Ord, Show, Generic)

-- | Cell formatting
--
-- TODO: The @extLst@ field is currently unsupported.
--
-- Section 18.8.45 "xf (Format)" (p. 1800)
data CellXf = CellXf {
    -- | A boolean value indicating whether the alignment formatting specified
    -- for this xf should be applied.
    _cellXfApplyAlignment    :: Maybe Bool

    -- | A boolean value indicating whether the border formatting specified for
    -- this xf should be applied.
  , _cellXfApplyBorder       :: Maybe Bool

    -- | A boolean value indicating whether the fill formatting specified for
    -- this xf should be applied.
  , _cellXfApplyFill         :: Maybe Bool

    -- | A boolean value indicating whether the font formatting specified for
    -- this xf should be applied.
  , _cellXfApplyFont         :: Maybe Bool

    -- | A boolean value indicating whether the number formatting specified for
    -- this xf should be applied.
  , _cellXfApplyNumberFormat :: Maybe Bool

    -- | A boolean value indicating whether the protection formatting specified
    -- for this xf should be applied.
  , _cellXfApplyProtection   :: Maybe Bool

    -- | Zero-based index of the border record used by this cell format.
    --
    -- (18.18.2, p. 2437).
  , _cellXfBorderId          :: Maybe Int

    -- | Zero-based index of the fill record used by this cell format.
    --
    -- (18.18.30, p. 2455)
  , _cellXfFillId            :: Maybe Int

    -- | Zero-based index of the font record used by this cell format.
    --
    -- An integer that represents a zero based index into the `styleSheetFonts`
    -- collection in the style sheet (18.18.32, p. 2456).
  , _cellXfFontId            :: Maybe Int

    -- | Id of the number format (numFmt) record used by this cell format.
    --
    -- This simple type defines the identifier to a style sheet number format
    -- entry in CT_NumFmts. Number formats are written to the styles part
    -- (18.18.47, p. 2468). See also 18.8.31 (p. 1784) for more information on
    -- number formats.
    --
  , _cellXfNumFmtId          :: Maybe Int

    -- | A boolean value indicating whether the cell rendering includes a pivot
    -- table dropdown button.
  , _cellXfPivotButton       :: Maybe Bool

    -- | A boolean value indicating whether the text string in a cell should be
    -- prefixed by a single quote mark (e.g., 'text). In these cases, the quote
    -- is not stored in the Shared Strings Part.
  , _cellXfQuotePrefix       :: Maybe Bool

    -- | For xf records contained in cellXfs this is the zero-based index of an
    -- xf record contained in cellStyleXfs corresponding to the cell style
    -- applied to the cell.
    --
    -- Not present for xf records contained in cellStyleXfs.
    --
    -- Used by xf records and cellStyle records to reference xf records defined
    -- in the cellStyleXfs collection. (18.18.10, p. 2442)
    -- TODO: the cellStyleXfs field of a style sheet not currently implemented.
  , _cellXfId                :: Maybe Int

    -- | Formatting information pertaining to text alignment in cells. There are
    -- a variety of choices for how text is aligned both horizontally and
    -- vertically, as well as indentation settings, and so on.
  , _cellXfAlignment         :: Maybe Alignment

    -- | Contains protection properties associated with the cell. Each cell has
    -- protection properties that can be set. The cell protection properties do
    -- not take effect unless the sheet has been protected.
  , _cellXfProtection        :: Maybe Protection
  }
  deriving (Eq, Ord, Show, Generic)

{-------------------------------------------------------------------------------
  Supporting record types
-------------------------------------------------------------------------------}

-- | Alignment
--
-- See 18.8.1 "alignment (Alignment)" (p. 1754)
data Alignment = Alignment {
    -- | Specifies the type of horizontal alignment in cells.
    _alignmentHorizontal      :: Maybe CellHorizontalAlignment

    -- | An integer value, where an increment of 1 represents 3 spaces.
    -- Indicates the number of spaces (of the normal style font) of indentation
    -- for text in a cell.
  , _alignmentIndent          :: Maybe Int

    -- | A boolean value indicating if the cells justified or distributed
    -- alignment should be used on the last line of text. (This is typical for
    -- East Asian alignments but not typical in other contexts.)
  , _alignmentJustifyLastLine :: Maybe Bool

    -- | An integer value indicating whether the reading order
    -- (bidirectionality) of the cell is leftto- right, right-to-left, or
    -- context dependent.
  , _alignmentReadingOrder    :: Maybe ReadingOrder

    -- | An integer value (used only in a dxf element) to indicate the
    -- additional number of spaces of indentation to adjust for text in a cell.
  , _alignmentRelativeIndent  :: Maybe Int

    -- | A boolean value indicating if the displayed text in the cell should be
    -- shrunk to fit the cell width. Not applicable when a cell contains
    -- multiple lines of text.
  , _alignmentShrinkToFit     :: Maybe Bool

    -- | Text rotation in cells. Expressed in degrees. Values range from 0 to
    -- 180. The first letter of the text is considered the center-point of the
    -- arc.
  , _alignmentTextRotation    :: Maybe Int

    -- | Vertical alignment in cells.
  , _alignmentVertical        :: Maybe CellVerticalAlignment

    -- | A boolean value indicating if the text in a cell should be line-wrapped
    -- within the cell.
  , _alignmentWrapText        :: Maybe Bool
  }
  deriving (Eq, Ord, Show, Generic)

-- | Expresses a single set of cell border formats (left, right, top, bottom,
-- diagonal). Color is optional. When missing, 'automatic' is implied.
--
-- See 18.8.4 "border (Border)" (p. 1759)
data Border = Border {
    -- | A boolean value indicating if the cell's diagonal border includes a
    -- diagonal line, starting at the top left corner of the cell and moving
    -- down to the bottom right corner of the cell.
    _borderDiagonalDown :: Maybe Bool

    -- | A boolean value indicating if the cell's diagonal border includes a
    -- diagonal line, starting at the bottom left corner of the cell and moving
    -- up to the top right corner of the cell.
  , _borderDiagonalUp   :: Maybe Bool

    -- | A boolean value indicating if left, right, top, and bottom borders
    -- should be applied only to outside borders of a cell range.
  , _borderOutline      :: Maybe Bool

    -- | Bottom border
  , _borderBottom       :: Maybe BorderStyle

    -- | Diagonal
  , _borderDiagonal     :: Maybe BorderStyle

    -- | Trailing edge border
    --
    -- See also 'borderRight'
  , _borderEnd          :: Maybe BorderStyle

    -- | Horizontal inner borders
  , _borderHorizontal   :: Maybe BorderStyle

    -- | Left border
    --
    -- NOTE: The spec does not formally list a 'left' border element, but the
    -- examples do mention 'left' and the scheme contains it too. See also 'borderStart'.
  , _borderLeft         :: Maybe BorderStyle

    -- | Right border
    --
    -- NOTE: The spec does not formally list a 'right' border element, but the
    -- examples do mention 'right' and the scheme contains it too. See also 'borderEnd'.
  , _borderRight        :: Maybe BorderStyle

    -- | Leading edge border
    --
    -- See also 'borderLeft'
  , _borderStart        :: Maybe BorderStyle

    -- | Top border
  , _borderTop          :: Maybe BorderStyle

    -- | Vertical inner border
  , _borderVertical     :: Maybe BorderStyle
  }
  deriving (Eq, Ord, Show, Generic)

-- | Border style
-- See @CT_BorderPr@ (p. 3934)
data BorderStyle = BorderStyle {
    _borderStyleColor :: Maybe Color
  , _borderStyleLine  :: Maybe LineStyle
  }
  deriving (Eq, Ord, Show, Generic)

-- | One of the colors associated with the data bar or color scale.
--
-- The 'indexed' attribute (used for backwards compatibility only) is not
-- modelled here.
--
-- See 18.3.1.15 "color (Data Bar Color)" (p. 1608)
data Color = Color {
    -- | A boolean value indicating the color is automatic and system color
    -- dependent.
    _colorAutomatic :: Maybe Bool

    -- | Standard Alpha Red Green Blue color value (ARGB).
    --
    -- This simple type's contents have a length of exactly 8 hexadecimal
    -- digit(s); see "18.18.86 ST_UnsignedIntHex (Hex Unsigned Integer)" (p.
    -- 2511).
  , _colorARGB      :: Maybe Text

    -- | A zero-based index into the <clrScheme> collection (20.1.6.2),
    -- referencing a particular <sysClr> or <srgbClr> value expressed in the
    -- Theme part.
  , _colorTheme     :: Maybe Int

    -- | Specifies the tint value applied to the color.
    --
    -- If tint is supplied, then it is applied to the RGB value of the color to
    -- determine the final color applied.
    --
    -- The tint value is stored as a double from -1.0 .. 1.0, where -1.0 means
    -- 100% darken and 1.0 means 100% lighten. Also, 0.0 means no change.
  , _colorTint      :: Maybe Double
  }
  deriving (Eq, Ord, Show, Generic)

-- | Converts a color name to its ARGB value; 657 available color names
colorToARGB :: Text -> Maybe Text
colorToARGB color
  | color == "white" = Just "FFFFFFFF"
  | color == "aliceblue" = Just "FFF0F8FF"
  | color == "antiquewhite" = Just "FFFAEBD7"
  | color == "antiquewhite1" = Just "FFFFEFDB"
  | color == "antiquewhite2" = Just "FFEEDFCC"
  | color == "antiquewhite3" = Just "FFCDC0B0"
  | color == "antiquewhite4" = Just "FF8B8378"
  | color == "aquamarine" = Just "FF7FFFD4"
  | color == "aquamarine1" = Just "FF7FFFD4"
  | color == "aquamarine2" = Just "FF76EEC6"
  | color == "aquamarine3" = Just "FF66CDAA"
  | color == "aquamarine4" = Just "FF458B74"
  | color == "azure" = Just "FFF0FFFF"
  | color == "azure1" = Just "FFF0FFFF"
  | color == "azure2" = Just "FFE0EEEE"
  | color == "azure3" = Just "FFC1CDCD"
  | color == "azure4" = Just "FF838B8B"
  | color == "beige" = Just "FFF5F5DC"
  | color == "bisque" = Just "FFFFE4C4"
  | color == "bisque1" = Just "FFFFE4C4"
  | color == "bisque2" = Just "FFEED5B7"
  | color == "bisque3" = Just "FFCDB79E"
  | color == "bisque4" = Just "FF8B7D6B"
  | color == "black" = Just "FF000000"
  | color == "blanchedalmond" = Just "FFFFEBCD"
  | color == "blue" = Just "FF0000FF"
  | color == "blue1" = Just "FF0000FF"
  | color == "blue2" = Just "FF0000EE"
  | color == "blue3" = Just "FF0000CD"
  | color == "blue4" = Just "FF00008B"
  | color == "blueviolet" = Just "FF8A2BE2"
  | color == "brown" = Just "FFA52A2A"
  | color == "brown1" = Just "FFFF4040"
  | color == "brown2" = Just "FFEE3B3B"
  | color == "brown3" = Just "FFCD3333"
  | color == "brown4" = Just "FF8B2323"
  | color == "burlywood" = Just "FFDEB887"
  | color == "burlywood1" = Just "FFFFD39B"
  | color == "burlywood2" = Just "FFEEC591"
  | color == "burlywood3" = Just "FFCDAA7D"
  | color == "burlywood4" = Just "FF8B7355"
  | color == "cadetblue" = Just "FF5F9EA0"
  | color == "cadetblue1" = Just "FF98F5FF"
  | color == "cadetblue2" = Just "FF8EE5EE"
  | color == "cadetblue3" = Just "FF7AC5CD"
  | color == "cadetblue4" = Just "FF53868B"
  | color == "chartreuse" = Just "FF7FFF00"
  | color == "chartreuse1" = Just "FF7FFF00"
  | color == "chartreuse2" = Just "FF76EE00"
  | color == "chartreuse3" = Just "FF66CD00"
  | color == "chartreuse4" = Just "FF458B00"
  | color == "chocolate" = Just "FFD2691E"
  | color == "chocolate1" = Just "FFFF7F24"
  | color == "chocolate2" = Just "FFEE7621"
  | color == "chocolate3" = Just "FFCD661D"
  | color == "chocolate4" = Just "FF8B4513"
  | color == "coral" = Just "FFFF7F50"
  | color == "coral1" = Just "FFFF7256"
  | color == "coral2" = Just "FFEE6A50"
  | color == "coral3" = Just "FFCD5B45"
  | color == "coral4" = Just "FF8B3E2F"
  | color == "cornflowerblue" = Just "FF6495ED"
  | color == "cornsilk" = Just "FFFFF8DC"
  | color == "cornsilk1" = Just "FFFFF8DC"
  | color == "cornsilk2" = Just "FFEEE8CD"
  | color == "cornsilk3" = Just "FFCDC8B1"
  | color == "cornsilk4" = Just "FF8B8878"
  | color == "cyan" = Just "FF00FFFF"
  | color == "cyan1" = Just "FF00FFFF"
  | color == "cyan2" = Just "FF00EEEE"
  | color == "cyan3" = Just "FF00CDCD"
  | color == "cyan4" = Just "FF008B8B"
  | color == "darkblue" = Just "FF00008B"
  | color == "darkcyan" = Just "FF008B8B"
  | color == "darkgoldenrod" = Just "FFB8860B"
  | color == "darkgoldenrod1" = Just "FFFFB90F"
  | color == "darkgoldenrod2" = Just "FFEEAD0E"
  | color == "darkgoldenrod3" = Just "FFCD950C"
  | color == "darkgoldenrod4" = Just "FF8B6508"
  | color == "darkgray" = Just "FFA9A9A9"
  | color == "darkgreen" = Just "FF006400"
  | color == "darkgrey" = Just "FFA9A9A9"
  | color == "darkkhaki" = Just "FFBDB76B"
  | color == "darkmagenta" = Just "FF8B008B"
  | color == "darkolivegreen" = Just "FF556B2F"
  | color == "darkolivegreen1" = Just "FFCAFF70"
  | color == "darkolivegreen2" = Just "FFBCEE68"
  | color == "darkolivegreen3" = Just "FFA2CD5A"
  | color == "darkolivegreen4" = Just "FF6E8B3D"
  | color == "darkorange" = Just "FFFF8C00"
  | color == "darkorange1" = Just "FFFF7F00"
  | color == "darkorange2" = Just "FFEE7600"
  | color == "darkorange3" = Just "FFCD6600"
  | color == "darkorange4" = Just "FF8B4500"
  | color == "darkorchid" = Just "FF9932CC"
  | color == "darkorchid1" = Just "FFBF3EFF"
  | color == "darkorchid2" = Just "FFB23AEE"
  | color == "darkorchid3" = Just "FF9A32CD"
  | color == "darkorchid4" = Just "FF68228B"
  | color == "darkred" = Just "FF8B0000"
  | color == "darksalmon" = Just "FFE9967A"
  | color == "darkseagreen" = Just "FF8FBC8F"
  | color == "darkseagreen1" = Just "FFC1FFC1"
  | color == "darkseagreen2" = Just "FFB4EEB4"
  | color == "darkseagreen3" = Just "FF9BCD9B"
  | color == "darkseagreen4" = Just "FF698B69"
  | color == "darkslateblue" = Just "FF483D8B"
  | color == "darkslategray" = Just "FF2F4F4F"
  | color == "darkslategray1" = Just "FF97FFFF"
  | color == "darkslategray2" = Just "FF8DEEEE"
  | color == "darkslategray3" = Just "FF79CDCD"
  | color == "darkslategray4" = Just "FF528B8B"
  | color == "darkslategrey" = Just "FF2F4F4F"
  | color == "darkturquoise" = Just "FF00CED1"
  | color == "darkviolet" = Just "FF9400D3"
  | color == "deeppink" = Just "FFFF1493"
  | color == "deeppink1" = Just "FFFF1493"
  | color == "deeppink2" = Just "FFEE1289"
  | color == "deeppink3" = Just "FFCD1076"
  | color == "deeppink4" = Just "FF8B0A50"
  | color == "deepskyblue" = Just "FF00BFFF"
  | color == "deepskyblue1" = Just "FF00BFFF"
  | color == "deepskyblue2" = Just "FF00B2EE"
  | color == "deepskyblue3" = Just "FF009ACD"
  | color == "deepskyblue4" = Just "FF00688B"
  | color == "dimgray" = Just "FF696969"
  | color == "dimgrey" = Just "FF696969"
  | color == "dodgerblue" = Just "FF1E90FF"
  | color == "dodgerblue1" = Just "FF1E90FF"
  | color == "dodgerblue2" = Just "FF1C86EE"
  | color == "dodgerblue3" = Just "FF1874CD"
  | color == "dodgerblue4" = Just "FF104E8B"
  | color == "firebrick" = Just "FFB22222"
  | color == "firebrick1" = Just "FFFF3030"
  | color == "firebrick2" = Just "FFEE2C2C"
  | color == "firebrick3" = Just "FFCD2626"
  | color == "firebrick4" = Just "FF8B1A1A"
  | color == "floralwhite" = Just "FFFFFAF0"
  | color == "forestgreen" = Just "FF228B22"
  | color == "gainsboro" = Just "FFDCDCDC"
  | color == "ghostwhite" = Just "FFF8F8FF"
  | color == "gold" = Just "FFFFD700"
  | color == "gold1" = Just "FFFFD700"
  | color == "gold2" = Just "FFEEC900"
  | color == "gold3" = Just "FFCDAD00"
  | color == "gold4" = Just "FF8B7500"
  | color == "goldenrod" = Just "FFDAA520"
  | color == "goldenrod1" = Just "FFFFC125"
  | color == "goldenrod2" = Just "FFEEB422"
  | color == "goldenrod3" = Just "FFCD9B1D"
  | color == "goldenrod4" = Just "FF8B6914"
  | color == "gray" = Just "FFBEBEBE"
  | color == "gray0" = Just "FF000000"
  | color == "gray1" = Just "FF030303"
  | color == "gray2" = Just "FF050505"
  | color == "gray3" = Just "FF080808"
  | color == "gray4" = Just "FF0A0A0A"
  | color == "gray5" = Just "FF0D0D0D"
  | color == "gray6" = Just "FF0F0F0F"
  | color == "gray7" = Just "FF121212"
  | color == "gray8" = Just "FF141414"
  | color == "gray9" = Just "FF171717"
  | color == "gray10" = Just "FF1A1A1A"
  | color == "gray11" = Just "FF1C1C1C"
  | color == "gray12" = Just "FF1F1F1F"
  | color == "gray13" = Just "FF212121"
  | color == "gray14" = Just "FF242424"
  | color == "gray15" = Just "FF262626"
  | color == "gray16" = Just "FF292929"
  | color == "gray17" = Just "FF2B2B2B"
  | color == "gray18" = Just "FF2E2E2E"
  | color == "gray19" = Just "FF303030"
  | color == "gray20" = Just "FF333333"
  | color == "gray21" = Just "FF363636"
  | color == "gray22" = Just "FF383838"
  | color == "gray23" = Just "FF3B3B3B"
  | color == "gray24" = Just "FF3D3D3D"
  | color == "gray25" = Just "FF404040"
  | color == "gray26" = Just "FF424242"
  | color == "gray27" = Just "FF454545"
  | color == "gray28" = Just "FF474747"
  | color == "gray29" = Just "FF4A4A4A"
  | color == "gray30" = Just "FF4D4D4D"
  | color == "gray31" = Just "FF4F4F4F"
  | color == "gray32" = Just "FF525252"
  | color == "gray33" = Just "FF545454"
  | color == "gray34" = Just "FF575757"
  | color == "gray35" = Just "FF595959"
  | color == "gray36" = Just "FF5C5C5C"
  | color == "gray37" = Just "FF5E5E5E"
  | color == "gray38" = Just "FF616161"
  | color == "gray39" = Just "FF636363"
  | color == "gray40" = Just "FF666666"
  | color == "gray41" = Just "FF696969"
  | color == "gray42" = Just "FF6B6B6B"
  | color == "gray43" = Just "FF6E6E6E"
  | color == "gray44" = Just "FF707070"
  | color == "gray45" = Just "FF737373"
  | color == "gray46" = Just "FF757575"
  | color == "gray47" = Just "FF787878"
  | color == "gray48" = Just "FF7A7A7A"
  | color == "gray49" = Just "FF7D7D7D"
  | color == "gray50" = Just "FF7F7F7F"
  | color == "gray51" = Just "FF828282"
  | color == "gray52" = Just "FF858585"
  | color == "gray53" = Just "FF878787"
  | color == "gray54" = Just "FF8A8A8A"
  | color == "gray55" = Just "FF8C8C8C"
  | color == "gray56" = Just "FF8F8F8F"
  | color == "gray57" = Just "FF919191"
  | color == "gray58" = Just "FF949494"
  | color == "gray59" = Just "FF969696"
  | color == "gray60" = Just "FF999999"
  | color == "gray61" = Just "FF9C9C9C"
  | color == "gray62" = Just "FF9E9E9E"
  | color == "gray63" = Just "FFA1A1A1"
  | color == "gray64" = Just "FFA3A3A3"
  | color == "gray65" = Just "FFA6A6A6"
  | color == "gray66" = Just "FFA8A8A8"
  | color == "gray67" = Just "FFABABAB"
  | color == "gray68" = Just "FFADADAD"
  | color == "gray69" = Just "FFB0B0B0"
  | color == "gray70" = Just "FFB3B3B3"
  | color == "gray71" = Just "FFB5B5B5"
  | color == "gray72" = Just "FFB8B8B8"
  | color == "gray73" = Just "FFBABABA"
  | color == "gray74" = Just "FFBDBDBD"
  | color == "gray75" = Just "FFBFBFBF"
  | color == "gray76" = Just "FFC2C2C2"
  | color == "gray77" = Just "FFC4C4C4"
  | color == "gray78" = Just "FFC7C7C7"
  | color == "gray79" = Just "FFC9C9C9"
  | color == "gray80" = Just "FFCCCCCC"
  | color == "gray81" = Just "FFCFCFCF"
  | color == "gray82" = Just "FFD1D1D1"
  | color == "gray83" = Just "FFD4D4D4"
  | color == "gray84" = Just "FFD6D6D6"
  | color == "gray85" = Just "FFD9D9D9"
  | color == "gray86" = Just "FFDBDBDB"
  | color == "gray87" = Just "FFDEDEDE"
  | color == "gray88" = Just "FFE0E0E0"
  | color == "gray89" = Just "FFE3E3E3"
  | color == "gray90" = Just "FFE5E5E5"
  | color == "gray91" = Just "FFE8E8E8"
  | color == "gray92" = Just "FFEBEBEB"
  | color == "gray93" = Just "FFEDEDED"
  | color == "gray94" = Just "FFF0F0F0"
  | color == "gray95" = Just "FFF2F2F2"
  | color == "gray96" = Just "FFF5F5F5"
  | color == "gray97" = Just "FFF7F7F7"
  | color == "gray98" = Just "FFFAFAFA"
  | color == "gray99" = Just "FFFCFCFC"
  | color == "gray100" = Just "FFFFFFFF"
  | color == "green" = Just "FF00FF00"
  | color == "green1" = Just "FF00FF00"
  | color == "green2" = Just "FF00EE00"
  | color == "green3" = Just "FF00CD00"
  | color == "green4" = Just "FF008B00"
  | color == "greenyellow" = Just "FFADFF2F"
  | color == "grey" = Just "FFBEBEBE"
  | color == "grey0" = Just "FF000000"
  | color == "grey1" = Just "FF030303"
  | color == "grey2" = Just "FF050505"
  | color == "grey3" = Just "FF080808"
  | color == "grey4" = Just "FF0A0A0A"
  | color == "grey5" = Just "FF0D0D0D"
  | color == "grey6" = Just "FF0F0F0F"
  | color == "grey7" = Just "FF121212"
  | color == "grey8" = Just "FF141414"
  | color == "grey9" = Just "FF171717"
  | color == "grey10" = Just "FF1A1A1A"
  | color == "grey11" = Just "FF1C1C1C"
  | color == "grey12" = Just "FF1F1F1F"
  | color == "grey13" = Just "FF212121"
  | color == "grey14" = Just "FF242424"
  | color == "grey15" = Just "FF262626"
  | color == "grey16" = Just "FF292929"
  | color == "grey17" = Just "FF2B2B2B"
  | color == "grey18" = Just "FF2E2E2E"
  | color == "grey19" = Just "FF303030"
  | color == "grey20" = Just "FF333333"
  | color == "grey21" = Just "FF363636"
  | color == "grey22" = Just "FF383838"
  | color == "grey23" = Just "FF3B3B3B"
  | color == "grey24" = Just "FF3D3D3D"
  | color == "grey25" = Just "FF404040"
  | color == "grey26" = Just "FF424242"
  | color == "grey27" = Just "FF454545"
  | color == "grey28" = Just "FF474747"
  | color == "grey29" = Just "FF4A4A4A"
  | color == "grey30" = Just "FF4D4D4D"
  | color == "grey31" = Just "FF4F4F4F"
  | color == "grey32" = Just "FF525252"
  | color == "grey33" = Just "FF545454"
  | color == "grey34" = Just "FF575757"
  | color == "grey35" = Just "FF595959"
  | color == "grey36" = Just "FF5C5C5C"
  | color == "grey37" = Just "FF5E5E5E"
  | color == "grey38" = Just "FF616161"
  | color == "grey39" = Just "FF636363"
  | color == "grey40" = Just "FF666666"
  | color == "grey41" = Just "FF696969"
  | color == "grey42" = Just "FF6B6B6B"
  | color == "grey43" = Just "FF6E6E6E"
  | color == "grey44" = Just "FF707070"
  | color == "grey45" = Just "FF737373"
  | color == "grey46" = Just "FF757575"
  | color == "grey47" = Just "FF787878"
  | color == "grey48" = Just "FF7A7A7A"
  | color == "grey49" = Just "FF7D7D7D"
  | color == "grey50" = Just "FF7F7F7F"
  | color == "grey51" = Just "FF828282"
  | color == "grey52" = Just "FF858585"
  | color == "grey53" = Just "FF878787"
  | color == "grey54" = Just "FF8A8A8A"
  | color == "grey55" = Just "FF8C8C8C"
  | color == "grey56" = Just "FF8F8F8F"
  | color == "grey57" = Just "FF919191"
  | color == "grey58" = Just "FF949494"
  | color == "grey59" = Just "FF969696"
  | color == "grey60" = Just "FF999999"
  | color == "grey61" = Just "FF9C9C9C"
  | color == "grey62" = Just "FF9E9E9E"
  | color == "grey63" = Just "FFA1A1A1"
  | color == "grey64" = Just "FFA3A3A3"
  | color == "grey65" = Just "FFA6A6A6"
  | color == "grey66" = Just "FFA8A8A8"
  | color == "grey67" = Just "FFABABAB"
  | color == "grey68" = Just "FFADADAD"
  | color == "grey69" = Just "FFB0B0B0"
  | color == "grey70" = Just "FFB3B3B3"
  | color == "grey71" = Just "FFB5B5B5"
  | color == "grey72" = Just "FFB8B8B8"
  | color == "grey73" = Just "FFBABABA"
  | color == "grey74" = Just "FFBDBDBD"
  | color == "grey75" = Just "FFBFBFBF"
  | color == "grey76" = Just "FFC2C2C2"
  | color == "grey77" = Just "FFC4C4C4"
  | color == "grey78" = Just "FFC7C7C7"
  | color == "grey79" = Just "FFC9C9C9"
  | color == "grey80" = Just "FFCCCCCC"
  | color == "grey81" = Just "FFCFCFCF"
  | color == "grey82" = Just "FFD1D1D1"
  | color == "grey83" = Just "FFD4D4D4"
  | color == "grey84" = Just "FFD6D6D6"
  | color == "grey85" = Just "FFD9D9D9"
  | color == "grey86" = Just "FFDBDBDB"
  | color == "grey87" = Just "FFDEDEDE"
  | color == "grey88" = Just "FFE0E0E0"
  | color == "grey89" = Just "FFE3E3E3"
  | color == "grey90" = Just "FFE5E5E5"
  | color == "grey91" = Just "FFE8E8E8"
  | color == "grey92" = Just "FFEBEBEB"
  | color == "grey93" = Just "FFEDEDED"
  | color == "grey94" = Just "FFF0F0F0"
  | color == "grey95" = Just "FFF2F2F2"
  | color == "grey96" = Just "FFF5F5F5"
  | color == "grey97" = Just "FFF7F7F7"
  | color == "grey98" = Just "FFFAFAFA"
  | color == "grey99" = Just "FFFCFCFC"
  | color == "grey100" = Just "FFFFFFFF"
  | color == "honeydew" = Just "FFF0FFF0"
  | color == "honeydew1" = Just "FFF0FFF0"
  | color == "honeydew2" = Just "FFE0EEE0"
  | color == "honeydew3" = Just "FFC1CDC1"
  | color == "honeydew4" = Just "FF838B83"
  | color == "hotpink" = Just "FFFF69B4"
  | color == "hotpink1" = Just "FFFF6EB4"
  | color == "hotpink2" = Just "FFEE6AA7"
  | color == "hotpink3" = Just "FFCD6090"
  | color == "hotpink4" = Just "FF8B3A62"
  | color == "indianred" = Just "FFCD5C5C"
  | color == "indianred1" = Just "FFFF6A6A"
  | color == "indianred2" = Just "FFEE6363"
  | color == "indianred3" = Just "FFCD5555"
  | color == "indianred4" = Just "FF8B3A3A"
  | color == "ivory" = Just "FFFFFFF0"
  | color == "ivory1" = Just "FFFFFFF0"
  | color == "ivory2" = Just "FFEEEEE0"
  | color == "ivory3" = Just "FFCDCDC1"
  | color == "ivory4" = Just "FF8B8B83"
  | color == "khaki" = Just "FFF0E68C"
  | color == "khaki1" = Just "FFFFF68F"
  | color == "khaki2" = Just "FFEEE685"
  | color == "khaki3" = Just "FFCDC673"
  | color == "khaki4" = Just "FF8B864E"
  | color == "lavender" = Just "FFE6E6FA"
  | color == "lavenderblush" = Just "FFFFF0F5"
  | color == "lavenderblush1" = Just "FFFFF0F5"
  | color == "lavenderblush2" = Just "FFEEE0E5"
  | color == "lavenderblush3" = Just "FFCDC1C5"
  | color == "lavenderblush4" = Just "FF8B8386"
  | color == "lawngreen" = Just "FF7CFC00"
  | color == "lemonchiffon" = Just "FFFFFACD"
  | color == "lemonchiffon1" = Just "FFFFFACD"
  | color == "lemonchiffon2" = Just "FFEEE9BF"
  | color == "lemonchiffon3" = Just "FFCDC9A5"
  | color == "lemonchiffon4" = Just "FF8B8970"
  | color == "lightblue" = Just "FFADD8E6"
  | color == "lightblue1" = Just "FFBFEFFF"
  | color == "lightblue2" = Just "FFB2DFEE"
  | color == "lightblue3" = Just "FF9AC0CD"
  | color == "lightblue4" = Just "FF68838B"
  | color == "lightcoral" = Just "FFF08080"
  | color == "lightcyan" = Just "FFE0FFFF"
  | color == "lightcyan1" = Just "FFE0FFFF"
  | color == "lightcyan2" = Just "FFD1EEEE"
  | color == "lightcyan3" = Just "FFB4CDCD"
  | color == "lightcyan4" = Just "FF7A8B8B"
  | color == "lightgoldenrod" = Just "FFEEDD82"
  | color == "lightgoldenrod1" = Just "FFFFEC8B"
  | color == "lightgoldenrod2" = Just "FFEEDC82"
  | color == "lightgoldenrod3" = Just "FFCDBE70"
  | color == "lightgoldenrod4" = Just "FF8B814C"
  | color == "lightgoldenrodyellow" = Just "FFFAFAD2"
  | color == "lightgray" = Just "FFD3D3D3"
  | color == "lightgreen" = Just "FF90EE90"
  | color == "lightgrey" = Just "FFD3D3D3"
  | color == "lightpink" = Just "FFFFB6C1"
  | color == "lightpink1" = Just "FFFFAEB9"
  | color == "lightpink2" = Just "FFEEA2AD"
  | color == "lightpink3" = Just "FFCD8C95"
  | color == "lightpink4" = Just "FF8B5F65"
  | color == "lightsalmon" = Just "FFFFA07A"
  | color == "lightsalmon1" = Just "FFFFA07A"
  | color == "lightsalmon2" = Just "FFEE9572"
  | color == "lightsalmon3" = Just "FFCD8162"
  | color == "lightsalmon4" = Just "FF8B5742"
  | color == "lightseagreen" = Just "FF20B2AA"
  | color == "lightskyblue" = Just "FF87CEFA"
  | color == "lightskyblue1" = Just "FFB0E2FF"
  | color == "lightskyblue2" = Just "FFA4D3EE"
  | color == "lightskyblue3" = Just "FF8DB6CD"
  | color == "lightskyblue4" = Just "FF607B8B"
  | color == "lightslateblue" = Just "FF8470FF"
  | color == "lightslategray" = Just "FF778899"
  | color == "lightslategrey" = Just "FF778899"
  | color == "lightsteelblue" = Just "FFB0C4DE"
  | color == "lightsteelblue1" = Just "FFCAE1FF"
  | color == "lightsteelblue2" = Just "FFBCD2EE"
  | color == "lightsteelblue3" = Just "FFA2B5CD"
  | color == "lightsteelblue4" = Just "FF6E7B8B"
  | color == "lightyellow" = Just "FFFFFFE0"
  | color == "lightyellow1" = Just "FFFFFFE0"
  | color == "lightyellow2" = Just "FFEEEED1"
  | color == "lightyellow3" = Just "FFCDCDB4"
  | color == "lightyellow4" = Just "FF8B8B7A"
  | color == "limegreen" = Just "FF32CD32"
  | color == "linen" = Just "FFFAF0E6"
  | color == "magenta" = Just "FFFF00FF"
  | color == "magenta1" = Just "FFFF00FF"
  | color == "magenta2" = Just "FFEE00EE"
  | color == "magenta3" = Just "FFCD00CD"
  | color == "magenta4" = Just "FF8B008B"
  | color == "maroon" = Just "FFB03060"
  | color == "maroon1" = Just "FFFF34B3"
  | color == "maroon2" = Just "FFEE30A7"
  | color == "maroon3" = Just "FFCD2990"
  | color == "maroon4" = Just "FF8B1C62"
  | color == "mediumaquamarine" = Just "FF66CDAA"
  | color == "mediumblue" = Just "FF0000CD"
  | color == "mediumorchid" = Just "FFBA55D3"
  | color == "mediumorchid1" = Just "FFE066FF"
  | color == "mediumorchid2" = Just "FFD15FEE"
  | color == "mediumorchid3" = Just "FFB452CD"
  | color == "mediumorchid4" = Just "FF7A378B"
  | color == "mediumpurple" = Just "FF9370DB"
  | color == "mediumpurple1" = Just "FFAB82FF"
  | color == "mediumpurple2" = Just "FF9F79EE"
  | color == "mediumpurple3" = Just "FF8968CD"
  | color == "mediumpurple4" = Just "FF5D478B"
  | color == "mediumseagreen" = Just "FF3CB371"
  | color == "mediumslateblue" = Just "FF7B68EE"
  | color == "mediumspringgreen" = Just "FF00FA9A"
  | color == "mediumturquoise" = Just "FF48D1CC"
  | color == "mediumvioletred" = Just "FFC71585"
  | color == "midnightblue" = Just "FF191970"
  | color == "mintcream" = Just "FFF5FFFA"
  | color == "mistyrose" = Just "FFFFE4E1"
  | color == "mistyrose1" = Just "FFFFE4E1"
  | color == "mistyrose2" = Just "FFEED5D2"
  | color == "mistyrose3" = Just "FFCDB7B5"
  | color == "mistyrose4" = Just "FF8B7D7B"
  | color == "moccasin" = Just "FFFFE4B5"
  | color == "navajowhite" = Just "FFFFDEAD"
  | color == "navajowhite1" = Just "FFFFDEAD"
  | color == "navajowhite2" = Just "FFEECFA1"
  | color == "navajowhite3" = Just "FFCDB38B"
  | color == "navajowhite4" = Just "FF8B795E"
  | color == "navy" = Just "FF000080"
  | color == "navyblue" = Just "FF000080"
  | color == "oldlace" = Just "FFFDF5E6"
  | color == "olivedrab" = Just "FF6B8E23"
  | color == "olivedrab1" = Just "FFC0FF3E"
  | color == "olivedrab2" = Just "FFB3EE3A"
  | color == "olivedrab3" = Just "FF9ACD32"
  | color == "olivedrab4" = Just "FF698B22"
  | color == "orange" = Just "FFFFA500"
  | color == "orange1" = Just "FFFFA500"
  | color == "orange2" = Just "FFEE9A00"
  | color == "orange3" = Just "FFCD8500"
  | color == "orange4" = Just "FF8B5A00"
  | color == "orangered" = Just "FFFF4500"
  | color == "orangered1" = Just "FFFF4500"
  | color == "orangered2" = Just "FFEE4000"
  | color == "orangered3" = Just "FFCD3700"
  | color == "orangered4" = Just "FF8B2500"
  | color == "orchid" = Just "FFDA70D6"
  | color == "orchid1" = Just "FFFF83FA"
  | color == "orchid2" = Just "FFEE7AE9"
  | color == "orchid3" = Just "FFCD69C9"
  | color == "orchid4" = Just "FF8B4789"
  | color == "palegoldenrod" = Just "FFEEE8AA"
  | color == "palegreen" = Just "FF98FB98"
  | color == "palegreen1" = Just "FF9AFF9A"
  | color == "palegreen2" = Just "FF90EE90"
  | color == "palegreen3" = Just "FF7CCD7C"
  | color == "palegreen4" = Just "FF548B54"
  | color == "paleturquoise" = Just "FFAFEEEE"
  | color == "paleturquoise1" = Just "FFBBFFFF"
  | color == "paleturquoise2" = Just "FFAEEEEE"
  | color == "paleturquoise3" = Just "FF96CDCD"
  | color == "paleturquoise4" = Just "FF668B8B"
  | color == "palevioletred" = Just "FFDB7093"
  | color == "palevioletred1" = Just "FFFF82AB"
  | color == "palevioletred2" = Just "FFEE799F"
  | color == "palevioletred3" = Just "FFCD6889"
  | color == "palevioletred4" = Just "FF8B475D"
  | color == "papayawhip" = Just "FFFFEFD5"
  | color == "peachpuff" = Just "FFFFDAB9"
  | color == "peachpuff1" = Just "FFFFDAB9"
  | color == "peachpuff2" = Just "FFEECBAD"
  | color == "peachpuff3" = Just "FFCDAF95"
  | color == "peachpuff4" = Just "FF8B7765"
  | color == "peru" = Just "FFCD853F"
  | color == "pink" = Just "FFFFC0CB"
  | color == "pink1" = Just "FFFFB5C5"
  | color == "pink2" = Just "FFEEA9B8"
  | color == "pink3" = Just "FFCD919E"
  | color == "pink4" = Just "FF8B636C"
  | color == "plum" = Just "FFDDA0DD"
  | color == "plum1" = Just "FFFFBBFF"
  | color == "plum2" = Just "FFEEAEEE"
  | color == "plum3" = Just "FFCD96CD"
  | color == "plum4" = Just "FF8B668B"
  | color == "powderblue" = Just "FFB0E0E6"
  | color == "purple" = Just "FFA020F0"
  | color == "purple1" = Just "FF9B30FF"
  | color == "purple2" = Just "FF912CEE"
  | color == "purple3" = Just "FF7D26CD"
  | color == "purple4" = Just "FF551A8B"
  | color == "red" = Just "FFFF0000"
  | color == "red1" = Just "FFFF0000"
  | color == "red2" = Just "FFEE0000"
  | color == "red3" = Just "FFCD0000"
  | color == "red4" = Just "FF8B0000"
  | color == "rosybrown" = Just "FFBC8F8F"
  | color == "rosybrown1" = Just "FFFFC1C1"
  | color == "rosybrown2" = Just "FFEEB4B4"
  | color == "rosybrown3" = Just "FFCD9B9B"
  | color == "rosybrown4" = Just "FF8B6969"
  | color == "royalblue" = Just "FF4169E1"
  | color == "royalblue1" = Just "FF4876FF"
  | color == "royalblue2" = Just "FF436EEE"
  | color == "royalblue3" = Just "FF3A5FCD"
  | color == "royalblue4" = Just "FF27408B"
  | color == "saddlebrown" = Just "FF8B4513"
  | color == "salmon" = Just "FFFA8072"
  | color == "salmon1" = Just "FFFF8C69"
  | color == "salmon2" = Just "FFEE8262"
  | color == "salmon3" = Just "FFCD7054"
  | color == "salmon4" = Just "FF8B4C39"
  | color == "sandybrown" = Just "FFF4A460"
  | color == "seagreen" = Just "FF2E8B57"
  | color == "seagreen1" = Just "FF54FF9F"
  | color == "seagreen2" = Just "FF4EEE94"
  | color == "seagreen3" = Just "FF43CD80"
  | color == "seagreen4" = Just "FF2E8B57"
  | color == "seashell" = Just "FFFFF5EE"
  | color == "seashell1" = Just "FFFFF5EE"
  | color == "seashell2" = Just "FFEEE5DE"
  | color == "seashell3" = Just "FFCDC5BF"
  | color == "seashell4" = Just "FF8B8682"
  | color == "sienna" = Just "FFA0522D"
  | color == "sienna1" = Just "FFFF8247"
  | color == "sienna2" = Just "FFEE7942"
  | color == "sienna3" = Just "FFCD6839"
  | color == "sienna4" = Just "FF8B4726"
  | color == "skyblue" = Just "FF87CEEB"
  | color == "skyblue1" = Just "FF87CEFF"
  | color == "skyblue2" = Just "FF7EC0EE"
  | color == "skyblue3" = Just "FF6CA6CD"
  | color == "skyblue4" = Just "FF4A708B"
  | color == "slateblue" = Just "FF6A5ACD"
  | color == "slateblue1" = Just "FF836FFF"
  | color == "slateblue2" = Just "FF7A67EE"
  | color == "slateblue3" = Just "FF6959CD"
  | color == "slateblue4" = Just "FF473C8B"
  | color == "slategray" = Just "FF708090"
  | color == "slategray1" = Just "FFC6E2FF"
  | color == "slategray2" = Just "FFB9D3EE"
  | color == "slategray3" = Just "FF9FB6CD"
  | color == "slategray4" = Just "FF6C7B8B"
  | color == "slategrey" = Just "FF708090"
  | color == "snow" = Just "FFFFFAFA"
  | color == "snow1" = Just "FFFFFAFA"
  | color == "snow2" = Just "FFEEE9E9"
  | color == "snow3" = Just "FFCDC9C9"
  | color == "snow4" = Just "FF8B8989"
  | color == "springgreen" = Just "FF00FF7F"
  | color == "springgreen1" = Just "FF00FF7F"
  | color == "springgreen2" = Just "FF00EE76"
  | color == "springgreen3" = Just "FF00CD66"
  | color == "springgreen4" = Just "FF008B45"
  | color == "steelblue" = Just "FF4682B4"
  | color == "steelblue1" = Just "FF63B8FF"
  | color == "steelblue2" = Just "FF5CACEE"
  | color == "steelblue3" = Just "FF4F94CD"
  | color == "steelblue4" = Just "FF36648B"
  | color == "tan" = Just "FFD2B48C"
  | color == "tan1" = Just "FFFFA54F"
  | color == "tan2" = Just "FFEE9A49"
  | color == "tan3" = Just "FFCD853F"
  | color == "tan4" = Just "FF8B5A2B"
  | color == "thistle" = Just "FFD8BFD8"
  | color == "thistle1" = Just "FFFFE1FF"
  | color == "thistle2" = Just "FFEED2EE"
  | color == "thistle3" = Just "FFCDB5CD"
  | color == "thistle4" = Just "FF8B7B8B"
  | color == "tomato" = Just "FFFF6347"
  | color == "tomato1" = Just "FFFF6347"
  | color == "tomato2" = Just "FFEE5C42"
  | color == "tomato3" = Just "FFCD4F39"
  | color == "tomato4" = Just "FF8B3626"
  | color == "turquoise" = Just "FF40E0D0"
  | color == "turquoise1" = Just "FF00F5FF"
  | color == "turquoise2" = Just "FF00E5EE"
  | color == "turquoise3" = Just "FF00C5CD"
  | color == "turquoise4" = Just "FF00868B"
  | color == "violet" = Just "FFEE82EE"
  | color == "violetred" = Just "FFD02090"
  | color == "violetred1" = Just "FFFF3E96"
  | color == "violetred2" = Just "FFEE3A8C"
  | color == "violetred3" = Just "FFCD3278"
  | color == "violetred4" = Just "FF8B2252"
  | color == "wheat" = Just "FFF5DEB3"
  | color == "wheat1" = Just "FFFFE7BA"
  | color == "wheat2" = Just "FFEED8AE"
  | color == "wheat3" = Just "FFCDBA96"
  | color == "wheat4" = Just "FF8B7E66"
  | color == "whitesmoke" = Just "FFF5F5F5"
  | color == "yellow" = Just "FFFFFF00"
  | color == "yellow1" = Just "FFFFFF00"
  | color == "yellow2" = Just "FFEEEE00"
  | color == "yellow3" = Just "FFCDCD00"
  | color == "yellow4" = Just "FF8B8B00"
  | color == "yellowgreen" = Just "FF9ACD32"
  | otherwise = Nothing

-- | This element specifies fill formatting.
--
-- TODO: Gradient fills (18.8.4) are currently unsupported. If we add them,
-- then the spec says (@CT_Fill@, p. 3935), _either_ a gradient _or_ a solid
-- fill pattern should be specified.
--
-- Section 18.8.20, "fill (Fill)" (p. 1768)
data Fill = Fill {
    _fillPattern :: Maybe FillPattern
  }
  deriving (Eq, Ord, Show, Generic)

-- | This element is used to specify cell fill information for pattern and solid
-- color cell fills. For solid cell fills (no pattern), fgColor is used. For
-- cell fills with patterns specified, then the cell fill color is specified by
-- the bgColor element.
--
-- Section 18.8.32 "patternFill (Pattern)" (p. 1793)
data FillPattern = FillPattern {
    _fillPatternBgColor :: Maybe Color
  , _fillPatternFgColor :: Maybe Color
  , _fillPatternType    :: Maybe PatternType
  }
  deriving (Eq, Ord, Show, Generic)

-- | This element defines the properties for one of the fonts used in this
-- workbook.
--
-- Section 18.2.22 "font (Font)" (p. 1769)
data Font = Font {
    -- | Displays characters in bold face font style.
    _fontBold          :: Maybe Bool

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
  , _fontCharset       :: Maybe Int

    -- | Color
  , _fontColor         :: Maybe Color

    -- | Macintosh compatibility setting. Represents special word/character
    -- rendering on Macintosh, when this flag is set. The effect is to condense
    -- the text (squeeze it together). SpreadsheetML applications are not
    -- required to render according to this flag.
  , _fontCondense      :: Maybe Bool

    -- | This element specifies a compatibility setting used for previous
    -- spreadsheet applications, resulting in special word/character rendering
    -- on those legacy applications, when this flag is set. The effect extends
    -- or stretches out the text. SpreadsheetML applications are not required to
    -- render according to this flag.
  , _fontExtend        :: Maybe Bool

    -- | The font family this font belongs to. A font family is a set of fonts
    -- having common stroke width and serif characteristics. This is system
    -- level font information. The font name overrides when there are
    -- conflicting values.
  , _fontFamily        :: Maybe FontFamily

    -- | Displays characters in italic font style. The italic style is defined
    -- by the font at a system level and is not specified by ECMA-376.
  , _fontItalic        :: Maybe Bool

    -- | This element specifies the face name of this font.
    --
    -- A string representing the name of the font. If the font doesn't exist
    -- (because it isn't installed on the system), or the charset not supported
    -- by that font, then another font should be substituted.
    --
    -- The string length for this attribute shall be 0 to 31 characters.
  , _fontName          :: Maybe Text

    -- | This element displays only the inner and outer borders of each
    -- character. This is very similar to Bold in behavior.
  , _fontOutline       :: Maybe Bool

    -- | Defines the font scheme, if any, to which this font belongs. When a
    -- font definition is part of a theme definition, then the font is
    -- categorized as either a major or minor font scheme component. When a new
    -- theme is chosen, every font that is part of a theme definition is updated
    -- to use the new major or minor font definition for that theme. Usually
    -- major fonts are used for styles like headings, and minor fonts are used
    -- for body and paragraph text.
  , _fontScheme        :: Maybe FontScheme

    -- | Macintosh compatibility setting. Represents special word/character
    -- rendering on Macintosh, when this flag is set. The effect is to render a
    -- shadow behind, beneath and to the right of the text. SpreadsheetML
    -- applications are not required to render according to this flag.
  , _fontShadow        :: Maybe Bool

    -- | This element draws a strikethrough line through the horizontal middle
    -- of the text.
  , _fontStrikeThrough :: Maybe Bool

    -- | This element represents the point size (1/72 of an inch) of the Latin
    -- and East Asian text.
  , _fontSize          :: Maybe Double

    -- | This element represents the underline formatting style.
  , _fontUnderline     :: Maybe FontUnderline

    -- | This element adjusts the vertical position of the text relative to the
    -- text's default appearance for this run. It is used to get 'superscript'
    -- or 'subscript' texts, and shall reduce the font size (if a smaller size
    -- is available) accordingly.
  , _fontVertAlign     :: Maybe FontVerticalAlignment
  }
  deriving (Eq, Ord, Show, Generic)

-- | A single dxf record, expressing incremental formatting to be applied.
--
-- Section 18.8.14, "dxf (Formatting)" (p. 1765)
data Dxf = Dxf
    { _dxfFont       :: Maybe Font
    -- TODO: numFmt
    , _dxfFill       :: Maybe Fill
    , _dxfAlignment  :: Maybe Alignment
    , _dxfBorder     :: Maybe Border
    , _dxfProtection :: Maybe Protection
    -- TODO: extList
    } deriving (Eq, Ord, Show, Generic)

type NumFmt = Text

-- | This element specifies number format properties which indicate
-- how to format and render the numeric value of a cell.
--
-- Section 18.8.30 "numFmt (Number Format)" (p. 1777)
data NumberFormat
    = StdNumberFormat ImpliedNumberFormat
    | UserNumberFormat NumFmt
    deriving (Eq, Ord, Show, Generic)

-- | Basic number format with predefined number of decimals
-- as format code of number format in xlsx should be less than 255 characters
-- number of decimals shouldn't be more than 253
fmtDecimals :: Int -> NumberFormat
fmtDecimals k = UserNumberFormat $ "0." <> T.replicate k "#"

-- | Basic number format with predefined number of decimals.
-- Works like 'fmtDecimals' with the only difference that extra zeroes are
-- displayed when number of digits after the point is less than the number
-- of digits specified in the format
fmtDecimalsZeroes :: Int -> NumberFormat
fmtDecimalsZeroes k = UserNumberFormat $ "0." <> T.replicate k "0"

-- | Implied number formats
--
-- /Note:/ This only implements the predefined values for 18.2.30 "All Languages",
-- other built-in format ids (with id < 'firstUserNumFmtId') are stored in 'NfOtherBuiltin'
data ImpliedNumberFormat =
    NfGeneral                         -- ^> 0 General
  | NfZero                            -- ^> 1 0
  | Nf2Decimal                        -- ^> 2 0.00
  | NfMax3Decimal                     -- ^> 3 #,##0
  | NfThousandSeparator2Decimal       -- ^> 4 #,##0.00
  | NfPercent                         -- ^> 9 0%
  | NfPercent2Decimal                 -- ^> 10 0.00%
  | NfExponent2Decimal                -- ^> 11 0.00E+00
  | NfSingleSpacedFraction            -- ^> 12 # ?/?
  | NfDoubleSpacedFraction            -- ^> 13 # ??/??
  | NfMmDdYy                          -- ^> 14 mm-dd-yy
  | NfDMmmYy                          -- ^> 15 d-mmm-yy
  | NfDMmm                            -- ^> 16 d-mmm
  | NfMmmYy                           -- ^> 17 mmm-yy
  | NfHMm12Hr                         -- ^> 18 h:mm AM/PM
  | NfHMmSs12Hr                       -- ^> 19 h:mm:ss AM/PM
  | NfHMm                             -- ^> 20 h:mm
  | NfHMmSs                           -- ^> 21 h:mm:ss
  | NfMdyHMm                          -- ^> 22 m/d/yy h:mm
  | NfThousandsNegativeParens         -- ^> 37 #,##0 ;(#,##0)
  | NfThousandsNegativeRed            -- ^> 38 #,##0 ;[Red](#,##0)
  | NfThousands2DecimalNegativeParens -- ^> 39 #,##0.00;(#,##0.00)
  | NfThousands2DecimalNegativeRed    -- ^> 40 #,##0.00;[Red](#,##0.00)
  | NfMmSs                            -- ^> 45 mm:ss
  | NfOptHMmSs                        -- ^> 46 [h]:mm:ss
  | NfMmSs1Decimal                    -- ^> 47 mmss.0
  | NfExponent1Decimal                -- ^> 48 ##0.0E+0
  | NfTextPlaceHolder                 -- ^> 49 @
  | NfOtherImplied Int                -- ^ other (non local-neutral?) built-in format (id < 164)
  deriving (Eq, Ord, Show, Generic)

stdNumberFormatId :: ImpliedNumberFormat -> Int
stdNumberFormatId NfGeneral                         = 0 -- General
stdNumberFormatId NfZero                            = 1 -- 0
stdNumberFormatId Nf2Decimal                        = 2 -- 0.00
stdNumberFormatId NfMax3Decimal                     = 3 -- #,##0
stdNumberFormatId NfThousandSeparator2Decimal       = 4 -- #,##0.00
stdNumberFormatId NfPercent                         = 9 -- 0%
stdNumberFormatId NfPercent2Decimal                 = 10 -- 0.00%
stdNumberFormatId NfExponent2Decimal                = 11 -- 0.00E+00
stdNumberFormatId NfSingleSpacedFraction            = 12 -- # ?/?
stdNumberFormatId NfDoubleSpacedFraction            = 13 -- # ??/??
stdNumberFormatId NfMmDdYy                          = 14 -- mm-dd-yy
stdNumberFormatId NfDMmmYy                          = 15 -- d-mmm-yy
stdNumberFormatId NfDMmm                            = 16 -- d-mmm
stdNumberFormatId NfMmmYy                           = 17 -- mmm-yy
stdNumberFormatId NfHMm12Hr                         = 18 -- h:mm AM/PM
stdNumberFormatId NfHMmSs12Hr                       = 19 -- h:mm:ss AM/PM
stdNumberFormatId NfHMm                             = 20 -- h:mm
stdNumberFormatId NfHMmSs                           = 21 -- h:mm:ss
stdNumberFormatId NfMdyHMm                          = 22 -- m/d/yy h:mm
stdNumberFormatId NfThousandsNegativeParens         = 37 -- #,##0 ;(#,##0)
stdNumberFormatId NfThousandsNegativeRed            = 38 -- #,##0 ;[Red](#,##0)
stdNumberFormatId NfThousands2DecimalNegativeParens = 39 -- #,##0.00;(#,##0.00)
stdNumberFormatId NfThousands2DecimalNegativeRed    = 40 -- #,##0.00;[Red](#,##0.00)
stdNumberFormatId NfMmSs                            = 45 -- mm:ss
stdNumberFormatId NfOptHMmSs                        = 46 -- [h]:mm:ss
stdNumberFormatId NfMmSs1Decimal                    = 47 -- mmss.0
stdNumberFormatId NfExponent1Decimal                = 48 -- ##0.0E+0
stdNumberFormatId NfTextPlaceHolder                 = 49 -- @
stdNumberFormatId (NfOtherImplied i)                = i

idToStdNumberFormat :: Int -> Maybe ImpliedNumberFormat
idToStdNumberFormat 0  = Just NfGeneral                         -- General
idToStdNumberFormat 1  = Just NfZero                            -- 0
idToStdNumberFormat 2  = Just Nf2Decimal                        -- 0.00
idToStdNumberFormat 3  = Just NfMax3Decimal                     -- #,##0
idToStdNumberFormat 4  = Just NfThousandSeparator2Decimal       -- #,##0.00
idToStdNumberFormat 9  = Just NfPercent                         -- 0%
idToStdNumberFormat 10 = Just NfPercent2Decimal                 -- 0.00%
idToStdNumberFormat 11 = Just NfExponent2Decimal                -- 0.00E+00
idToStdNumberFormat 12 = Just NfSingleSpacedFraction            -- # ?/?
idToStdNumberFormat 13 = Just NfDoubleSpacedFraction            -- # ??/??
idToStdNumberFormat 14 = Just NfMmDdYy                          -- mm-dd-yy
idToStdNumberFormat 15 = Just NfDMmmYy                          -- d-mmm-yy
idToStdNumberFormat 16 = Just NfDMmm                            -- d-mmm
idToStdNumberFormat 17 = Just NfMmmYy                           -- mmm-yy
idToStdNumberFormat 18 = Just NfHMm12Hr                         -- h:mm AM/PM
idToStdNumberFormat 19 = Just NfHMmSs12Hr                       -- h:mm:ss AM/PM
idToStdNumberFormat 20 = Just NfHMm                             -- h:mm
idToStdNumberFormat 21 = Just NfHMmSs                           -- h:mm:ss
idToStdNumberFormat 22 = Just NfMdyHMm                          -- m/d/yy h:mm
idToStdNumberFormat 37 = Just NfThousandsNegativeParens         -- #,##0 ;(#,##0)
idToStdNumberFormat 38 = Just NfThousandsNegativeRed            -- #,##0 ;[Red](#,##0)
idToStdNumberFormat 39 = Just NfThousands2DecimalNegativeParens -- #,##0.00;(#,##0.00)
idToStdNumberFormat 40 = Just NfThousands2DecimalNegativeRed    -- #,##0.00;[Red](#,##0.00)
idToStdNumberFormat 45 = Just NfMmSs                            -- mm:ss
idToStdNumberFormat 46 = Just NfOptHMmSs                        -- [h]:mm:ss
idToStdNumberFormat 47 = Just NfMmSs1Decimal                    -- mmss.0
idToStdNumberFormat 48 = Just NfExponent1Decimal                -- ##0.0E+0
idToStdNumberFormat 49 = Just NfTextPlaceHolder                 -- @
idToStdNumberFormat i  = if i < firstUserNumFmtId then Just (NfOtherImplied i) else Nothing

firstUserNumFmtId :: Int
firstUserNumFmtId = 164

-- | Protection properties
--
-- Contains protection properties associated with the cell. Each cell has
-- protection properties that can be set. The cell protection properties do not
-- take effect unless the sheet has been protected.
--
-- Section 18.8.33, "protection (Protection Properties)", p. 1793
data Protection = Protection {
    _protectionHidden :: Maybe Bool
  , _protectionLocked :: Maybe Bool
  }
  deriving (Eq, Ord, Show, Generic)

{-------------------------------------------------------------------------------
  Enumerations
-------------------------------------------------------------------------------}

-- | Horizontal alignment in cells
--
-- See 18.18.40 "ST_HorizontalAlignment (Horizontal Alignment Type)" (p. 2459)
data CellHorizontalAlignment =
    CellHorizontalAlignmentCenter
  | CellHorizontalAlignmentCenterContinuous
  | CellHorizontalAlignmentDistributed
  | CellHorizontalAlignmentFill
  | CellHorizontalAlignmentGeneral
  | CellHorizontalAlignmentJustify
  | CellHorizontalAlignmentLeft
  | CellHorizontalAlignmentRight
  deriving (Eq, Ord, Show, Generic)

-- | Vertical alignment in cells
--
-- See 18.18.88 "ST_VerticalAlignment (Vertical Alignment Types)" (p. 2512)
data CellVerticalAlignment =
    CellVerticalAlignmentBottom
  | CellVerticalAlignmentCenter
  | CellVerticalAlignmentDistributed
  | CellVerticalAlignmentJustify
  | CellVerticalAlignmentTop
  deriving (Eq, Ord, Show, Generic)

-- | Font family
--
-- See 18.8.18 "family (Font Family)" (p. 1766)
-- and 17.18.30 "ST_FontFamily (Font Family Value)" (p. 1388)
data FontFamily =
    -- | Family is not applicable
    FontFamilyNotApplicable

    -- | Proportional font with serifs
  | FontFamilyRoman

    -- | Proportional font without serifs
  | FontFamilySwiss

    -- | Monospace font with or without serifs
  | FontFamilyModern

    -- | Script font designed to mimic the appearance of handwriting
  | FontFamilyScript

    -- | Novelty font
  | FontFamilyDecorative
  deriving (Eq, Ord, Show, Generic)

-- | Font scheme
--
-- See 18.18.33 "ST_FontScheme (Font scheme Styles)" (p. 2456)
data FontScheme =
    -- | This font is the major font for this theme.
    FontSchemeMajor

    -- | This font is the minor font for this theme.
  | FontSchemeMinor

    -- | This font is not a theme font.
  | FontSchemeNone
  deriving (Eq, Ord, Show, Generic)

-- | Font underline property
--
-- See 18.4.13 "u (Underline)", p 1728
data FontUnderline =
    FontUnderlineSingle
  | FontUnderlineDouble
  | FontUnderlineSingleAccounting
  | FontUnderlineDoubleAccounting
  | FontUnderlineNone
  deriving (Eq, Ord, Show, Generic)

-- | Vertical alignment
--
-- See 22.9.2.17 "ST_VerticalAlignRun (Vertical Positioning Location)" (p. 3794)
data FontVerticalAlignment =
    FontVerticalAlignmentBaseline
  | FontVerticalAlignmentSubscript
  | FontVerticalAlignmentSuperscript
  deriving (Eq, Ord, Show, Generic)

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
  deriving (Eq, Ord, Show, Generic)

-- | Indicates the style of fill pattern being used for a cell format.
--
-- Section 18.18.55 "ST_PatternType (Pattern Type)" (p. 2472)
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
  deriving (Eq, Ord, Show, Generic)

-- | Reading order
--
-- See 18.8.1 "alignment (Alignment)" (p. 1754, esp. p. 1755)
data ReadingOrder =
    ReadingOrderContextDependent
  | ReadingOrderLeftToRight
  | ReadingOrderRightToLeft
  deriving (Eq, Ord, Show, Generic)

{-------------------------------------------------------------------------------
  Lenses
-------------------------------------------------------------------------------}

makeLenses ''StyleSheet
makeLenses ''CellXf
makeLenses ''Dxf

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
    , _styleSheetDxfs    = []
    , _styleSheetNumFmts = M.empty
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

instance Default Dxf where
    def = Dxf
          { _dxfFont       = Nothing
          , _dxfFill       = Nothing
          , _dxfAlignment  = Nothing
          , _dxfBorder     = Nothing
          , _dxfProtection = Nothing
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
  toElement nm StyleSheet{..} = elementListSimple nm elements
    where
      elements = [ countedElementList "numFmts" $ map (toElement "numFmt") numFmts
                 , countedElementList "fonts"   $ map (toElement "font")   _styleSheetFonts
                 , countedElementList "fills"   $ map (toElement "fill")   _styleSheetFills
                 , countedElementList "borders" $ map (toElement "border") _styleSheetBorders
                   -- TODO: cellStyleXfs
                 , countedElementList "cellXfs" $ map (toElement "xf")     _styleSheetCellXfs
                 -- TODO: cellStyles
                 , countedElementList "dxfs"    $ map (toElement "dxf")    _styleSheetDxfs
                 -- TODO: tableStyles
                 -- TODO: colors
                 -- TODO: extLst
                 ]
      numFmts = map NumFmtPair $ M.toList _styleSheetNumFmts

-- | See @CT_Xf@, p. 4486
instance ToElement CellXf where
  toElement nm CellXf{..} = Element {
      elementName       = nm
    , elementNodes      = map NodeElement . catMaybes $ [
          toElement "alignment"  <$> _cellXfAlignment
        , toElement "protection" <$> _cellXfProtection
          -- TODO: extLst
        ]
    , elementAttributes = M.fromList . catMaybes $ [
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

-- | See @CT_Dxf@, p. 3937
instance ToElement Dxf where
    toElement nm Dxf{..} = Element
        { elementName       = nm
        , elementNodes      = map NodeElement $
                              catMaybes [ toElement "font"       <$> _dxfFont
                                        , toElement "fill"       <$> _dxfFill
                                        , toElement "alignment"  <$> _dxfAlignment
                                        , toElement "border"     <$> _dxfBorder
                                        , toElement "protection" <$> _dxfProtection
                                        ]
        , elementAttributes = M.empty
        }

-- | See @CT_CellAlignment@, p. 4482
instance ToElement Alignment where
  toElement nm Alignment{..} = Element {
      elementName       = nm
    , elementNodes      = []
    , elementAttributes = M.fromList . catMaybes $ [
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
    , elementAttributes = M.fromList . catMaybes $ [
          "diagonalUp"   .=? _borderDiagonalUp
        , "diagonalDown" .=? _borderDiagonalDown
        , "outline"      .=? _borderOutline
        ]
    , elementNodes      = map NodeElement . catMaybes $ [
          toElement "start"      <$> _borderStart
        , toElement "end"        <$> _borderEnd
        , toElement "left"       <$> _borderLeft
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
    , elementAttributes = M.fromList . catMaybes $ [
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
    , elementAttributes = M.fromList . catMaybes $ [
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
    , elementAttributes = M.empty
    , elementNodes      = map NodeElement . catMaybes $ [
          toElement "patternFill" <$> _fillPattern
        ]
    }

-- | See @CT_PatternFill@, p. 4484
instance ToElement FillPattern where
  toElement nm FillPattern{..} = Element {
      elementName       = nm
    , elementAttributes = M.fromList . catMaybes $ [
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
    , elementAttributes = M.empty -- all properties specified as child nodes
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
    , elementAttributes = M.fromList . catMaybes $ [
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
  toAttrVal FontFamilyNotApplicable = "0"
  toAttrVal FontFamilyRoman         = "1"
  toAttrVal FontFamilySwiss         = "2"
  toAttrVal FontFamilyModern        = "3"
  toAttrVal FontFamilyScript        = "4"
  toAttrVal FontFamilyDecorative    = "5"

instance ToAttrVal FontScheme where
  toAttrVal FontSchemeMajor = "major"
  toAttrVal FontSchemeMinor = "minor"
  toAttrVal FontSchemeNone  = "none"

-- See @ST_UnderlineValues@, p. 3940
instance ToAttrVal FontUnderline where
  toAttrVal FontUnderlineSingle           = "single"
  toAttrVal FontUnderlineDouble           = "double"
  toAttrVal FontUnderlineSingleAccounting = "singleAccounting"
  toAttrVal FontUnderlineDoubleAccounting = "doubleAccounting"
  toAttrVal FontUnderlineNone             = "none"

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

{-------------------------------------------------------------------------------
  Parsing
-------------------------------------------------------------------------------}
-- | See @CT_Stylesheet@, p. 4482
instance FromCursor StyleSheet where
  fromCursor cur = do
    let
      _styleSheetFonts = cur $/ element (n_ "fonts") &/ element (n_ "font") >=> fromCursor
      _styleSheetFills = cur $/ element (n_ "fills") &/ element (n_ "fill") >=> fromCursor
      _styleSheetBorders = cur $/ element (n_ "borders") &/ element (n_ "border") >=> fromCursor
         -- TODO: cellStyleXfs
      _styleSheetCellXfs = cur $/ element (n_ "cellXfs") &/ element (n_ "xf") >=> fromCursor
         -- TODO: cellStyles
      _styleSheetDxfs = cur $/ element (n_ "dxfs") &/ element (n_ "dxf") >=> fromCursor
      _styleSheetNumFmts = M.fromList . map unNumFmtPair $
          cur $/ element (n_ "numFmts")&/ element (n_ "numFmt") >=> fromCursor
         -- TODO: tableStyles
         -- TODO: colors
         -- TODO: extLst
    return StyleSheet{..}

-- | See @CT_Font@, p. 4489
instance FromCursor Font where
  fromCursor cur = do
    _fontName         <- maybeElementValue (n_ "name") cur
    _fontCharset      <- maybeElementValue (n_ "charset") cur
    _fontFamily       <- maybeElementValue (n_ "family") cur
    _fontBold         <- maybeBoolElementValue (n_ "b") cur
    _fontItalic       <- maybeBoolElementValue (n_ "i") cur
    _fontStrikeThrough<- maybeBoolElementValue (n_ "strike") cur
    _fontOutline      <- maybeBoolElementValue (n_ "outline") cur
    _fontShadow       <- maybeBoolElementValue (n_ "shadow") cur
    _fontCondense     <- maybeBoolElementValue (n_ "condense") cur
    _fontExtend       <- maybeBoolElementValue (n_ "extend") cur
    _fontColor        <- maybeFromElement  (n_ "color") cur
    _fontSize         <- maybeElementValue (n_ "sz") cur
    _fontUnderline    <- maybeElementValueDef (n_ "u") FontUnderlineSingle cur
    _fontVertAlign    <- maybeElementValue (n_ "vertAlign") cur
    _fontScheme       <- maybeElementValue (n_ "scheme") cur
    return Font{..}

-- | See 18.18.94 "ST_FontFamily (Font Family)" (p. 2517)
instance FromAttrVal FontFamily where
  fromAttrVal "0" = readSuccess FontFamilyNotApplicable
  fromAttrVal "1" = readSuccess FontFamilyRoman
  fromAttrVal "2" = readSuccess FontFamilySwiss
  fromAttrVal "3" = readSuccess FontFamilyModern
  fromAttrVal "4" = readSuccess FontFamilyScript
  fromAttrVal "5" = readSuccess FontFamilyDecorative
  fromAttrVal t   = invalidText "FontFamily" t

-- | See @CT_Color@, p. 4484
instance FromCursor Color where
  fromCursor cur = do
    _colorAutomatic <- maybeAttribute "auto" cur
    _colorARGB      <- maybeAttribute "rgb" cur
    _colorTheme     <- maybeAttribute "theme" cur
    _colorTint      <- maybeAttribute "tint" cur
    return Color{..}

-- See @ST_UnderlineValues@, p. 3940
instance FromAttrVal FontUnderline where
  fromAttrVal "single"           = readSuccess FontUnderlineSingle
  fromAttrVal "double"           = readSuccess FontUnderlineDouble
  fromAttrVal "singleAccounting" = readSuccess FontUnderlineSingleAccounting
  fromAttrVal "doubleAccounting" = readSuccess FontUnderlineDoubleAccounting
  fromAttrVal "none"             = readSuccess FontUnderlineNone
  fromAttrVal t                  = invalidText "FontUnderline" t

instance FromAttrVal FontVerticalAlignment where
  fromAttrVal "baseline"    = readSuccess FontVerticalAlignmentBaseline
  fromAttrVal "subscript"   = readSuccess FontVerticalAlignmentSubscript
  fromAttrVal "superscript" = readSuccess FontVerticalAlignmentSuperscript
  fromAttrVal t             = invalidText "FontVerticalAlignment" t

instance FromAttrVal FontScheme where
  fromAttrVal "major" = readSuccess FontSchemeMajor
  fromAttrVal "minor" = readSuccess FontSchemeMinor
  fromAttrVal "none"  = readSuccess FontSchemeNone
  fromAttrVal t       = invalidText "FontScheme" t

-- | See @CT_Fill@, p. 4484
instance FromCursor Fill where
  fromCursor cur = do
    _fillPattern <- maybeFromElement (n_ "patternFill") cur
    return Fill{..}

-- | See @CT_PatternFill@, p. 4484
instance FromCursor FillPattern where
  fromCursor cur = do
    _fillPatternType <- maybeAttribute "patternType" cur
    _fillPatternFgColor <- maybeFromElement (n_ "fgColor") cur
    _fillPatternBgColor <- maybeFromElement (n_ "bgColor") cur
    return FillPattern{..}

instance FromAttrVal PatternType where
  fromAttrVal "darkDown"        = readSuccess PatternTypeDarkDown
  fromAttrVal "darkGray"        = readSuccess PatternTypeDarkGray
  fromAttrVal "darkGrid"        = readSuccess PatternTypeDarkGrid
  fromAttrVal "darkHorizontal"  = readSuccess PatternTypeDarkHorizontal
  fromAttrVal "darkTrellis"     = readSuccess PatternTypeDarkTrellis
  fromAttrVal "darkUp"          = readSuccess PatternTypeDarkUp
  fromAttrVal "darkVertical"    = readSuccess PatternTypeDarkVertical
  fromAttrVal "gray0625"        = readSuccess PatternTypeGray0625
  fromAttrVal "gray125"         = readSuccess PatternTypeGray125
  fromAttrVal "lightDown"       = readSuccess PatternTypeLightDown
  fromAttrVal "lightGray"       = readSuccess PatternTypeLightGray
  fromAttrVal "lightGrid"       = readSuccess PatternTypeLightGrid
  fromAttrVal "lightHorizontal" = readSuccess PatternTypeLightHorizontal
  fromAttrVal "lightTrellis"    = readSuccess PatternTypeLightTrellis
  fromAttrVal "lightUp"         = readSuccess PatternTypeLightUp
  fromAttrVal "lightVertical"   = readSuccess PatternTypeLightVertical
  fromAttrVal "mediumGray"      = readSuccess PatternTypeMediumGray
  fromAttrVal "none"            = readSuccess PatternTypeNone
  fromAttrVal "solid"           = readSuccess PatternTypeSolid
  fromAttrVal t                 = invalidText "PatternType" t

-- | See @CT_Border@, p. 4483
instance FromCursor Border where
  fromCursor cur = do
    _borderDiagonalUp   <- maybeAttribute "diagonalUp" cur
    _borderDiagonalDown <- maybeAttribute "diagonalDown" cur
    _borderOutline      <- maybeAttribute "outline" cur
    _borderStart      <- maybeFromElement (n_ "start") cur
    _borderEnd        <- maybeFromElement (n_ "end") cur
    _borderLeft       <- maybeFromElement (n_ "left") cur
    _borderRight      <- maybeFromElement (n_ "right") cur
    _borderTop        <- maybeFromElement (n_ "top") cur
    _borderBottom     <- maybeFromElement (n_ "bottom") cur
    _borderDiagonal   <- maybeFromElement (n_ "diagonal") cur
    _borderVertical   <- maybeFromElement (n_ "vertical") cur
    _borderHorizontal <- maybeFromElement (n_ "horizontal") cur
    return Border{..}

instance FromCursor BorderStyle where
  fromCursor cur = do
    _borderStyleLine  <- maybeAttribute "style" cur
    _borderStyleColor <- maybeFromElement (n_ "color") cur
    return BorderStyle{..}

instance FromAttrVal LineStyle where
  fromAttrVal "dashDot"          = readSuccess LineStyleDashDot
  fromAttrVal "dashDotDot"       = readSuccess LineStyleDashDotDot
  fromAttrVal "dashed"           = readSuccess LineStyleDashed
  fromAttrVal "dotted"           = readSuccess LineStyleDotted
  fromAttrVal "double"           = readSuccess LineStyleDouble
  fromAttrVal "hair"             = readSuccess LineStyleHair
  fromAttrVal "medium"           = readSuccess LineStyleMedium
  fromAttrVal "mediumDashDot"    = readSuccess LineStyleMediumDashDot
  fromAttrVal "mediumDashDotDot" = readSuccess LineStyleMediumDashDotDot
  fromAttrVal "mediumDashed"     = readSuccess LineStyleMediumDashed
  fromAttrVal "none"             = readSuccess LineStyleNone
  fromAttrVal "slantDashDot"     = readSuccess LineStyleSlantDashDot
  fromAttrVal "thick"            = readSuccess LineStyleThick
  fromAttrVal "thin"             = readSuccess LineStyleThin
  fromAttrVal t                  = invalidText "LineStyle" t

-- | See @CT_Xf@, p. 4486
instance FromCursor CellXf where
  fromCursor cur = do
    _cellXfAlignment  <- maybeFromElement (n_ "alignment") cur
    _cellXfProtection <- maybeFromElement (n_ "protection") cur
    _cellXfNumFmtId          <- maybeAttribute "numFmtId" cur
    _cellXfFontId            <- maybeAttribute "fontId" cur
    _cellXfFillId            <- maybeAttribute "fillId" cur
    _cellXfBorderId          <- maybeAttribute "borderId" cur
    _cellXfId                <- maybeAttribute "xfId" cur
    _cellXfQuotePrefix       <- maybeAttribute "quotePrefix" cur
    _cellXfPivotButton       <- maybeAttribute "pivotButton" cur
    _cellXfApplyNumberFormat <- maybeAttribute "applyNumberFormat" cur
    _cellXfApplyFont         <- maybeAttribute "applyFont" cur
    _cellXfApplyFill         <- maybeAttribute "applyFill" cur
    _cellXfApplyBorder       <- maybeAttribute "applyBorder" cur
    _cellXfApplyAlignment    <- maybeAttribute "applyAlignment" cur
    _cellXfApplyProtection   <- maybeAttribute "applyProtection" cur
    return CellXf{..}

-- | See @CT_Dxf@, p. 3937
instance FromCursor Dxf where
    fromCursor cur = do
      _dxfFont         <- maybeFromElement (n_ "font") cur
      _dxfFill         <- maybeFromElement (n_ "fill") cur
      _dxfAlignment    <- maybeFromElement (n_ "alignment") cur
      _dxfBorder       <- maybeFromElement (n_ "border") cur
      _dxfProtection   <- maybeFromElement (n_ "protection") cur
      return Dxf{..}

-- | See @CT_CellAlignment@, p. 4482
instance FromCursor Alignment where
  fromCursor cur = do
    _alignmentHorizontal      <- maybeAttribute "horizontal" cur
    _alignmentVertical        <- maybeAttribute "vertical" cur
    _alignmentTextRotation    <- maybeAttribute "textRotation" cur
    _alignmentWrapText        <- maybeAttribute "wrapText" cur
    _alignmentRelativeIndent  <- maybeAttribute "relativeIndent" cur
    _alignmentIndent          <- maybeAttribute "indent" cur
    _alignmentJustifyLastLine <- maybeAttribute "justifyLastLine" cur
    _alignmentShrinkToFit     <- maybeAttribute "shrinkToFit" cur
    _alignmentReadingOrder    <- maybeAttribute "readingOrder" cur
    return Alignment{..}

instance FromAttrVal CellHorizontalAlignment where
  fromAttrVal "center"           = readSuccess CellHorizontalAlignmentCenter
  fromAttrVal "centerContinuous" = readSuccess CellHorizontalAlignmentCenterContinuous
  fromAttrVal "distributed"      = readSuccess CellHorizontalAlignmentDistributed
  fromAttrVal "fill"             = readSuccess CellHorizontalAlignmentFill
  fromAttrVal "general"          = readSuccess CellHorizontalAlignmentGeneral
  fromAttrVal "justify"          = readSuccess CellHorizontalAlignmentJustify
  fromAttrVal "left"             = readSuccess CellHorizontalAlignmentLeft
  fromAttrVal "right"            = readSuccess CellHorizontalAlignmentRight
  fromAttrVal t                  = invalidText "CellHorizontalAlignment" t

instance FromAttrVal CellVerticalAlignment where
  fromAttrVal "bottom"      = readSuccess CellVerticalAlignmentBottom
  fromAttrVal "center"      = readSuccess CellVerticalAlignmentCenter
  fromAttrVal "distributed" = readSuccess CellVerticalAlignmentDistributed
  fromAttrVal "justify"     = readSuccess CellVerticalAlignmentJustify
  fromAttrVal "top"         = readSuccess CellVerticalAlignmentTop
  fromAttrVal t             = invalidText "CellVerticalAlignment" t

instance FromAttrVal ReadingOrder where
  fromAttrVal "0" = readSuccess ReadingOrderContextDependent
  fromAttrVal "1" = readSuccess ReadingOrderLeftToRight
  fromAttrVal "2" = readSuccess ReadingOrderRightToLeft
  fromAttrVal t   = invalidText "ReadingOrder" t

-- | See @CT_CellProtection@, p. 4484
instance FromCursor Protection where
  fromCursor cur = do
    _protectionLocked <- maybeAttribute "locked" cur
    _protectionHidden <- maybeAttribute "hidden" cur
    return Protection{..}
