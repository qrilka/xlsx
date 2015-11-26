{-# LANGUAGE OverloadedStrings #-}
{-# LANGUAGE TemplateHaskell   #-}
{-# LANGUAGE RecordWildCards   #-}
{-# OPTIONS_GHC -Wall #-}
module Codec.Xlsx.PageSetup (
    -- * Main types
    PageSetup(..)
  , renderPageSetup
    -- * Enumerations
  , CellComments(..)
  , PrintErrors(..)
  , Orientation(..)
  , PageOrder(..)
  , PaperSize(..)
    -- * Lenses
    -- ** PageSetup
  , pageSetupBlackAndWhite
  , pageSetupCellComments
  , pageSetupCopies
  , pageSetupDraft
  , pageSetupErrors
  , pageSetupFirstPageNumber
  , pageSetupFitToHeight
  , pageSetupFitToWidth
  , pageSetupHorizontalDpi
  , pageSetupId
  , pageSetupOrientation
  , pageSetupPageOrder
  , pageSetupPaperHeight
  , pageSetupPaperSize
  , pageSetupPaperWidth
  , pageSetupScale
  , pageSetupUseFirstPageNumber
  , pageSetupUsePrinterDefaults
  , pageSetupVerticalDpi
  ) where

import Control.Lens (makeLenses)
import Data.Default
import Data.Maybe (catMaybes)
import Text.XML
import qualified Data.Map as Map

import Codec.Xlsx.Types
import Codec.Xlsx.Writer.Internal

{-------------------------------------------------------------------------------
  Main types
-------------------------------------------------------------------------------}

data PageSetup = PageSetup {
    -- | Print black and white.
    _pageSetupBlackAndWhite :: Maybe Bool

    -- | This attribute specifies how to print cell comments.
  , _pageSetupCellComments :: Maybe CellComments

    -- | Number of copies to print.
  , _pageSetupCopies :: Maybe Int

    -- | Print without graphics.
  , _pageSetupDraft :: Maybe Bool

     -- | Specifies how to print cell values for cells with errors.
  , _pageSetupErrors :: Maybe PrintErrors

     -- | Page number for first printed page. If no value is specified, then
     -- 'automatic' is assumed.
  , _pageSetupFirstPageNumber :: Maybe Int

     -- | Number of vertical pages to fit on.
  , _pageSetupFitToHeight :: Maybe Int

     -- | Number of horizontal pages to fit on.
  , _pageSetupFitToWidth :: Maybe Int

     -- | Horizontal print resolution of the device.
  , _pageSetupHorizontalDpi :: Maybe Int

     -- | Relationship Id of the devMode printer settings part.
     --
     -- (Explicit reference to a parent XML element.)
     --
     -- See 22.8.2.1 "ST_RelationshipId (Explicit Relationship ID)" (p. 4324)
  , _pageSetupId :: Maybe String

     -- | Orientation of the page.
  , _pageSetupOrientation :: Maybe Orientation

     -- | Order of printed pages
  , _pageSetupPageOrder :: Maybe PageOrder

     -- | Height of custom paper as a number followed by a unit identifier.
     --
     -- When paperHeight and paperWidth are specified, paperSize shall be ignored.
     -- Examples: @"297mm"@, @"11in"@.
     --
     -- See 22.9.2.12 "ST_PositiveUniversalMeasure (Positive Universal Measurement)" (p. 4336)
  , _pageSetupPaperHeight :: Maybe String

     -- | Pager size
     --
     -- When paperHeight, paperWidth, and paperUnits are specified, paperSize
     -- should be ignored.
  , _pageSetupPaperSize :: Maybe PaperSize

     -- | Width of custom paper as a number followed by a unit identifier
     --
     -- Examples: @21cm@, @8.5in@
     --
     -- When paperHeight and paperWidth are specified, paperSize shall be
     -- ignored.
  , _pageSetupPaperWidth :: Maybe String

     -- | Print scaling.
     --
     -- This attribute is restricted to values ranging from 10 to 400.
     -- This setting is overridden when fitToWidth and/or fitToHeight are in
     -- use.
  , _pageSetupScale :: Maybe Int

     -- | Use '_pageSetupFirstPageNumber' value for first page number, and do
     -- not auto number the pages.
  , _pageSetupUseFirstPageNumber :: Maybe Bool

     -- | Use the printer’s defaults settings for page setup values and don't
     -- use the default values specified in the schema.
     --
     -- Example: If dpi is not present or specified in the XML, the application
     -- must not assume 600dpi as specified in the schema as a default and
     -- instead must let the printer specify the default dpi.
  , _pageSetupUsePrinterDefaults :: Maybe Bool

    -- | Vertical print resolution of the device.
  , _pageSetupVerticalDpi :: Maybe Int
  }
  deriving (Show, Eq, Ord)

{-------------------------------------------------------------------------------
  Enumerations
-------------------------------------------------------------------------------}

-- | Cell comments
--
-- These enumerations specify how cell comments shall be displayed for paper
-- printing purposes.
--
-- See 18.18.5 "ST_CellComments (Cell Comments)" (p. 2676).
data CellComments =
    -- | Print cell comments as displayed
    CellCommentsAsDisplayed

    -- | Print cell comments at end of document
  | CellCommentsAtEnd

    -- | Do not print cell comments
  | CellCommentsNone
  deriving (Show, Eq, Ord)

-- | Print errors
--
-- This enumeration specifies how to display cells with errors when printing the
-- worksheet.
data PrintErrors =
     -- | Display cell errors as blank
     PrintErrorsBlank

     -- | Display cell errors as dashes
   | PrintErrorsDash

     -- | Display cell errors as displayed on screen
   | PrintErrorsDisplayed

     -- | Display cell errors as @#N/A@
   | PrintErrorsNA
  deriving (Show, Eq, Ord)

-- | Print orientation for this sheet
data Orientation =
    OrientationDefault
  | OrientationLandscape
  | OrientationPortrait
  deriving (Show, Eq, Ord)

-- | Specifies printed page order
data PageOrder =
    -- | Order pages vertically first, then move horizontally
    PageOrderDownThenOver

    -- | Order pages horizontally first, then move vertically
  | PageOrderOverThenDown
  deriving (Show, Eq, Ord)

-- | Paper size
data PaperSize =
    PaperA2                      -- ^ A2 paper (420 mm by 594 mm)
  | PaperA3                      -- ^ A3 paper (297 mm by 420 mm)
  | PaperA3Extra                 -- ^ A3 extra paper (322 mm by 445 mm)
  | PaperA3ExtraTransverse       -- ^ A3 extra transverse paper (322 mm by 445 mm)
  | PaperA3Transverse            -- ^ A3 transverse paper (297 mm by 420 mm)
  | PaperA4                      -- ^ A4 paper (210 mm by 297 mm)
  | PaperA4Extra                 -- ^ A4 extra paper (236 mm by 322 mm)
  | PaperA4Plus                  -- ^ A4 plus paper (210 mm by 330 mm)
  | PaperA4Small                 -- ^ A4 small paper (210 mm by 297 mm)
  | PaperA4Transverse            -- ^ A4 transverse paper (210 mm by 297 mm)
  | PaperA5                      -- ^ A5 paper (148 mm by 210 mm)
  | PaperA5Extra                 -- ^ A5 extra paper (174 mm by 235 mm)
  | PaperA5Transverse            -- ^ A5 transverse paper (148 mm by 210 mm)
  | PaperB4                      -- ^ B4 paper (250 mm by 353 mm)
  | PaperB5                      -- ^ B5 paper (176 mm by 250 mm)
  | PaperC                       -- ^ C paper (17 in. by 22 in.)
  | PaperD                       -- ^ D paper (22 in. by 34 in.)
  | PaperE                       -- ^ E paper (34 in. by 44 in.)
  | PaperExecutive               -- ^ Executive paper (7.25 in. by 10.5 in.)
  | PaperFanfoldGermanLegal      -- ^ German legal fanfold (8.5 in. by 13 in.)
  | PaperFanfoldGermanStandard   -- ^ German standard fanfold (8.5 in. by 12 in.)
  | PaperFanfoldUsStandard       -- ^ US standard fanfold (14.875 in. by 11 in.)
  | PaperFolio                   -- ^ Folio paper (8.5 in. by 13 in.)
  | PaperIsoB4                   -- ^ ISO B4 (250 mm by 353 mm)
  | PaperIsoB5Extra              -- ^ ISO B5 extra paper (201 mm by 276 mm)
  | PaperJapaneseDoublePostcard  -- ^ Japanese double postcard (200 mm by 148 mm)
  | PaperJisB5Transverse         -- ^ JIS B5 transverse paper (182 mm by 257 mm)
  | PaperLedger                  -- ^ Ledger paper (17 in. by 11 in.)
  | PaperLegal                   -- ^ Legal paper (8.5 in. by 14 in.)
  | PaperLegalExtra              -- ^ Legal extra paper (9.275 in. by 15 in.)
  | PaperLetter                  -- ^ Letter paper (8.5 in. by 11 in.)
  | PaperLetterExtra             -- ^ Letter extra paper (9.275 in. by 12 in.)
  | PaperLetterExtraTransverse   -- ^ Letter extra transverse paper (9.275 in. by 12 in.)
  | PaperLetterPlus              -- ^ Letter plus paper (8.5 in. by 12.69 in.)
  | PaperLetterSmall             -- ^ Letter small paper (8.5 in. by 11 in.)
  | PaperLetterTransverse        -- ^ Letter transverse paper (8.275 in. by 11 in.)
  | PaperNote                    -- ^ Note paper (8.5 in. by 11 in.)
  | PaperQuarto                  -- ^ Quarto paper (215 mm by 275 mm)
  | PaperStandard9_11            -- ^ Standard paper (9 in. by 11 in.)
  | PaperStandard10_11           -- ^ Standard paper (10 in. by 11 in.)
  | PaperStandard10_14           -- ^ Standard paper (10 in. by 14 in.)
  | PaperStandard11_17           -- ^ Standard paper (11 in. by 17 in.)
  | PaperStandard15_11           -- ^ Standard paper (15 in. by 11 in.)
  | PaperStatement               -- ^ Statement paper (5.5 in. by 8.5 in.)
  | PaperSuperA                  -- ^ SuperA/SuperA/A4 paper (227 mm by 356 mm)
  | PaperSuperB                  -- ^ SuperB/SuperB/A3 paper (305 mm by 487 mm)
  | PaperTabloid                 -- ^ Tabloid paper (11 in. by 17 in.)
  | PaperTabloidExtra            -- ^ Tabloid extra paper (11.69 in. by 18 in.)
  | Envelope6_3_4                -- ^ 6 3/4 envelope (3.625 in. by 6.5 in.)
  | Envelope9                    -- ^ #9 envelope (3.875 in. by 8.875 in.)
  | Envelope10                   -- ^ #10 envelope (4.125 in. by 9.5 in.)
  | Envelope11                   -- ^ #11 envelope (4.5 in. by 10.375 in.)
  | Envelope12                   -- ^ #12 envelope (4.75 in. by 11 in.)
  | Envelope14                   -- ^ #14 envelope (5 in. by 11.5 in.)
  | EnvelopeB4                   -- ^ B4 envelope (250 mm by 353 mm)
  | EnvelopeB5                   -- ^ B5 envelope (176 mm by 250 mm)
  | EnvelopeB6                   -- ^ B6 envelope (176 mm by 125 mm)
  | EnvelopeC3                   -- ^ C3 envelope (324 mm by 458 mm)
  | EnvelopeC4                   -- ^ C4 envelope (229 mm by 324 mm)
  | EnvelopeC5                   -- ^ C5 envelope (162 mm by 229 mm)
  | EnvelopeC6                   -- ^ C6 envelope (114 mm by 162 mm)
  | EnvelopeC65                  -- ^ C65 envelope (114 mm by 229 mm)
  | EnvelopeDL                   -- ^ DL envelope (110 mm by 220 mm)
  | EnvelopeInvite               -- ^ Invite envelope (220 mm by 220 mm)
  | EnvelopeItaly                -- ^ Italy envelope (110 mm by 230 mm)
  | EnvelopeMonarch              -- ^ Monarch envelope (3.875 in. by 7.5 in.).
  deriving (Show, Eq, Ord)

{-------------------------------------------------------------------------------
  Default instances
-------------------------------------------------------------------------------}

instance Default PageSetup where
  def = PageSetup {
      _pageSetupBlackAndWhite      = Nothing
    , _pageSetupCellComments       = Nothing
    , _pageSetupCopies             = Nothing
    , _pageSetupDraft              = Nothing
    , _pageSetupErrors             = Nothing
    , _pageSetupFirstPageNumber    = Nothing
    , _pageSetupFitToHeight        = Nothing
    , _pageSetupFitToWidth         = Nothing
    , _pageSetupHorizontalDpi      = Nothing
    , _pageSetupId                 = Nothing
    , _pageSetupOrientation        = Nothing
    , _pageSetupPageOrder          = Nothing
    , _pageSetupPaperHeight        = Nothing
    , _pageSetupPaperSize          = Nothing
    , _pageSetupPaperWidth         = Nothing
    , _pageSetupScale              = Nothing
    , _pageSetupUseFirstPageNumber = Nothing
    , _pageSetupUsePrinterDefaults = Nothing
    , _pageSetupVerticalDpi        = Nothing
   }

{-------------------------------------------------------------------------------
  Lenses
-------------------------------------------------------------------------------}

makeLenses ''PageSetup

{-------------------------------------------------------------------------------
  Rendering
-------------------------------------------------------------------------------}

-- | Render page setup
renderPageSetup :: PageSetup -> RawPageSetup
renderPageSetup = RawPageSetup . NodeElement . toElement "pageSetup"

-- | See @CT_PageSetup@, p. 4471
instance ToElement PageSetup where
  toElement nm PageSetup{..} = Element {
      elementName       = nm
    , elementNodes      = []
    , elementAttributes = Map.fromList . catMaybes $ [
          "paperSize"          .=? _pageSetupPaperSize
        , "paperHeight"        .=? _pageSetupPaperHeight
        , "paperWidth"         .=? _pageSetupPaperWidth
        , "scale"              .=? _pageSetupScale
        , "firstPageNumber"    .=? _pageSetupFirstPageNumber
        , "fitToWidth"         .=? _pageSetupFitToWidth
        , "fitToHeight"        .=? _pageSetupFitToHeight
        , "pageOrder"          .=? _pageSetupPageOrder
        , "orientation"        .=? _pageSetupOrientation
        , "usePrinterDefaults" .=? _pageSetupUsePrinterDefaults
        , "blackAndWhite"      .=? _pageSetupBlackAndWhite
        , "draft"              .=? _pageSetupDraft
        , "cellComments"       .=? _pageSetupCellComments
        , "useFirstPageNumber" .=? _pageSetupUseFirstPageNumber
        , "errors"             .=? _pageSetupErrors
        , "horizontalDpi"      .=? _pageSetupHorizontalDpi
        , "verticalDpi"        .=? _pageSetupVerticalDpi
        , "copies"             .=? _pageSetupCopies
        , "id"                 .=? _pageSetupId
        ]
    }

-- | See @ST_CellComments@, p. 4472
instance ToAttrVal CellComments where
  toAttrVal CellCommentsNone        = "none"
  toAttrVal CellCommentsAsDisplayed = "asDisplayed"
  toAttrVal CellCommentsAtEnd       = "atEnd"

-- | See @ST_PrintError@, p. 4472
instance ToAttrVal PrintErrors where
  toAttrVal PrintErrorsDisplayed = "displayed"
  toAttrVal PrintErrorsBlank     = "blank"
  toAttrVal PrintErrorsDash      = "dash"
  toAttrVal PrintErrorsNA        = "NA"

-- | See @ST_Orientation@, p. 4472
instance ToAttrVal Orientation where
  toAttrVal OrientationDefault   = "default"
  toAttrVal OrientationPortrait  = "portrait"
  toAttrVal OrientationLandscape = "landscape"

-- | See @ST_PageOrder@, p. 4472
instance ToAttrVal PageOrder where
  toAttrVal PageOrderDownThenOver = "downThenOver"
  toAttrVal PageOrderOverThenDown = "overThenDown"

-- | See @paperSize@ (attribute of @pageSetup@), p. 1837
instance ToAttrVal PaperSize where
  toAttrVal PaperLetter                 = "1"
  toAttrVal PaperLetterSmall            = "2"
  toAttrVal PaperTabloid                = "3"
  toAttrVal PaperLedger                 = "4"
  toAttrVal PaperLegal                  = "5"
  toAttrVal PaperStatement              = "6"
  toAttrVal PaperExecutive              = "7"
  toAttrVal PaperA3                     = "8"
  toAttrVal PaperA4                     = "9"
  toAttrVal PaperA4Small                = "10"
  toAttrVal PaperA5                     = "11"
  toAttrVal PaperB4                     = "12"
  toAttrVal PaperB5                     = "13"
  toAttrVal PaperFolio                  = "14"
  toAttrVal PaperQuarto                 = "15"
  toAttrVal PaperStandard10_14          = "16"
  toAttrVal PaperStandard11_17          = "17"
  toAttrVal PaperNote                   = "18"
  toAttrVal Envelope9                   = "19"
  toAttrVal Envelope10                  = "20"
  toAttrVal Envelope11                  = "21"
  toAttrVal Envelope12                  = "22"
  toAttrVal Envelope14                  = "23"
  toAttrVal PaperC                      = "24"
  toAttrVal PaperD                      = "25"
  toAttrVal PaperE                      = "26"
  toAttrVal EnvelopeDL                  = "27"
  toAttrVal EnvelopeC5                  = "28"
  toAttrVal EnvelopeC3                  = "29"
  toAttrVal EnvelopeC4                  = "30"
  toAttrVal EnvelopeC6                  = "31"
  toAttrVal EnvelopeC65                 = "32"
  toAttrVal EnvelopeB4                  = "33"
  toAttrVal EnvelopeB5                  = "34"
  toAttrVal EnvelopeB6                  = "35"
  toAttrVal EnvelopeItaly               = "36"
  toAttrVal EnvelopeMonarch             = "37"
  toAttrVal Envelope6_3_4               = "38"
  toAttrVal PaperFanfoldUsStandard      = "39"
  toAttrVal PaperFanfoldGermanStandard  = "40"
  toAttrVal PaperFanfoldGermanLegal     = "41"
  toAttrVal PaperIsoB4                  = "42"
  toAttrVal PaperJapaneseDoublePostcard = "43"
  toAttrVal PaperStandard9_11           = "44"
  toAttrVal PaperStandard10_11          = "45"
  toAttrVal PaperStandard15_11          = "46"
  toAttrVal EnvelopeInvite              = "47"
  toAttrVal PaperLetterExtra            = "50"
  toAttrVal PaperLegalExtra             = "51"
  toAttrVal PaperTabloidExtra           = "52"
  toAttrVal PaperA4Extra                = "53"
  toAttrVal PaperLetterTransverse       = "54"
  toAttrVal PaperA4Transverse           = "55"
  toAttrVal PaperLetterExtraTransverse  = "56"
  toAttrVal PaperSuperA                 = "57"
  toAttrVal PaperSuperB                 = "58"
  toAttrVal PaperLetterPlus             = "59"
  toAttrVal PaperA4Plus                 = "60"
  toAttrVal PaperA5Transverse           = "61"
  toAttrVal PaperJisB5Transverse        = "62"
  toAttrVal PaperA3Extra                = "63"
  toAttrVal PaperA5Extra                = "64"
  toAttrVal PaperIsoB5Extra             = "65"
  toAttrVal PaperA2                     = "66"
  toAttrVal PaperA3Transverse           = "67"
  toAttrVal PaperA3ExtraTransverse      = "68"
