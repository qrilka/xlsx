{-# LANGUAGE OverloadedStrings   #-}
{-# LANGUAGE ScopedTypeVariables #-}

module Stylesheet where

import Codec.Xlsx (Cell (..), CellValue (..), Worksheet, atCell, atSheet,
                   cellValueAt, def, fromXlsx, renderStyleSheet, xlStyles)
import Codec.Xlsx.Types (Alignment, Border, CellXf (..), Fill, Font,
                         LineStyle (LineStyleThin),
                         PatternType (PatternTypeSolid), StyleSheet)
import Codec.Xlsx.Types.StyleSheet (alignmentWrapText, borderBottom, borderLeft,
                                    borderRight, borderStyleColor,
                                    borderStyleLine, borderTop, cellXfAlignment,
                                    cellXfApplyBorder, cellXfApplyFill,
                                    cellXfApplyFont, cellXfBorderId,
                                    cellXfFillId, cellXfFontId, colorARGB,
                                    fillPattern, fillPatternFgColor,
                                    fillPatternType, fontColor, fontItalic,
                                    fontSize, minimalStyleSheet,
                                    styleSheetBorders, styleSheetCellXfs,
                                    styleSheetFills, styleSheetFonts)
import Data.Text (Text)
import Data.Time.Clock.POSIX (getPOSIXTime)
import Lens.Micro (ix, (%~), (&), (.~), (?~))

import qualified Data.ByteString.Lazy as L

main :: IO ()
main = do
    ct <- getPOSIXTime
    let name :: String = "example"
        stylesheetEnv = carStylesheetEnv
        ws = def
                 & addHeaders stylesheetEnv
                 & addNote stylesheetEnv
                 & addItems
        xlsx = def
                 & atSheet "List1" ?~ ws
                 & xlStyles .~ renderStyleSheet (sseStyleSheet stylesheetEnv)

    L.writeFile (name <> ".xlsx") $ fromXlsx ct xlsx

-- | In terms of styling, we'll ensure that headers will:
--   1. Have white text (#FFFFFF)
--   2. Be filled with color #1A2B44 (dark blue)
--   3. Have border around each cell
addHeaders :: StyleSheetEnv -> Worksheet -> Worksheet
addHeaders stylesheetEnv sheet =
    let headerStyleId = Just $ sseHeaderStyleId stylesheetEnv
    in sheet
        & atCell (1, 1) ?~ Cell headerStyleId (Just $ CellText "Make") Nothing Nothing                 -- A simple text
        & atCell (1, 2) ?~ Cell headerStyleId (Just $ CellText "Model") Nothing Nothing                -- An Integer
        & atCell (1, 3) ?~ Cell headerStyleId (Just $ CellText "Year") Nothing Nothing                 -- Data for Year will be a Maybe value
        & atCell (1, 4) ?~ Cell headerStyleId (Just $ CellText "Sales (in thousands)") Nothing Nothing -- Decimal number

-- | We'll add some note about what each column represents.
-- Each cell of this row will:
--   1. Be filled with the color #EAF4FE (light blue)
--   2. Have italicized text
--   3. Contain wrapped text
--   4. Have border around each cell
addNote :: StyleSheetEnv -> Worksheet -> Worksheet
addNote stylesheetEnv sheet =
    let noteStyleId = Just $ sseNoteStyleId stylesheetEnv
    in sheet
        & atCell (2, 1) ?~ Cell noteStyleId (Just $ CellText "Manufacturer (e.g., Toyota, Ford, Kia)") Nothing Nothing
        & atCell (2, 2) ?~ Cell noteStyleId (Just $ CellText "Car name/variant (e.g., Corolla, Civic)") Nothing Nothing
        & atCell (2, 3) ?~ Cell noteStyleId (Just $ CellText "Year of sale/production") Nothing Nothing
        & atCell (2, 4) ?~ Cell noteStyleId (Just $ CellText "Number of units sold in that year (000s)") Nothing Nothing

-- We'll make a 3 x 4 sheet with car data
addItems :: Worksheet -> Worksheet
addItems sheet =
    sheet
        & cellValueAt (3, 1) ?~ CellText "Toyota"
        & cellValueAt (3, 2) ?~ CellText "Corolla"
        & maybe id (\year -> cellValueAt (3, 3) ?~ CellDouble year) (Just 2021)
        & cellValueAt (3, 4) ?~ CellDouble 120.5

        & cellValueAt (4, 1) ?~ CellText "Ford"
        & cellValueAt (4, 2) ?~ CellText "Mustang"
        & maybe id (\year -> cellValueAt (4, 3) ?~ CellDouble year) Nothing
        & cellValueAt (4, 4) ?~ CellDouble 85.3

        & cellValueAt (5, 1) ?~ CellText "Audi"
        & cellValueAt (5, 2) ?~ CellText "A4"
        & maybe id (\year -> cellValueAt (5, 3) ?~ CellDouble year) (Just 2019)
        & cellValueAt (5, 4) ?~ CellDouble 34.9

-- | Data type to capture the stylesheet
-- along with references to CellXfs.
--
-- We define a data type like this in order to
-- pass around references to different @CellXf@s in
-- a much more flexible way. Because if we don't
-- then we will have to refer the different @CellXf@
-- with an @Int@ and then we will have arithmetic all
-- over the place.
data StyleSheetEnv = StyleSheetEnv
    { sseStyleSheet    :: !StyleSheet
    , sseHeaderStyleId :: !Int
    , sseNoteStyleId   :: !Int
    }

-- | High level function that builds our stylesheet
carStylesheetEnv :: StyleSheetEnv
carStylesheetEnv =
    let ss =
          minimalStyleSheet
              -- When we append @[headerFill, noteFill]@ to styleSheetCellXfs, @headerFill@ ends up at
              -- index 2 and @noteFill@ at index 3 (because of @fillNone@ and @fillGray125@ already existing at
              -- index 0 and 1 respectively). These two fills are required and reserved by excel.
              -- Later, when a cell uses cellStyle = Just 1, that means:
              --   look up styleSheetCellXfs !! 1 = headerCellXf → which points to fillId 2 (@headerFill@).
              & styleSheetCellXfs %~ (++ [headerCellXf, noteCellXf])
              & styleSheetFills %~ (++ [headerFill, noteFill])
              & styleSheetFonts . ix 0 .~ defaultFontSize
              & styleSheetFonts %~ (++ [headerFont, noteFont])
              & styleSheetBorders %~ (++ [thinSolidBorder])
    in StyleSheetEnv ss 1 2

-- | CellXf that references @headerFill@.
-- NOTE: FillId starts from 0 inside styleSheetFills.
-- minimalStyleSheet already contains 2 mandatory default fills at index 0 and 1 (required
-- and reserved by excel). Therefore, we need to define @FillId@s that start with ID 2.
headerCellXf :: CellXf
headerCellXf =
    def
        & cellXfFillId ?~ 2
        & cellXfApplyFill ?~ True
        & cellXfFontId ?~ 1
        & cellXfApplyFont ?~ True
        & cellXfBorderId ?~ 1
        & cellXfApplyBorder ?~ True

-- | CellXf that references @noteFill@.
noteCellXf :: CellXf
noteCellXf =
    def
        & cellXfFillId ?~ 3
        & cellXfApplyFill ?~ True
        & cellXfFontId ?~ 2
        & cellXfApplyFont ?~ True
        & cellXfAlignment ?~ wrapAlign
        & cellXfBorderId ?~ 1
        & cellXfApplyBorder ?~ True


wrapAlign :: Alignment
wrapAlign = def & alignmentWrapText ?~ True

thinSolidBorder :: Border
thinSolidBorder =
    def
        & borderTop ?~ solidSide
        & borderBottom ?~ solidSide
        & borderLeft ?~ solidSide
        & borderRight ?~ solidSide
  where
    solidSide =
        def
            & borderStyleColor ?~ (def & colorARGB ?~ "000000")
            & borderStyleLine ?~ LineStyleThin

-- | Because the default font size provided along with @minimalStyleSheet@ is
-- set to 11, we need to define our own in order to override that.
defaultFontSize :: Font
defaultFontSize = def & fontSize ?~ 12

-- | Font color used by the header row
headerFont :: Font
headerFont =
    def
        & fontColor ?~ (def & colorARGB ?~ "FFFFFF")

-- | Font color and style used by the note row
noteFont :: Font
noteFont =
    def
        & fontColor ?~ (def & colorARGB ?~ "45779B")
        & fontItalic ?~ True

noteFill, headerFill :: Fill
noteFill = solidFill "EAF4FE"
headerFill = solidFill "1A2B44"

solidFill :: Text -> Fill
solidFill color =
    def
        & fillPattern
            ?~ ( def
                    & fillPatternFgColor ?~ (def & colorARGB ?~ color)
                    & fillPatternType ?~ PatternTypeSolid
               )
