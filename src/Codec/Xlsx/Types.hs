{-# LANGUAGE NoMonomorphismRestriction #-}
{-# LANGUAGE OverloadedStrings #-}
{-# LANGUAGE RecordWildCards #-}
{-# LANGUAGE TemplateHaskell #-}
{-# LANGUAGE DeriveGeneric #-}
module Codec.Xlsx.Types (
    -- * The main types
    Xlsx(..)
    , Styles(..)
    , DefinedNames(..)
    , ColumnsProperties(..)
    , PageSetup(..)
    , Worksheet(..)
    , CellMap
    , CellValue(..)
    , CellFormula(..)
    , FormulaExpression(..)
    , Cell.SharedFormulaIndex(..)
    , Cell.SharedFormulaOptions(..)
    , Cell(..)
    , RowHeight(..)
    , RowProperties (..)
    -- * Lenses
    -- ** Workbook
    , xlSheets
    , xlStyles
    , xlDefinedNames
    , xlCustomProperties
    , xlDateBase
    -- ** Worksheet
    , wsColumnsProperties
    , wsRowPropertiesMap
    , wsCells
    , wsDrawing
    , wsMerges
    , wsSheetViews
    , wsPageSetup
    , wsConditionalFormattings
    , wsDataValidations
    , wsPivotTables
    , wsAutoFilter
    , wsTables
    , wsProtection
    , wsSharedFormulas
    -- ** Cells
    , Cell.cellValue
    , Cell.cellStyle
    , Cell.cellComment
    , Cell.cellFormula
    -- * Style helpers
    , emptyStyles
    , renderStyleSheet
    , parseStyleSheet
    -- * Misc
    , simpleCellFormula
    , sharedFormulaByIndex
    , def
    , toRows
    , fromRows
    , module X
    ) where

import Control.Exception (SomeException, toException)
import Control.Lens.TH
import Control.DeepSeq (NFData)
import qualified Data.ByteString.Lazy as L
import Data.Default
import Data.Function (on)
import Data.List (groupBy)
import Data.Map (Map)
import qualified Data.Map as M
import Data.Maybe (catMaybes, isJust)
import Data.Text (Text)
import GHC.Generics (Generic)
import Text.XML (parseLBS, renderLBS)
import Text.XML.Cursor

import Codec.Xlsx.Parser.Internal
import Codec.Xlsx.Types.AutoFilter as X
import Codec.Xlsx.Types.Cell as Cell
import Codec.Xlsx.Types.Comment as X
import Codec.Xlsx.Types.Common as X
import Codec.Xlsx.Types.ConditionalFormatting as X
import Codec.Xlsx.Types.DataValidation as X
import Codec.Xlsx.Types.Drawing as X
import Codec.Xlsx.Types.Drawing.Chart as X
import Codec.Xlsx.Types.Drawing.Common as X
import Codec.Xlsx.Types.PageSetup as X
import Codec.Xlsx.Types.PivotTable as X
import Codec.Xlsx.Types.Protection as X
import Codec.Xlsx.Types.RichText as X
import Codec.Xlsx.Types.SheetState as X
import Codec.Xlsx.Types.SheetViews as X
import Codec.Xlsx.Types.StyleSheet as X
import Codec.Xlsx.Types.Table as X
import Codec.Xlsx.Types.Variant as X
import Codec.Xlsx.Writer.Internal

-- | Height of a row in points (1/72in)
data RowHeight
  = CustomHeight    !Double
    -- ^ Row height is set by the user
  | AutomaticHeight !Double
    -- ^ Row height is set automatically by the program
  deriving (Eq, Ord, Show, Read, Generic)
instance NFData RowHeight

-- | Properties of a row. See §18.3.1.73 "row (Row)" for more details
data RowProperties = RowProps
  { rowHeight       :: Maybe RowHeight
    -- ^ Row height in points
  , rowStyle        :: Maybe Int
    -- ^ Style to be applied to row
  , rowHidden       :: Bool
    -- ^ Whether row is visible or not
  } deriving (Eq, Ord, Show, Read, Generic)
instance NFData RowProperties

instance Default RowProperties where
  def = RowProps { rowHeight       = Nothing
                 , rowStyle        = Nothing
                 , rowHidden       = False
                 }

-- | Column range (from cwMin to cwMax) properties
data ColumnsProperties = ColumnsProperties
  { cpMin :: Int
  -- ^ First column affected by this 'ColumnWidth' record.
  , cpMax :: Int
  -- ^ Last column affected by this 'ColumnWidth' record.
  , cpWidth :: Maybe Double
  -- ^ Column width measured as the number of characters of the
  -- maximum digit width of the numbers 0, 1, 2, ..., 9 as rendered in
  -- the normal style's font.
  --
  -- See longer description in Section 18.3.1.13 "col (Column Width &
  -- Formatting)" (p. 1605)
  , cpStyle :: Maybe Int
  -- ^ Default style for the affected column(s). Affects cells not yet
  -- allocated in the column(s).  In other words, this style applies
  -- to new columns.
  , cpHidden :: Bool
  -- ^ Flag indicating if the affected column(s) are hidden on this
  -- worksheet.
  , cpCollapsed :: Bool
  -- ^ Flag indicating if the outlining of the affected column(s) is
  -- in the collapsed state.
  , cpBestFit :: Bool
  -- ^ Flag indicating if the specified column(s) is set to 'best
  -- fit'.
  } deriving (Eq, Show, Generic)
instance NFData ColumnsProperties

instance FromCursor ColumnsProperties where
  fromCursor c = do
    cpMin <- fromAttribute "min" c
    cpMax <- fromAttribute "max" c
    cpWidth <- maybeAttribute "width" c
    cpStyle <- maybeAttribute "style" c
    cpHidden <- fromAttributeDef "hidden" False c
    cpCollapsed <- fromAttributeDef "collapsed" False c
    cpBestFit <- fromAttributeDef "bestFit" False c
    return ColumnsProperties {..}

instance FromXenoNode ColumnsProperties where
  fromXenoNode root = parseAttributes root $ do
    cpMin <- fromAttr "min"
    cpMax <- fromAttr "max"
    cpWidth <- maybeAttr "width"
    cpStyle <- maybeAttr "style"
    cpHidden <- fromAttrDef "hidden" False
    cpCollapsed <- fromAttrDef "collapsed" False
    cpBestFit <- fromAttrDef "bestFit" False
    return ColumnsProperties {..}

-- | Xlsx worksheet
data Worksheet = Worksheet
  { _wsColumnsProperties :: [ColumnsProperties] -- ^ column widths
  , _wsRowPropertiesMap :: Map Int RowProperties -- ^ custom row properties (height, style) map
  , _wsCells :: CellMap -- ^ data mapped by (row, column) pairs
  , _wsDrawing :: Maybe Drawing -- ^ SpreadsheetML Drawing
  , _wsMerges :: [Range] -- ^ list of cell merges
  , _wsSheetViews :: Maybe [SheetView]
  , _wsPageSetup :: Maybe PageSetup
  , _wsConditionalFormattings :: Map SqRef ConditionalFormatting
  , _wsDataValidations :: Map SqRef DataValidation
  , _wsPivotTables :: [PivotTable]
  , _wsAutoFilter :: Maybe AutoFilter
  , _wsTables :: [Table]
  , _wsProtection :: Maybe SheetProtection
  , _wsSharedFormulas :: Map SharedFormulaIndex SharedFormulaOptions
  } deriving (Eq, Show, Generic)
instance NFData Worksheet

makeLenses ''Worksheet

instance Default Worksheet where
  def =
    Worksheet
    { _wsColumnsProperties = []
    , _wsRowPropertiesMap = M.empty
    , _wsCells = M.empty
    , _wsDrawing = Nothing
    , _wsMerges = []
    , _wsSheetViews = Nothing
    , _wsPageSetup = Nothing
    , _wsConditionalFormattings = M.empty
    , _wsDataValidations = M.empty
    , _wsPivotTables = []
    , _wsAutoFilter = Nothing
    , _wsTables = []
    , _wsProtection = Nothing
    , _wsSharedFormulas = M.empty
    }

newtype Styles = Styles {unStyles :: L.ByteString}
            deriving (Eq, Show, Generic)
instance NFData Styles

-- | Structured representation of Xlsx file (currently a subset of its contents)
data Xlsx = Xlsx
  { _xlSheets :: [(Text, SheetState, Worksheet)]
  , _xlStyles :: Styles
  , _xlDefinedNames :: DefinedNames
  , _xlCustomProperties :: Map Text Variant
  , _xlDateBase :: DateBase
  -- ^ date base to use when converting serial value (i.e. 'CellDouble d')
  -- into date-time. Default value is 'DateBase1900'
  --
  -- See also 18.17.4.1 "Date Conversion for Serial Date-Times" (p. 2067)
  } deriving (Eq, Show, Generic)
instance NFData Xlsx


-- | Defined names
--
-- Each defined name consists of a name, an optional local sheet ID, and a value.
--
-- This element defines the collection of defined names for this workbook.
-- Defined names are descriptive names to represent cells, ranges of cells,
-- formulas, or constant values. Defined names can be used to represent a range
-- on any worksheet.
--
-- Excel also defines a number of reserved names with a special interpretation:
--
-- * @_xlnm.Print_Area@ specifies the workbook's print area.
--   Example value: @SheetName!$A:$A,SheetName!$1:$4@
-- * @_xlnm.Print_Titles@ specifies the row(s) or column(s) to repeat
--   at the top of each printed page.
-- * @_xlnm.Sheet_Title@:refers to a sheet title.
--
-- and others. See Section 18.2.6, "definedNames (Defined Names)" (p. 1728) of
-- the spec (second edition).
--
-- NOTE: Right now this is only a minimal implementation of defined names.
newtype DefinedNames = DefinedNames [(Text, Maybe Text, Text)]
  deriving (Eq, Show, Generic)
instance NFData DefinedNames

makeLenses ''Xlsx

instance Default Xlsx where
    def = Xlsx [] emptyStyles def M.empty DateBase1900

instance Default DefinedNames where
    def = DefinedNames []

emptyStyles :: Styles
emptyStyles = Styles "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\
\<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"></styleSheet>"

-- | Render 'StyleSheet'
--
-- This is used to render a structured 'StyleSheet' into a raw XML 'Styles'
-- document. Actually /replacing/ 'Styles' with 'StyleSheet' would mean we
-- would need to write a /parser/ for 'StyleSheet' as well (and would moreover
-- require that we support the full style sheet specification, which is still
-- quite a bit of work).
renderStyleSheet :: StyleSheet -> Styles
renderStyleSheet = Styles . renderLBS def . toDocument

-- | Parse 'StyleSheet'
--
-- This is used to parse raw 'Styles' into structured 'StyleSheet'
-- currently not all of the style sheet specification is supported
-- so parser (and the data model) is to be completed
parseStyleSheet :: Styles -> Either SomeException StyleSheet
parseStyleSheet (Styles bs) = parseLBS def bs >>= parseDoc
  where
    parseDoc doc = case fromCursor (fromDocument doc) of
      [stylesheet] -> Right stylesheet
      _ -> Left . toException $ ParseException "Could not parse style sheets"

-- | converts cells mapped by (row, column) into rows which contain
-- row index and cells as pairs of column indices and cell values
toRows :: CellMap -> [(Int, [(Int, Cell)])]
toRows cells =
    map extractRow $ groupBy ((==) `on` (fst . fst)) $ M.toList cells
  where
    extractRow row@(((x,_),_):_) =
        (x, map (\((_,y),v) -> (y,v)) row)
    extractRow _ = error "invalid CellMap row"

-- | reverse to 'toRows'
fromRows :: [(Int, [(Int, Cell)])] -> CellMap
fromRows rows = M.fromList $ concatMap mapRow rows
  where
    mapRow (r, cells) = map (\(c, v) -> ((r, c), v)) cells


instance ToElement ColumnsProperties where
  toElement nm ColumnsProperties {..} = leafElement nm attrs
    where
      attrs =
        ["min" .= cpMin, "max" .= cpMax] ++
        catMaybes
          [ "style" .=? (justNonDef 0 =<< cpStyle)
          , "width" .=? cpWidth
          , "customWidth" .=? justTrue (isJust cpWidth)
          , "hidden" .=? justTrue cpHidden
          , "collapsed" .=? justTrue cpCollapsed
          , "bestFit" .=? justTrue cpBestFit
          ]
