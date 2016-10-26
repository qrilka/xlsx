{-# LANGUAGE NoMonomorphismRestriction #-}
{-# LANGUAGE OverloadedStrings         #-}
{-# LANGUAGE RecordWildCards           #-}
{-# LANGUAGE TemplateHaskell           #-}
module Codec.Xlsx.Types (
    -- * The main types
    Xlsx(..)
    , Styles(..)
    , DefinedNames(..)
    , ColumnsWidth(..)
    , PageSetup(..)
    , Worksheet(..)
    , CellMap
    , CellValue(..)
    , CellFormula(..)
    , Cell(..)
    , RowProperties (..)
    , Range
    -- * Lenses
    -- ** Workbook
    , xlSheets
    , xlStyles
    , xlDefinedNames
    , xlCustomProperties
    -- ** Worksheet
    , wsColumns
    , wsRowPropertiesMap
    , wsCells
    , wsDrawing
    , wsMerges
    , wsSheetViews
    , wsPageSetup
    , wsConditionalFormattings
    , wsDataValidations
    -- ** Cells
    , cellValue
    , cellStyle
    , cellComment
    , cellFormula
    -- * Style helpers
    , emptyStyles
    , renderStyleSheet
    , parseStyleSheet
    -- * Misc
    , simpleCellFormula
    , def
    , mkRange
    , fromRange
    , toRows
    , fromRows
    , module X
    ) where

import           Control.Exception                      (SomeException,
                                                         toException)
import           Control.Lens.TH
import qualified Data.ByteString.Lazy                   as L
import           Data.Default
import           Data.Function                          (on)
import           Data.List                              (groupBy)
import           Data.Map                               (Map)
import qualified Data.Map                               as M
import           Data.Maybe                             (catMaybes)
import           Data.Monoid                            ((<>))
import           Data.Text                              (Text)
import qualified Data.Text                              as T
import           Text.XML                               (Element (..), parseLBS,
                                                         renderLBS)
import           Text.XML.Cursor

import           Codec.Xlsx.Parser.Internal
import           Codec.Xlsx.Types.Comment               as X
import           Codec.Xlsx.Types.Common                as X
import           Codec.Xlsx.Types.ConditionalFormatting as X
import           Codec.Xlsx.Types.DataValidation        as X
import           Codec.Xlsx.Types.Drawing               as X
import           Codec.Xlsx.Types.Drawing.Chart         as X
import           Codec.Xlsx.Types.Drawing.Common        as X
import           Codec.Xlsx.Types.PageSetup             as X
import           Codec.Xlsx.Types.RichText              as X
import           Codec.Xlsx.Types.SheetViews            as X
import           Codec.Xlsx.Types.StyleSheet            as X
import           Codec.Xlsx.Types.Variant               as X
import           Codec.Xlsx.Writer.Internal

-- | Cell values include text, numbers and booleans,
-- standard includes date format also but actually dates
-- are represented by numbers with a date format assigned
-- to a cell containing it
data CellValue = CellText   Text
               | CellDouble Double
               | CellBool   Bool
               | CellRich   [RichTextRun]
               deriving (Eq, Show)

-- | Formula for the cell.
--
-- TODO: array. dataTable and shared formula types support
--
-- See 18.3.1.40 "f (Formula)" (p. 1636)
data CellFormula
    = NormalCellFormula
      { _cellfExpression    :: Formula
      , _cellfAssignsToName :: Bool
      -- ^ Specifies that this formula assigns a value to a name.
      , _cellfCalculate     :: Bool
      -- ^ Indicates that this formula needs to be recalculated
      -- the next time calculation is performed.
      -- [/Example/: This is always set on volatile functions,
      -- like =RAND(), and circular references. /end example/]
      } deriving (Eq, Show)

simpleCellFormula :: Text -> CellFormula
simpleCellFormula expr = NormalCellFormula
    { _cellfExpression    = Formula expr
    , _cellfAssignsToName = False
    , _cellfCalculate     = False
    }

-- | Currently cell details include only cell values and style ids
-- (e.g. formulas from @\<f\>@ and inline strings from @\<is\>@
-- subelements are ignored)
data Cell = Cell
    { _cellStyle   :: Maybe Int
    , _cellValue   :: Maybe CellValue
    , _cellComment :: Maybe Comment
    , _cellFormula :: Maybe CellFormula
    } deriving (Eq, Show)

makeLenses ''Cell

instance Default Cell where
    def = Cell Nothing Nothing Nothing Nothing

-- | Map containing cell values which are indexed by row and column
-- if you need to use more traditional (x,y) indexing please you could
-- use corresponding accessors from ''Codec.Xlsx.Lens''
type CellMap = Map (Int, Int) Cell

data RowProperties = RowProps { rowHeight :: Maybe Double, rowStyle::Maybe Int}
                   deriving (Read,Eq,Show,Ord)

-- | Column range (from cwMin to cwMax) width
data ColumnsWidth = ColumnsWidth
    { cwMin   :: Int
    , cwMax   :: Int
    , cwWidth :: Double
    , cwStyle :: Int
    } deriving (Eq, Show)

instance FromCursor ColumnsWidth where
    fromCursor c = do
      cwMin <- decimal =<< attribute "min" c
      cwMax <- decimal =<< attribute "max" c
      cwWidth <- rational =<< attribute "width" c
      cwStyle <- decimal =<< attribute "style" c
      return ColumnsWidth{..}

-- | Excel range (e.g. @D13:H14@)
type Range = Text

-- | Xlsx worksheet
data Worksheet = Worksheet
    { _wsColumns                :: [ColumnsWidth]          -- ^ column widths
    , _wsRowPropertiesMap       :: Map Int RowProperties   -- ^ custom row properties (height, style) map
    , _wsCells                  :: CellMap                 -- ^ data mapped by (row, column) pairs
    , _wsDrawing                :: Maybe Drawing           -- ^ SpreadsheetML Drawing
    , _wsMerges                 :: [Range]                 -- ^ list of cell merges
    , _wsSheetViews             :: Maybe [SheetView]
    , _wsPageSetup              :: Maybe PageSetup
    , _wsConditionalFormattings :: Map SqRef ConditionalFormatting
    , _wsDataValidations        :: Map SqRef DataValidation
    } deriving (Eq, Show)

makeLenses ''Worksheet

instance Default Worksheet where
    def = Worksheet [] M.empty M.empty Nothing [] Nothing Nothing M.empty M.empty

newtype Styles = Styles {unStyles :: L.ByteString}
            deriving (Eq, Show)

-- | Structured representation of Xlsx file (currently a subset of its contents)
data Xlsx = Xlsx
    { _xlSheets           :: [(Text, Worksheet)]
    , _xlStyles           :: Styles
    , _xlDefinedNames     :: DefinedNames
    , _xlCustomProperties :: Map Text Variant
    } deriving (Eq, Show)

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
  deriving (Eq, Show)

makeLenses ''Xlsx

instance Default Xlsx where
    def = Xlsx [] emptyStyles def M.empty

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

-- | Render range
--
-- > mkRange (2, 4) (6, 8) == "D2:H6"
mkRange :: (Int, Int) -> (Int, Int) -> Range
mkRange fr to = T.concat [mkCellRef fr, T.pack ":", mkCellRef to]

-- | reverse to 'mkRange'
--
-- /Warning:/ the function isn't total and will throw an error if
-- incorrect value will get passed
fromRange :: Range -> ((Int, Int), (Int, Int))
fromRange t = case T.split (==':') t of
    [from, to] -> (fromCellRef from, fromCellRef to)
    _ -> error $ "invalid range " <> show t

{-------------------------------------------------------------------------------
  Parsing
-------------------------------------------------------------------------------}

instance FromCursor CellFormula where
    fromCursor cur = do
        t <- fromAttributeDef "t" "normal" cur
        typedCellFormula t cur

typedCellFormula :: Text -> Cursor -> [CellFormula]
typedCellFormula "normal" cur = do
    _cellfExpression <- fromCursor cur
    _cellfAssignsToName <- fromAttributeDef "bx" False cur
    _cellfCalculate <- fromAttributeDef "ca" False cur
    return NormalCellFormula{..}
typedCellFormula _ _ = fail "parseable cell formula type was not found"

{-------------------------------------------------------------------------------
  Rendering
-------------------------------------------------------------------------------}

instance ToElement CellFormula where
    toElement nm NormalCellFormula{..} =
        let formulaEl = toElement nm _cellfExpression
        in formulaEl
           { elementAttributes =
                   M.fromList $ catMaybes [ "bx" .=? justTrue _cellfAssignsToName
                                          , "ca" .=? justTrue _cellfCalculate ]
           }
