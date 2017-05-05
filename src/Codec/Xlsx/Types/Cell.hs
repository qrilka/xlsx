{-# LANGUAGE OverloadedStrings #-}
{-# LANGUAGE RecordWildCards #-}
{-# LANGUAGE TemplateHaskell #-}
{-# LANGUAGE DeriveGeneric #-}
module Codec.Xlsx.Types.Cell
  ( CellFormula(..)
  , simpleCellFormula
  , Cell(..)
  , cellStyle
  , cellValue
  , cellComment
  , cellFormula
  , CellMap
  ) where

import Control.Lens.TH (makeLenses)
import Data.Default
import Data.Map (Map)
import qualified Data.Map as M
import Data.Maybe (catMaybes)
import Data.Text (Text)
import Text.XML
import Text.XML.Cursor

import Codec.Xlsx.Parser.Internal
import Codec.Xlsx.Types.Comment
import Codec.Xlsx.Types.Common
import Codec.Xlsx.Writer.Internal

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
      } deriving (Eq, Show, Generic)

simpleCellFormula :: Text -> CellFormula
simpleCellFormula expr = NormalCellFormula
    { _cellfExpression    = Formula expr
    , _cellfAssignsToName = False
    , _cellfCalculate     = False
    }

-- | Currently cell details include cell values, style ids and cell
-- formulas (inline strings from @\<is\>@ subelements are ignored)
data Cell = Cell
    { _cellStyle   :: Maybe Int
    , _cellValue   :: Maybe CellValue
    , _cellComment :: Maybe Comment
    , _cellFormula :: Maybe CellFormula
    } deriving (Eq, Show, Generic)

instance Default Cell where
    def = Cell Nothing Nothing Nothing Nothing

makeLenses ''Cell

-- | Map containing cell values which are indexed by row and column
-- if you need to use more traditional (x,y) indexing please you could
-- use corresponding accessors from ''Codec.Xlsx.Lens''
type CellMap = Map (Int, Int) Cell

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
  return NormalCellFormula {..}
typedCellFormula _ _ = fail "parseable cell formula type was not found"

{-------------------------------------------------------------------------------
  Rendering
-------------------------------------------------------------------------------}

instance ToElement CellFormula where
  toElement nm NormalCellFormula {..} =
    let formulaEl = toElement nm _cellfExpression
    in formulaEl
       { elementAttributes =
           M.fromList $
           catMaybes
             [ "bx" .=? justTrue _cellfAssignsToName
             , "ca" .=? justTrue _cellfCalculate
             ]
       }
