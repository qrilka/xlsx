module Codec.Xlsx.Lens
    ( ixSheet
    , atSheet
    , ixCell
    , atCell
    , cellValueAt
 ) where

import Codec.Xlsx.Types
import Control.Applicative
import Control.Lens
import Data.Text

ixSheet :: Applicative f
        => Text
        -> (Worksheet -> f Worksheet)
        -> Xlsx
        -> f Xlsx
ixSheet s = xlSheets . ix s

atSheet :: Functor f
        => Text
        -> (Maybe Worksheet -> f (Maybe Worksheet))
        -> Xlsx
        -> f Xlsx
atSheet s = xlSheets . at s

ixCell :: Applicative f
       => (Int, Int)
       -> (Cell -> f Cell)
       -> Worksheet
       -> f Worksheet
ixCell i = wsCells . ix i

atCell :: Functor f
       => (Int, Int)
       -> (Maybe Cell -> f (Maybe Cell))
       -> Worksheet
       -> f Worksheet
atCell i = wsCells . at i

cellValueAt :: Functor f
           => (Int, Int)
           -> (Maybe CellValue -> f (Maybe CellValue))
           -> Worksheet
           -> f Worksheet
cellValueAt i = atCell i . non def . cellValue
