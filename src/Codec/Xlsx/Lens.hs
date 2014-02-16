module Codec.Xlsx.Lens
    ( ixSheet
    , ixCell ) where

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

ixCell :: Applicative f
       => (Int, Int)
       -> (Cell -> f Cell)
       -> Worksheet
       -> f Worksheet
ixCell i = wsCells . ix i
