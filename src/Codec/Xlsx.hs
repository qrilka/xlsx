-- | This module provides solution for parsing and writing Microsoft
-- Open Office XML Workbook format i.e. *.xlsx files
--
-- As a simple example you could read cell B3 from the 1st sheet of workbook \"report.xlsx\"
-- using the following code:
--
-- > {-# LANGUAGE OverloadedStrings #-}
-- > module Read where
-- > import Codec.Xlsx
-- > import qualified Data.ByteString.Lazy as L
-- > import Control.Lens
-- >
-- > main :: IO ()
-- > main = do
-- >   bs <- L.readFile "report.xlsx"
-- >   let value = toXlsx bs ^? ixSheet "List1" .
-- >               ixCell (3,2) . cellValue . _Just
-- >   putStrLn $ "Cell B3 contains " ++ show value
--
-- And the following example module shows a way to construct and write xlsx file
--
-- > {-# LANGUAGE OverloadedStrings #-}
-- > module Write where
-- > import Codec.Xlsx
-- > import Control.Lens
-- > import qualified Data.ByteString.Lazy as L
-- > import Data.Time.Clock.POSIX
-- >
-- > main :: IO ()
-- > main = do
-- >   ct <- getPOSIXTime
-- >   let
-- >       sheet = def & cellValueAt (1,2) ?~ CellDouble 42.0
-- >                   & cellValueAt (3,2) ?~ CellText "foo"
-- >       xlsx = def & atSheet "List1" ?~ sheet
-- >   L.writeFile "example.xlsx" $ fromXlsx ct xlsx
module Codec.Xlsx
    ( module X
    ) where

import Codec.Xlsx.Types as X
import Codec.Xlsx.Parser as X
import Codec.Xlsx.Writer as X
import Codec.Xlsx.Lens as X
