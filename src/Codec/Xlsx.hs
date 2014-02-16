-- | This module provides solution for parsing and writing MIcrosoft
-- Open Office XML Workbook format i.e. *.xlsx files
--
-- As a simple example you could read cell B3 from the 1st sheet of workbook "report.xlsx"
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
module Codec.Xlsx
    ( module X
    ) where

import Codec.Xlsx.Types as X
import Codec.Xlsx.Parser as X
import Codec.Xlsx.Writer as X
import Codec.Xlsx.Lens as X
