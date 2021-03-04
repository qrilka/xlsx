{-# LANGUAGE OverloadedStrings #-}
{-# LANGUAGE LambdaCase #-}
module Main (main) where

import Prelude hiding (foldl)
import Codec.Xlsx
import Criterion.Main
import qualified Data.ByteString as BS
import qualified Data.ByteString.Lazy as LB
import Conduit
import Codec.Xlsx.Parser.Stream
import Data.Text(Text)
import qualified Data.Text as Text
import qualified Data.Text.IO as Text
import qualified Data.Conduit.Combinators as C

celVal :: CellValue -> Text
celVal  = \case
  CellText x -> x
  CellDouble x -> Text.pack $ show x
  CellBool x -> Text.pack $ show x
  CellRich x -> Text.pack $ show x
  CellError x -> Text.pack $ show x

format :: SheetItem -> Text
format x = Text.pack $ show (_si_sheet x, _si_row_index x, maybe "" celVal . _cellValue <$> _si_cell_row x)

main :: IO ()
main = do
  let filename =
        "../temp/policy-bordereau-template.xlsx"
        -- "data/testInput.xlsx"
        -- "data/6000.rows.x.26.cols.xlsx"
  bs <- BS.readFile filename
  state <- runResourceT $ runConduit $ sourceFile filename .| parseSharedStrings
  print state

-- -  x <- runResourceT $ runConduit $ readXlsx (sourceFile filename) .| C.filter (("sheet1.xml" ==) . _si_sheet) .| C.foldl (\a b -> b : a) []
-- -  Text.putStrLn $ Text.intercalate "\n" $ format <$> x
-- -  pure ()
