{-# LANGUAGE OverloadedStrings #-}
module Main (main) where

import Prelude hiding (foldl)
import Codec.Xlsx
import Criterion.Main
import qualified Data.ByteString as BS
import qualified Data.ByteString.Lazy as LB
import Conduit
import Data.Conduit.Combinators hiding (print)
import Codec.Xlsx.Parser.Stream

main :: IO ()
main = do
  let filename = "data/testInput.xlsx"
        -- "data/6000.rows.x.26.cols.xlsx"
  bs <- BS.readFile filename
  x <- runResourceT $ runConduit $ sourceFile filename .| readXlsx .| foldl (\a b -> b : a) []
  print x
