{-# LANGUAGE OverloadedStrings #-}
module Main (main) where

import Codec.Xlsx
import Criterion.Main
import qualified Data.ByteString as BS
import qualified Data.ByteString.Lazy as LB

main :: IO ()
main = do
  let filename = "data/testInput.xlsx"
        -- "data/6000.rows.x.26.cols.xlsx"
  bs <- BS.readFile filename
  let bs' = LB.fromStrict bs
  defaultMain
    [ bgroup
        "readFile"
        [ bench "with xlsx" $ nf toXlsx bs'
        , bench "with xlsx fast" $ nf toXlsxFast bs'
        ]
    ]
