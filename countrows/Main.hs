{-# LANGUAGE OverloadedStrings #-}
module Main (main) where

import Codec.Xlsx.Parser.Stream
import System.Environment

main :: IO ()
main = do
  args <- getArgs
  case args of
    (fname:_) -> do
      mNum <- runXlsxM fname $ countRowsInSheet 1
      putStrLn $ show mNum
    _ -> error "usage: count-rows FILE"
