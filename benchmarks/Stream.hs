{-# LANGUAGE BangPatterns      #-}
{-# LANGUAGE BlockArguments    #-}
{-# LANGUAGE LambdaCase        #-}
{-# LANGUAGE OverloadedStrings #-}
module Main (main) where

import           Codec.Xlsx
import           Codec.Xlsx.Parser.Stream
import           Conduit
import qualified Conduit                  as C
import           Criterion.Main
import qualified Data.ByteString          as BS
import qualified Data.ByteString.Lazy     as LB
import qualified Data.Conduit.Combinators as C
import qualified Data.Map                 as Map
import           Data.Text                (Text)
import qualified Data.Text                as Text
import qualified Data.Text.IO             as Text
import           Debug.Trace
import           Prelude                  hiding (foldl)

celVal :: CellValue -> Text
celVal  = \case
  CellText x   -> x
  CellDouble x -> Text.pack $ show x
  CellBool x   -> Text.pack $ show x
  CellRich x   -> Text.pack $ show x
  CellError x  -> Text.pack $ show x

format :: SheetItem -> Text
format x = Text.pack $ show (_si_sheet x, _si_row_index x, maybe "" celVal . _cellValue <$> _si_cell_row x)

parseFileStream :: FilePath -> IO ()
parseFileStream filepath = do
  input <- runResourceT $ readXlsxC (sourceFile filepath)
  x <- runResourceT $ runConduit (parseConduit input)
  -- ssState <- runResourceT $ C.runConduit $ sourceFile filepath .| parseSharedStrings .| C.foldMap (uncurry Map.singleton)
  -- putStrLn $ "ssState: " <> show ssState
  pure ()
  where
    parseConduit input =
         input
      .| C.filter (("sheet1.xml" ==) . _si_sheet)
      .| C.foldl (\a b -> a <> [b]) []


parseFile :: FilePath -> IO ()
parseFile filepath = do
  contents <- LB.readFile filepath
  let !_ = toXlsx contents
  pure ()

parseFileFast :: FilePath -> IO ()
parseFileFast filepath = do
  contents <- LB.readFile filepath
  let !_ = toXlsxFast contents
  pure ()

-- main :: IO ()
-- main = parseFile "data/6000.rows.x.26.cols.xlsx"
-- main = parseFileStream "data/6000.rows.x.26.cols.xlsx"
-- main = parseFileFast "data/6000.rows.x.26.cols.xlsx"

main :: IO ()
main = let
  files =
    [ "data/simple.xlsx"
    , "data/testInput.xlsx"
    , "data/6000.rows.x.26.cols.xlsx"
    ]
  benchFiles parse filename = bench ("parsing " <> filename) . nfIO $ parse filename
  in defaultMain
      [ bgroup "stream" $ benchFiles parseFileStream <$> files
      , bgroup "canonical" $ benchFiles parseFile <$> files
      , bgroup "canonical fast" $ benchFiles parseFileFast <$> files
      ]
