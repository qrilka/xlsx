{-# LANGUAGE BangPatterns      #-}
{-# LANGUAGE BlockArguments    #-}
{-# LANGUAGE LambdaCase        #-}
{-# LANGUAGE OverloadedStrings #-}
module Main (main) where

import Codec.Xlsx.Writer.Stream
import           Codec.Xlsx
import           Codec.Xlsx.Parser.Stream
import           Conduit
import qualified Conduit                  as C
import           Criterion.Main
import qualified Data.ByteString          as BS
import qualified Data.ByteString.Builder          as BS
import qualified Data.ByteString.Lazy     as LB
import qualified Data.Conduit.Combinators as C
import qualified Data.Map                 as Map
import           Data.Text                (Text)
import qualified Data.Text                as Text
import qualified Data.Text.IO             as Text
import           Debug.Trace
import           Prelude                  hiding (foldl)
import qualified Data.ByteString.Lazy as BS
import Data.Text.Encoding as Text
import Codec.Xlsx.Parser.Stream
import Control.Monad
import Data.Void

main :: IO ()
main = do
  readStr  <- runResourceT $ readXlsxC (C.sourceFile "data/simple.xlsx")
  runConduitRes $ void (writeXlsx readStr) .| C.sinkFile "out/out.zip"
  pure ()


extremlyUnsafe :: Void
extremlyUnsafe = error "evaluated extremly unsafe"
