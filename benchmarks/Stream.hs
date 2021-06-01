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

main :: IO ()
main = do
  out  <- runConduit $ writeSstXml map' .| C.fold
  Text.putStrLn (Text.decodeUtf8 $ BS.toStrict $ BS.toLazyByteString out)
  where
    map' = Map.fromList [("a", 1), ("b", 2), ("c", 3), ("d", 10), ("e", 0), ("f", -3), ("z", 11)]
