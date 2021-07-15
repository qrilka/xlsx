{-# LANGUAGE BangPatterns      #-}
{-# LANGUAGE BlockArguments    #-}
{-# LANGUAGE LambdaCase        #-}
{-# LANGUAGE OverloadedStrings #-}
module Main (main) where

import Codec.Xlsx.Writer.Stream
import           Codec.Xlsx.Parser.Stream
import           Conduit
import qualified Conduit                  as C
import qualified Data.Map                 as Map
import           Prelude                  hiding (foldl)
import           Control.Monad
import           Data.DList as DL
import           Data.IORef

main :: IO ()
main = do
  -- xxx: figure out a better api here..
  ref <- liftIO $ newIORef mempty
  runXlsxM "data/simple.xlsx" $ do
    WorkbookInfo sheets <- getWorkbookInfo
    forM_ (Map.keys sheets) $ \sheetNum -> do
      readSheet UseHexpat sheetNum $ \x -> modifyIORef' ref (`DL.snoc` x)
  let nonStreamingSource = do
        gathered <- liftIO $ readIORef ref
        yieldMany $ DL.toList gathered
  liftIO $ runConduitRes $ void (writeXlsx nonStreamingSource) .| C.sinkFile "out/out.zip"
  pure ()
