{-# LANGUAGE LambdaCase #-}

-- | Read .xlsx as a stream
module Codec.Xlsx.Parser.Stream
  ( readXlsx
  , SheetItem(..)
  , XlsxItem(..)
  ) where

import Data.Conduit
import Codec.Archive.Zip.Conduit.UnZip
import Codec.Xlsx.Types.Cell
import Conduit
import qualified Data.ByteString as BS

data SheetItem = MkSheetItem {
  siCellRow :: CellRow
  } deriving Show

data XlsxItem = MkXlsxItem {
  xiSheet :: SheetItem
  } deriving Show

readXlsx :: MonadIO m => MonadThrow m
  => PrimMonad m
  => ConduitT BS.ByteString XlsxItem m ()
readXlsx = (() <$ unZipStream)
    .| do
      await >>= goRead

goRead :: MonadIO m => MonadThrow m
  => PrimMonad m
  => Maybe (Either ZipEntry BS.ByteString) -> ConduitT (Either ZipEntry BS.ByteString) XlsxItem m ()
goRead = \case
  Just (Left path) -> do
   liftIO $ print path
   await >>= goRead
  Just (Right fdata) -> do
   yield (MkXlsxItem $ MkSheetItem mempty)
   await >>= goRead
  Nothing ->  pure ()
