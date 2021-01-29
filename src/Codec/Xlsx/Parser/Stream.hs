{-# LANGUAGE LambdaCase #-}
{-# LANGUAGE FlexibleContexts #-}
{-# LANGUAGE TemplateHaskell #-}

-- | Read .xlsx as a stream
module Codec.Xlsx.Parser.Stream
  ( readXlsx
  , SheetItem(..)
  , XlsxItem(..)
  ) where

import Control.Monad.State.Lazy
import Data.Conduit(ConduitT)
import Conduit(PrimMonad, MonadThrow, yield, await, (.|))
import qualified Conduit as C
import qualified Data.Conduit as C
import qualified Data.Conduit.Combinators as C
import Codec.Archive.Zip.Conduit.UnZip
import Codec.Xlsx.Types.Cell
import qualified Data.ByteString as BS
import Text.XML.Stream.Parse
import Data.XML.Types
import Data.Text (Text)
import qualified Data.Text.Encoding as Text
import Control.Lens

data SheetItem = MkSheetItem {
  _si_cell_row :: CellRow
  } deriving Show

data XlsxItem = MkXlsxItem {
  _xi_sheet :: SheetItem
  } deriving Show

data PsFiles = UnkownFile Text
             | Sheet Text
             | InitialNoFile


data PipeState = MkPipeState
  { _ps_file :: PsFiles
  }
makeLenses 'MkPipeState

readXlsx :: MonadIO m => MonadThrow m
  => PrimMonad m
  => ConduitT BS.ByteString XlsxItem m ()
readXlsx = (() <$ unZipStream)
    .| (C.evalStateLC (MkPipeState InitialNoFile) $ (await >>= tagFiles)
      .| parseBytes def
      .| parseSheet)

-- | there are various files in the excell file, which is a glorified zip folder
-- here we tag them with things we know, and push it into the state monad.
-- we need a state monad to make the excell parsing conduit to function
tagFiles ::
  MonadState PipeState m
  => MonadIO m
  => MonadThrow m
  => PrimMonad m
  => Maybe (Either ZipEntry BS.ByteString) -> ConduitT (Either ZipEntry BS.ByteString) BS.ByteString m ()
tagFiles = \case
  Just (Left zipEntry) -> do
   liftIO $ print zipEntry
   -- bsf_file .= either id Text.decodeUtf8 (zipEntryName zipEntry)
   await >>= tagFiles
  Just (Right fdata) -> do
   yield fdata
   await >>= tagFiles
  Nothing -> pure ()

parseSheet ::
  MonadIO m
  => MonadThrow m
  => PrimMonad m
  => MonadState PipeState m
  => ConduitT Event XlsxItem  m ()
parseSheet = C.mapM $ \x -> do
    -- liftIO $ print x
    pure $ MkXlsxItem $ MkSheetItem mempty
