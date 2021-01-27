
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
readXlsx = ((() <$ unZipStream)
    .| do
      x  <- await
      liftIO $ print x
      yield (MkXlsxItem $ MkSheetItem mempty)
      )
