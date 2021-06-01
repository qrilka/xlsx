{-# LANGUAGE BangPatterns        #-}
{-# LANGUAGE ConstraintKinds     #-}
{-# LANGUAGE DataKinds           #-}
{-# LANGUAGE DeriveAnyClass      #-}
{-# LANGUAGE DeriveGeneric       #-}
{-# LANGUAGE DerivingVia         #-}
{-# LANGUAGE FlexibleContexts    #-}
{-# LANGUAGE LambdaCase          #-}
{-# LANGUAGE MultiWayIf          #-}
{-# LANGUAGE NamedFieldPuns      #-}
{-# LANGUAGE OverloadedStrings   #-}
{-# LANGUAGE RankNTypes          #-}
{-# LANGUAGE RecordWildCards     #-}
{-# LANGUAGE ScopedTypeVariables #-}
{-# LANGUAGE StandaloneDeriving  #-}
{-# LANGUAGE StrictData          #-}
{-# LANGUAGE TemplateHaskell     #-}
{-# LANGUAGE TypeApplications    #-}

-- | writes excell files from a stream, which allows creation of
--   massive excell files while remaining in constant memory.
module Codec.Xlsx.Writer.Stream
  ( writeSharedStrings
  , writeXlsx

  -- tests by hand
  , writeSstXml
  -- automated testing
  , getSetNumber
  , initialSharedString
  , string_map
  ) where

import           Codec.Archive.Zip.Conduit.UnZip
import           Codec.Xlsx.Types.Cell
import           Codec.Xlsx.Types.Common
import           Conduit                         (PrimMonad, await, yield, (.|))
import qualified Conduit                         as C
import           Control.Lens
import           Control.Monad.Catch
import           Control.Monad.Except
import           Control.Monad.State.Strict
import           Data.Bifunctor
import Data.ByteString(ByteString)
import Data.ByteString.Builder(Builder)
import qualified Data.ByteString                 as BS
import qualified Data.ByteString.Lazy                 as LBS
import           Data.Conduit                    (ConduitT)
import qualified Data.Conduit.Combinators        as C
import qualified Data.Conduit.List as CL
import Codec.Archive.Zip.Conduit.Zip
import           Data.Foldable
import           Data.Map.Strict                 (Map)
import qualified Data.Map.Strict                 as Map
import           Data.Text                       (Text)
import qualified Data.Text                       as Text
import qualified Data.Text.Encoding              as Text
import qualified Data.Text.Read                  as Read
import           Data.XML.Types
import           Debug.Trace
import           GHC.Generics
import           NoThunks.Class
import           Text.Read
import           Text.XML.Stream.Render
import           Codec.Xlsx.Parser.Stream
import Data.Maybe
import Data.List
import Data.Word
import Data.Coerce
import Data.Time
import Codec.Xlsx.Writer.Internal(toAttrVal)

newtype SharedStringState = MkSharedStringState
  { _string_map :: Map Text Int
  }
makeLenses 'MkSharedStringState

initialSharedString :: SharedStringState
initialSharedString = MkSharedStringState mempty

-- properties:
-- for a list of [text], every unique text get's a unique number.
--
getSetNumber :: MonadState SharedStringState m => Text -> m (Text,Int)
getSetNumber current = do
  strings  <- use string_map

  let mIdx :: Maybe Int
      mIdx = strings ^? ix current

      idx :: Int
      idx = fromMaybe (length strings) mIdx

      newMap :: Map Text Int
      newMap = at current ?~ idx $ strings

  string_map .= newMap
  pure (current, idx)

mapFold :: MonadState SharedStringState m => SheetItem  -> m [(Text,Int)]
mapFold  row =
  traverse getSetNumber items
  where

    items :: [Text]
    items = row ^.. si_cell_row . traversed . cellValue . _Just . _CellText

-- | creates a unique number for every encountered string in the stream
--   This is used for creating a required structure in the xlsx format
--   called shared strings. Every string get's transformed into a number
writeSharedStrings :: Monad m  =>
  ConduitT SheetItem (Text, Int) m (Map Text Int)
writeSharedStrings = fmap (view string_map) $ C.execStateC initialSharedString $
  CL.mapFoldableM mapFold

-- TODO maybe should use bimap instead: https://hackage.haskell.org/package/bimap-0.4.0/docs/Data-Bimap.html
-- it guarantees uniqueness of both text and int
writeXlsx :: MonadThrow m => PrimMonad m  =>
  Map Text Int ->
  ConduitT SheetItem ByteString m Word64
writeXlsx sstable =
  combinedFiles sstable .| zipStream (defaultZipOptions { zipOpt64 = True})

combinedFiles :: PrimMonad m  =>
  Map Text Int ->
  ConduitT SheetItem (ZipEntry, ZipData m) m ()
combinedFiles sstable = do
  yield (zipEntry "xl/sharedStrings.xml", ZipDataSource $ writeSstXml sstable .| C.builderToByteString)
  writeWorkSheet sstable .| writeEvents .| C.builderToByteString .| C.map (\x -> (zipEntry "xl/worksheets/sheet1.xml", ZipDataByteString $ LBS.fromStrict x))

zipEntry :: Text -> ZipEntry
zipEntry x = ZipEntry
  { zipEntryName = Left x
  , zipEntryTime = LocalTime (toEnum 0) midnight
  , zipEntrySize = Nothing
  , zipEntryExternalAttributes = Nothing
  }

writeSstXml  ::  PrimMonad m  =>  Map Text Int  -> forall i. ConduitT i Builder m ()
writeSstXml sstable = writeSst sstable .| writeEvents

writeSst ::  Monad m  => Map Text Int  -> forall i.  ConduitT i Event m ()
writeSst sstable = do
  yield EventBeginDocument
  yield $ EventBeginElement "sst" []
  traverse (\(e, _)  -> do
                yield $ EventBeginElement "si" []
                yield $ EventBeginElement "t" []
                yield $ EventContent (ContentText e)
                yield $ EventEndElement "t"
                yield $ EventEndElement "si"
                ) $ sortBy (\(_, i) (_, y :: Int) -> compare i y) $ Map.toList sstable
  yield $ EventEndElement "sst"
  yield EventEndDocument

writeEvents ::  PrimMonad m => ConduitT Event Builder m ()
writeEvents = renderBuilder (def {rsPretty=True})

writeWorkSheet :: Monad m => Map Text Int  -> ConduitT SheetItem Event m ()
writeWorkSheet sstable = do
  yield EventBeginDocument
  yield $ EventBeginElement "worksheet" []
  yield $ EventBeginElement "sheetData" []
  C.concatMap (mapItem sstable)
  yield $ EventEndElement "sheetData"
  yield $ EventEndElement "worksheet"
  yield EventEndDocument



mapItem :: Map Text Int -> SheetItem -> [Event]
mapItem sstable sheetItem = do
  [EventBeginElement "row"  [("r", [ContentText $ toAttrVal rowIx])]]
  itraverse (mapCell sstable rowIx) $ sheetItem ^. si_cell_row
  [EventEndElement "row"]

  where
    rowIx = sheetItem ^. si_row_index

mapCell :: Map Text Int -> RowIndex -> ColIndex -> Cell -> [Event]
mapCell sstable rix cix cell =
  [ EventBeginElement "c"  [("r", [ContentText ref])]
  , EventBeginElement "v"  []
  , EventContent $ ContentText $ renderCell sstable cell
  , EventEndElement "v"
  , EventEndElement "c"
  ]
  where
    ref :: Text
    ref = coerce $ singleCellRef (rix, cix)

renderCell :: Map Text Int -> Cell -> Text
renderCell sstable cell =  renderValue sstable val
  where
    val :: CellValue
    val = fromMaybe (CellText mempty) $ cell ^? cellValue . _Just

renderValue :: Map Text Int -> CellValue -> Text
renderValue sstable = \case
  CellText x -> let
    int :: Int
    int = fromMaybe (error "could not find in sstable") $ sstable ^? ix x
    in toAttrVal int
  CellDouble x -> toAttrVal x
  CellBool b -> toAttrVal b
  CellRich _ -> error "rich text is not supported yet"
  CellError err  -> toAttrVal err
