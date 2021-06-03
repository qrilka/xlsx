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
  ( writeXlsx
  , writeXlsxWithSharedStrings
  , sharedStrings
  , sharedStringsStream

  -- automated testing -- TODO move to internal
  , getSetNumber
  , initialSharedString
  , string_map

  ) where

import           Codec.Archive.Zip.Conduit.UnZip
import           Codec.Xlsx.Types.Cell
import           Codec.Xlsx.Types.Common
import           Conduit                         (PrimMonad, yield, (.|))
import qualified Conduit                         as C
import           Control.Lens
import           Control.Monad.Catch
import           Control.Monad.State.Strict
import Data.ByteString(ByteString)
import Data.ByteString.Builder(Builder)
import qualified Data.ByteString.Lazy                 as LBS
import qualified Data.ByteString.Builder as BS
import           Data.Conduit                    (ConduitT)
import qualified Data.Conduit.Combinators        as C
import qualified Data.Conduit.List as CL
import Codec.Archive.Zip.Conduit.Zip
import           Data.Map.Strict                 (Map)
import qualified Data.Map.Strict                 as Map
import           Data.Text                       (Text)
import           Data.XML.Types
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

-- | Process sheetItems into shared strings structure to be put into
--   'writeXlsxWithSharedStrings'
sharedStrings :: Monad m  => ConduitT SheetItem b m (Map Text Int)
sharedStrings = void sharedStringsStream .| CL.foldMap (uncurry Map.singleton)

-- | creates a unique number for every encountered string in the stream
--   This is used for creating a required structure in the xlsx format
--   called shared strings. Every string get's transformed into a number
--
--   exposed to allow further processing, we also know the map after processing
--   but I don't think conduit provides a way of getting that out.
--   use 'sharedStrings' to just get the map
sharedStringsStream :: Monad m  =>
  ConduitT SheetItem (Text, Int) m (Map Text Int)
sharedStringsStream = fmap (view string_map) $ C.execStateC initialSharedString $
  CL.mapFoldableM mapFold

-- | Transform a SheetItem stream into a stream that creates the xlsx file format
--   (to be consumed by sinkfile for example)
--  This first runs 'sharedStrings' and then 'writeXlsxWithSharedStrings'.
--  If you want xlsx files this is the most obvious function to use.
--  the others are exposed in case you can cache the shared strings for example.
writeXlsx :: MonadThrow m => PrimMonad m  =>
  ConduitT () SheetItem m ()  ->
  ConduitT () ByteString m Word64
writeXlsx sheetC = do
    sstrings  <- sheetC .| sharedStrings
    sheetC .|  writeXlsxWithSharedStrings sstrings


recomendedZipOptions :: ZipOptions
recomendedZipOptions = defaultZipOptions {
  zipOpt64 = False -- TODO renable
  -- There is a magick number in the zip archive package,
  -- https://hackage.haskell.org/package/zip-archive-0.4.1/docs/src/Codec.Archive.Zip.html#local-6989586621679055672
  -- if we enable 64bit the number doesn't align causing the test to fail.
  -- I'll make the test pass, and after that I'll make the
  }

-- TODO provide a safe default interface and also add this in case you need perfromance
-- TODO maybe should use bimap instead: https://hackage.haskell.org/package/bimap-0.4.0/docs/Data-Bimap.html
-- it guarantees uniqueness of both text and int
writeXlsxWithSharedStrings :: MonadThrow m => PrimMonad m  =>
  Map Text Int ->
  ConduitT SheetItem ByteString m Word64
writeXlsxWithSharedStrings sstable = do
  res  <- combinedFiles sstable .| zipStream recomendedZipOptions
  -- yield (LBS.toStrict $ BS.toLazyByteString $ BS.word32LE 0x06054b50) -- insert magic number for fun.
  pure res

combinedFiles :: PrimMonad m  =>
  Map Text Int ->
  ConduitT SheetItem (ZipEntry, ZipData m) m ()
combinedFiles sstable = do
  yield (zipEntry "xl/sharedStrings.xml", ZipDataSource $ writeSstXml sstable .| C.builderToByteString)
  yield (zipEntry "[Content_Types].xml", ZipDataSource $ writeContentTypes .| writeEvents .| C.builderToByteString)
  writeWorkSheet sstable .| writeEvents .| C.builderToByteString .| C.map (\x -> (zipEntry "xl/worksheets/sheet1.xml", ZipDataByteString $ LBS.fromStrict x))

el :: Monad m => Name -> Monad m => forall i.  ConduitT i Event m () -> ConduitT i Event m ()
el x = tag x mempty

-- | required by excell.
writeContentTypes :: Monad m => forall i.  ConduitT i Event m ()
writeContentTypes = do
  yield EventBeginDocument
  el "Types" $ pure ()
  yield EventEndDocument

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
  el "sst" $
    void $ traverse (el "si" .  el "t" . content . fst
                  ) $ sortBy (\(_, i) (_, y :: Int) -> compare i y) $ Map.toList sstable
  yield EventEndDocument

writeEvents ::  PrimMonad m => ConduitT Event Builder m ()
writeEvents = renderBuilder (def {rsPretty=True})

writeWorkSheet :: Monad m => Map Text Int  -> ConduitT SheetItem Event m ()
writeWorkSheet sstable = do
  yield EventBeginDocument
  el "worksheet" $
    el "sheetData" $ C.concatMap (mapItem sstable)
  yield EventEndDocument



mapItem :: Map Text Int -> SheetItem -> [Event]
mapItem sstable sheetItem = do
  void [EventBeginElement "row"  [("r", [ContentText $ toAttrVal rowIx])]]
  void $ itraverse (mapCell sstable rowIx) $ sheetItem ^. si_cell_row
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
