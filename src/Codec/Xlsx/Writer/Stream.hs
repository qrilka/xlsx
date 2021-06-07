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
--   large excell files while remaining in constant memory.
--
--   This module uses the Clark notation a lot for xml namespaces:
--   <https://hackage.haskell.org/package/xml-types-0.3.8/docs/Data-XML-Types.html#t:Name>
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
import Codec.Xlsx.Parser.Internal(n_)
import Codec.Xlsx.Types.Internal.Relationships(odr, pr)

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
  yield (zipEntry "xl/sharedStrings.xml", ZipDataSource $ writeSst sstable .| eventsToBS)
  yield (zipEntry "[Content_Types].xml", ZipDataSource $ writeContentTypes .| eventsToBS)
  yield (zipEntry "xl/workbook.xml", ZipDataSource $ writeWorkBook .| eventsToBS)
  yield (zipEntry "xl/_rels/workbook.xml.rels", ZipDataSource $ writeRels .| eventsToBS)
  writeWorkSheet sstable .| eventsToBS .| C.map (\x -> (zipEntry "xl/worksheets/sheet1.xml", ZipDataByteString $ LBS.fromStrict x))

el :: Monad m => Name -> Monad m => forall i.  ConduitT i Event m () -> ConduitT i Event m ()
el x = tag x mempty

override :: Monad m => Text -> Text -> forall i.  ConduitT i Event m ()
override content' part =
    tag "{http://schemas.openxmlformats.org/package/2006/content-types}Override"
      (attr "{http://schemas.openxmlformats.org/package/2006/content-types}ContentType" content'
       <> attr "{http://schemas.openxmlformats.org/package/2006/content-types}PartName" part) $ pure ()


-- | required by excell.
writeContentTypes :: Monad m => forall i.  ConduitT i Event m ()
writeContentTypes = doc "{http://schemas.openxmlformats.org/package/2006/content-types}Types" $ do
    override "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml" "/xl/workbook.xml"
    override "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml" "/xl/sharedStrings.xml"
    override "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml" "/xl/worksheets/sheet1.xml"
    override "application/vnd.openxmlformats-package.relationships+xml" "/xl/_rels/workbook.xml.rels"

-- | required by excell.
writeWorkBook :: Monad m => forall i.  ConduitT i Event m ()
writeWorkBook = doc (n_ "workbook") $ do
    el (n_ "sheets") $ do
      tag (n_ "sheet")
        (attr (n_ "name") "Sheet1"
         <> attr (n_ "sheetId") "1" <>
         attr (pr "id") "rId3") $
        pure ()

doc :: Monad m => Name ->  forall i.  ConduitT i Event m () -> ConduitT i Event m ()
doc root docM = do
  yield EventBeginDocument
  el root docM
  yield EventEndDocument

relationship :: Monad m => Text -> Text -> Text ->  forall i.  ConduitT i Event m ()
relationship target id' type' =
  tag (pr "Relationship")
    (attr "Type" type'
      <> attr "Id" id'
      <> attr "Target" target
    ) $ pure ()

writeRels :: Monad m => forall i.  ConduitT i Event m ()
writeRels = doc (pr "Relationships") $  do
  relationship "sharedStrings.xml" "rId1" "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings"
  relationship "worksheets/sheet1.xml" "rId3" "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"



zipEntry :: Text -> ZipEntry
zipEntry x = ZipEntry
  { zipEntryName = Left x
  , zipEntryTime = LocalTime (toEnum 0) midnight
  , zipEntrySize = Nothing
  , zipEntryExternalAttributes = Nothing
  }

eventsToBS :: PrimMonad m  => ConduitT Event ByteString m ()
eventsToBS = writeEvents .| C.builderToByteString

writeSst ::  Monad m  => Map Text Int  -> forall i.  ConduitT i Event m ()
writeSst sstable = doc (n_ "sst") $
    void $ traverse (el (n_ "si") .  el (n_ "t") . content . fst
                  ) $ sortBy (\(_, i) (_, y :: Int) -> compare i y) $ Map.toList sstable

writeEvents ::  PrimMonad m => ConduitT Event Builder m ()
writeEvents = renderBuilder (def {rsPretty=True})

writeWorkSheet :: Monad m => Map Text Int  -> ConduitT SheetItem Event m ()
writeWorkSheet sstable = doc (n_ "worksheet") $ do
    el (n_ "sheetData") $ C.concatMap (mapItem sstable)



mapItem :: Map Text Int -> SheetItem -> [Event]
mapItem sstable sheetItem =
  [EventBeginElement (n_ "row")  [((n_ "r"), [ContentText $ toAttrVal rowIx])]]
   <>
  (ifoldMap (mapCell sstable rowIx) $ sheetItem ^. si_cell_row)
   <>
  [EventEndElement (n_ "row")]

  where
    rowIx = sheetItem ^. si_row_index

mapCell :: Map Text Int -> RowIndex -> ColIndex -> Cell -> [Event]
mapCell sstable rix cix cell =
  [ EventBeginElement (n_ "c")  [((n_ "r"), [ContentText ref])]
  , EventBeginElement (n_ "v")  []
  , EventContent $ ContentText $ renderCell sstable cell
  , EventEndElement (n_ "v")
  , EventEndElement (n_ "c")
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
