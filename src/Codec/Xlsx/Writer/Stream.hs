{-# LANGUAGE CPP                 #-}
{-# LANGUAGE ConstraintKinds     #-}
{-# LANGUAGE DataKinds           #-}
{-# LANGUAGE FlexibleContexts    #-}
{-# LANGUAGE LambdaCase          #-}
{-# LANGUAGE OverloadedStrings   #-}
{-# LANGUAGE RankNTypes          #-}
{-# LANGUAGE ScopedTypeVariables #-}
{-# LANGUAGE StrictData          #-}
{-# LANGUAGE TemplateHaskell     #-}

-- | Writes Excel files from a stream, which allows creation of
--   large Excel files while remaining in constant memory.
module Codec.Xlsx.Writer.Stream
  ( writeXlsx
  , writeXlsxWithSharedStrings
  , SheetWriteSettings(..)
  , defaultSettings
  , wsSheetView
  , wsZip
  , wsColumnProperties
  , wsRowProperties
  , wsStyles
  -- *** Shared strings
  , sharedStrings
  , sharedStringsStream
  ) where

import Codec.Archive.Zip.Conduit.UnZip
import Codec.Archive.Zip.Conduit.Zip
import Codec.Xlsx.Parser.Internal (n_)
import Codec.Xlsx.Parser.Stream
import Codec.Xlsx.Types (ColumnsProperties (..), RowProperties (..),
                         Styles (..), _AutomaticHeight, _CustomHeight,
                         emptyStyles, rowHeightLens)
import Codec.Xlsx.Types.Cell
import Codec.Xlsx.Types.Common
import Codec.Xlsx.Types.Internal.Relationships (odr, pr)
import Codec.Xlsx.Types.SheetViews
import Codec.Xlsx.Writer.Internal (nonEmptyElListSimple, toAttrVal, toElement,
                                   txtd, txti)
import Codec.Xlsx.Writer.Internal.Stream
import Conduit (PrimMonad, yield, (.|))
import qualified Conduit as C
#ifdef USE_MICROLENS
import Data.Traversable.WithIndex
import Lens.Micro.Platform
#else
import Control.Lens
#endif
import Control.Monad.Catch
import Control.Monad.Reader.Class
import Control.Monad.State.Strict
import Data.ByteString (ByteString)
import Data.ByteString.Builder (Builder)
import Data.Coerce
import Data.Conduit (ConduitT)
import qualified Data.Conduit.List as CL
import Data.Foldable (fold, traverse_)
import Data.List
import Data.Map.Strict (Map)
import qualified Data.Map.Strict as Map
import Data.Maybe
import Data.Text (Text)
import qualified Data.Text as Text
import Data.Time
import Data.Word
import Data.XML.Types
import Text.Printf
import Text.XML (toXMLElement)
import qualified Text.XML as TXML
import Text.XML.Stream.Render
import Text.XML.Unresolved (elementToEvents)


upsertSharedStrings :: MonadState SharedStringState m => RowItem -> m [(Text,Int)]
upsertSharedStrings row =
  traverse upsertSharedString items
  where
    items :: [Text]
    items = row ^.. ri_cell_row . traversed . cellValue . _Just . _CellText

-- | Process sheetItems into shared strings structure to be put into
--   'writeXlsxWithSharedStrings'
sharedStrings :: Monad m  => ConduitT RowItem b m (Map Text Int)
sharedStrings = void sharedStringsStream .| CL.foldMap (uncurry Map.singleton)

-- | creates a unique number for every encountered string in the stream
--   This is used for creating a required structure in the xlsx format
--   called shared strings. Every string get's transformed into a number
--
--   exposed to allow further processing, we also know the map after processing
--   but I don't think conduit provides a way of getting that out.
--   use 'sharedStrings' to just get the map
sharedStringsStream :: Monad m  =>
  ConduitT RowItem (Text, Int) m (Map Text Int)
sharedStringsStream = fmap (view string_map) $ C.execStateC initialSharedString $
  CL.mapFoldableM upsertSharedStrings

-- | Settings for writing a single sheet.
data SheetWriteSettings = MkSheetWriteSettings
  { _wsSheetView        :: [SheetView]
  , _wsZip              :: ZipOptions -- ^ Enable zipOpt64=True if you intend writing large xlsx files, zip needs 64bit for files over 4gb.
  , _wsColumnProperties :: [ColumnsProperties]
  , _wsRowProperties    :: Map Int RowProperties
  , _wsStyles           :: Styles
  }
instance Show  SheetWriteSettings where
  -- ZipOptions lacks a show instance-}
  show (MkSheetWriteSettings s _ y r _) = printf "MkSheetWriteSettings{ _wsSheetView=%s, _wsColumnProperties=%s, _wsZip=defaultZipOptions, _wsRowProperties=%s }" (show s) (show y) (show r)
makeLenses ''SheetWriteSettings

defaultSettings :: SheetWriteSettings
defaultSettings = MkSheetWriteSettings
  { _wsSheetView = []
  , _wsColumnProperties = []
  , _wsRowProperties = mempty
  , _wsStyles = emptyStyles
  , _wsZip = defaultZipOptions {
  zipOpt64 = False
  -- There is a magick number in the zip archive package,
  -- https://hackage.haskell.org/package/zip-archive-0.4.1/docs/src/Codec.Archive.Zip.html#local-6989586621679055672
  -- if we enable 64bit the number doesn't align causing the test to fail.
  }
  }



-- | Transform a 'RowItem' stream into a stream that creates the xlsx file format
--   (to be consumed by sinkfile for example)
--  This first runs 'sharedStrings' and then 'writeXlsxWithSharedStrings'.
--  If you want xlsx files this is the most obvious function to use.
--  the others are exposed in case you can cache the shared strings for example.
--
--  Note that the current implementation concatenates everything into a single sheet.
--  In other words there is no tab support yet.
writeXlsx :: MonadThrow m
    => PrimMonad m
    => SheetWriteSettings -- ^ use 'defaultSettings'
    -> ConduitT () RowItem m () -- ^ the conduit producing sheetitems
    -> ConduitT () ByteString m Word64 -- ^ result conduit producing xlsx files
writeXlsx settings sheetC = do
    sstrings  <- sheetC .| sharedStrings
    writeXlsxWithSharedStrings settings sstrings sheetC


-- TODO maybe should use bimap instead: https://hackage.haskell.org/package/bimap-0.4.0/docs/Data-Bimap.html
-- it guarantees uniqueness of both text and int
-- | This write Excel file with a shared strings lookup table.
--   It appears that it is optional.
--   Failed lookups will result in valid xlsx.
--   There are several conditions on shared strings,
--
--      1. Every text to int is unique on both text and int.
--      2. Every Int should have a gap no greater than 1. [("xx", 3), ("yy", 4)] is okay, whereas [("xx", 3), ("yy", 5)] is not.
--      3. It's expected this starts from 0.
--
--   Use 'sharedStringsStream' to get a good shared strings table.
--   This is provided because the user may have a more efficient way of
--   constructing this table than the library can provide,
--   for example through database operations.
writeXlsxWithSharedStrings :: MonadThrow m => PrimMonad m
    => SheetWriteSettings
    -> Map Text Int -- ^ shared strings table
    -> ConduitT () RowItem m ()
    -> ConduitT () ByteString m Word64
writeXlsxWithSharedStrings settings sharedStrings' items =
  combinedFiles settings sharedStrings' items .| zipStream (settings ^. wsZip)

-- massive amount of boilerplate needed for excel to function
boilerplate :: forall m . PrimMonad m  => SheetWriteSettings -> Map Text Int -> [(ZipEntry,  ZipData m)]
boilerplate settings sharedStrings' =
  [ (zipEntry "xl/sharedStrings.xml", ZipDataSource $ writeSst sharedStrings' .| eventsToBS)
  , (zipEntry "[Content_Types].xml", ZipDataSource $ writeContentTypes .| eventsToBS)
  , (zipEntry "xl/workbook.xml", ZipDataSource $ writeWorkbook .| eventsToBS)
  , (zipEntry "xl/styles.xml", ZipDataByteString $ coerce $ settings ^. wsStyles)
  , (zipEntry "xl/_rels/workbook.xml.rels", ZipDataSource $ writeWorkbookRels .| eventsToBS)
  , (zipEntry "_rels/.rels", ZipDataSource $ writeRootRels .| eventsToBS)
  ]

combinedFiles :: PrimMonad m
  => SheetWriteSettings
  -> Map Text Int
  -> ConduitT () RowItem m ()
  -> ConduitT () (ZipEntry, ZipData m) m ()
combinedFiles settings sharedStrings' items =
  C.yieldMany $
    boilerplate settings  sharedStrings' <>
    [(zipEntry "xl/worksheets/sheet1.xml", ZipDataSource $
       items .| C.runReaderC settings (writeWorkSheet sharedStrings') .| eventsToBS )]

el :: Monad m => Name -> Monad m => forall i.  ConduitT i Event m () -> ConduitT i Event m ()
el x = tag x mempty

--   Clark notation is used a lot for xml namespaces in this module:
--   <https://hackage.haskell.org/package/xml-types-0.3.8/docs/Data-XML-Types.html#t:Name>
--   Name has an IsString instance which parses it
override :: Monad m => Text -> Text -> forall i.  ConduitT i Event m ()
override content' part =
    tag "{http://schemas.openxmlformats.org/package/2006/content-types}Override"
      (attr "ContentType" content'
       <> attr "PartName" part) $ pure ()


-- | required by Excel.
writeContentTypes :: Monad m => forall i.  ConduitT i Event m ()
writeContentTypes = doc "{http://schemas.openxmlformats.org/package/2006/content-types}Types" $ do
    override "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml" "/xl/workbook.xml"
    override "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml" "/xl/sharedStrings.xml"
    override "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml" "/xl/styles.xml"
    override "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml" "/xl/worksheets/sheet1.xml"
    override "application/vnd.openxmlformats-package.relationships+xml" "/xl/_rels/workbook.xml.rels"
    override "application/vnd.openxmlformats-package.relationships+xml" "/_rels/.rels"

-- | required by Excel.
writeWorkbook :: Monad m => forall i.  ConduitT i Event m ()
writeWorkbook = doc (n_ "workbook") $ do
    el (n_ "sheets") $ do
      tag (n_ "sheet")
        (attr "name" "Sheet1"
         <> attr "sheetId" "1" <>
         attr (odr "id") "rId3") $
        pure ()

doc :: Monad m => Name ->  forall i.  ConduitT i Event m () -> ConduitT i Event m ()
doc root docM = do
  yield EventBeginDocument
  el root docM
  yield EventEndDocument

relationship :: Monad m => Text -> Int -> Text ->  forall i.  ConduitT i Event m ()
relationship target id' type' =
  tag (pr "Relationship")
    (attr "Type" type'
      <> attr "Id" (Text.pack $ "rId" <> show id')
      <> attr "Target" target
    ) $ pure ()

writeWorkbookRels :: Monad m => forall i.  ConduitT i Event m ()
writeWorkbookRels = doc (pr "Relationships") $  do
  relationship "sharedStrings.xml" 1 "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings"
  relationship "worksheets/sheet1.xml" 3 "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"
  relationship "styles.xml" 2 "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"

writeRootRels :: Monad m => forall i.  ConduitT i Event m ()
writeRootRels = doc (pr "Relationships") $  do
  relationship "xl/workbook.xml" 1 "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
  -- relationship "docProps/core.xml" "rId2" "http://schemas.openxmlformats.org/officeDocument/2006/relationships/metadata/core-properties"
  -- relationship "docProps/app.xml" "rId3" "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties"


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
writeSst sharedStrings' = doc (n_ "sst") $
    void $ traverse (el (n_ "si") .  el (n_ "t") . content . fst
                  ) $ sortBy (\(_, i) (_, y :: Int) -> compare i y) $ Map.toList sharedStrings'

writeEvents ::  PrimMonad m => ConduitT Event Builder m ()
writeEvents = renderBuilder (def {rsPretty=False})

sheetViews :: forall m . MonadReader SheetWriteSettings m => forall i . ConduitT i Event m ()
sheetViews = do
  sheetView <- view wsSheetView

  unless (null sheetView) $ el (n_ "sheetViews") $ do
    let
        view' :: [Element]
        view' = setNameSpaceRec spreadSheetNS . toXMLElement .  toElement (n_ "sheetView") <$> sheetView
    -- tag (n_ "sheetView") (attr "topLeftCell" "D10" <> attr "tabSelected" "1") $ pure ()

    C.yieldMany $ elementToEvents =<< view'

spreadSheetNS :: Text
spreadSheetNS = fold $ nameNamespace $ n_ ""

setNameSpaceRec :: Text -> Element -> Element
setNameSpaceRec space xelm =
    xelm {elementName = ((elementName xelm ){nameNamespace =
                                    Just space })
      , elementNodes = elementNodes xelm <&> \case
                                    NodeElement x -> NodeElement (setNameSpaceRec space x)
                                    y -> y
    }

columns :: MonadReader SheetWriteSettings m => ConduitT RowItem Event m ()
columns = do
  colProps <- view wsColumnProperties
  let cols :: Maybe TXML.Element
      cols = nonEmptyElListSimple (n_ "cols") $ map (toElement (n_ "col")) colProps
  traverse_ (C.yieldMany . elementToEvents . toXMLElement) cols

writeWorkSheet :: MonadReader SheetWriteSettings  m => Map Text Int  -> ConduitT RowItem Event m ()
writeWorkSheet sharedStrings' = doc (n_ "worksheet") $ do
    sheetViews
    columns
    el (n_ "sheetData") $ C.awaitForever (mapRow sharedStrings')

mapRow :: MonadReader SheetWriteSettings m => Map Text Int -> RowItem -> ConduitT RowItem Event m ()
mapRow sharedStrings' sheetItem = do
  mRowProp <- preview $ wsRowProperties . ix rowIx . rowHeightLens . _Just . failing _CustomHeight _AutomaticHeight
  let rowAttr :: Attributes
      rowAttr = ixAttr <> fold (attr "ht" . txtd <$> mRowProp)
  tag (n_ "row") rowAttr $
    void $ itraverse (mapCell sharedStrings' rowIx) (sheetItem ^. ri_cell_row)
  where
    rowIx = sheetItem ^. ri_row_index
    ixAttr = attr "r" $ toAttrVal rowIx

mapCell :: Monad m => Map Text Int -> Int -> Int -> Cell -> ConduitT RowItem Event m ()
mapCell sharedStrings' rix cix cell =
  when (has (cellValue . _Just) cell || has (cellStyle . _Just) cell) $
  tag (n_ "c") celAttr $
    when (has (cellValue . _Just) cell) $
    el (n_ "v") $
      content $ renderCell sharedStrings' cell
  where
    celAttr  = attr "r" ref <>
      renderCellType sharedStrings' cell
      <> foldMap (attr "s" . txti) (cell ^. cellStyle)
    ref :: Text
    ref = coerce $ singleCellRef (rix, cix)

renderCellType :: Map Text Int -> Cell -> Attributes
renderCellType sharedStrings' cell =
  maybe mempty
  (attr "t" . renderType sharedStrings')
  $ cell ^? cellValue . _Just

renderCell :: Map Text Int -> Cell -> Text
renderCell sharedStrings' cell =  renderValue sharedStrings' val
  where
    val :: CellValue
    val = fromMaybe (CellText mempty) $ cell ^? cellValue . _Just

renderValue :: Map Text Int -> CellValue -> Text
renderValue sharedStrings' = \case
  CellText x ->
    -- if we can't find it in the sst, print the string
    maybe x toAttrVal $ sharedStrings' ^? ix x
  CellDouble x -> toAttrVal x
  CellBool b -> toAttrVal b
  CellRich _ -> error "rich text is not supported yet"
  CellError err  -> toAttrVal err


renderType :: Map Text Int -> CellValue -> Text
renderType sharedStrings' = \case
  CellText x ->
    maybe "str" (const "s") $ sharedStrings' ^? ix x
  CellDouble _ -> "n"
  CellBool _ -> "b"
  CellRich _ -> "r"
  CellError _ -> "e"
