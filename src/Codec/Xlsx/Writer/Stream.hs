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
  , writeXlsxMultipleSheets
  , writeXlsxWithSharedStrings
  , writeXlsxWithConfig
  , SheetWriteSettings(..)
  , defaultSettings
  , wsSheetView
  , wsZip
  , wsColumnProperties
  , wsRowProperties
  , wsStyles
  , WriteXlsxConfig
  , defWriteXlsxConfig
  , setWriteSettings
  , setSharedStringsMap
  , setNamesAndConduits
  , setNameInfosAndConduits
  -- *** Shared strings
  , sharedStrings
  , sharedStringsStream
  ) where

import Codec.Archive.Zip.Conduit.UnZip
import Codec.Archive.Zip.Conduit.Zip
import Codec.Xlsx.Parser.Internal (addSmlNamespace)
import Codec.Xlsx.Parser.Stream
import Codec.Xlsx.Types (ColumnsProperties (..), RowProperties (..),
                         Styles (..), _AutomaticHeight, _CustomHeight,
                         emptyStyles, rowHeightLens, SheetState (..))
import Codec.Xlsx.Types.Cell
import Codec.Xlsx.Types.Common
import Codec.Xlsx.Types.Internal.Relationships (odr, pr, odRelNs)
import Codec.Xlsx.Types.SheetViews
import Codec.Xlsx.Writer.Internal (nonEmptyElListSimple, toAttrVal, toElement,
                                   txtd, txti)
import Codec.Xlsx.Writer.Internal.Stream
import Conduit (PrimMonad, yield, (.|))
import qualified Conduit as C
import Codec.Xlsx.LensCompat
import Control.Monad
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
import Data.Bifunctor (Bifunctor(second))


upsertSharedStrings :: MonadState SharedStringState m => Row -> m [(Text,Int)]
upsertSharedStrings row =
  traverse upsertSharedString items
  where
    items :: [Text]
    items = row ^.. ri_cell_row . traversed . cellValue . _Just . _CellText

-- | Process sheetItems into shared strings structure to be put into
--   'writeXlsxWithSharedStrings'
sharedStrings :: Monad m  => ConduitT Row b m (Map Text Int)
sharedStrings = void sharedStringsStream .| CL.foldMap (uncurry Map.singleton)

-- | creates a unique number for every encountered string in the stream
--   This is used for creating a required structure in the xlsx format
--   called shared strings. Every string get's transformed into a number
--
--   exposed to allow further processing, we also know the map after processing
--   but I don't think conduit provides a way of getting that out.
--   use 'sharedStrings' to just get the map
sharedStringsStream :: Monad m  =>
  ConduitT Row (Text, Int) m (Map Text Int)
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



-- | Transform a 'Row' stream into a stream that creates the xlsx file format
--   (to be consumed by sinkfile for example)
--  This first runs 'sharedStrings' and then 'writeXlsxWithSharedStrings'.
--  If you want xlsx files this is the most obvious function to use.
--  the others are exposed in case you can cache the shared strings for example.
-- See also 'writeXlsxMultipleSheets'.
writeXlsx :: MonadThrow m
    => PrimMonad m
    => SheetWriteSettings -- ^ use 'defaultSettings'
    -> ConduitT () Row m () -- ^ the conduit producing sheetitems
    -> ConduitT () ByteString m Word64 -- ^ result conduit producing xlsx files
writeXlsx settings sheetC = do
    writeXlsxWithConfig (defWriteXlsxConfig
                        & setWriteSettings settings
                        & setNamesAndConduits [("Sheet1", sheetC)])

-- | Same as 'writeXlsx' but write to multiple sheets.
writeXlsxMultipleSheets :: MonadThrow m
    => PrimMonad m
    => SheetWriteSettings
    -- ^ use 'defaultSettings'.
    -- Currently, all sheets will use the same setting.
    -> [(Text, ConduitT () Row m ())]
    -- ^ the conduits producing sheetitems for each sheet
    -> ConduitT () ByteString m Word64 -- ^ result conduit producing xlsx files
writeXlsxMultipleSheets settings sheets = do
    writeXlsxWithConfig (defWriteXlsxConfig
                        & setWriteSettings settings
                        & setNamesAndConduits sheets)

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
    -> [(Text, ConduitT () Row m ())]
    -> ConduitT () ByteString m Word64
writeXlsxWithSharedStrings settings sharedStrings' sheets =
  writeXlsxWithConfig (defWriteXlsxConfig
                      & setWriteSettings settings
                      & setSharedStringsMap (Just sharedStrings')
                      & setNamesAndConduits sheets)

-- | Configure how to write to an xlsx. Provides setters but no getters to
-- hide internals for future compatibility
data WriteXlsxConfig m = MkWriteXlsxConfig
  { _writeSettings :: SheetWriteSettings
  , _sharedStringsMap :: Maybe (Map Text Int)
  , _infosAndConduits :: [((Text, SheetState), ConduitT () Row m ())]
  }

-- | Create a blank config. Uses the default write setting config, derives shared
-- strings from the conduits and has no conduits or sheet informations within.
defWriteXlsxConfig :: WriteXlsxConfig m
defWriteXlsxConfig = MkWriteXlsxConfig defaultSettings mempty []

-- | Overwrite the existing sheet write settings.
setWriteSettings :: SheetWriteSettings -> WriteXlsxConfig m -> WriteXlsxConfig m
setWriteSettings ws (MkWriteXlsxConfig _ ssm iac) = MkWriteXlsxConfig ws ssm iac

-- | Set the shared strings map. If set to Nothing, derives the shared strings
-- from the conduits.
setSharedStringsMap :: Maybe (Map Text Int) -> WriteXlsxConfig m -> WriteXlsxConfig m
setSharedStringsMap ssm (MkWriteXlsxConfig ws _ iac) = MkWriteXlsxConfig ws ssm iac

-- | Set the sheet infos and conduits for each sheet.
setNameInfosAndConduits :: [((Text, SheetState), ConduitT () Row m ())] -> WriteXlsxConfig m -> WriteXlsxConfig m
setNameInfosAndConduits iac (MkWriteXlsxConfig ws ssm _) = MkWriteXlsxConfig ws ssm iac

-- | Set the names and conduits for each sheet. The sheet info is derived from
-- the ordering.
setNamesAndConduits :: [(Text, ConduitT () Row m ())] -> WriteXlsxConfig m -> WriteXlsxConfig m
setNamesAndConduits nacs (MkWriteXlsxConfig ws ssm _) = MkWriteXlsxConfig ws ssm infosAndConduits'
  where
  infosAndConduits' = nacs <&> \(n, c) -> ((n, Visible), c)

-- | Write to an XLSX using the 'WriteXlsxConfig' structure.
writeXlsxWithConfig :: MonadThrow m => PrimMonad m
    => WriteXlsxConfig m
    -> ConduitT () ByteString m Word64
writeXlsxWithConfig (MkWriteXlsxConfig ws ssm iac) = do
  let rowConduits = foldr ((>>) . snd) mempty iac
  sstrings <- maybe (rowConduits .| sharedStrings) pure ssm
  combinedFiles ws sstrings iac .| zipStream (ws ^. wsZip)

combinedFiles :: PrimMonad m
  => SheetWriteSettings
  -> Map Text Int
  -> [((Text, SheetState), ConduitT () Row m ())]
  -> ConduitT () (ZipEntry, ZipData m) m ()
combinedFiles settings sharedStrings' sheets =
  let indicesAndSheets = zip [1..] sheets
      zippedSheets = map (\(sheetId, (_, rowConduit)) ->
        ( zipEntry ("xl/worksheets/sheet" <> Text.pack (show sheetId) <> ".xml")
        , ZipDataSource $ rowConduit .| C.runReaderC settings (writeWorkSheet sharedStrings') .| eventsToBS
        )) indicesAndSheets
  in
  C.yieldMany $
    [ (zipEntry "xl/sharedStrings.xml", ZipDataSource $ writeSst sharedStrings' .| eventsToBS)
    , (zipEntry "[Content_Types].xml", ZipDataSource $ writeContentTypes .| eventsToBS)
    , (zipEntry "xl/workbook.xml", ZipDataSource $ writeWorkbook (map (second fst) indicesAndSheets)
    .| renderBuilder ( def {rsNamespaces=[("r", odRelNs)]})
    .| C.builderToByteString)
    , (zipEntry "xl/styles.xml", ZipDataByteString $ coerce $ settings ^. wsStyles)
    , (zipEntry "xl/_rels/workbook.xml.rels", ZipDataSource $ writeWorkbookRels (length sheets) .| eventsToBS)
    , (zipEntry "_rels/.rels", ZipDataSource $ writeRootRels .| eventsToBS)
    ] <> zippedSheets

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
writeWorkbook :: Monad m => [(Int, (Text, SheetState))] -> forall i.  ConduitT i Event m ()
writeWorkbook sheetInfos =
  let addSheet (sheetId, (sheetName, visibility)) = tag (addSmlNamespace "sheet")
          (attr "name" sheetName
          <> attr "sheetId" (Text.pack $ show sheetId)
          <> attr (odr "id") ("rId" <> (Text.pack $ show (sheetId + 2)))
          <> attr "state" (case visibility of
                              Visible    -> "visible"
                              Hidden     -> "hidden"
                              VeryHidden -> "veryHidden")
          ) $ pure ()
  in doc (addSmlNamespace "workbook") $
    el (addSmlNamespace "sheets") $ mapM_ addSheet sheetInfos

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

writeWorkbookRels :: Monad m => forall i. Int -> ConduitT i Event m ()
writeWorkbookRels sheetCount = doc (pr "Relationships") $ do
  relationship "sharedStrings.xml" 1 "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings"
  relationship "styles.xml" 2 "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"
  let schema = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"
  mapM_ (\sheetId -> relationship ("worksheets/sheet" <> (Text.pack (show sheetId)) <> ".xml") (sheetId + 2) schema) [1..sheetCount]

writeRootRels :: Monad m => forall i.  ConduitT i Event m ()
writeRootRels = doc (pr "Relationships") $
  relationship "xl/workbook.xml" 1 "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"


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
writeSst sharedStrings' = doc (addSmlNamespace "sst") $
    void $ traverse (el (addSmlNamespace "si") .  el (addSmlNamespace "t") . content . fst
                  ) $ sortBy (\(_, i) (_, y :: Int) -> compare i y) $ Map.toList sharedStrings'

writeEvents ::  PrimMonad m => ConduitT Event Builder m ()
writeEvents = renderBuilder def

sheetViews :: forall m . MonadReader SheetWriteSettings m => forall i . ConduitT i Event m ()
sheetViews = do
  sheetView <- view wsSheetView

  unless (null sheetView) $ el (addSmlNamespace "sheetViews") $ do
    let
        view' :: [Element]
        view' = setNameSpaceRec spreadSheetNS . toXMLElement .  toElement (addSmlNamespace "sheetView") <$> sheetView

    C.yieldMany $ elementToEvents =<< view'

spreadSheetNS :: Text
spreadSheetNS = fold $ nameNamespace $ addSmlNamespace ""

setNameSpaceRec :: Text -> Element -> Element
setNameSpaceRec space xelm =
    xelm {elementName = ((elementName xelm ){nameNamespace =
                                    Just space })
      , elementNodes = elementNodes xelm <&> \case
                                    NodeElement x -> NodeElement (setNameSpaceRec space x)
                                    y -> y
    }

columns :: MonadReader SheetWriteSettings m => ConduitT Row Event m ()
columns = do
  colProps <- view wsColumnProperties
  let cols :: Maybe TXML.Element
      cols = nonEmptyElListSimple (addSmlNamespace "cols") $ map (toElement (addSmlNamespace "col")) colProps
  traverse_ (C.yieldMany . elementToEvents . toXMLElement) cols

writeWorkSheet :: MonadReader SheetWriteSettings  m => Map Text Int  -> ConduitT Row Event m ()
writeWorkSheet sharedStrings' = doc (addSmlNamespace "worksheet") $ do
    sheetViews
    columns
    el (addSmlNamespace "sheetData") $ C.awaitForever (mapRow sharedStrings')

mapRow :: MonadReader SheetWriteSettings m => Map Text Int -> Row -> ConduitT Row Event m ()
mapRow sharedStrings' sheetItem = do
  mRowProp <- preview $ wsRowProperties . ix (unRowIndex rowIx) . rowHeightLens . _Just . failing _CustomHeight _AutomaticHeight
  let rowAttr :: Attributes
      rowAttr = ixAttr <> fold (attr "ht" . txtd <$> mRowProp)
  tag (addSmlNamespace "row") rowAttr $
    void $ itraverse (mapCell sharedStrings' rowIx) (sheetItem ^. ri_cell_row)
  where
    rowIx = sheetItem ^. ri_row_index
    ixAttr = attr "r" $ toAttrVal rowIx

mapCell ::
  Monad m => Map Text Int -> RowIndex -> Int -> Cell -> ConduitT Row Event m ()
mapCell sharedStrings' rix cix' cell =
  when (has (cellValue . _Just) cell || has (cellStyle . _Just) cell) $
  tag (addSmlNamespace "c") celAttr $
    when (has (cellValue . _Just) cell) $
    el (addSmlNamespace "v") $
      content $ renderCell sharedStrings' cell
  where
    cix = ColumnIndex cix'
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
