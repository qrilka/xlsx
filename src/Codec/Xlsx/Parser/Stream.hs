{-# LANGUAGE PackageImports #-}
{-# LANGUAGE BangPatterns        #-}
{-# LANGUAGE ConstraintKinds     #-}
{-# LANGUAGE DataKinds           #-}
{-# LANGUAGE DeriveAnyClass      #-}
{-# LANGUAGE DeriveGeneric       #-}
{-# LANGUAGE GeneralisedNewtypeDeriving #-}
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
{-# LANGUAGE UndecidableInstances #-}
-- |
-- Module      : Codex.Xlsx.Parser.Stream
-- Description : Stream parser for xlsx files
-- Copyright   :
--   (c) Adam, 2021
--   (c) Supercede, 2021
-- License     : MIT
-- Stability   : experimental
-- Portability : POSIX
--
-- Parse @.xlsx@ sheets in constant memory.
--
-- All actions on an xlsx file run inside the 'XlsxM' monad, and must
-- be run with 'runXlsxM'. XlsxM is not a monad transformer, a design
-- inherited from the "zip" package's ZipArchive monad.
--
-- Inside the XlsxM monad, you can stream 'SheetItem's (a row) from a
-- particular sheet, using 'readSheet', which is callback-based and tied to IO.
--
{-# LANGUAGE TypeApplications    #-}
module Codec.Xlsx.Parser.Stream
  ( XlsxM
  , runXlsxM
  , WorkbookInfo(..)
  , wiSheets
  , getWorkbookInfo
  , CellRow
  , SheetItem(..)
  , readSheet
  , countRowsInSheet
  , collectItems
  -- ** `SheetItem` lenses
  , si_sheet
  , si_row_index
  , si_cell_row
  -- * Errors
  , SheetErrors(..)
  {-- TODO: check if these errors could be thrown exposed to API users or not
  , AddCellErrors(..)
  , CoordinateErrors(..)
  , TypeError(..)
  --}
  ) where

import qualified "zip" Codec.Archive.Zip as Zip
import qualified Data.Vector as V
import           Codec.Xlsx.Types.Cell
import           Codec.Xlsx.Types.Common
import           Conduit                         (PrimMonad, (.|))
import qualified Conduit                         as C
import           Control.Lens
import           Control.Monad.Catch
import           Control.Monad.Except
import           Control.Monad.Reader
import           Control.Monad.State.Strict
import           Data.Bifunctor
import           Data.ByteString (ByteString)
import qualified Data.ByteString                 as BS
import           Data.Conduit                    (ConduitT)
import qualified Data.DList as DL
import           Data.Foldable
import           Data.IORef
import Data.IntMap.Strict (IntMap)
import qualified Data.IntMap.Strict as IntMap
import Data.Map.Strict (Map)
import qualified Data.Map.Strict as M
import           Data.Text                       (Text)
import qualified Data.Text                       as Text
import qualified Data.Text.Read                  as Read
import qualified Data.Text.Lazy.Builder as TB
import qualified Data.Text.Lazy as LT
import           Data.Traversable                (for)
import           Data.XML.Types
import           GHC.Generics
import           NoThunks.Class
import Codec.Xlsx.Parser.Internal

import Text.XML.Expat.SAX as Hexpat
import Text.XML.Expat.Internal.IO as Hexpat
import Control.Monad.Base
import Control.Monad.Trans.Control
import qualified Codec.Xlsx.Parser.Stream.HexpatInternal as HexpatInternal

type CellRow = IntMap Cell

-- | Sheet item
--
-- The current sheet at a time, every sheet is constructed of these items.
data SheetItem = MkSheetItem
  { _si_sheet     :: Text      -- ^ The sheet number
  , _si_row_index :: Int       -- ^ Row number
  , _si_cell_row  :: ~CellRow  -- ^ Row itself
  } deriving stock (Generic, Show)
makeLenses 'MkSheetItem

deriving via AllowThunksIn
  '[ "_si_cell_row"
   ] SheetItem
  instance NoThunks SheetItem

type SharedStringMap = V.Vector Text

-- | Type of the excel value
--
-- Note: Some values are untyped and rules of their type resolution are not known.
-- They may be treated simply as strings as well as they may be context-dependent.
-- By far we do not bother with it.
data ExcelValueType
  = TS      -- ^ shared string
  | TStr    -- ^ original string
  | TN      -- ^ number
  | TB      -- ^ boolean
  | TE      -- ^ excell error, the sheet can contain error values, for example if =1/0, causes division by zero
  | Untyped -- ^ Not all values are types
  deriving stock (Generic, Show)
  deriving anyclass NoThunks

-- | State for parsing sheets
data SheetState = MkSheetState
  { _ps_row            :: ~CellRow        -- ^ Current row
  , _ps_sheet_name     :: Text            -- ^ Current sheet name
  , _ps_cell_row_index :: Int             -- ^ Current row number
  , _ps_cell_col_index :: Int             -- ^ Current column number
  , _ps_is_in_val      :: Bool            -- ^ Flag for indexing wheter the parser is in value or not
  , _ps_shared_strings :: SharedStringMap -- ^ Shared string map
  , _ps_type           :: ExcelValueType  -- ^ The last detected value type

  , _ps_text_buf :: Text
  -- ^ for hexpat only, which can break up char data into multiple events
  , _ps_worksheet_ended :: Bool
  -- ^ For hexpat only, which can throw errors right at the end of the sheet
  -- rather than ending gracefully.
  } deriving stock (Generic, Show)
makeLenses 'MkSheetState

deriving via AllowThunksIn
  '[ "_ps_row"
   ] SheetState
  instance NoThunks SheetState

-- | State for parsing shared strings
data SharedStringState = MkSharedStringState
  { _ss_string    :: TB.Builder -- ^ String we are parsing
  , _ss_list :: DL.DList Text -- ^ list of shared strings
  } deriving stock (Generic, Show)
makeLenses 'MkSharedStringState

type HasSheetState = MonadState SheetState
type HasSharedStringState = MonadState SharedStringState

-- | Information about the workbook contained in xl/workbook.xml
-- (currently a subset)
data WorkbookInfo = WorkbookInfo
  { _wiSheets :: Map Int Text
  -- ^ Map from the sheet number (AKA sheetId) to the sheet name
  } deriving Show
makeLenses 'WorkbookInfo

data XlsxMState = MkXlsxMState
  { _xs_shared_strings :: IORef (Maybe (V.Vector Text))
  , _xs_workbook_info :: IORef (Maybe WorkbookInfo)
  }

newtype XlsxM a = XlsxM {_unXlsxM :: ReaderT XlsxMState Zip.ZipArchive a}
  deriving newtype
    ( Functor,
      Applicative,
      Monad,
      MonadIO,
      MonadCatch,
      MonadMask,
      MonadThrow,
      MonadReader XlsxMState,
      MonadBase IO,
      MonadBaseControl IO
    )

-- | Initial parsing state
initialSheetState :: SheetState
initialSheetState = MkSheetState
  { _ps_row             = mempty
  , _ps_sheet_name      = mempty
  , _ps_cell_row_index  = 0
  , _ps_cell_col_index  = 0
  , _ps_is_in_val       = False
  , _ps_shared_strings  = mempty
  , _ps_type            = Untyped
  , _ps_text_buf = mempty
  , _ps_worksheet_ended = False
  }

-- | Initial parsing state
initialSharedString :: SharedStringState
initialSharedString = MkSharedStringState
  { _ss_string = mempty
  , _ss_list = mempty
  }

-- | Parse shared string entry from xml event and return it once
-- we've reached the end of given element
{-# SCC parseSharedString #-}
parseSharedString
  :: ( MonadThrow m
     , HasSharedStringState m
     )
  => HexpatEvent -> m (Maybe Text)
parseSharedString = \case
  StartElement "t" _ -> Nothing <$ (ss_string .= mempty)
  EndElement "t" -> Just . LT.toStrict . TB.toLazyText <$> gets _ss_string
  CharacterData txt -> Nothing <$ (ss_string <>= TB.fromText txt)
  _ -> pure Nothing

-- | Run a series of actions on an Xlsx file
runXlsxM :: MonadIO m => FilePath -> XlsxM a -> m a
runXlsxM xlsxFile (XlsxM act) = do
  env0 <- liftIO $ MkXlsxMState <$> newIORef Nothing <*> newIORef Nothing
  Zip.withArchive xlsxFile $ runReaderT act env0

liftZip :: Zip.ZipArchive a -> XlsxM a
liftZip = XlsxM . ReaderT . const

{-# SCC getOrParseSharedStrings #-}
getOrParseSharedStrings :: XlsxM (V.Vector Text)
getOrParseSharedStrings = do
  sharedStringsRef <- asks _xs_shared_strings
  mSharedStrings <- liftIO $ readIORef sharedStringsRef
  case mSharedStrings of
    Just strs -> pure strs
    Nothing -> do
      sharedStrsSel <- liftZip $ Zip.mkEntrySelector "xl/sharedStrings.xml"
      let state0 = initialSharedString
      byteSrc <- liftZip $ Zip.getEntrySource sharedStrsSel
      st <- runExpat state0 byteSrc $ \evs -> forM_ evs $ \ev -> do
        mTxt <- parseSharedString ev
        for_ mTxt $ \txt ->
          ss_list %= (`DL.snoc` txt)
      let sharedStrings = V.fromList $ DL.toList $ _ss_list st
      liftIO $ writeIORef sharedStringsRef $ Just sharedStrings
      pure sharedStrings

-- | Returns information about the workbook, found in
-- xl/workbook.xml. The result is cached so the XML will only be
-- decompressed and parsed once inside a larger XlsxM action.
getWorkbookInfo :: XlsxM WorkbookInfo
getWorkbookInfo = do
  ref <- asks _xs_workbook_info
  mInfo <- liftIO $ readIORef ref
  case mInfo of
    Just info -> pure info
    Nothing -> do
      sel <- liftZip $ Zip.mkEntrySelector "xl/workbook.xml"
      src <- liftZip $ Zip.getEntrySource sel
      sheets <- runExpat mempty src $ \evs -> forM_ evs $ \case
        StartElement ("sheet" :: ByteString) attrs -> do
          nm <- maybe (throwM MalformedWorkbook) pure $ lookup ("name" :: ByteString) attrs
          sheetId <- maybe (throwM MalformedWorkbook) pure $ lookup "sheetId" attrs
          sheetNum <- either (const $ throwM MalformedWorkbook) pure $ eitherDecimal sheetId
          modify' (M.insert sheetNum nm)
        _ -> pure ()
      let info = WorkbookInfo sheets
      liftIO $ writeIORef ref $ Just info
      pure info

type HexpatEvent = SAXEvent ByteString Text

-- If the given sheet number exists, returns Just a conduit source of the stream
-- of XML events in a particular sheet. Returns Nothing when the sheet doesn't
-- exist.
{-# SCC getSheetXmlSource #-}
getSheetXmlSource ::
  (PrimMonad m, MonadThrow m, C.MonadResource m) =>
  -- | The sheet number
  Int ->
  XlsxM (Maybe (ConduitT () ByteString m ()))
getSheetXmlSource sheetNumber = do
  -- XXX: The Zip library may throw exceptions that aren't exposed from this
  -- module, so downstream library users would need to add the 'zip' package to
  -- handle them. Consider re-wrapping zip library exceptions, or just
  -- re-export them?
  sheetSel <- liftZip $ Zip.mkEntrySelector $ "xl/worksheets/sheet" <> show sheetNumber <> ".xml"
  sheetExists <- liftZip $ Zip.doesEntryExist sheetSel
  if not sheetExists
    then pure Nothing
    else Just <$> liftZip (Zip.getEntrySource sheetSel)

{-# SCC runExpat #-}
runExpat :: forall state tag text.
  (GenericXMLString tag, GenericXMLString text) =>
  state ->
  ConduitT () ByteString (C.ResourceT IO) () ->
  ([SAXEvent tag text] -> StateT state IO ()) ->
  XlsxM state
runExpat initialState byteSource handler = liftIO $ do
  -- Set up state
  ref <- newIORef initialState
  -- Set up parser and callbacks
  (parseChunk, _getLoc) <- Hexpat.hexpatNewParser Nothing Nothing False
  let noExtra _ offset = pure ((), offset)
      {-# SCC processChunk #-}
      {-# INLINE processChunk #-}
      processChunk isFinalChunk chunk = do
        (buf, len, mError) <- parseChunk chunk isFinalChunk
        saxen <- HexpatInternal.parseBuf buf len noExtra
        case mError of
          Just err -> error $ "expat error: " <> show err
          Nothing -> do
            state0 <- liftIO $ readIORef ref
            state1 <-
              {-# SCC "runExpat_runStateT_call" #-}
              execStateT (handler $ map fst saxen) state0
            writeIORef ref state1
  C.runConduitRes $
    byteSource .|
    C.awaitForever (liftIO . processChunk False)
  processChunk True BS.empty
  readIORef ref

runExpatForSheet ::
  SheetState ->
  ConduitT () ByteString (C.ResourceT IO) () ->
  (SheetItem -> IO ()) ->
  XlsxM ()
runExpatForSheet initState byteSource inner =
  void $ runExpat initState byteSource handler
  where
    sheetName = _ps_sheet_name initState
    handler evs = forM_ evs $ \ev -> do
      parseRes <- runExceptT $ matchHexpatEvent ev
      case parseRes of
        Left err -> error $ "error after matchHexpatEvent: " <> show err
        Right (Just cellRow)
          | not (IntMap.null cellRow) -> do
              rowNum <- use ps_cell_row_index
              liftIO $ inner $ MkSheetItem sheetName rowNum cellRow
        _ -> pure ()

-- FIXME: hack to be compatible with the zip-stream parser (no longer
-- present), which as of this writing sets the sheet name to be
-- "/sheetN.xml"
mkSheetName :: Int -> Text
mkSheetName sheetNumber =
     "/sheet"  <> Text.pack (show sheetNumber) <> ".xml"

-- | this will collect the sheetitems in a list.
--   useful for cases were memory is of no concern but a sheetitem
--   type in a list is needed.
collectItems ::
  Int ->
  XlsxM [SheetItem]
collectItems  sheetNumber = do
  res <- liftIO $ newIORef []
  void $ readSheet sheetNumber $ \item -> liftIO (modifyIORef' res (item :))
  liftIO $ readIORef res

readSheet ::
  -- | Sheet number
  Int ->
  -- | Function to consume the sheet's rows
  (SheetItem -> IO ()) ->
  -- | Returns False if sheet doesn't exist, or True otherwise
  XlsxM Bool
readSheet sheetNumber inner = do
  mSrc :: Maybe (ConduitT () ByteString (C.ResourceT IO) ()) <-
    getSheetXmlSource sheetNumber
  case mSrc of
    Nothing -> pure False
    Just sourceSheetXml -> do
      sharedStrs <- getOrParseSharedStrings
      let sheetState0 = initialSheetState
            & ps_shared_strings .~ sharedStrs
            & ps_sheet_name .~ sheetName
      runExpatForSheet sheetState0 sourceSheetXml inner
      pure True
  where
    sheetName = mkSheetName sheetNumber

-- | Returns number of rows in the given sheet (identified by sheet
-- number), or Nothing if the sheet does not exist. Does not perform a
-- full parse of the XML into 'SheetItem's, so it should be more
-- efficient than counting via 'readSheet'.
countRowsInSheet :: Int -> XlsxM (Maybe Int)
countRowsInSheet sheetNumber = do
  mSrc :: Maybe (ConduitT () ByteString (C.ResourceT IO) ()) <-
    getSheetXmlSource sheetNumber
  for mSrc $ \sourceSheetXml -> do
    runExpat @Int @ByteString @ByteString 0 sourceSheetXml $ \evs ->
      forM_ evs $ \case
        StartElement "row" _ -> modify' (+1)
        _ -> pure ()

-- | Return row from the state and empty it
popRow :: HasSheetState m => m CellRow
popRow = do
  row <- use ps_row
  ps_row .= mempty
  pure row

data AddCellErrors
  = ReadError -- ^ Could not read current cell value
      Text    -- ^ Original value
      String  -- ^ Error message
  | SharedStringNotFound -- ^ Could not find string by index in shared string table
      Int                -- ^ Given index
      (V.Vector Text)      -- ^ Given shared strings to lookup in
  deriving Show

-- | Parse the given value
--
-- If it's a string, we try to get it our of a shared string table
{-# SCC parseValue #-}
parseValue :: SharedStringMap -> Text -> ExcelValueType -> Either AddCellErrors CellValue
parseValue sstrings txt = \case
  TS -> do
    (idx, _) <- ReadError txt `first` Read.decimal @Int txt
    string <- maybe (Left $ SharedStringNotFound idx sstrings) Right $ {-# SCC "sstrings_lookup_scc" #-}  sstrings ^? ix idx
    Right $ CellText string
  TStr -> pure $ CellText txt
  TN -> bimap (ReadError txt) (CellDouble . fst) $ Read.double txt
  TE -> bimap (ReadError txt) (CellError . fst) $ fromAttrVal txt
  TB | txt == "1" -> Right $ CellBool True
     | txt == "0" -> Right $ CellBool False
     | otherwise -> Left $ ReadError txt "Could not read Excel boolean value (expected 0 or 1)"
  Untyped -> Right (parseUntypedValue txt)

-- TODO: some of the cells are untyped and we need to test whether
-- they all are strings or something more complicated
parseUntypedValue :: Text -> CellValue
parseUntypedValue = CellText

-- | Adds a cell to row in state monad
{-# SCC addCellToRow #-}
addCellToRow
  :: ( MonadError SheetErrors m
     , HasSheetState m
     )
  => Text -> m ()
addCellToRow txt = do
  st <- get
  when (_ps_is_in_val st) $ do
    val <- liftEither $ first ParseCellError $ parseValue (_ps_shared_strings st) txt (_ps_type st)
    put $ st { _ps_row = IntMap.insert (_ps_cell_col_index st)
                         (Cell { _cellStyle   = Nothing
                               , _cellValue   = Just val
                               , _cellComment = Nothing
                               , _cellFormula = Nothing
                               }) $ _ps_row st}

data SheetErrors
  = ParseCoordinateError CoordinateErrors -- ^ Error while parsing coordinates
  | ParseTypeError TypeError              -- ^ Error while parsing types
  | ParseCellError AddCellErrors          -- ^ Error while parsing cells
  | HexpatParseError Hexpat.XMLParseError
  deriving stock Show
  deriving anyclass Exception

type SheetValue = (ByteString, Text)
type SheetValues = [SheetValue]

data CoordinateErrors
  = CoordinateNotFound SheetValues         -- ^ If the coordinate was not specified in "r" attribute
  | NoListElement SheetValue SheetValues   -- ^ If the value is empty for some reason
  | NoTextContent Content SheetValues      -- ^ If the value has something besides @ContentText@ inside
  | DecodeFailure Text SheetValues         -- ^ If malformed coordinate text was passed
  deriving stock Show
  deriving anyclass Exception

data TypeError
  = TypeNotFound SheetValues
  | TypeNoListElement SheetValue SheetValues
  | UnkownType Text SheetValues
  | TypeNoTextContent Content SheetValues
  deriving Show
  deriving anyclass Exception

data WorkbookError = MalformedWorkbook
  deriving Show
  deriving anyclass Exception

{-# SCC matchHexpatEvent #-}
matchHexpatEvent ::
  ( MonadError SheetErrors m,
    HasSheetState m
  ) =>
  HexpatEvent ->
  m (Maybe CellRow)
matchHexpatEvent ev = case ev of
  CharacterData txt -> {-# SCC "handle_CharData" #-} do
    inVal <- use ps_is_in_val
    when inVal $
      {-# SCC "append_text_buf" #-} ps_text_buf <>= txt
    pure Nothing
  StartElement "c" attrs -> Nothing <$ (setCoord attrs >> setType attrs)
  StartElement "v" _ -> Nothing <$ (ps_is_in_val .= True)
  EndElement "v" -> {-# SCC "handle_EndElementV" #-} do
    txt <- gets _ps_text_buf
    {-# SCC "call_addCellToRow" #-} addCellToRow txt
    {-# SCC "call_modify_endelemV" #-} modify' $ \st ->
      st { _ps_is_in_val = False
         , _ps_text_buf = mempty
         }
    pure Nothing
  -- If beginning of row, empty the state and return nothing [why exactly?]
  StartElement "row" _ -> Nothing <$ popRow
  -- If at the end of the row, we have collected the whole row into
  -- the current state. Empty the state and return the row.
  EndElement "row" -> Just <$> popRow
  StartElement "worksheet" _ -> ps_worksheet_ended .= False >> pure Nothing
  EndElement "worksheet" -> ps_worksheet_ended .= True >> pure Nothing
  -- Skip everything else, e.g. the formula elements <f>
  FailDocument err -> do
    -- this event is emitted at the end the xml stream (possibly
    -- because the xml files in xlsx archives don't end in a
    -- newline, but that's a guess), so we use state to determine if
    -- it's expected.
    finished <- use ps_worksheet_ended
    unless finished $
      throwError $ HexpatParseError err
    pure Nothing
  _ -> pure Nothing

-- | Update state coordinates accordingly to @parseCoordinates@
{-# SCC setCoord #-}
setCoord
  :: ( MonadError SheetErrors m
     , HasSheetState m
     )
  => SheetValues -> m ()
setCoord list = do
  coordinates <- liftEither $ first ParseCoordinateError $ parseCoordinates list
  ps_cell_col_index .= (coordinates ^. _2)
  ps_cell_row_index .= (coordinates ^. _1)

-- | Parse type from values and update state accordingly
setType
  :: ( MonadError SheetErrors m
     , HasSheetState m
 )
  => SheetValues -> m ()
setType list = do
  type' <- liftEither $ first ParseTypeError $ parseType list
  ps_type .= type'

-- | Find sheet value by its name
findName :: ByteString -> SheetValues -> Maybe SheetValue
findName name = find ((name ==) . fst)
{-# INLINE findName #-}

-- | Parse value type
{-# SCC parseType #-}
parseType :: SheetValues -> Either TypeError ExcelValueType
parseType list =
  case findName "t" list of
    Nothing -> pure Untyped
    Just (_nm, valText)->
      case valText of
        "n"   -> Right TN
        "s"   -> Right TS
        "str" -> Right TStr
        "b"   -> Right TB
        "e"   -> Right TE
        other -> Left $ UnkownType other list

-- | Parse coordinates from a list of xml elements if such were found on "r" key
{-# SCC parseCoordinates #-}
parseCoordinates :: SheetValues -> Either CoordinateErrors (Int, Int)
parseCoordinates list = do
  (_nm, valText) <- maybe (Left $ CoordinateNotFound list) Right $ findName "r" list
  maybe (Left $ DecodeFailure valText list) Right $ fromSingleCellRef $ CellRef valText
