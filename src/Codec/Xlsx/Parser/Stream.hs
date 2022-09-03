{-# LANGUAGE CPP                        #-}
{-# LANGUAGE ConstraintKinds            #-}
{-# LANGUAGE DeriveAnyClass             #-}
{-# LANGUAGE DeriveGeneric              #-}
{-# LANGUAGE DerivingStrategies         #-}
{-# LANGUAGE FlexibleContexts           #-}
{-# LANGUAGE GeneralizedNewtypeDeriving #-}
{-# LANGUAGE LambdaCase                 #-}
{-# LANGUAGE OverloadedStrings          #-}
{-# LANGUAGE PackageImports             #-}
{-# LANGUAGE RankNTypes                 #-}
{-# LANGUAGE RecordWildCards            #-}
{-# LANGUAGE ScopedTypeVariables        #-}
{-# LANGUAGE StrictData                 #-}
{-# LANGUAGE TemplateHaskell            #-}
{-# LANGUAGE TypeApplications           #-}
{-# LANGUAGE UndecidableInstances       #-}

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
-- particular sheet, using 'readSheetByIndex', which is callback-based and tied to IO.
--
module Codec.Xlsx.Parser.Stream
  ( XlsxM
  , runXlsxM
  , WorkbookInfo(..)
  , SheetInfo(..)
  , wiSheets
  , getWorkbookInfo
  , CellRow
  , readSheet
  , countRowsInSheet
  , collectItems
  -- ** Index
  , SheetIndex
  , makeIndex
  , makeIndexFromName
  -- ** SheetItem
  , SheetItem(..)
  , si_sheet_index
  , si_row
  -- ** Row
  , Row(..)
  , ri_row_index
  , ri_cell_row
  -- * Errors
  , SheetErrors(..)
  , AddCellErrors(..)
  , CoordinateErrors(..)
  , TypeError(..)
  , WorkbookError(..)
  ) where

import qualified "zip" Codec.Archive.Zip as Zip
import Codec.Xlsx.Types.Cell
import Codec.Xlsx.Types.Common
import Codec.Xlsx.Types.Internal (RefId (..))
import Codec.Xlsx.Types.Internal.Relationships (Relationship (..),
                                                Relationships (..))
import Conduit (PrimMonad, (.|))
import qualified Conduit as C
import qualified Data.Vector as V
#ifdef USE_MICROLENS
import Lens.Micro
import Lens.Micro.GHC ()
import Lens.Micro.Mtl
import Lens.Micro.Platform
import Lens.Micro.TH
#else
import Control.Lens
#endif
import Codec.Xlsx.Parser.Internal
import Codec.Xlsx.Parser.Internal.Memoize
import Control.DeepSeq
import Control.Monad.Catch
import Control.Monad.Except
import Control.Monad.Reader
import Control.Monad.State.Strict
import Data.Bifunctor
import Data.ByteString (ByteString)
import qualified Data.ByteString as BS
import Data.Conduit (ConduitT)
import qualified Data.DList as DL
import Data.Foldable
import Data.IntMap.Strict (IntMap)
import qualified Data.IntMap.Strict as IntMap
import Data.IORef
import qualified Data.Map.Strict as M
import Data.Text (Text)
import qualified Data.Text as T
import qualified Data.Text.Lazy as LT
import qualified Data.Text.Lazy.Builder as TB
import qualified Data.Text.Read as Read
import Data.Traversable (for)
import Data.XML.Types
import GHC.Generics

import qualified Codec.Xlsx.Parser.Stream.HexpatInternal as HexpatInternal
import Control.Monad.Base
import Control.Monad.Trans.Control
import Text.XML.Expat.Internal.IO as Hexpat
import Text.XML.Expat.SAX as Hexpat

#ifdef USE_MICROLENS
(<>=) :: (MonadState s m, Monoid a) => ASetter' s a -> a -> m ()
l <>= a = modify (l <>~ a)
#else
#endif

type CellRow = IntMap Cell

-- | Sheet item
--
-- The current sheet at a time, every sheet is constructed of these items.
data SheetItem = MkSheetItem
  { _si_sheet_index :: Int       -- ^ The sheet number
  , _si_row         :: ~Row
  } deriving stock (Generic, Show)
    deriving anyclass NFData

data Row = MkRow
  { _ri_row_index :: Int       -- ^ Row number
  , _ri_cell_row  :: ~CellRow  -- ^ Row itself
  } deriving stock (Generic, Show)
    deriving anyclass NFData

makeLenses 'MkSheetItem
makeLenses 'MkRow

type SharedStringsMap = V.Vector Text

-- | Type of the excel value
--
-- Note: Some values are untyped and rules of their type resolution are not known.
-- They may be treated simply as strings as well as they may be context-dependent.
-- By far we do not bother with it.
data ExcelValueType
  = TS      -- ^ shared string
  | TStr    -- ^ either an inline string ("inlineStr") or a formula string ("str")
  | TN      -- ^ number
  | TB      -- ^ boolean
  | TE      -- ^ excell error, the sheet can contain error values, for example if =1/0, causes division by zero
  | Untyped -- ^ Not all values have types
  deriving stock (Generic, Show)

-- | State for parsing sheets
data SheetState = MkSheetState
  { _ps_row             :: ~CellRow        -- ^ Current row
  , _ps_sheet_index     :: Int             -- ^ Current sheet ID (AKA 'sheetInfoSheetId')
  , _ps_cell_row_index  :: Int             -- ^ Current row number
  , _ps_cell_col_index  :: Int             -- ^ Current column number
  , _ps_cell_style      :: Maybe Int
  , _ps_is_in_val       :: Bool            -- ^ Flag for indexing wheter the parser is in value or not
  , _ps_shared_strings  :: SharedStringsMap -- ^ Shared string map
  , _ps_type            :: ExcelValueType  -- ^ The last detected value type

  , _ps_text_buf        :: Text
  -- ^ for hexpat only, which can break up char data into multiple events
  , _ps_worksheet_ended :: Bool
  -- ^ For hexpat only, which can throw errors right at the end of the sheet
  -- rather than ending gracefully.
  } deriving stock (Generic, Show)
makeLenses 'MkSheetState

-- | State for parsing shared strings
data SharedStringsState = MkSharedStringsState
  { _ss_string :: TB.Builder -- ^ String we are parsing
  , _ss_list   :: DL.DList Text -- ^ list of shared strings
  } deriving stock (Generic, Show)
makeLenses 'MkSharedStringsState

type HasSheetState = MonadState SheetState
type HasSharedStringsState = MonadState SharedStringsState

-- | Represents sheets from the workbook.xml file. E.g.
-- <sheet name="Data" sheetId="1" state="hidden" r:id="rId2" /
data SheetInfo = SheetInfo
  { sheetInfoName    :: Text,
    -- | The r:id attribute value.
    sheetInfoRelId   :: RefId,
    -- | The sheetId attribute value
    sheetInfoSheetId :: Int
  } deriving (Show, Eq)

-- | Information about the workbook contained in xl/workbook.xml
-- (currently a subset)
data WorkbookInfo = WorkbookInfo
  { _wiSheets :: [SheetInfo]
  } deriving Show
makeLenses 'WorkbookInfo

data XlsxMState = MkXlsxMState
  { _xs_shared_strings :: Memoized (V.Vector Text)
  , _xs_workbook_info  :: Memoized WorkbookInfo
  , _xs_relationships  :: Memoized Relationships
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
  , _ps_sheet_index     = 0
  , _ps_cell_row_index  = 0
  , _ps_cell_col_index  = 0
  , _ps_is_in_val       = False
  , _ps_shared_strings  = mempty
  , _ps_type            = Untyped
  , _ps_text_buf        = mempty
  , _ps_worksheet_ended = False
  , _ps_cell_style      = Nothing
  }

-- | Initial parsing state
initialSharedStrings :: SharedStringsState
initialSharedStrings = MkSharedStringsState
  { _ss_string = mempty
  , _ss_list = mempty
  }

-- | Parse shared string entry from xml event and return it once
-- we've reached the end of given element
{-# SCC parseSharedStrings #-}
parseSharedStrings
  :: ( MonadThrow m
     , HasSharedStringsState m
     )
  => HexpatEvent -> m (Maybe Text)
parseSharedStrings = \case
  StartElement "t" _ -> Nothing <$ (ss_string .= mempty)
  EndElement "t"     -> Just . LT.toStrict . TB.toLazyText <$> gets _ss_string
  CharacterData txt  -> Nothing <$ (ss_string <>= TB.fromText txt)
  _                  -> pure Nothing

-- | Run a series of actions on an Xlsx file
runXlsxM :: MonadIO m => FilePath -> XlsxM a -> m a
runXlsxM xlsxFile (XlsxM act) = liftIO $ do
  -- TODO: don't run the withArchive multiple times but use liftWith or runInIO instead
  _xs_workbook_info  <- memoizeRef (Zip.withArchive xlsxFile readWorkbookInfo)
  _xs_relationships  <- memoizeRef (Zip.withArchive xlsxFile readWorkbookRelationships)
  _xs_shared_strings <- memoizeRef (Zip.withArchive xlsxFile parseSharedStringss)
  Zip.withArchive xlsxFile $ runReaderT act $ MkXlsxMState{..}

liftZip :: Zip.ZipArchive a -> XlsxM a
liftZip = XlsxM . ReaderT . const

parseSharedStringss :: Zip.ZipArchive (V.Vector Text)
parseSharedStringss = do
      sharedStrsSel <- Zip.mkEntrySelector "xl/sharedStrings.xml"
      hasSharedStrs <- Zip.doesEntryExist sharedStrsSel
      if not hasSharedStrs
        then pure mempty
        else do
          let state0 = initialSharedStrings
          byteSrc <- Zip.getEntrySource sharedStrsSel
          st <- liftIO $ runExpat state0 byteSrc $ \evs -> forM_ evs $ \ev -> do
            mTxt <- parseSharedStrings ev
            for_ mTxt $ \txt ->
              ss_list %= (`DL.snoc` txt)
          pure $ V.fromList $ DL.toList $ _ss_list st

{-# SCC getOrParseSharedStringss #-}
getOrParseSharedStringss :: XlsxM (V.Vector Text)
getOrParseSharedStringss = runMemoized =<< asks _xs_shared_strings

readWorkbookInfo :: Zip.ZipArchive WorkbookInfo
readWorkbookInfo = do
   sel <- Zip.mkEntrySelector "xl/workbook.xml"
   src <- Zip.getEntrySource sel
   sheets <- liftIO $ runExpat [] src $ \evs -> forM_ evs $ \case
     StartElement ("sheet" :: ByteString) attrs -> do
       nm <- lookupBy "name" attrs
       sheetId <- lookupBy "sheetId" attrs
       rId <- lookupBy "r:id" attrs
       sheetNum <- either (throwM . ParseDecimalError sheetId) pure $ eitherDecimal sheetId
       modify' (SheetInfo nm (RefId rId) sheetNum :)
     _ -> pure ()
   pure $ WorkbookInfo sheets

lookupBy :: MonadThrow m => ByteString -> [(ByteString, Text)] -> m Text
lookupBy fields attrs = maybe (throwM $ LookupError attrs fields) pure $ lookup fields attrs

-- | Returns information about the workbook, found in
-- xl/workbook.xml. The result is cached so the XML will only be
-- decompressed and parsed once inside a larger XlsxM action.
getWorkbookInfo :: XlsxM WorkbookInfo
getWorkbookInfo = runMemoized =<< asks _xs_workbook_info

readWorkbookRelationships :: Zip.ZipArchive Relationships
readWorkbookRelationships = do
   sel <- Zip.mkEntrySelector "xl/_rels/workbook.xml.rels"
   src <- Zip.getEntrySource sel
   liftIO $ fmap Relationships $ runExpat mempty src $ \evs -> forM_ evs $ \case
     StartElement ("Relationship" :: ByteString) attrs -> do
       rId <- lookupBy "Id" attrs
       rTarget <- lookupBy "Target" attrs
       rType <- lookupBy "Type" attrs
       modify' $ M.insert (RefId rId) $
         Relationship { relType = rType,
                        relTarget = T.unpack rTarget
                       }
     _ -> pure ()

-- | Gets relationships for the workbook (this means the filenames in
-- the relationships map are relative to "xl/" base path within the
-- zip file.
--
-- The relationships xml file will only be parsed once when called
-- multiple times within a larger XlsxM action.
getWorkbookRelationships :: XlsxM Relationships
getWorkbookRelationships = runMemoized =<< asks _xs_relationships

type HexpatEvent = SAXEvent ByteString Text

relIdToEntrySelector :: RefId -> XlsxM (Maybe Zip.EntrySelector)
relIdToEntrySelector rid = do
  Relationships rels <- getWorkbookRelationships
  for (M.lookup rid rels) $ \rel -> do
    Zip.mkEntrySelector $ "xl/" <> relTarget rel

sheetIdToRelId :: Int -> XlsxM (Maybe RefId)
sheetIdToRelId sheetId = do
  WorkbookInfo sheets <- getWorkbookInfo
  pure $ sheetInfoRelId <$> find ((== sheetId) . sheetInfoSheetId) sheets

sheetIdToEntrySelector :: Int -> XlsxM (Maybe Zip.EntrySelector)
sheetIdToEntrySelector sheetId = do
  sheetIdToRelId sheetId >>= \case
    Nothing  -> pure Nothing
    Just rid -> relIdToEntrySelector rid

-- If the given sheet number exists, returns Just a conduit source of the stream
-- of XML events in a particular sheet. Returns Nothing when the sheet doesn't
-- exist.
{-# SCC getSheetXmlSource #-}
getSheetXmlSource ::
  (PrimMonad m, MonadThrow m, C.MonadResource m) =>
  Int ->
  XlsxM (Maybe (ConduitT () ByteString m ()))
getSheetXmlSource sheetId = do
  -- TODO: The Zip library may throw exceptions that aren't exposed from this
  -- module, so downstream library users would need to add the 'zip' package to
  -- handle them. Consider re-wrapping zip library exceptions, or just
  -- re-export them?
  mSheetSel <- sheetIdToEntrySelector sheetId
  sheetExists <- maybe (pure False) (liftZip . Zip.doesEntryExist) mSheetSel
  case mSheetSel of
    Just sheetSel
      | sheetExists ->
          Just <$> liftZip (Zip.getEntrySource sheetSel)
    _ -> pure Nothing

{-# SCC runExpat #-}
runExpat :: forall state tag text.
  (GenericXMLString tag, GenericXMLString text) =>
  state ->
  ConduitT () ByteString (C.ResourceT IO) () ->
  ([SAXEvent tag text] -> StateT state IO ()) ->
  IO state
runExpat initialState byteSource handler = do
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
  void $ liftIO $ runExpat initState byteSource handler
  where
    sheetName = _ps_sheet_index initState
    handler evs = forM_ evs $ \ev -> do
      parseRes <- runExceptT $ matchHexpatEvent ev
      case parseRes of
        Left err -> throwM err
        Right (Just cellRow)
          | not (IntMap.null cellRow) -> do
              rowNum <- use ps_cell_row_index
              liftIO $ inner $ MkSheetItem sheetName $ MkRow rowNum cellRow
        _ -> pure ()

-- | this will collect the sheetitems in a list.
--   useful for cases were memory is of no concern but a sheetitem
--   type in a list is needed.
collectItems ::
  SheetIndex ->
  XlsxM [SheetItem]
collectItems sheetId = do
 res <- liftIO $ newIORef []
 void $ readSheet sheetId $ \item ->
   liftIO (modifyIORef' res (item :))
 fmap reverse $ liftIO $ readIORef res

-- | datatype representing a sheet index, looking it up by name
--   can be done with 'makeIndexFromName', which is the preferred approach.
--   although 'makeIndex' is available in case it's already known.
newtype SheetIndex = MkSheetIndex Int
 deriving newtype NFData

-- | This does *no* checking if the index exists or not.
--   you could have index out of bounds issues because of this.
makeIndex :: Int -> SheetIndex
makeIndex = MkSheetIndex

-- | Look up the index of a case insensitive sheet name
makeIndexFromName :: Text -> XlsxM (Maybe SheetIndex)
makeIndexFromName sheetName = do
  wi <- getWorkbookInfo
  -- The Excel UI does not allow a user to create two sheets whose
  -- names differ only in alphabetic case (at least for ascii...)
  let sheetNameCI = T.toLower sheetName
      findRes :: Maybe SheetInfo
      findRes = find ((== sheetNameCI) . T.toLower . sheetInfoName) $ _wiSheets wi
  pure $ makeIndex . sheetInfoSheetId <$> findRes

readSheet ::
  SheetIndex ->
  -- | Function to consume the sheet's rows
  (SheetItem -> IO ()) ->
  -- | Returns False if sheet doesn't exist, or True otherwise
  XlsxM Bool
readSheet (MkSheetIndex sheetId) inner = do
  mSrc :: Maybe (ConduitT () ByteString (C.ResourceT IO) ()) <-
    getSheetXmlSource sheetId
  let
  case mSrc of
    Nothing -> pure False
    Just sourceSheetXml -> do
      sharedStrs <- getOrParseSharedStringss
      let sheetState0 = initialSheetState
            & ps_shared_strings .~ sharedStrs
            & ps_sheet_index .~ sheetId
      runExpatForSheet sheetState0 sourceSheetXml inner
      pure True

-- | Returns number of rows in the given sheet (identified by the
-- sheet's ID, AKA the sheetId attribute, AKA 'sheetInfoSheetId'), or Nothing
-- if the sheet does not exist. Does not perform a full parse of the
-- XML into 'SheetItem's, so it should be more efficient than counting
-- via 'readSheetByIndex'.
countRowsInSheet :: SheetIndex -> XlsxM (Maybe Int)
countRowsInSheet (MkSheetIndex sheetId) = do
  mSrc :: Maybe (ConduitT () ByteString (C.ResourceT IO) ()) <-
    getSheetXmlSource sheetId
  for mSrc $ \sourceSheetXml -> do
    liftIO $ runExpat @Int @ByteString @ByteString 0 sourceSheetXml $ \evs ->
      forM_ evs $ \case
        StartElement "row" _ -> modify' (+1)
        _                    -> pure ()

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
  | SharedStringsNotFound -- ^ Could not find string by index in shared string table
      Int                -- ^ Given index
      (V.Vector Text)      -- ^ Given shared strings to lookup in
  deriving Show

-- | Parse the given value
--
-- If it's a string, we try to get it our of a shared string table
{-# SCC parseValue #-}
parseValue :: SharedStringsMap -> Text -> ExcelValueType -> Either AddCellErrors CellValue
parseValue sstrings txt = \case
  TS -> do
    (idx, _) <- ReadError txt `first` Read.decimal @Int txt
    string <- maybe (Left $ SharedStringsNotFound idx sstrings) Right $ {-# SCC "sstrings_lookup_scc" #-}  (sstrings ^? ix idx)
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
  style <- use ps_cell_style
  when (_ps_is_in_val st) $ do
    val <- liftEither $ first ParseCellError $ parseValue (_ps_shared_strings st) txt (_ps_type st)
    put $ st { _ps_row = IntMap.insert (_ps_cell_col_index st)
                         (Cell { _cellStyle   = style
                               , _cellValue   = Just val
                               , _cellComment = Nothing
                               , _cellFormula = Nothing
                               }) $ _ps_row st}

data SheetErrors
  = ParseCoordinateError CoordinateErrors -- ^ Error while parsing coordinates
  | ParseTypeError TypeError              -- ^ Error while parsing types
  | ParseCellError AddCellErrors          -- ^ Error while parsing cells
  | ParseStyleErrors StyleError
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

data WorkbookError = LookupError { lookup_attrs :: [(ByteString, Text)], lookup_field :: ByteString }
                   | ParseDecimalError Text String
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
      {-# SCC "append_text_buf" #-} (ps_text_buf <>= txt)
    pure Nothing
  StartElement "c" attrs -> Nothing <$ (setCoord attrs *> setType attrs *> setStyle attrs)
  StartElement "is" _ -> Nothing <$ (ps_is_in_val .= True)
  EndElement "is" -> Nothing <$ finaliseCellValue
  StartElement "v" _ -> Nothing <$ (ps_is_in_val .= True)
  EndElement "v" -> Nothing <$ finaliseCellValue
  -- If beginning of row, empty the state and return nothing.
  -- We don't know if there is anything in the state, the user may have
  -- decided to <row> <row> (not closing). In any case it's the beginning of a new row
  -- so we clear the state.
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

{-# INLINE finaliseCellValue #-}
finaliseCellValue ::
  ( MonadError SheetErrors m, HasSheetState m ) => m ()
finaliseCellValue = do
  txt <- gets _ps_text_buf
  addCellToRow txt
  modify' $ \st ->
    st { _ps_is_in_val = False
       , _ps_text_buf = mempty
       }

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

setStyle :: (MonadError SheetErrors m, HasSheetState m) => SheetValues -> m ()
setStyle list = do
  style <- liftEither $ first ParseStyleErrors $ parseStyle list
  ps_cell_style .= style

data StyleError = InvalidStyleRef { seInput:: Text,  seErrorMsg :: String}
  deriving Show

parseStyle :: SheetValues -> Either StyleError (Maybe Int)
parseStyle list =
  case findName "s" list of
    Nothing -> pure Nothing
    Just (_nm, valTex) -> case Read.decimal valTex of
      Left err        -> Left (InvalidStyleRef valTex err)
      Right (i, _rem) -> pure $ Just i

-- | Parse value type
{-# SCC parseType #-}
parseType :: SheetValues -> Either TypeError ExcelValueType
parseType list =
  case findName "t" list of
    -- NB: According to format specification default value for cells without
    -- `t` attribute is a `n` - number.
    --
    -- <xsd:complexType name="CT_Cell" from spec (see the `CellValue` spec reference)>
    --  ..
    --  <xsd:attribute name="t" type="ST_CellType" use="optional" default="n"/>
    -- </xsd:complexType>
    Nothing -> Right TN
    Just (_nm, valText)->
      case valText of
        "n"         -> Right TN
        "s"         -> Right TS
         -- "Cell containing a formula string". Probably shouldn't be TStr..
        "str"       -> Right TStr
        "inlineStr" -> Right TStr
        "b"         -> Right TB
        "e"         -> Right TE
        other       -> Left $ UnkownType other list

-- | Parse coordinates from a list of xml elements if such were found on "r" key
{-# SCC parseCoordinates #-}
parseCoordinates :: SheetValues -> Either CoordinateErrors (Int, Int)
parseCoordinates list = do
  (_nm, valText) <- maybe (Left $ CoordinateNotFound list) Right $ findName "r" list
  maybe (Left $ DecodeFailure valText list) Right $ fromSingleCellRef $ CellRef valText
