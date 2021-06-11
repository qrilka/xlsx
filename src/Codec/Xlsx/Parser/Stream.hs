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

-- |
-- Module      : Codex.Xlsx.Parser.Stream
-- Description : Stream parser for xlsx files
-- Copyright   :
--   (c) Jappie Klooster, 2021
--   Adam, 2021
-- License     : MIT
-- Stability   : experimental
-- Portability : POSIX
--
-- Parse @.xlsx@ using conduit streaming capabilities.
--
-- == Goal
-- The prime goal of this module was to be able to use stream processing
-- capabilities of `conduit` library to be able to parse @.xlsx@ files in
-- constant memory. Below I explain why this is problematic.
--
-- == Implementation
-- A @.xlsx@ file is essentially an archive with a bunch of @.xml@'s underneath.
--
-- Here's a sample of how data looks internally:
--
-- @
-- .
-- ├── [Content_Types].xml
-- ├── _rels
-- ├── docProps
-- │   ├── app.xml
-- │   └── core.xml
-- └── xl
--     ├── _rels
--     │   └── workbook.xml.rels
--     ├── calcChain.xml
--     ├── sharedStrings.xml
--     ├── styles.xml
--     ├── theme
--     │   └── theme1.xml
--     ├── workbook.xml
--     └── worksheets
--         ├── sheet1.xml
--         └── sheet2.xml
-- @
--
-- Here we mostly pay attention to @sharedStrings.xml@ and @workesheets/@ subdirectory.
--
-- The @sharedStrings.xml@ file stores all the unique strings in all of the worksheets
-- @.xlsx@ file and all (actually not all, some of the numbers are not stored in shared
-- strings table for reason I could find explanation for) strings are referenced from this
-- shared strings table. This is used mostly to safe space and causes a problem for us to
-- use conduits in first place, since to be able to parse sheets we have to store this table
-- in memory which does not agree with our only requirement.
--
-- The parsing is then separated to two stages:
--
--   1. We parse shared string table in @parseSharedStringsC@ conduit and collect entries
--      using strict state monad. This is done that way because @xml-types@ library uses
--      eventful approach on separating xml structures.
--   2. We parse the sheets using shared string table that is loaded onto memory in
--      @readXlsxWithSharedStringMapC@.
--
-- Well, actually we do not parse things in this module, just construct conduit for it,
-- for the example usage of this parser you can refer to @benchmarks/Stream.hs@ file or read
-- the subsection explaining its @parseFileStream@ below.
--
-- == How to use this parser
--
-- The starting point of parsing @.xlsx@ via conduits is `readXlsxC` parser.
-- Basically what it does is that it constructs a conduit parser for some
-- existing conduit (or sink if you prefer in this case) that produces `BS.ByteString`
-- which is then passed to our parser conduits and wrapped into some @PrimMonad m@ monad.
--
-- Consider the type of conduits
--
-- @
-- ConduitT a b m r
--          ^------- The value conduit consumes
--            ^------ The intermediate result that is passed to other conduits in the pipeline
--              ^----- The base monad that tells what kind of effects we can perform with the value
--                ^---- The result of the entire pipeline
-- @
--
-- If you're doing it in @IO@ it could be simply invoked from @sourceFile@ that constructs
-- the needed conduit from the source filepath it was given and later the conduit can be
-- evaluated by calling @runResourceT@ that takes the given @ResourceT@
-- (produced by `readXlsxC` function) and returns the value in base monad it was specified.
--
-- In types
--
-- @
-- parseFileC :: FilePath -> IO ByteString
-- parseFileC filePath = do
--   (inputC :: ConduitT () SheetItem (ResourceT IO) ()) <-
--       (sourceFile filepath :: ConduitT () ByteString (ResourceT IO) ())
--        -- ^ Construct a conduit that returns @ByteString@ with @IO@ based effects
--     & \s -> (readXlsxC s :: ResourceT IO (ConduitT () SheetItem (ResourceT IO) ()))
--        -- ^ Construct a conduit that parses the given chunk of @ByteString@ value
--        -- and parses it to @SheetItem@
--     & \s -> (runResourceT s :: IO (ConduitT () SheetItem (ResourceT IO) ()))
--        -- ^ This is needed because we want to allocate the needed resources and be able
--        -- to safely clean up after
--   res <-
--         (parseConduit inputC :: ConduitT () Void (ResourceT IO) [SheetItem])
--         -- ^ Construct a conduit that only returns resulting value out of the
--         -- given @inputC@ conduit so that we can collect all the values it parses
--       & \s -> (runConduit s :: ResourceT IO [SheetItem])
--       & \s -> (runResourceT s :: IO [SheetItem])
--   pure res
--   where
--     -- Fold all the intermediate results of the given conduit and
--     -- return the conduit that only returns resulting value.
--     parseConduit
--       :: Monad m
--       => ConduitM a SheetItem m () -> ConduitM () Void m [SheetItem]
--
-- @
{-# LANGUAGE TypeApplications    #-}
module Codec.Xlsx.Parser.Stream
  ( -- * Parsers
    readXlsxC
  , readXlsxWithSharedStringMapC
  , parseSharedStringsC
  , parseSharedStringsIntoMapC
  -- * Structs
  , CellRow
  , SheetItem(..)
  -- ** `SheetItem` lens
  , si_sheet
  , si_row_index
  , si_cell_row

  -- ** API based on 'zip' library for selective sheet reading
  , XlsxM
  , runXlsxM
  , getSheetSource
  , getOrParseSharedStrings
  ) where

import           Codec.Archive.Zip.Conduit.UnZip
import qualified "zip" Codec.Archive.Zip as Zip
import           Codec.Xlsx.Types.Cell
import           Codec.Xlsx.Types.Common
import           Conduit                         (PrimMonad, await, yield, (.|))
import qualified Conduit                         as C
import           Control.Lens
import           Control.Monad.Catch
import           Control.Monad.Except
import           Control.Monad.Reader
import           Control.Monad.State.Strict
import           Data.Bifunctor
import qualified Data.ByteString                 as BS
import           Data.Conduit                    (ConduitT)
import qualified Data.Conduit.Combinators        as C
import qualified Data.Conduit.List               as CL
import           Data.Foldable
import           Data.IORef
import           Data.Map.Strict                 (Map)
import qualified Data.Map.Strict                 as Map
import           Data.Text                       (Text)
import qualified Data.Text                       as Text
import qualified Data.Text.Encoding              as Text
import qualified Data.Text.Read                  as Read
import           Data.XML.Types
import           GHC.Generics
import           NoThunks.Class
import           Text.Read                       hiding (lift)
import           Text.XML.Stream.Parse
import Codec.Xlsx.Parser.Internal

type CellRow = Map Int Cell

-- | Sheet item
--
-- The current sheet at a time, every sheet is constructed of these items.
data SheetItem = MkSheetItem
  { _si_sheet     :: Text      -- ^ Name of the sheet
  , _si_row_index :: Int       -- ^ Row number
  , _si_cell_row  :: ~CellRow  -- ^ Row itself
  } deriving stock (Generic, Show)
makeLenses 'MkSheetItem

deriving via AllowThunksIn
  '[ "_si_cell_row"
   ] SheetItem
  instance NoThunks SheetItem

-- | File tags that we match with zip entries to be able to separate them and use needed
-- http://officeopenxml.com/anatomyofOOXML-xlsx.php
data PsFileTag =
              Ignored Text    -- ^ Files that appear in workbook, but get ignored
             | Sheet Text     -- ^ Sheet itself
             | InitialNoFile  -- ^ Initial state, not related to excel files and hence should be moved to some other data
             | SharedStrings  -- ^ Shared strings file
             | Styles
             | Workbook       -- ^ Contains names of the tabs and shared string ids
             | ContentTypes
             | Relationships
             | FormulaChain   -- ^ Contains all the formulas in the sheet
             | SheetRel Text  -- ^ Contains relations between internal ids and paths
  deriving stock (Generic, Show)
  deriving anyclass NoThunks

makePrisms ''PsFileTag

type SharedStringMap = Map Int Text

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
  { _ps_file           :: PsFileTag       -- ^ File tag that we associate from `tagFilesC`
  , _ps_row            :: ~CellRow        -- ^ Current row
  , _ps_sheet_name     :: Text            -- ^ Current sheet name
  , _ps_cell_row_index :: Int             -- ^ Current row number
  , _ps_cell_col_index :: Int             -- ^ Current column number
  , _ps_is_in_val      :: Bool            -- ^ Flag for indexing wheter the parser is in value or not
  , _ps_shared_strings :: SharedStringMap -- ^ Shared string map
  , _ps_type           :: ExcelValueType  -- ^ The last detected value type
  } deriving stock (Generic, Show)
makeLenses 'MkSheetState

deriving via AllowThunksIn
  '[ "_ps_row"
   ] SheetState
  instance NoThunks SheetState

-- | State for parsing shared strings
data SharedStringState = MkSharedStringState
  { _ss_file      :: PsFileTag -- ^ File tag of the current file we are in
  , _ss_shared_ix :: Int       -- ^ Id of a shared string, int is okay, because bigger int blows up memory anyway
  , _ss_string    :: Text      -- ^ String we are parsing
  } deriving stock (Generic, Show)
    deriving anyclass NoThunks
makeLenses 'MkSharedStringState

type HasSheetState = MonadState SheetState
type HasSharedStringState = MonadState SharedStringState

class HasPSFiles a where
  fileLens :: Lens' a PsFileTag

instance HasPSFiles SheetState where
  fileLens = ps_file
instance HasPSFiles SharedStringState where
  fileLens = ss_file

--
-- Here begins the types for the part of this module that uses the "zip" library parser ---
--
-- TODO: decide whether to remove one or the other approaches, or unify them, or separate modules
data XlsxMState = MkXlsxMState
  { _xs_shared_strings :: IORef (Maybe (Map Int Text))
--  , _xs_sheet_names :: IORef (Maybe (Map Int Text))
  }
makeLenses 'MkXlsxMState

newtype XlsxM a = XlsxM { unXlsxM :: ReaderT XlsxMState Zip.ZipArchive a }
  deriving newtype (Functor, Applicative, Monad, MonadIO, MonadCatch, MonadMask, MonadThrow, MonadReader XlsxMState)
--- Here ends the types for the zip library/XlsxM approach ---

-- | Initial parsing state
initialSheetState :: SheetState
initialSheetState = MkSheetState
  { _ps_file            = InitialNoFile
  , _ps_row             = mempty
  , _ps_sheet_name      = mempty
  , _ps_cell_row_index  = 0
  , _ps_cell_col_index  = 0
  , _ps_is_in_val       = False
  , _ps_shared_strings  = mempty
  , _ps_type            = Untyped
  }

-- | Initial parsing state
initialSharedString :: SharedStringState
initialSharedString = MkSharedStringState
  { _ss_file            = InitialNoFile
  , _ss_shared_ix       = 0
  , _ss_string          = mempty
  }

-- | Classify zip entries
getFileTag :: Text -> PsFileTag
getFileTag = \case
  "xl/sharedStrings.xml" -> SharedStrings
  "xl/styles.xml"        -> Styles
  "xl/workbook.xml"      -> Workbook
  "xl/calcChain.xml"     -> FormulaChain
  "[Content_Types].xml"  -> ContentTypes
  "_rels/.rels"          -> Relationships
  unknown
    | ws `Text.isPrefixOf` unknown ->
      let unknown' = Text.drop (Text.length ws) unknown
      in if
      | rel `Text.isPrefixOf` unknown' ->
        SheetRel $ Text.drop (Text.length rel) unknown'
      | otherwise -> Sheet unknown'
    | otherwise -> Ignored unknown
  where
    ws = "xl/worksheets"
    rel = "_rels/"

-- | Conduit for parsing shared strings
--
-- This is essentially parsing a zip stream with each entry stored as bytestring
-- that we tag before proceeding with parsing the elements.
parseSharedStringsC
  :: ( MonadThrow m
     , PrimMonad m
     )
  => ConduitT BS.ByteString (Int, Text) m ()
parseSharedStringsC = (() <$ unZipStream) .| C.evalStateLC initialSharedString go
  where
    go = (await >>= tagFilesC)
         .| C.filterM (const $ has _SharedStrings <$> use fileLens)
         .| parseBytes def
         .| C.concatMapM parseSharedString

-- | Conduit that returns shared strings map
--
-- This is used mostly for testing.
parseSharedStringsIntoMapC
  :: ( MonadThrow m
     , PrimMonad m
     )
  => ConduitT BS.ByteString C.Void m (Map Int Text)
parseSharedStringsIntoMapC = parseSharedStringsC .| C.foldMap (uncurry Map.singleton)

-- | Parse shared string entry from xml event and return it once
-- we've reached the end of given element
parseSharedString
  :: ( MonadThrow m
     , HasSharedStringState m
     )
  => Event -> m (Maybe (Int, Text))
parseSharedString = \case
  EventBeginElement Name {nameLocalName = "t"} _ -> Nothing <$ (ss_string .= "")
  EventEndElement Name {nameLocalName = "t"} -> do
    !idx <- use ss_shared_ix
    ss_shared_ix += 1
    !txt <- use ss_string
    pure $ Just (idx, txt)
  EventContent (ContentText txt) -> Nothing <$ (ss_string <>= txt)
  _ -> pure Nothing

-- | Parse sheet with shared strings map provided
--
-- This is essentially parsing of a zip stream in which we ignore all unrecognised
-- files, literally just parsing @worksheets/*.xml@ here.
readXlsxWithSharedStringMapC
  :: ( MonadThrow m
     , PrimMonad m
     )
  => SharedStringMap
  -> ConduitT BS.ByteString SheetItem m ()
readXlsxWithSharedStringMapC sstate = (() <$ unZipStream) .| C.evalStateLC initial parseFileCState
    where
      parseFileCState = (await >>= tagFilesC)
        .| C.filterM (const $ not . has _Ignored <$> use fileLens)
        .| parseBytes def
        .| parseFileC
      initial = set ps_shared_strings sstate initialSheetState

-- | Parse xlsx file from a bytestring
--
-- This first reads the shared string table, then it provides another conduit
-- for parsing the rest of the data. Reading happens twice. All shared strings
-- are stored into memory wrapped in state monad.
readXlsxC
  :: ( MonadThrow m
     , PrimMonad m
     )
  => ConduitT () BS.ByteString m ()
  -> m (ConduitT () SheetItem m ())
readXlsxC input = do
    ssState <- C.runConduit $
         input
      .| parseSharedStringsC
      .| C.foldMap (uncurry Map.singleton)
    pure $
         input
      .| readXlsxWithSharedStringMapC ssState

-- | Run a series of actions on an Xlsx file
runXlsxM :: MonadIO m => FilePath -> XlsxM a -> m a
runXlsxM xlsxFile (XlsxM act) = do
  env0 <- liftIO $ MkXlsxMState <$> newIORef Nothing
  Zip.withArchive xlsxFile $ runReaderT act env0

liftZip :: Zip.ZipArchive a -> XlsxM a
liftZip = XlsxM . ReaderT . const

getOrParseSharedStrings :: XlsxM (Map Int Text)
getOrParseSharedStrings = do
  sharedStringsRef <- asks _xs_shared_strings
  mSharedStrings <- liftIO $ readIORef sharedStringsRef
  case mSharedStrings of
    Just strs -> pure strs
    Nothing -> do
      sharedStrsSel <- liftZip $ Zip.mkEntrySelector "xl/sharedStrings.xml"
      sharedStrings <- liftZip $ Zip.sourceEntry sharedStrsSel $
        parseBytes def
        .| C.evalStateC
             (initialSharedString & ss_file .~ SharedStrings)
             (C.concatMapM parseSharedString)
        .| C.foldMap (uncurry Map.singleton)
      liftIO $ writeIORef sharedStringsRef $ Just sharedStrings
      pure sharedStrings

-- | If the given sheet number exists, returns Just a conduit source
-- of the stream of XML events (using the xml-conduit packgae) in a
-- particular sheet. Returns Nothing when the sheet doesn't exist.
--
-- This is a lower level API for full control over the XML in a given
-- sheet. For a higher level API, see 'getSheetSource'.
getSheetXmlSource ::
  (PrimMonad m, MonadIO m, MonadThrow m, C.MonadResource m) =>
  -- | The sheet number
  Int ->
  XlsxM (Maybe (ConduitT () Event m ()))
getSheetXmlSource sheetNumber = do
  -- TODO: consider re-wrapping zip library exceptions, or just re-export them?
  sheetSel <- liftZip $ Zip.mkEntrySelector $ "xl/worksheets/sheet" <> show sheetNumber <> ".xml"
  sheetExists <- liftZip $ Zip.doesEntryExist sheetSel
  if not sheetExists
    then pure Nothing
    else do
      sourceSheet <- liftZip $ Zip.getEntrySource sheetSel
      pure $ Just $ sourceSheet .| parseBytes def
 where
   sheetName = mkSheetName sheetNumber

-- FIXME: hack to be compatible with the other parser,
-- which as of this writing sets the sheet name to be
-- "/sheetN.xml"
mkSheetName :: Int -> Text
mkSheetName sheetNumber =
     "/sheet"  <> Text.pack (show sheetNumber) <> ".xml"

-- | If the given sheet number exists, returns Just a conduit source
-- of the stream of XML events (using the xml-conduit packgae) in a
-- particular sheet. Returns Nothing when the sheet doesn't exist.
--
-- May throw SheetErrors if there is a parsing error occurs in the
-- underlying XML parser.
getSheetSource ::
  (PrimMonad m, MonadIO m, MonadThrow m, C.MonadResource m) =>
  -- | The sheet number
  Int ->
  XlsxM (Maybe (ConduitT () SheetItem m ()))
getSheetSource sheetNumber = do
  getSheetXmlSource sheetNumber >>= \case
    Nothing -> pure Nothing
    Just sourceSheetXml -> do
      sharedStrs <- getOrParseSharedStrings
      let
        sheetName = mkSheetName sheetNumber
        sheetState0 = initialSheetState
          & ps_shared_strings .~ sharedStrs
          & ps_file .~ Sheet sheetName
          & ps_sheet_name .~ sheetName
        xmlEventToSheetItem event = do
          -- A version of readSheetC that doesn't include logic for
          -- streaming multiple sheets.
          -- XXX: unify the two approaches, or support both?
          parseRes <- runExceptT $ matchEvent sheetName event
          rix' <- use ps_cell_row_index
          case parseRes of
            Left err -> C.throwM err
            Right (Just cellRow)
              | not (Map.null cellRow) ->
                pure $ Just $ MkSheetItem sheetName rix' cellRow
            _ -> pure Nothing
      pure $ Just $ sourceSheetXml
        .| C.evalStateC sheetState0 (CL.mapMaybeM xmlEventToSheetItem)

-- | Tag zip entries with `PsFileTags`
--
-- Essentiall @.xlsx@ file is a zip archive with bunch of @.xml@s stored
-- inside. This conduit ensures that we recognise all the needed files to
-- work with later. All of the tags are pushed in a state monad.
tagFilesC
  :: ( HasPSFiles env
     , MonadState env m
     , MonadThrow m
     )
  => Maybe (Either ZipEntry BS.ByteString)
  -> ConduitT (Either ZipEntry BS.ByteString) BS.ByteString m ()
tagFilesC = \case
  Just (Left zipEntry) -> do
   let !filePath = either id Text.decodeUtf8 (zipEntryName zipEntry)
   fileLens .= getFileTag filePath
   await >>= tagFilesC
  Just (Right fdata) -> do
    -- strange behaviour indeed, why it does not await already yielded data?
    yield fdata
    await >>= tagFilesC
  Nothing -> pure ()

-- | Parse a xml event
parseFileC
  :: ( MonadThrow m
     , HasSheetState m
     )
  => ConduitT Event SheetItem m ()
parseFileC = await >>= parseEventLoop
  where
    parseEventLoop
      :: ( MonadThrow m
         , HasSheetState m
         )
      => Maybe Event
      -> ConduitT Event SheetItem m ()
    parseEventLoop = \case
      Nothing -> pure ()
      Just evt -> use fileLens >>= \case
        Sheet name -> do
          prevSheetName <- use ps_sheet_name
          rix <- use ps_cell_row_index
          when (prevSheetName /= name) $
            -- TODO: this is the last part of the file name containing
            -- a particular sheet, not the actual sheet name as it
            -- appears in the Excel UI.
            ps_sheet_name .= name
          parseSheetC evt name
        _          -> await >>= parseEventLoop

-- | Parse sheet
--
-- We extract previous sheet name from the state and as well
-- as a current row index. If we're still in the same sheet
-- then we're returning currently stored row.
-- If we're on a different sheet, then parse the event and
-- return newly generated sheet item.
parseSheetC
  :: ( MonadThrow m
     , HasSheetState m
     )
  => Event
  -> Text
  -> ConduitT Event SheetItem m ()
parseSheetC event name = do
  prevSheetName <- use ps_sheet_name
  rix <- use ps_cell_row_index
  -- only consider yielding the sheet item if this is a _new_ sheet name.
  -- The idea is to here yield the entire excel row in one chunk. i.e. we've started a previous
  unless (prevSheetName == name) $ do
    -- Take the ps_row from state, set the ps_row to null, yield if not empty.
    -- The point is to 'get rid' of the final sheet item in the previous worksheet.
    popRow >>= yieldSheetItem name rix

  -- Now continue, either on the current mid-way row or from the beginning after yielding the above.
  -- match the event
  parseRes <- runExceptT $ matchEvent name event
  rix' <- use ps_cell_row_index
  case parseRes of
    Left err -> C.throwM err
    Right mResult -> do
      -- Yield the result if Just
      traverse_ (yieldSheetItem name rix') mResult
      parseFileC

-- | @SheetItem@ constructor
yieldSheetItem
  :: MonadThrow m
  => Text
  -> Int
  -> CellRow
  -> ConduitT Event SheetItem m ()
yieldSheetItem name rix' row =
  unless (row == mempty) $
    yield (MkSheetItem name rix' row)

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
      (Map Int Text)     -- ^ Given map to lookup in
  deriving Show

-- | Parse the given value
--
-- If it's a string, we try to get it our of a shared string table
parseValue :: SharedStringMap -> Text -> ExcelValueType -> Either AddCellErrors CellValue
parseValue sstrings txt = \case
  TS -> do
    (idx, _) <- ReadError txt `first` Read.decimal @Int txt
    string <- maybe (Left $ SharedStringNotFound idx sstrings) Right $ sstrings ^? ix idx
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
addCellToRow
  :: ( MonadError SheetErrors m
     , HasSheetState m
     )
  => Text -> m ()
addCellToRow txt = do
  inVal <- use ps_is_in_val
  when inVal $ do
    type' <- use ps_type
    sstrings <- use ps_shared_strings
    val <- liftEither $ first ParseCellError $ parseValue sstrings txt type'

    col <- use ps_cell_col_index
    ps_row <>= (Map.singleton col $ Cell
      { _cellStyle   = Nothing
      , _cellValue   = Just val
      , _cellComment = Nothing
      , _cellFormula = Nothing
      })

data SheetErrors
  = ParseCoordinateError CoordinateErrors -- ^ Error while parsing coordinates
  | ParseTypeError TypeError              -- ^ Error while parsing types
  | ParseCellError AddCellErrors          -- ^ Error while parsing cells
  deriving stock Show
  deriving anyclass Exception

type SheetValue = (Name, [Content])
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

contentTextPrims :: Prism' Content Text
contentTextPrims = prism' ContentText (\case ContentText x -> Just x
                                             _             -> Nothing)

-- | Update state coordinates accordingly to @parseCoordinates@
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
findName :: Text -> SheetValues -> Maybe SheetValue
findName name = find ((name ==) . nameLocalName . fst)

-- | Parse value type
parseType :: SheetValues -> Either TypeError ExcelValueType
parseType list =
  case findName "t" list of
    Nothing -> pure Untyped
    Just nameValPair -> do
      valContent <- maybe (Left $ TypeNoListElement nameValPair list) Right $  nameValPair ^? _2 . ix 0
      valText    <- maybe (Left $ TypeNoTextContent valContent list)  Right $ valContent ^? contentTextPrims
      case valText of
        "n"   -> Right TN
        "s"   -> Right TS
        "str" -> Right TStr
        "b"   -> Right TB
        "e"   -> Right TE
        other -> Left $ UnkownType other list

-- | Parse coordinates from a list of xml elements if such were found on "r" key
parseCoordinates :: SheetValues -> Either CoordinateErrors (Int, Int)
parseCoordinates list = do
      nameValPair <- maybe (Left $ CoordinateNotFound list)        Right $ findName "r" list
      valContent  <- maybe (Left $ NoListElement nameValPair list) Right $ nameValPair ^? _2 . ix 0
      valText     <- maybe (Left $ NoTextContent valContent list)  Right $ valContent ^? contentTextPrims
      maybe (Left $ DecodeFailure valText list) Right $ fromSingleCellRef $ CellRef valText

-- | Update state accordingly the xml event given, basically parse a xml and return row
-- if we've ended parsing.
-- (I wish xml docs were more elaborate on their structures)
matchEvent
  :: ( MonadError SheetErrors m
     , HasSheetState m
     )
  => Text -> Event -> m (Maybe CellRow)
matchEvent _currentSheet = \case
  EventContent (ContentText txt)                    -> Nothing <$ addCellToRow txt
  EventBeginElement Name{nameLocalName = "c"} vals  -> Nothing <$ (setCoord vals >> setType vals)
  EventBeginElement Name {nameLocalName = "v"} _    -> Nothing <$ (ps_is_in_val .= True)
  EventEndElement Name {nameLocalName = "v"}        -> Nothing <$ (ps_is_in_val .= False)
  -- If beginning of row, empty the state and return nothing [why exactly?]
  EventBeginElement Name {nameLocalName = "row"} _  -> Nothing <$ popRow
  -- If at the end of the row, we have collected the whole row into
  -- the current state. Empty the state and return the row.
  EventEndElement Name{nameLocalName = "row"}       -> Just <$> popRow
  -- Skip everything else, e.g. the formula elements <f>
  _                                                 -> pure Nothing
