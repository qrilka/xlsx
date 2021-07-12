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
{-# OPTIONS_GHC -fno-full-laziness #-}

-- This module must be compiled with -fno-full-laziness, otherwise
-- over-aggressive caching by GHC during -O2 builds mean that the
-- first call to `getEntrySource <sheetNum>` will be cached, and all
-- subsequent calls (_even_ those with a different sheet number) will
-- return the original sheet source.

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
  ( -- * Parser API based on "zip-stream" package
  -- * Parser API based on "zip" package
  XlsxM
  , runXlsxM
  , getSheetSource
  , sourceSheet
  , getOrParseSharedStrings
  , countRowsInSheet

  -- * Types shared by both parser approaches
  , CellRow
  , SheetItem(..)
  -- ** `SheetItem` lens
  , si_sheet
  , si_row_index
  , si_cell_row

  -- * Errors that either parser can throw
  , SheetErrors(..)
  {-- TODO: check if these errors could be exposed to users or not
  , AddCellErrors(..)
  , CoordinateErrors(..)
  , TypeError(..)
  --}

  -- * Temporary:
  , XmlLib(..)
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
import           Data.Time.Clock
import           Data.Traversable                (for)
import           Data.XML.Types
import           GHC.Generics
import           NoThunks.Class
import           Text.XML.Stream.Parse
import Codec.Xlsx.Parser.Internal

import Text.XML.Expat.SAX as Hexpat
import Text.XML.Expat.Internal.IO as Hexpat
import qualified Text.XML.LibXML.SAX as LX
import Control.Monad.Base
import Control.Monad.Trans.Control
import qualified Codec.Xlsx.Parser.Stream.HexpatInternal as HexpatInternal

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
  { _ss_string    :: Text      -- ^ String we are parsing
  } deriving stock (Generic, Show)
    deriving anyclass NoThunks
makeLenses 'MkSharedStringState

type HasSheetState = MonadState SheetState
type HasSharedStringState = MonadState SharedStringState

data XlsxMState = MkXlsxMState
  { _xs_shared_strings :: IORef (Maybe (V.Vector Text))
--  , _xs_sheet_names :: IORef (Maybe (Map Int Text))
  }
makeLenses 'MkXlsxMState

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
  , _ps_text_buf = ""
  , _ps_worksheet_ended = False
  }

-- | Initial parsing state
initialSharedString :: SharedStringState
initialSharedString = MkSharedStringState
  { _ss_string = mempty
  }

-- | Parse shared string entry from xml event and return it once
-- we've reached the end of given element
{-# SCC parseSharedString #-}
parseSharedString
  :: ( MonadThrow m
     , HasSharedStringState m
     )
  => Event -> m (Maybe Text)
parseSharedString = \case
  EventBeginElement Name {nameLocalName = "t"} _ -> Nothing <$ (ss_string .= "")
  EventEndElement Name {nameLocalName = "t"} -> do
    !txt <- use ss_string
    pure $ Just txt
  EventContent (ContentText txt) -> Nothing <$ (ss_string <>= txt)
  _ -> pure Nothing

-- | Run a series of actions on an Xlsx file
runXlsxM :: MonadIO m => FilePath -> XlsxM a -> m a
runXlsxM xlsxFile (XlsxM act) = do
  env0 <- liftIO $ MkXlsxMState <$> newIORef Nothing
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
      t0 <- liftIO getCurrentTime
      sharedStrings <- {-# SCC "eval_sharedStrings" #-} liftZip $ Zip.sourceEntry sharedStrsSel $
        ( {-# SCC "sharedStrings_parseBytes" #-} parseBytes def)
        .| C.evalStateC state0 (C.concatMapM parseSharedString)
        -- C.sinkVector uses a mutable vector internally and runs an
        -- 'unsafeFreeze' at the end. The vector's size is doubled
        -- whenever capacity is reached via Vector.Mutable.grow (which
        -- copies rather than extends, copying pointers rather than
        -- the underlying text values)
        .| C.sinkVector
      t1 <- liftIO getCurrentTime
      liftIO $ putStrLn $ "Took " <> show (t1 `diffUTCTime` t0) <> " to read shared strings via xml-conduit"
      liftIO $ writeIORef sharedStringsRef $ Just sharedStrings
      pure sharedStrings

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

{-# SCC runLibxml #-}
runLibxml ::
--  (PrimMonad m, MonadIO m, MonadThrow m, C.MonadResource m) =>
  SheetState ->
  ConduitT () ByteString (C.ResourceT IO) () ->
  (SheetItem -> IO ()) ->
  XlsxM ()
runLibxml initialState byteSource inner = liftIO $ do
  -- Set up state
  ref <- newIORef initialState
  -- Set up parser and callbacks
  p <- LX.newParserIO Nothing
  LX.setCallback p LX.parsedBeginElement (\el -> onEvent ref . EventBeginElement el)
  LX.setCallback p LX.parsedEndElement (onEvent ref . EventEndElement)
  LX.setCallback p LX.parsedCharacters (onEvent ref . EventContent . ContentText)
  -- Set up conduit
  C.runConduitRes $
    byteSource .|
    C.runReaderC ref (C.awaitForever $ liftIO . LX.parseBytes p)
  LX.parseComplete p
  where
    {-# SCC onEvent #-}
    {-# INLINE onEvent #-}
    onEvent stateRef ev = do
      state0 <- liftIO $ readIORef stateRef
      (parseRes, state1) <-
        {-# SCC "libxmlC_runStateT_call" #-} (`runStateT` state0) $ runExceptT $ matchEvent ev
      writeIORef stateRef state1
      case parseRes of
        Left err -> C.throwM err
        Right (Just cellRow)
          | not (Map.null cellRow) ->
            inner $ MkSheetItem (_ps_sheet_name state1) (_ps_cell_row_index state1) cellRow
        _ -> pure ()
      pure True

{-# SCC runExpat #-}
runExpat ::
  SheetState ->
  ConduitT () ByteString (C.ResourceT IO) () ->
  (SheetItem -> IO ()) ->
  XlsxM ()
runExpat initialState byteSource inner = liftIO $ do
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
          Nothing -> onEvents ref $ map fst saxen
  C.runConduitRes $
    byteSource .|
    C.awaitForever (liftIO . processChunk False)
  processChunk True BS.empty
  where
    sheetName = _ps_sheet_name initialState
    {-# SCC onEvents #-}
    {-# INLINE onEvents #-}
    onEvents stateRef evs = do
      state0 <- liftIO $ readIORef stateRef
      state1 <-
        {-# SCC "runExpat_runStateT_call" #-} (`execStateT` state0) $ do
          forM_ evs $ \ev -> do
            parseRes <- runExceptT $ matchHexpatEvent ev
            case parseRes of
              Left err -> error $ "error after matchHexpatEvent: " <> show err
              Right (Just cellRow)
                | not (Map.null cellRow) -> do
                    rowNum <- use ps_cell_row_index
                    liftIO $ inner $ MkSheetItem sheetName rowNum cellRow
              _ -> pure ()
      writeIORef stateRef state1

-- FIXME: hack to be compatible with the zip-stream parser (no longer
-- present), which as of this writing sets the sheet name to be
-- "/sheetN.xml"
mkSheetName :: Int -> Text
mkSheetName sheetNumber =
     "/sheet"  <> Text.pack (show sheetNumber) <> ".xml"

data XmlLib = UseXmlConduit | UseHexpat | UseLibxml deriving (Show, Eq, Read)

{-# SCC sourceSheet #-}
sourceSheet ::
  -- | Which XML library to use
  XmlLib ->
  -- | Sheet number
  Int ->
  -- | Function to consume the sheet's rows
  (SheetItem -> IO ()) ->
  -- | Returns False if sheet doesn't exist, or True otherwise
  XlsxM Bool
sourceSheet xmlLib sheetNumber inner = do
  mSrc :: Maybe (ConduitT () ByteString (C.ResourceT IO) ()) <-
    getSheetXmlSource sheetNumber
  case mSrc of
    Nothing -> pure False
    Just sourceSheetXml -> do
      sharedStrs <- getOrParseSharedStrings
      let sheetState0 = initialSheetState
            & ps_shared_strings .~ sharedStrs
            & ps_sheet_name .~ sheetName
      case xmlLib of
        UseXmlConduit -> error "use getSheetSource"
        UseLibxml -> runLibxml sheetState0 sourceSheetXml inner
        UseHexpat -> runExpat sheetState0 sourceSheetXml inner
      pure True
  where
    sheetName = mkSheetName sheetNumber

-- | If the given sheet number exists, returns Just a conduit source
-- of the stream of 'SheetItem's in a particular sheet. Returns
-- Nothing when the sheet doesn't exist.
--
-- May throw SheetErrors if a parsing error occurs in the
-- underlying XML parser.
--
-- The xml-conduit package is used to parse the underlying XML. In the
-- authors' tests, consuming large excel sheets into 'SheetItem's is
-- 1.6x slower than the 'sourceSheet' IO-based callback API which uses
-- expat for xml parsing. Your mileage may vary.
{-# SCC getSheetSource  #-}
getSheetSource ::
  (PrimMonad m, MonadThrow m, C.MonadResource m) =>
  -- | The sheet number
  Int ->
  XlsxM (Maybe (ConduitT () SheetItem m ()))
getSheetSource sheetNumber = do
  mSrc <- getSheetXmlSource sheetNumber
  for mSrc $ \sourceSheetXml -> do
    sharedStrs <- getOrParseSharedStrings
    let sheetState0 = initialSheetState
          & ps_shared_strings .~ sharedStrs
          & ps_sheet_name .~ sheetName
    pure $ sourceSheetXml
      .| parseBytes def
      .| C.evalStateC sheetState0
         (CL.mapMaybeM ({-# SCC "xmlEventToSheetItemXmlConduit_scc" #-} xmlEventToSheetItem))
  where
    sheetName = mkSheetName sheetNumber
    {-# SCC xmlEventToSheetItem #-}
    xmlEventToSheetItem event = do
      -- XXX: unify the two approaches, or support both?
      parseRes <- {-# SCC "runExceptT_scc" #-} runExceptT $ matchEvent event
      rix' <- use ps_cell_row_index
      case parseRes of
        Left err -> C.throwM err
        Right (Just cellRow)
          | not (Map.null cellRow) ->
            pure $ Just $ MkSheetItem sheetName rix' cellRow
        _ -> pure Nothing

-- | Returns number of rows in the given sheet (identified by sheet
-- number), or Nothing if the sheet does not exist. Does not perform a
-- full parse of the XML, so it should be more efficient than counting
-- via 'getSheetSource'.
countRowsInSheet :: Int -> XmlLib -> XlsxM (Maybe Int)
countRowsInSheet sheetNumber lib = do
  case lib of
    UseHexpat -> do
      ref <- liftIO $ newIORef 0
      exists <- sourceSheet UseHexpat sheetNumber $ const $ modifyIORef' ref (+1)
      if exists then liftIO $ Just <$> readIORef ref else pure Nothing
    UseLibxml -> do
      ref <- liftIO $ newIORef 0
      exists <- sourceSheet UseLibxml sheetNumber $ const $ modifyIORef' ref (+1)
      if exists then liftIO $ Just <$> readIORef ref else pure Nothing
    UseXmlConduit -> do
      mSrc <- getSheetSource sheetNumber
      for mSrc $ \src -> liftIO $ C.runConduitRes $
        src .| C.length

-- | Return row from the state and empty it
{-# SCC popRow #-}
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
  | HexpatParseError Hexpat.XMLParseError
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
findName :: Text -> SheetValues -> Maybe SheetValue
findName name = find ((name ==) . nameLocalName . fst)

-- | Parse value type
{-# SCC parseType #-}
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
{-# SCC parseCoordinates #-}
parseCoordinates :: SheetValues -> Either CoordinateErrors (Int, Int)
parseCoordinates list = do
      nameValPair <- maybe (Left $ CoordinateNotFound list)        Right $ findName "r" list
      valContent  <- maybe (Left $ NoListElement nameValPair list) Right $ nameValPair ^? _2 . ix 0
      valText     <- maybe (Left $ NoTextContent valContent list)  Right $ valContent ^? contentTextPrims
      maybe (Left $ DecodeFailure valText list) Right $ fromSingleCellRef $ CellRef valText

-- | Update state accordingly the xml event given, basically parse a xml and return row
-- if we've ended parsing.
-- (I wish xml docs were more elaborate on their structures)
{-# SCC matchEvent #-}
matchEvent
  :: ( MonadError SheetErrors m
     , HasSheetState m
     )
  =>  Event -> m (Maybe CellRow)
matchEvent = \case
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


{-
======== hexpat
-}

{-# SCC matchHexpatEvent #-}
matchHexpatEvent ::
  ( MonadError SheetErrors m,
    HasSheetState m
  ) =>
  HexpatEvent ->
  m (Maybe CellRow)
matchHexpatEvent ev = case ev of
  CharacterData txt -> do
    inVal <- use ps_is_in_val
    when inVal $  do
      res <- use ps_text_buf
      ps_text_buf .= (res <> txt)
    pure Nothing
  StartElement "c" attrs -> Nothing <$ (setCoord' attrs >> setType' attrs)
  StartElement "v" _ -> Nothing <$ (ps_is_in_val .= True)
  EndElement "v" -> do
    txt <- use ps_text_buf
    addCellToRow txt
    ps_is_in_val .= False
    ps_text_buf .= ""
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

type SheetValue' = (ByteString, Text)
type SheetValues' = [SheetValue']

-- | Update state coordinates accordingly to @parseCoordinates@
{-# SCC setCoord' #-}
setCoord'
  :: ( MonadError SheetErrors m
     , HasSheetState m
     )
  => SheetValues' -> m ()
setCoord' list = do
  coordinates <- liftEither $ first ParseCoordinateError $ parseCoordinates' list
  ps_cell_col_index .= (coordinates ^. _2)
  ps_cell_row_index .= (coordinates ^. _1)

-- | Parse type from values and update state accordingly
setType'
  :: ( MonadError SheetErrors m
     , HasSheetState m
     )
  => SheetValues' -> m ()
setType' list = do
  type' <- liftEither $ first ParseTypeError $ parseType' list
  ps_type .= type'

-- | Find sheet value by its name
findName' :: ByteString -> SheetValues' -> Maybe SheetValue'
findName' name = find ((name ==) . fst)
{-# INLINE findName' #-}

-- | Parse value type
{-# SCC parseType' #-}
parseType' :: SheetValues' -> Either TypeError ExcelValueType
parseType' list =
  case findName' "t" list of
    Nothing -> pure Untyped
    Just (_nm, valText)->
      case valText of
        "n"   -> Right TN
        "s"   -> Right TS
        "str" -> Right TStr
        "b"   -> Right TB
        "e"   -> Right TE
        other -> Left $ UnkownType other $ asSheetValues list

asSheetValues :: [(ByteString, b)] -> [(Name, [Content])]
asSheetValues = map f
  where f (s, _) =
          (Name (Text.decodeUtf8 s)
           Nothing Nothing, [ContentText $ Text.decodeUtf8 s])

-- | Parse coordinates from a list of xml elements if such were found on "r" key
{-# SCC parseCoordinates' #-}
parseCoordinates' :: SheetValues' -> Either CoordinateErrors (Int, Int)
parseCoordinates' list = do
  (_nm, valText) <- maybe (Left $ CoordinateNotFound $ asSheetValues list) Right $ findName' "r" list
  maybe (Left $ DecodeFailure valText $ asSheetValues list) Right $ fromSingleCellRef $ CellRef valText
