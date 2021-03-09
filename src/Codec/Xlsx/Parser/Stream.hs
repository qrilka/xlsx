{-# LANGUAGE FlexibleContexts    #-}
{-# LANGUAGE LambdaCase          #-}
{-# LANGUAGE NamedFieldPuns      #-}
{-# LANGUAGE OverloadedStrings   #-}
{-# LANGUAGE RankNTypes          #-}
{-# LANGUAGE RecordWildCards     #-}
{-# LANGUAGE ScopedTypeVariables #-}
{-# LANGUAGE TemplateHaskell     #-}
{-# LANGUAGE DeriveAnyClass #-}

-- | Read .xlsx as a stream
module Codec.Xlsx.Parser.Stream
  ( readXlsx
  , readXlsxWithState
  , SheetItem(..)
  , parseSharedStrings
  , parseStringsIntoMap
  , CellRow
  , si_sheet
  , si_row_index
  , si_cell_row
  ) where

import           Codec.Archive.Zip.Conduit.UnZip
import           Codec.Xlsx.Types.Cell
import           Codec.Xlsx.Types.Common
import           Conduit                         (PrimMonad, await,
                                                  yield, (.|))
import qualified Conduit                         as C
import           Control.Lens
import           Control.Monad.Except
import           Control.Monad.State.Lazy
import           Data.Bifunctor
import qualified Data.ByteString                 as BS
import           Data.Conduit                    (ConduitT)
import qualified Data.Conduit.Combinators        as C
import           Data.Foldable
import           Data.Map                        (Map)
import qualified Data.Map                        as Map
import           Data.Text                       (Text)
import qualified Data.Text                       as Text
import qualified Data.Text.Encoding              as Text
import           Data.XML.Types
import           Text.Read
import           Text.XML.Stream.Parse
import Control.Monad.Catch

-- TODO NonNull
type CellRow = Map Int Cell

data SheetItem = MkSheetItem
  { _si_sheet     :: Text
  , _si_row_index :: Int
  , _si_cell_row  :: CellRow
  } deriving Show
makeLenses 'MkSheetItem

-- http://officeopenxml.com/anatomyofOOXML-xlsx.php
data PsFiles =
             -- | Files that are encountered in the workbook but not part
             --   of this sumtype get this tag. They are then ignored.
              UnkownFile Text
             -- | files with actual data (rows and cells)
             | Sheet Text
             -- | Initial state, unrelated to excell format
             | InitialNoFile
              -- | stores strings which are replaced by id's in sheet to save space
             | SharedStrings
             | Styles
              -- | contains the names in the tabs and some internal id
             | Workbook
             | ContentTypes
             | Relationships
              -- | contains relations ship between internal id and path
             | SheetRel Text
  deriving Show

makePrisms ''PsFiles

decodeFiles :: Text -> PsFiles
decodeFiles = \case
  "xl/sharedStrings.xml" -> SharedStrings
  "xl/styles.xml"        -> Styles
  "xl/workbook.xml"      -> Workbook
  "[Content_Types].xml"  -> ContentTypes
  "_rels/.rels"          -> Relationships
  unkown                 -> let
      ws = "xl/worksheets/"
      wsL = Text.length ws
    in
    if Text.take wsL unkown == ws then
      let
        known = Text.drop wsL unkown
        rel = "_rels/"
        relL = Text.length rel
      in
        if Text.take relL known == rel then
          SheetRel $ Text.drop relL known
        else
          Sheet known
      else UnkownFile unkown

data TVal = T_S -- ^ string
          | T_N -- ^ number
          | T_B -- ^ boolean


data SheetState = MkSheetState
  { _ps_file           :: PsFiles
  , _ps_row            :: CellRow
  , _ps_sheet_name     :: Text
  , _ps_cell_row_index :: Int
  , _ps_cell_col_index :: Int
  , _ps_is_in_val      :: Bool
  , _ps_shared_strings :: Map Int Text
  , _ps_t              :: TVal -- the last detected value
  }
makeLenses 'MkSheetState

data SharedStringState = MkSharedStringState
  { _ss_file           :: PsFiles
  , _ss_shared_ix      :: Int -- int is okay, because bigger then int blows up memory anyway
  } deriving Show
makeLenses 'MkSharedStringState

class HasPSFiles a where
  fileLens :: Lens' a PsFiles

instance HasPSFiles SheetState where
  fileLens = ps_file
instance HasPSFiles SharedStringState where
  fileLens = ss_file

initialSheetState :: SheetState
initialSheetState = MkSheetState
  { _ps_file            = InitialNoFile
  , _ps_row             = mempty
  , _ps_sheet_name      = mempty
  , _ps_cell_row_index  = 0
  , _ps_cell_col_index  = 0
  , _ps_is_in_val       = False
  , _ps_shared_strings  = mempty
  , _ps_t               = T_S
  }

initialSharedString :: SharedStringState
initialSharedString = MkSharedStringState
  { _ss_file            = InitialNoFile
  , _ss_shared_ix       = 0
  }

parseSharedStrings :: forall m . MonadThrow m
  => PrimMonad m
  => ConduitT BS.ByteString (Int, Text) m ()
parseSharedStrings = (() <$ unZipStream)
          .| (C.evalStateLC initialSharedString $
             (await >>= tagFiles)
              .| C.filterM (const $ has _SharedStrings <$> use fileLens)
              .| parseBytes def
              .| C.concatMapM parseString
             )

parseStringsIntoMap :: forall m . MonadThrow m
  => PrimMonad m
  => ConduitT BS.ByteString C.Void m (Map Int Text)
parseStringsIntoMap = parseSharedStrings .| C.foldMap (uncurry Map.singleton)

parseString :: MonadThrow m
  => MonadState SharedStringState m
  => Event -> m (Maybe (Int, Text))
parseString = \case
  EventContent (ContentText txt) -> do
    idx <- use ss_shared_ix
    ss_shared_ix += 1
    pure $ Just (idx, txt)
  _ -> pure Nothing


-- TODO figure out how to allow user to lookup shared string instead
-- of always reading into memory.
readXlsxWithState :: forall m . MonadThrow m
  => PrimMonad m
  => Map Int Text -> ConduitT BS.ByteString SheetItem m ()
readXlsxWithState sstate =
      (() <$ unZipStream)
      .| (C.evalStateLC initial $ (await >>= tagFiles)
      .| C.filterM (const $ not . has _UnkownFile <$> use fileLens)
      .| parseBytes def
      .| parseFiles)
    where
      initial = set ps_shared_strings sstate initialSheetState


-- | first reads the share string table, then provides another conuit to be run again for the remaining data.
--   reading happens twice. All shared strings will be read into memory.
readXlsx ::
  forall m . MonadThrow m
  => PrimMonad m
  => ConduitT () BS.ByteString m () -> m (ConduitT () SheetItem m ())
readXlsx input = do
    ssState <- C.runConduit $ input .| parseSharedStrings .| C.foldMap (uncurry Map.singleton)
    pure $ input .| readXlsxWithState ssState

-- | there are various files in the excell file, which is a glorified zip folder
-- here we tag them with things we know, and push it into the state monad.
-- we need a state monad to make the excell parsing conduit to function
tagFiles ::
  HasPSFiles env
  => MonadState env m
  => MonadThrow m
  => Maybe (Either ZipEntry BS.ByteString) -> ConduitT (Either ZipEntry BS.ByteString) BS.ByteString m ()
tagFiles = \case
  Just (Left zipEntry) -> do
   let filePath = either id Text.decodeUtf8 (zipEntryName zipEntry)
   fileLens .= decodeFiles filePath
   await >>= tagFiles
  Just (Right fdata) -> do
   yield fdata
   await >>= tagFiles
  Nothing -> pure ()

parseFiles ::
   MonadThrow m
  => MonadState SheetState m
  => ConduitT Event SheetItem  m ()
parseFiles = await >>= parseFileLoop

-- we significantly
parseFileLoop ::
   MonadThrow m
  => MonadState SheetState m
  => Maybe Event
  -> ConduitT Event SheetItem  m ()
parseFileLoop = \case
  Nothing -> pure ()
  Just evt -> do
    file <- use ps_file
    case file of
      Sheet name -> parseSheet evt name
      _          -> await >>= parseFileLoop

parseSheet ::
   MonadThrow m
  => MonadState SheetState m
  => Event -> Text -> ConduitT Event SheetItem m ()
parseSheet evt name = do
        prevSheetName <- use ps_sheet_name
        rix <- use ps_cell_row_index
        unless (prevSheetName == name) $ do
          ps_sheet_name .= name
          popRow >>= yieldSheetItem name rix

        parseRes <- runExceptT $ matchEvent name evt
        rix' <- use ps_cell_row_index
        case parseRes of
          Left err -> C.throwM err
          Right mResult -> do
            traverse_ (yieldSheetItem name rix') mResult
            await >>= parseFileLoop

yieldSheetItem :: MonadThrow m
  => Text -> Int -> CellRow ->  ConduitT Event SheetItem m ()
yieldSheetItem name rix' row =
  unless (row == mempty) $ yield $ MkSheetItem name rix' row

popRow :: MonadState SheetState m => m CellRow
popRow = do
  row <- use ps_row
  ps_row .= mempty
  pure row

data AddCellErrors = ReadError String
                   | SharedStringNotFound Int (Map Int Text)
                   deriving Show

parseValue :: Map Int Text -> Text -> TVal -> Either AddCellErrors CellValue
parseValue sstrings txt = \case
  T_S -> do
    (idx :: Int) <- first ReadError $ readEither $ Text.unpack txt
    string <- maybe (Left $ SharedStringNotFound idx sstrings) Right $ sstrings ^? ix idx
    Right $ CellText string
  T_N -> bimap ReadError CellDouble $ readEither $ Text.unpack txt
  T_B -> bimap ReadError CellBool $ readEither $ Text.unpack txt

addCell :: MonadError SheetErrors m => MonadState SheetState m => Text -> m ()
addCell txt = do
   inVal <- use ps_is_in_val
   when inVal $ do
      type' <- use ps_t
      sstrings <- use ps_shared_strings
      val <- liftEither $ first MkCell $ parseValue sstrings txt type'

      col <- use ps_cell_col_index
      ps_row <>= (Map.singleton col $ Cell
        { _cellStyle   = Nothing
        , _cellValue   = Just val
        , _cellComment = Nothing
        , _cellFormula = Nothing
        })


data SheetErrors = MkCoordinate CoordinateErrors
                 | MkType TypeError
                 | MkCell AddCellErrors

  deriving (Show, Exception)


data CoordinateErrors = CoordinateNotFound [(Name, [Content])]
                      | NoListElement (Name, [Content]) [(Name, [Content])]
                      | NoTextContent Content [(Name, [Content])]
                      | DecodeFailure Text [(Name, [Content])]
  deriving Show

data TypeError = TypeNotFound [(Name, [Content])]
              | TypeNoListElement (Name, [Content]) [(Name, [Content])]
              | UnkownType Text [(Name, [Content])]
              | TypeNoTextContent Content [(Name, [Content])]
              deriving Show

contentTextPrims :: Prism' Content Text
contentTextPrims = prism' ContentText (\case ContentText x -> Just x
                                             _ -> Nothing)

setCoord :: MonadError SheetErrors m => MonadState SheetState m => [(Name, [Content])] -> m ()
setCoord list = do
  coordinates <- liftEither $ first MkCoordinate $ parseCoordinates list
  ps_cell_col_index .= (coordinates ^. _2)
  ps_cell_row_index .= (coordinates ^. _1)

setType :: MonadError SheetErrors m => MonadState SheetState m => [(Name, [Content])] -> m ()
setType list = do
  type' <- liftEither $ first MkType $ parseType list
  ps_t .= type'

findName :: Text -> [(Name, [Content])] -> Maybe (Name, [Content])
findName name = find ((name ==) . nameLocalName . fst)

parseType :: [(Name, [Content])] -> Either TypeError TVal
parseType list = do
    nameValPair <- maybe (Left (TypeNotFound list)) Right $ findName "t" list
    valContent <- maybe (Left $ TypeNoListElement nameValPair list) Right $  nameValPair ^? _2 . ix 0
    valText <- maybe (Left $ TypeNoTextContent valContent list) Right $ valContent ^? contentTextPrims
    case valText of
      "n"   -> Right T_N
      "s"   -> Right T_S
      "b"   -> Right T_S
      other -> Left $ UnkownType other list

parseCoordinates :: [(Name, [Content])] -> Either CoordinateErrors (Int, Int)
parseCoordinates list = do
      nameValPair <- maybe (Left $ CoordinateNotFound list) Right $ findName "r" list
      valContent <- maybe (Left $ NoListElement nameValPair list) Right $  nameValPair ^? _2 . ix 0
      valText <- maybe (Left $ NoTextContent valContent list) Right $ valContent ^? contentTextPrims
      maybe (Left $ DecodeFailure valText list) Right $ fromSingleCellRef $ CellRef valText

matchEvent :: MonadError SheetErrors m => MonadState SheetState m => Text -> Event -> m (Maybe CellRow)
matchEvent _currentSheet = \case
  EventContent (ContentText txt)                       -> Nothing <$ addCell txt
  EventBeginElement Name{nameLocalName = "c", ..} vals  -> Nothing <$ (setCoord vals >> setType vals)
  EventBeginElement Name {nameLocalName = "v"} _       -> Nothing <$ (ps_is_in_val .= True)
  EventEndElement Name {nameLocalName = "v"}           -> Nothing <$ (ps_is_in_val .= False)
  -- EventBeginElement Name {nameLocalName = "c"} _    -> Nothing <$ ps_is_in_cell .= True
  -- EventEndElement Name {nameLocalName = "c"}        -> Nothing <$ ps_is_in_cell .= False
  EventBeginElement Name {nameLocalName = "row"} _     -> Nothing <$ popRow
  EventEndElement Name{nameLocalName = "row", ..}       -> Just <$> popRow
  _ -> pure Nothing
