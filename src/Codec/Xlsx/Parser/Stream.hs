{-# LANGUAGE LambdaCase #-}
{-# LANGUAGE FlexibleContexts #-}
{-# LANGUAGE TemplateHaskell #-}
{-# LANGUAGE OverloadedStrings #-}
{-# LANGUAGE RecordWildCards #-}
{-# LANGUAGE NamedFieldPuns #-}

-- | Read .xlsx as a stream
module Codec.Xlsx.Parser.Stream
  ( readXlsx
  , SheetItem(..)
  ) where

import qualified Data.Map as Map
import Data.Map (Map)
import Data.Bifunctor
import Codec.Xlsx.Types.Common
import Control.Monad.Except
import Data.Foldable
import Control.Monad.State.Lazy
import Data.Conduit(ConduitT)
import Conduit(PrimMonad, MonadThrow, yield, await, (.|))
import qualified Conduit as C
import qualified Data.Conduit.Combinators as C
import Codec.Archive.Zip.Conduit.UnZip
import Codec.Xlsx.Types.Cell
import qualified Data.ByteString as BS
import Text.XML.Stream.Parse
import Data.XML.Types
import Data.Text (Text)
import qualified Data.Text.Encoding as Text
import qualified Data.Text as Text
import Control.Lens

data SheetItem = MkSheetItem
  { _si_sheet     :: Text
  , _si_row_index :: Int
  , _si_cell_row  :: CellRow
  } deriving Show

-- http://officeopenxml.com/anatomyofOOXML-xlsx.php
data PsFiles = UnkownFile Text
             | Sheet Text
             | InitialNoFile
             | SharedStrings
             | Styles
             | Workbook
             | ContentTypes
             | Relationships
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

data PipeState = MkPipeState
  { _ps_file            :: PsFiles
  , _ps_row             :: CellRow
  , _ps_sheet_name      :: Text
  , _ps_cell_row_index  :: Int
  , _ps_cell_col_index  :: Int
  , _ps_is_in_val       :: Bool
  , _ps_shared_strings  :: Map Int Text
  , _ps_shared_ix       :: Int -- int is okay, because bigger then int blows up memory anyway
  }
makeLenses 'MkPipeState

initialState :: PipeState
initialState = MkPipeState
  { _ps_file            = InitialNoFile
  , _ps_row             = mempty
  , _ps_sheet_name      = mempty
  , _ps_cell_row_index  = 0
  , _ps_cell_col_index  = 0
  , _ps_is_in_val       = False
  , _ps_shared_strings  = mempty
  , _ps_shared_ix       = 0
  }

readXlsx :: MonadIO m => MonadThrow m
  => PrimMonad m
  => ConduitT BS.ByteString SheetItem m ()
readXlsx = (() <$ unZipStream)
    .| (C.evalStateLC initialState $ (await >>= tagFiles)
      .| C.filterM (const $ not . has _UnkownFile <$> use ps_file)
      .| parseBytes def
      .| parseFiles)

-- | there are various files in the excell file, which is a glorified zip folder
-- here we tag them with things we know, and push it into the state monad.
-- we need a state monad to make the excell parsing conduit to function
tagFiles ::
  MonadState PipeState m
  => MonadIO m
  => MonadThrow m
  => PrimMonad m
  => Maybe (Either ZipEntry BS.ByteString) -> ConduitT (Either ZipEntry BS.ByteString) BS.ByteString m ()
tagFiles = \case
  Just (Left zipEntry) -> do
   let filePath = either id Text.decodeUtf8 (zipEntryName zipEntry)
   ps_file .= decodeFiles filePath
   await >>= tagFiles
  Just (Right fdata) -> do
   yield fdata
   await >>= tagFiles
  Nothing -> pure ()

parseFiles ::
  MonadIO m
  => MonadThrow m
  => PrimMonad m
  => MonadState PipeState m
  => ConduitT Event SheetItem  m ()
parseFiles = await >>= parseFileLoop

-- we significantly
parseFileLoop ::
  MonadIO m
  => MonadThrow m
  => PrimMonad m
  => MonadState PipeState m
  => Maybe Event
  -> ConduitT Event SheetItem  m ()
parseFileLoop = \case
  Nothing -> pure ()
  Just evt -> do
    file <- use ps_file
    case file of
      Sheet name -> parseSheet evt name
      SharedStrings -> parseStrings evt
      _ -> await >>= parseFileLoop

parseStrings ::
  MonadIO m
  => MonadThrow m
  => PrimMonad m
  => MonadState PipeState m
  => Event -> ConduitT Event SheetItem m ()
parseStrings evt = do
    parseString evt
    await >>= parseFileLoop

parseString ::
  MonadIO m
  => MonadThrow m
  => PrimMonad m
  => MonadState PipeState m
  => Event -> m ()
parseString = \case
  EventContent (ContentText txt) -> do
    idx <- use ps_shared_ix
    ps_shared_strings %= set (at idx) (Just txt)
  _ -> pure ()

parseSheet ::
  MonadIO m
  => MonadThrow m
  => PrimMonad m
  => MonadState PipeState m
  => Event -> Text -> ConduitT Event SheetItem m ()
parseSheet evt name = do
        prevSheetName <- use ps_sheet_name
        rix <- use ps_cell_row_index
        unless (prevSheetName == name) $ do
          ps_sheet_name .= name
          popRow >>= yield . MkSheetItem name rix

        liftIO $ print (name, evt)
        parseRes <- runExceptT $ matchEvent name evt
        rix' <- use ps_cell_row_index
        case parseRes of
          Left err -> liftIO $ print err
          Right mResult -> do
            traverse_ (\x -> do
                          liftIO $ print x
                          yield $ MkSheetItem name rix' x) mResult
            await >>= parseFileLoop

popRow :: MonadState PipeState m => m CellRow
popRow = do
  row <- use ps_row
  ps_row .= mempty
  pure row

addCell :: MonadState PipeState m => Text -> m ()
addCell txt = do
   inVal <- use ps_is_in_val
   when inVal $ do
      col <- use ps_cell_col_index
      ps_row <>= (Map.singleton col $ Cell
        { _cellStyle   = Nothing
        , _cellValue   = Just $ CellText txt -- TODO type
        , _cellComment = Nothing
        , _cellFormula = Nothing
        })


newtype PipeErrors = MkCoordinate CoordinateErrors
  deriving Show

data CoordinateErrors = CoordinateNotFound [(Name, [Content])]
                      | NoListElement (Name, [Content]) [(Name, [Content])]
                      | NoTextContent Content [(Name, [Content])]
                      | DecodeFailure Text [(Name, [Content])]
  deriving Show

contentTextPrims :: Prism' Content Text
contentTextPrims = prism' ContentText (\case ContentText x -> Just x
                                             _ -> Nothing)

setCoord :: MonadError PipeErrors m => MonadState PipeState m => [(Name, [Content])] -> m ()
setCoord list = do
  coordinates <- liftEither $ first MkCoordinate $ parseCoordinates list
  ps_cell_col_index .= (coordinates ^. _2)
  ps_cell_row_index .= (coordinates ^. _1)

findName :: Text -> [(Name, [Content])] -> Maybe (Name, [Content])
findName name = find ((name ==) . nameLocalName . fst)


parseCoordinates :: [(Name, [Content])] -> Either CoordinateErrors (Int, Int)
parseCoordinates list = do
      nameValPair <- maybe (Left $ CoordinateNotFound list) Right $ findName "r" list
      valContent <- maybe (Left $ NoListElement nameValPair list) Right $  nameValPair ^? _2 . ix 0
      valText <- maybe (Left $ NoTextContent valContent list) Right $ valContent ^? contentTextPrims
      maybe (Left $ DecodeFailure valText list) Right $ fromSingleCellRef $ CellRef valText

matchEvent :: MonadIO m => MonadError PipeErrors m => MonadState PipeState m => Text -> Event -> m (Maybe CellRow)
matchEvent _currentSheet = \case
  EventContent (ContentText txt)                       -> Nothing <$ addCell txt
  EventBeginElement Name{nameLocalName = "c", ..} vals  -> Nothing <$ setCoord vals
  EventBeginElement Name {nameLocalName = "v"} _       -> Nothing <$ (ps_is_in_val .= True)
  EventEndElement Name {nameLocalName = "v"}           -> Nothing <$ (ps_is_in_val .= False)
  -- EventBeginElement Name {nameLocalName = "c"} _    -> Nothing <$ ps_is_in_cell .= True
  -- EventEndElement Name {nameLocalName = "c"}        -> Nothing <$ ps_is_in_cell .= False
  EventBeginElement Name {nameLocalName = "row"} _     -> Nothing <$ popRow
  EventEndElement Name{nameLocalName = "row", ..}       -> Just <$> popRow
  _ -> pure Nothing
