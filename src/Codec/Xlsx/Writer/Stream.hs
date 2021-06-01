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
  ( writeSharedStrings
  , writeXlsx
  , writeSstXml
  -- testing
  , getSetNumber
  , initialSharedString
  , string_map
  ) where

import           Codec.Archive.Zip.Conduit.UnZip
import           Codec.Xlsx.Types.Cell
import           Codec.Xlsx.Types.Common
import           Conduit                         (PrimMonad, await, yield, (.|))
import qualified Conduit                         as C
import           Control.Lens
import           Control.Monad.Catch
import           Control.Monad.Except
import           Control.Monad.State.Strict
import           Data.Bifunctor
import Data.ByteString(ByteString)
import Data.ByteString.Builder(Builder)
import qualified Data.ByteString                 as BS
import           Data.Conduit                    (ConduitT)
import qualified Data.Conduit.Combinators        as C
import qualified Data.Conduit.List as C
import           Data.Foldable
import           Data.Map.Strict                 (Map)
import qualified Data.Map.Strict                 as Map
import           Data.Text                       (Text)
import qualified Data.Text                       as Text
import qualified Data.Text.Encoding              as Text
import qualified Data.Text.Read                  as Read
import           Data.XML.Types
import           Debug.Trace
import           GHC.Generics
import           NoThunks.Class
import           Text.Read
import           Text.XML.Stream.Render
import           Codec.Xlsx.Parser.Stream
import Data.Maybe
import Data.List

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

-- | creates a unique number for every encountered string in the stream
--   This is used for creating a required structure in the xlsx format
--   called shared strings. Every string get's transformed into a number
writeSharedStrings :: Monad m  =>
  ConduitT SheetItem (Text, Int) m (Map Text Int)
writeSharedStrings = fmap (view string_map) $ C.execStateC initialSharedString $
  C.mapFoldableM mapFold

-- TODO maybe should use bimap instead: https://hackage.haskell.org/package/bimap-0.4.0/docs/Data-Bimap.html
-- it guarantees uniqueness of both text and int
writeXlsx :: PrimMonad m  =>
  Map Text Int ->
  ConduitT SheetItem Builder m ()
writeXlsx sstable = error "x"

writeSstXml  ::  PrimMonad m  =>  Map Text Int  -> forall i. ConduitT i Builder m ()
writeSstXml sstable = writeSst sstable .| writeEvents

writeSst ::  Monad m  => Map Text Int  -> forall i.  ConduitT i Event m ()
writeSst sstable = do
  yield EventBeginDocument
  yield $ EventBeginElement "sst" []
  traverse (\(e, _)  -> do
                yield $ EventBeginElement "si" []
                yield $ EventBeginElement "t" []
                yield $ EventContent (ContentText e)
                yield $ EventEndElement "t"
                yield $ EventEndElement "si"
                ) $ sortBy (\(_, i) (_, y :: Int) -> compare i y) $ Map.toList sstable
  yield $ EventEndElement "sst"
  yield EventEndDocument

writeDataRows :: Monad m  => Map Text Int -> ConduitT SheetItem Event m ()
writeDataRows = error "i"

writeEvents ::  PrimMonad m => ConduitT Event Builder m ()
writeEvents = renderBuilder (def {rsPretty=True})
