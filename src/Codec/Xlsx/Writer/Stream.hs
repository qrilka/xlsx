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
import           Text.XML.Stream.Parse
import           Codec.Xlsx.Parser.Stream
import Data.Maybe

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
writeSharedStrings :: ( MonadThrow m , PrimMonad m)
  => ConduitT SheetItem (Text, Int) m (Map Text Int)
writeSharedStrings = fmap (view string_map) $ C.execStateC initialSharedString $
  C.mapFoldableM mapFold
