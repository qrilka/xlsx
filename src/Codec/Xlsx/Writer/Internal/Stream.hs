{-# LANGUAGE ConstraintKinds     #-}
{-# LANGUAGE DataKinds           #-}
{-# LANGUAGE DeriveAnyClass      #-}
{-# LANGUAGE DeriveGeneric       #-}
{-# LANGUAGE FlexibleContexts    #-}
{-# LANGUAGE LambdaCase          #-}
{-# LANGUAGE NamedFieldPuns      #-}
{-# LANGUAGE OverloadedStrings   #-}
{-# LANGUAGE RankNTypes          #-}
{-# LANGUAGE RecordWildCards     #-}
{-# LANGUAGE ScopedTypeVariables #-}
{-# LANGUAGE StandaloneDeriving  #-}
{-# LANGUAGE StrictData          #-}
{-# LANGUAGE TemplateHaskell     #-}
{-# LANGUAGE TypeApplications    #-}
{-# LANGUAGE CPP #-}

-- | Internal stream related functions.
--   These are exported because they're tested like this.
--   It's not expected a user would need this.
module Codec.Xlsx.Writer.Internal.Stream
  ( getSetNumber
  , initialSharedString
  , string_map
  , SharedStringState(..)
  ) where


#ifdef USE_MICROLENS
import           Lens.Micro.Platform
#else
import           Control.Lens
#endif
import           Control.Monad.State.Strict
import           Data.Map.Strict            (Map)
import           Data.Maybe
import           Data.Text                  (Text)

newtype SharedStringState = MkSharedStringState
  { _string_map :: Map Text Int
  }
makeLenses 'MkSharedStringState

initialSharedString :: SharedStringState
initialSharedString = MkSharedStringState mempty

-- properties:
-- for a list of [text], every unique text get's a unique number.
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

