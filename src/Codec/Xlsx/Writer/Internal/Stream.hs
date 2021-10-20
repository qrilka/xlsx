{-# LANGUAGE CPP              #-}
{-# LANGUAGE FlexibleContexts #-}
{-# LANGUAGE TemplateHaskell  #-}

-- | Internal stream related functions.
--   These are exported because they're tested like this.
--   It's not expected a user would need this.
module Codec.Xlsx.Writer.Internal.Stream
  ( upsertSharedString
  , initialSharedString
  , string_map
  , SharedStringState(..)
  ) where


#ifdef USE_MICROLENS
import Lens.Micro.Platform
#else
import Control.Lens
#endif
import Control.Monad.State.Strict
import Data.Map.Strict (Map)
import Data.Maybe
import Data.Text (Text)

newtype SharedStringState = MkSharedStringState
  { _string_map :: Map Text Int
  }
makeLenses 'MkSharedStringState

initialSharedString :: SharedStringState
initialSharedString = MkSharedStringState mempty

-- properties:
-- for a list of [text], every unique text gets a unique number.
upsertSharedString :: MonadState SharedStringState m => Text -> m (Text,Int)
upsertSharedString current = do
  strings  <- use string_map

  let mIdx :: Maybe Int
      mIdx = strings ^? ix current

      idx :: Int
      idx = fromMaybe (length strings) mIdx

      newMap :: Map Text Int
      newMap = at current ?~ idx $ strings

  string_map .= newMap
  pure (current, idx)

