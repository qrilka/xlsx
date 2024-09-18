{-# LANGUAGE CPP              #-}
{-# LANGUAGE BangPatterns     #-}
{-# LANGUAGE FlexibleContexts #-}
{-# LANGUAGE TemplateHaskell  #-}
{-# LANGUAGE ViewPatterns     #-}

-- | Internal stream related functions.
--   These are exported because they're tested like this.
--   It's not expected a user would need this.
module Codec.Xlsx.Writer.Internal.Stream
  ( upsertSharedString
  , initialSharedString
  , string_map
  , T(..)
  , SharedStringState(..)
  ) where


#ifdef USE_MICROLENS
import Lens.Micro.Platform
#else
import Control.Lens
#endif
import Control.Monad.State.Strict
import Data.Char
import Data.Map.Strict (Map)
import Data.Text (Text)
import qualified Data.Text as Text

import Codec.Xlsx.Writer.Internal (cleanText)

data T = T !Text !Int

newtype SharedStringState = MkSharedStringState
  { _string_map :: Map Text T
  }
makeLenses 'MkSharedStringState

initialSharedString :: SharedStringState
initialSharedString = MkSharedStringState mempty

-- properties:
-- for a list of [text], every unique text gets a unique number.
upsertSharedString :: MonadState SharedStringState m => Text -> m (Text, Int)
upsertSharedString (cleanText -> current) = do
  strings  <- use string_map

  case strings ^? ix current of
    Just (T old i) -> pure (old, i)
    Nothing -> do
      let !idx = length strings
      string_map .= (strings & ix current .~ T current idx)
      pure (current, idx)
