{-# LANGUAGE OverloadedStrings #-}
module Codec.Xlsx.Parser.Internal.Util
  ( boolean
  , decimal
  , rational
  ) where

import Data.Text (Text)
import qualified Data.Text as T
import qualified Data.Text.Read as T

decimal :: (Monad m, Integral a) => Text -> m a
decimal t = case T.signed T.decimal $ t of
  Right (d, leftover) | T.null leftover -> return d
  _ -> fail $ "invalid decimal" ++ show t

rational :: Monad m => Text -> m Double
rational t = case T.signed T.rational t of
  Right (r, leftover) | T.null leftover -> return r
  _ -> fail $ "invalid rational: " ++ show t

boolean :: Monad m => Text -> m Bool
boolean t = case T.strip t of
    "true"  -> return True
    "false" -> return False
    _       -> fail $ "invalid boolean: " ++ show t
