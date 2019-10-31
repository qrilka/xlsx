{-# LANGUAGE CPP #-}
module Codec.Xlsx.Parser.Internal.Util
  ( boolean
  , eitherBoolean
  , decimal
  , eitherDecimal
  , rational
  , eitherRational
  ) where

import Data.Text (Text)
import qualified Data.Text as T
import qualified Data.Text.Read as T

#if (MIN_VERSION_base(4,13,0))
#define FAIL_MONAD MonadFail
#else
#define FAIL_MONAD Monad
#endif

decimal :: (FAIL_MONAD m, Integral a) => Text -> m a
decimal = fromEither . eitherDecimal

eitherDecimal :: (Integral a) => Text -> Either String a
eitherDecimal t = case T.signed T.decimal t of
  Right (d, leftover) | T.null leftover -> return d
  _ -> Left $ "invalid decimal" ++ show t

rational :: (FAIL_MONAD m) => Text -> m Double
rational = fromEither . eitherRational

eitherRational :: Text -> Either String Double
eitherRational t = case T.signed T.rational t of
  Right (r, leftover) | T.null leftover -> return r
  _ -> Left $ "invalid rational: " ++ show t

boolean :: (FAIL_MONAD m) => Text -> m Bool
boolean = fromEither . eitherBoolean

eitherBoolean :: Text -> Either String Bool
eitherBoolean t = case T.unpack $ T.strip t of
    "true"  -> return True
    "false" -> return False
    _       -> Left $ "invalid boolean: " ++ show t

fromEither :: (FAIL_MONAD m) => Either String b -> m b
fromEither (Left a) = fail a
fromEither (Right b) = return b
