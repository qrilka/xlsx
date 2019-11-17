module Codec.Xlsx.Parser.Internal.Util
  ( boolean
  , eitherBoolean
  , decimal
  , eitherDecimal
  , rational
  , eitherRational
  ) where

import Data.Text (Text)
import Control.Monad.Fail (MonadFail)
import qualified Data.Text as T
import qualified Data.Text.Read as T
import qualified Control.Monad.Fail as F

decimal :: (MonadFail m, Integral a) => Text -> m a
decimal = fromEither . eitherDecimal

eitherDecimal :: (Integral a) => Text -> Either String a
eitherDecimal t = case T.signed T.decimal t of
  Right (d, leftover) | T.null leftover -> Right d
  _ -> Left $ "invalid decimal: " ++ show t

rational :: (MonadFail m) => Text -> m Double
rational = fromEither . eitherRational

eitherRational :: Text -> Either String Double
eitherRational t = case T.signed T.rational t of
  Right (r, leftover) | T.null leftover -> Right r
  _ -> Left $ "invalid rational: " ++ show t

boolean :: (MonadFail m) => Text -> m Bool
boolean = fromEither . eitherBoolean

eitherBoolean :: Text -> Either String Bool
eitherBoolean t = case T.unpack $ T.strip t of
    "true"  -> Right True
    "false" -> Right False
    _       -> Left $ "invalid boolean: " ++ show t

fromEither :: (MonadFail m) => Either String b -> m b
fromEither (Left a) = F.fail a
fromEither (Right b) = return b
