{-# LANGUAGE CPP                #-}
{-# LANGUAGE DeriveDataTypeable #-}
{-# LANGUAGE OverloadedStrings  #-}
module Codec.Xlsx.Parser.Internal
    ( ParseException(..)
    , n
    , FromCursor(..)
    , FromAttrVal(..)
    , fromAttribute
    , maybeAttribute
    , maybeElementValue
    , maybeFromElement
    , readSuccess
    , readFailure
    , invalidText
    , defaultReadFailure
    , boolean
    , decimal
    , rational
    ) where

import           Control.Exception       (Exception)
import           Data.Text               (Text)
import qualified Data.Text               as T
import qualified Data.Text.Read          as T
import           Data.Typeable           (Typeable)
import           Text.XML
import           Text.XML.Cursor

#if !MIN_VERSION_base(4,8,0)
import           Control.Applicative
#endif

data ParseException = ParseException String
                    deriving (Show, Typeable)

instance Exception ParseException

class FromCursor a where
    fromCursor :: Cursor -> [a]

class FromAttrVal a where
    fromAttrVal :: T.Reader a

instance FromAttrVal Text where
    fromAttrVal = readSuccess

instance FromAttrVal Int where
    fromAttrVal = T.decimal

instance FromAttrVal Double where
    fromAttrVal = T.rational

instance FromAttrVal Bool where
    fromAttrVal x | x == "1" || x == "true"  = readSuccess True
                  | x == "0" || x == "false" = readSuccess False
                  | otherwise                = defaultReadFailure

-- | required attribute parsing
fromAttribute :: FromAttrVal a => Name -> Cursor -> [a]
fromAttribute name cursor =
    attribute name cursor >>= runReader fromAttrVal

-- | parsing optional attributes
maybeAttribute :: FromAttrVal a => Name -> Cursor -> [Maybe a]
maybeAttribute name cursor =
    case attribute name cursor of
      [attr] -> Just <$> runReader fromAttrVal attr
      _ -> [Nothing]

maybeElementValue :: FromAttrVal a => Name -> Cursor -> [Maybe a]
maybeElementValue name cursor =
  case cursor $/ element name of
    [cursor'] -> maybeAttribute "val" cursor'
    _ -> [Nothing]

maybeFromElement :: FromCursor a => Name -> Cursor -> [Maybe a]
maybeFromElement name cursor = case cursor $/ element name of
  [cursor'] -> Just <$> fromCursor cursor'
  _ -> [Nothing]


readSuccess :: a -> Either String (a, Text)
readSuccess x = Right (x, T.empty)

readFailure :: Text -> Either String (a, Text)
readFailure = Left . T.unpack

invalidText :: Text -> Text -> Either String (a, Text)
invalidText what txt = readFailure $ T.concat ["Invalid ", what, ": '", txt , "'"]

defaultReadFailure :: Either String (a, Text)
defaultReadFailure = Left "invalid text"

runReader :: T.Reader a -> Text -> [a]
runReader reader t = case reader t of
  Right (r, _) -> [r]
  _ -> []

-- | Add sml namespace to name
n :: Text -> Name
n x = Name
  { nameLocalName = x
  , nameNamespace = Just "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  , namePrefix = Nothing
  }

decimal :: Monad m => Text -> m Int
decimal t = case T.decimal t of
  Right (d, _) -> return d
  _ -> fail $ "invalid decimal" ++ show t

rational :: Monad m => Text -> m Double
rational t = case T.rational t of
  Right (r, _) -> return r
  _ -> fail $ "invalid rational: " ++ show t

boolean :: Monad m => Text -> m Bool
boolean t = case T.strip t of
    "true"  -> return True
    "false" -> return False
    _       -> fail $ "invalid boolean: " ++ show t
