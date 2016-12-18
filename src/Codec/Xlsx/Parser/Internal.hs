{-# LANGUAGE CPP                #-}
{-# LANGUAGE DeriveDataTypeable #-}
{-# LANGUAGE OverloadedStrings  #-}
module Codec.Xlsx.Parser.Internal
    ( ParseException(..)
    , n_
    , nodeElNameIs
    , FromCursor(..)
    , FromAttrVal(..)
    , fromAttribute
    , fromAttributeDef
    , maybeAttribute
    , fromElementValue
    , maybeElementValue
    , maybeElementValueDef
    , maybeBoolElementValue
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
import           Data.Maybe
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

nodeElNameIs :: Node -> Name -> Bool
nodeElNameIs (NodeElement el) name = elementName el == name
nodeElNameIs _ _                   = False

class FromCursor a where
    fromCursor :: Cursor -> [a]

class FromAttrVal a where
    fromAttrVal :: T.Reader a

instance FromAttrVal Text where
    fromAttrVal = readSuccess

instance FromAttrVal Int where
    fromAttrVal = T.signed T.decimal

instance FromAttrVal Integer where
    fromAttrVal = T.signed T.decimal

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

-- | parsing optional attributes with defaults
fromAttributeDef :: FromAttrVal a => Name -> a -> Cursor -> [a]
fromAttributeDef name defVal cursor =
    case attribute name cursor of
      [attr] -> runReader fromAttrVal attr
      _      -> [defVal]

-- | parsing optional attributes
maybeAttribute :: FromAttrVal a => Name -> Cursor -> [Maybe a]
maybeAttribute name cursor =
    case attribute name cursor of
      [attr] -> Just <$> runReader fromAttrVal attr
      _ -> [Nothing]

fromElementValue :: FromAttrVal a => Name -> Cursor -> [a]
fromElementValue name cursor =
    cursor $/ element name >=> fromAttribute "val"

maybeElementValue :: FromAttrVal a => Name -> Cursor -> [Maybe a]
maybeElementValue name cursor =
  case cursor $/ element name of
    [cursor'] -> maybeAttribute "val" cursor'
    _ -> [Nothing]

maybeElementValueDef :: FromAttrVal a => Name -> a -> Cursor -> [Maybe a]
maybeElementValueDef name defVal cursor =
  case cursor $/ element name of
    [cursor'] -> Just . fromMaybe defVal <$> maybeAttribute "val" cursor'
    _ -> [Nothing]

maybeBoolElementValue :: Name -> Cursor -> [Maybe Bool]
maybeBoolElementValue name cursor = maybeElementValueDef name True cursor

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
  Right (r, leftover) | T.null leftover -> [r]
  _ -> []

-- | Add sml namespace to name
n_ :: Text -> Name
n_ x = Name
  { nameLocalName = x
  , nameNamespace = Just "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  , namePrefix = Just "n"
  }

decimal :: (Monad m, Integral a) => Text -> m a
decimal t = case T.signed T.decimal $ t of
  Right (d, leftover) | T.null leftover -> return d
  _ -> fail $ "invalid decimal" ++ show t

rational :: Monad m => Text -> m Double
rational t = case T.rational t of
  Right (r, leftover) | T.null leftover -> return r
  _ -> fail $ "invalid rational: " ++ show t

boolean :: Monad m => Text -> m Bool
boolean t = case T.strip t of
    "true"  -> return True
    "false" -> return False
    _       -> fail $ "invalid boolean: " ++ show t
