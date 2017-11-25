{-# LANGUAGE DeriveDataTypeable #-}
{-# LANGUAGE DeriveGeneric #-}
{-# LANGUAGE FunctionalDependencies #-}
{-# LANGUAGE MultiParamTypeClasses #-}
{-# LANGUAGE OverloadedStrings #-}
{-# LANGUAGE ScopedTypeVariables #-}
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
  , attrValIs
  , contentOrEmpty
  , readSuccess
  , readFailure
  , invalidText
  , defaultReadFailure
  , module Codec.Xlsx.Parser.Internal.Util
  , module Codec.Xlsx.Parser.Internal.Fast
  ) where

import Control.Exception (Exception)
import Data.Maybe
import Data.Text (Text)
import qualified Data.Text as T
import qualified Data.Text.Read as T
import Data.Typeable (Typeable)
import GHC.Generics (Generic)
import Text.XML
import Text.XML.Cursor

import Codec.Xlsx.Parser.Internal.Fast
import Codec.Xlsx.Parser.Internal.Util

data ParseException = ParseException String
                    deriving (Show, Typeable, Generic)

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

attrValIs :: (Eq a, FromAttrVal a) => Name -> a -> Axis
attrValIs n v c =
  case fromAttribute n c of
    [x] | x == v -> [c]
    _ -> []

contentOrEmpty :: Cursor -> [Text]
contentOrEmpty c =
  case c $/ content of
    [t] -> [t]
    [] -> [""]
    _ -> error "invalid item: more than one text node encountered"

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
