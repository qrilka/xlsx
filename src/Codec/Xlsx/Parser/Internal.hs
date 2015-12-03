{-# LANGUAGE CPP               #-}
{-# LANGUAGE OverloadedStrings #-}
module Codec.Xlsx.Parser.Internal
    ( n
    , parseSharedStrings
    , FromCursor(..)
    , FromAttrVal(..)
    , fromAttribute
    , maybeAttribute
    , readSuccess
    , readFailure
    , invalidText
    , defaultReadFailure
    , decimal
    , rational
    ) where

import qualified Data.IntMap as IM
import Data.Text (Text)
import qualified Data.Text as T
import qualified Data.Text.Read as T
import Data.XML.Types
import Text.XML.Cursor

#if !MIN_VERSION_base(4,8,0)
import Control.Applicative
#endif

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

parseSharedStrings :: Cursor -> IM.IntMap Text
parseSharedStrings c = IM.fromAscList $ zip [0..] (c $/ element (n"si") >=> parseT)
    where
      -- it's  either <t> or <r>s with <t> inside
      parseT c' = [T.concat $ c' $| orSelf (child >=> (element (n"r"))) &/ element (n"t") &/ content]

decimal :: Monad m => Text -> m Int
decimal t = case T.decimal t of
  Right (d, _) -> return d
  _ -> fail "invalid decimal"

rational :: Monad m => Text -> m Double
rational t = case T.rational t of
  Right (r, _) -> return r
  _ -> fail "invalid rational"
