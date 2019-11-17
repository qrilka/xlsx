{-# LANGUAGE BangPatterns #-}
{-# LANGUAGE LambdaCase #-}
{-# LANGUAGE MultiParamTypeClasses #-}
{-# LANGUAGE MultiWayIf #-}
{-# LANGUAGE OverloadedStrings #-}
{-# LANGUAGE TupleSections #-}
module Codec.Xlsx.Parser.Internal.Fast
  ( FromXenoNode(..)
  , collectChildren
  , maybeChild
  , requireChild
  , childList
  , maybeFromChild
  , fromChild
  , fromChildList
  , maybeParse
  , requireAndParse
  , childListAny
  , maybeElementVal
  , toAttrParser
  , parseAttributes
  , FromAttrBs(..)
  , unexpectedAttrBs
  , maybeAttrBs
  , maybeAttr
  , fromAttr
  , fromAttrDef
  , contentBs
  , contentX
  , nsPrefixes
  , addPrefix
  ) where

import Control.Applicative
import Control.Arrow (second)
import Control.Exception (Exception, throw)
import Control.Monad (ap, forM, join, liftM)
import Data.Bifunctor (first)
import Data.Bits ((.|.), shiftL)
import Data.ByteString (ByteString)
import qualified Data.ByteString as BS
import qualified Data.ByteString.Unsafe as SU
import Data.Char (chr)
import Data.Maybe
import Data.Monoid ((<>))
import Data.Text (Text)
import qualified Data.Text as T
import qualified Data.Text.Encoding as T
import Data.Word (Word8)
import Xeno.DOM hiding (parse)

import Codec.Xlsx.Parser.Internal.Util

class FromXenoNode a where
  fromXenoNode :: Node -> Either Text a

newtype ChildCollector a = ChildCollector
  { runChildCollector :: [Node] -> Either Text ([Node], a)
  }

instance Functor ChildCollector where
  fmap f a = ChildCollector $ \ns ->
    second f <$> runChildCollector a ns

instance Applicative ChildCollector where
  pure a = ChildCollector $ \ns ->
    return (ns, a)
  cf <*> ca = ChildCollector $ \ns -> do
    (ns', f) <- runChildCollector cf ns
    (ns'', a) <- runChildCollector ca ns'
    return (ns'', f a)

instance Alternative ChildCollector where
  empty = ChildCollector $ \_ -> Left "ChildCollector.empty"
  ChildCollector f <|> ChildCollector g = ChildCollector $ \ns ->
    either (const $ g ns) Right (f ns)

instance Monad ChildCollector where
  return = pure
  ChildCollector f >>= g = ChildCollector $ \ns ->
    either Left (\(!ns', f') -> runChildCollector (g f') ns') (f ns)

toChildCollector :: Either Text a -> ChildCollector a
toChildCollector unlifted =
  case unlifted of
    Right a -> return a
    Left e -> ChildCollector $ \_ -> Left e

collectChildren :: Node -> ChildCollector a -> Either Text a
collectChildren n c = snd <$> runChildCollector c (children n)

maybeChild :: ByteString -> ChildCollector (Maybe Node)
maybeChild nm =
  ChildCollector $ \case
    (n:ns)
      | name n == nm -> pure (ns, Just n)
    ns -> pure (ns, Nothing)

requireChild :: ByteString -> ChildCollector Node
requireChild nm =
  ChildCollector $ \case
    (n:ns)
      | name n == nm -> pure (ns, n)
    _ ->
      Left $ "required element " <> T.pack (show nm) <> " was not found"

childList :: ByteString -> ChildCollector [Node]
childList nm = do
  mNode <- maybeChild nm
  case mNode of
    Just n -> (n:) <$> childList nm
    Nothing -> return []

maybeFromChild :: (FromXenoNode a) => ByteString -> ChildCollector (Maybe a)
maybeFromChild nm = do
  mNode <- maybeChild nm
  mapM (toChildCollector . fromXenoNode) mNode

fromChild :: (FromXenoNode a) => ByteString -> ChildCollector a
fromChild nm = do
  n <- requireChild nm
  case fromXenoNode n of
    Right a -> return a
    Left e -> ChildCollector $ \_ -> Left e

fromChildList :: (FromXenoNode a) => ByteString -> ChildCollector [a]
fromChildList nm = do
  mA <- maybeFromChild nm
  case mA of
    Just a -> (a:) <$> fromChildList nm
    Nothing -> return []

maybeParse :: ByteString -> (Node -> Either Text a) -> ChildCollector (Maybe a)
maybeParse nm parse = maybeChild nm >>= (toChildCollector . mapM parse)

requireAndParse :: ByteString -> (Node -> Either Text a) -> ChildCollector a
requireAndParse nm parse = requireChild nm >>= (toChildCollector . parse)

childListAny :: (FromXenoNode a) => Node -> Either Text [a]
childListAny = mapM fromXenoNode . children

maybeElementVal :: (FromAttrBs a) => ByteString -> ChildCollector (Maybe a)
maybeElementVal nm = do
  mN <- maybeChild nm
  fmap join . forM mN $ \n ->
    toChildCollector . parseAttributes n $ maybeAttr "val"

-- Stolen from XML Conduit
newtype AttrParser a = AttrParser
  { runAttrParser :: [(ByteString, ByteString)] -> Either Text ( [( ByteString
                                                                  , ByteString)]
                                                               , a)
  }

instance Monad AttrParser where
  return a = AttrParser $ \as -> Right (as, a)
  (AttrParser f) >>= g =
    AttrParser $ \as ->
      either Left (\(as', f') -> runAttrParser (g f') as') (f as)
instance Applicative AttrParser where
    pure = return
    (<*>) = ap
instance Functor AttrParser where
  fmap = liftM

attrError :: Text -> AttrParser a
attrError err = AttrParser $ \_ -> Left err

toAttrParser :: Either Text a -> AttrParser a
toAttrParser unlifted =
  case unlifted of
    Right a -> return a
    Left e -> AttrParser $ \_ -> Left e

maybeAttrBs :: ByteString -> AttrParser (Maybe ByteString)
maybeAttrBs attrName = AttrParser $ go id
  where
    go front [] = Right (front [], Nothing)
    go front (a@(nm, val):as) =
      if nm == attrName
        then Right (front as, Just val)
        else go (front . (:) a) as

requireAttrBs :: ByteString -> AttrParser ByteString
requireAttrBs nm = do
  mVal <- maybeAttrBs nm
  case mVal of
    Just val -> return val
    Nothing -> attrError $ "attribute " <> T.pack (show nm) <> " is required"

unexpectedAttrBs :: Text -> ByteString -> Either Text a
unexpectedAttrBs typ val =
  Left $ "Unexpected value for " <> typ <> ": " <> T.pack (show val)

fromAttr :: FromAttrBs a => ByteString -> AttrParser a
fromAttr nm = do
  bs <- requireAttrBs nm
  toAttrParser $ fromAttrBs bs

maybeAttr :: FromAttrBs a => ByteString -> AttrParser (Maybe a)
maybeAttr nm = do
  mBs <- maybeAttrBs nm
  forM mBs (toAttrParser . fromAttrBs)

fromAttrDef :: FromAttrBs a => ByteString -> a -> AttrParser a
fromAttrDef nm defVal = fromMaybe defVal <$> maybeAttr nm

parseAttributes :: Node -> AttrParser a -> Either Text a
parseAttributes n attrParser =
  case runAttrParser attrParser (attributes n) of
    Left e -> Left e
    Right (_, a) -> return a

class FromAttrBs a where
  fromAttrBs :: ByteString -> Either Text a

instance FromAttrBs ByteString where
  fromAttrBs = pure

instance FromAttrBs Bool where
    fromAttrBs x | x == "1" || x == "true"  = return True
                 | x == "0" || x == "false" = return False
                 | otherwise                = unexpectedAttrBs "boolean" x

instance FromAttrBs Int where
  -- it appears that parser in text is more optimized than the one in
  -- attoparsec at least as of text-1.2.2.2 and attoparsec-0.13.1.0
  fromAttrBs = first T.pack . eitherDecimal . T.decodeLatin1

instance FromAttrBs Double where
  -- as for rationals
  fromAttrBs = first T.pack . eitherRational . T.decodeLatin1

instance FromAttrBs Text where
  fromAttrBs = replaceEntititesBs

replaceEntititesBs :: ByteString -> Either Text Text
replaceEntititesBs str =
  T.decodeUtf8 . BS.concat <$> findAmp 0
  where
    findAmp :: Int -> Either Text [ByteString]
    findAmp index =
      case elemIndexFrom ampersand str index of
        Nothing -> if BS.null text then return [] else return [text]
          where text = BS.drop index str
        Just fromAmp ->
          if BS.null text
             then checkEntity fromAmp
             else (text:) <$> checkEntity fromAmp
          where text = substring str index fromAmp
    checkEntity index =
      case elemIndexFrom semicolon str index of
        Just fromSemi | fromSemi >= index + 3 -> do
                          entity <- checkElementVal (index + 1) (fromSemi - index - 1)
                          (BS.singleton entity:) <$> findAmp (fromSemi + 1)
        _ -> Left "Unending entity"
    checkElementVal index len =
      if | len == 2
           && s_index this 0 == 108 -- l
           && s_index this 1 == 116 -- t
            -> return 60 -- '<'
         | len == 2
           && s_index this 0 == 103 -- g
           && s_index this 1 == 116 -- t
            -> return 62 -- '>'
         | len == 3
           && s_index this 0 ==  97 -- a
           && s_index this 1 == 109 -- m
           && s_index this 2 == 112 -- p
            -> return 38 -- '&'
         | len == 4
           && s_index this 0 == 113 -- q
           && s_index this 1 == 117 -- u
           && s_index this 2 == 111 -- o
           && s_index this 3 == 116 -- t
            -> return 34 -- '"'
         | len == 4
           && s_index this 0 ==  97 -- a
           && s_index this 1 == 112 -- p
           && s_index this 2 == 111 -- o
           && s_index this 3 == 115 -- s
           -> return 39 -- '\''
         |    s_index this 0 == 35  -- '#'
           ->
           if s_index this 1 == 120 -- 'x'
              then toEnum <$> checkHexadecimal (index + 2) (len - 2)
              else toEnum <$> checkDecimal (index + 1) (len - 1)
         | otherwise -> Left $ "Bad entity " <> T.pack (show $ (substring str (index-1) (index+len+1)))
      where
        this = BS.drop index str
    checkDecimal index len = BS.foldl' go (Right 0) (substring str index (index + len))
      where
        go :: Either Text Int -> Word8 -> Either Text Int
        go prev c = do
          a <- prev
          if c >= 48 && c <= 57
            then return $ a * 10 + fromIntegral (c - 48)
            else Left $ "Expected decimal digit but encountered " <> T.pack (show (chr $ fromIntegral c))
    checkHexadecimal index len = BS.foldl' go (Right 0) (substring str index (index + len))
      where
        go :: Either Text Int -> Word8 -> Either Text Int
        go prev c = do
          a <- prev
          if | c >= 48 && c <= 57
               -> return $ (a `shiftL` 4) .|. fromIntegral (c - 48)
             | c >= 97 && c <= 122
               -> return $ (a `shiftL` 4) .|. fromIntegral (c - 87)
             | c >= 65 && c <= 90
               -> return $ (a `shiftL` 4) .|. fromIntegral (c - 55)
             | otherwise
               ->
               Left $ "Expected hexadecimal digit but encountered " <> T.pack (show (chr $ fromIntegral c))
    ampersand = 38
    semicolon = 59

data EntityReplaceException = EntityReplaceException deriving Show

instance Exception EntityReplaceException

-- | /O(1)/ 'ByteString' index (subscript) operator, starting from 0.
s_index :: ByteString -> Int -> Word8
s_index ps n
    | n < 0             = throw EntityReplaceException
    | n >= BS.length ps = throw EntityReplaceException
    | otherwise         = ps `SU.unsafeIndex` n
{-# INLINE s_index #-}

-- | Get index of an element starting from offset.
elemIndexFrom :: Word8 -> ByteString -> Int -> Maybe Int
elemIndexFrom c str offset = fmap (+ offset) (BS.elemIndex c (BS.drop offset str))
-- Without the INLINE below, the whole function is twice as slow and
-- has linear allocation. See git commit with this comment for
-- results.
{-# INLINE elemIndexFrom #-}

-- | Get a substring of a string.
substring :: ByteString -> Int -> Int -> ByteString
substring s start end = BS.take (end - start) (BS.drop start s)
{-# INLINE substring #-}

newtype NsPrefixes = NsPrefixes [(ByteString, ByteString)]

nsPrefixes :: Node -> NsPrefixes
nsPrefixes root =
  NsPrefixes . flip mapMaybe (attributes root) $ \(nm, val) ->
    (val, ) <$> BS.stripPrefix "xmlns:" nm

addPrefix :: NsPrefixes -> ByteString -> (ByteString -> ByteString)
addPrefix (NsPrefixes prefixes) ns =
  maybe id (\prefix nm -> BS.concat [prefix, ":", nm]) $ Prelude.lookup ns prefixes

contentBs :: Node -> ByteString
contentBs n = BS.concat . map toBs $ contents n
  where
    toBs (Element _) = BS.empty
    toBs (Text bs) = bs
    toBs (CData bs) = bs

contentX :: Node -> Either Text Text
contentX = replaceEntititesBs . contentBs
