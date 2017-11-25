{-# LANGUAGE OverloadedStrings #-}
{-# LANGUAGE RecordWildCards   #-}
{-# LANGUAGE DeriveGeneric #-}
module Codec.Xlsx.Types.Internal.ContentTypes where

import Control.Arrow
import Data.Foldable (asum)
import Data.Map (Map)
import qualified Data.Map as M
import Data.Text (Text)
import qualified Data.Text as T
import GHC.Generics (Generic)
import System.FilePath.Posix (takeExtension)
import Text.XML
import Text.XML.Cursor

import Codec.Xlsx.Parser.Internal

data CtDefault = CtDefault
    { dfltExtension   :: FilePath
    , dfltContentType :: Text
    } deriving (Eq, Show, Generic)

data Override = Override
    { ovrPartName    :: FilePath
    , ovrContentType :: Text
    } deriving (Eq, Show, Generic)

data ContentTypes = ContentTypes
    { ctDefaults :: Map FilePath Text
    , ctTypes    :: Map FilePath Text
    } deriving (Eq, Show, Generic)

lookup :: FilePath -> ContentTypes -> Maybe Text
lookup path ContentTypes{..} =
    asum [ flip M.lookup ctDefaults =<< ext, M.lookup path ctTypes ]
  where
    ext = case takeExtension path of
        '.':e -> Just e
        _     -> Nothing

{-------------------------------------------------------------------------------
  Parsing
-------------------------------------------------------------------------------}
instance FromCursor ContentTypes where
    fromCursor cur = do
        let ds = M.fromList . map (dfltExtension &&& dfltContentType) $
                 cur $/ element (ct"Default") >=> fromCursor
            ts = M.fromList . map (ovrPartName &&& ovrContentType) $
                 cur $/ element (ct"Override") >=> fromCursor
        return (ContentTypes ds ts)

instance FromCursor CtDefault where
   fromCursor cur = do
       dfltExtension <- T.unpack <$> attribute "Extension" cur
       dfltContentType <- attribute "ContentType" cur
       return CtDefault{..}

instance FromCursor Override where
   fromCursor cur = do
       ovrPartName <- T.unpack <$> attribute "PartName" cur
       ovrContentType <- attribute "ContentType" cur
       return Override{..}

-- | Add package relationship namespace to name
ct :: Text -> Name
ct x = Name
  { nameLocalName = x
  , nameNamespace = Just contentTypesNs
  , namePrefix = Nothing
  }

contentTypesNs :: Text
contentTypesNs = "http://schemas.openxmlformats.org/package/2006/content-types"
