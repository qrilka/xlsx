{-# LANGUAGE CPP               #-}
{-# LANGUAGE OverloadedStrings #-}
{-# LANGUAGE RecordWildCards   #-}
module Codec.Xlsx.Types.Internal.ContentTypes where

import           Control.Arrow
import           Data.Map                   (Map)
import qualified Data.Map                   as M
import           Data.Text                  (Text)
import qualified Data.Text                  as T
import           Text.XML
import           Text.XML.Cursor

#if !MIN_VERSION_base(4,8,0)
import           Control.Applicative
#endif

import           Codec.Xlsx.Parser.Internal

data Override = Override
    { ovrPartName    :: FilePath
    , ovrContentType :: Text
    } deriving (Eq, Show)

newtype ContentTypes = ContentTypes
    { ctTypes :: Map FilePath Text
    } deriving (Eq, Show)

lookup :: FilePath -> ContentTypes -> Maybe Text
lookup path = M.lookup path . ctTypes

{-------------------------------------------------------------------------------
  Parsing
-------------------------------------------------------------------------------}
instance FromCursor ContentTypes where
    fromCursor cur = do
        let items = cur $/ element (ct"Override") >=> fromCursor
        return . ContentTypes . M.fromList $ map (ovrPartName &&& ovrContentType) items

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
