{-# LANGUAGE DeriveGeneric #-}
{-# LANGUAGE OverloadedStrings #-}
module Codec.Xlsx.Types.Internal where

import Control.Arrow
import Data.Monoid ((<>))
import Data.Text (Text)
import GHC.Generics (Generic)

import Codec.Xlsx.Parser.Internal
import Codec.Xlsx.Writer.Internal

newtype RefId = RefId { unRefId :: Text } deriving (Eq, Ord, Show, Generic)

instance ToAttrVal RefId where
    toAttrVal = toAttrVal . unRefId

instance FromAttrVal RefId where
    fromAttrVal t = first RefId <$> fromAttrVal t

instance FromAttrBs RefId where
  fromAttrBs = fmap RefId . fromAttrBs

unsafeRefId :: Int -> RefId
unsafeRefId num = RefId $ "rId" <> txti num
