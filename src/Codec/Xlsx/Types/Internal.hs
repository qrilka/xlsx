module Codec.Xlsx.Types.Internal where

import           Data.Text (Text)

import Codec.Xlsx.Writer.Internal

newtype RefId = RefId { unRefId :: Text } deriving (Show, Eq, Ord)

instance ToAttrVal RefId where
  toAttrVal = toAttrVal . unRefId
