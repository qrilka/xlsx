module Codec.Xlsx.Types.Internal where

import           Control.Arrow
import           Data.Text                  (Text)

import           Codec.Xlsx.Parser.Internal
import           Codec.Xlsx.Writer.Internal

newtype RefId = RefId { unRefId :: Text } deriving (Show, Eq, Ord)

instance ToAttrVal RefId where
    toAttrVal = toAttrVal . unRefId

instance FromAttrVal RefId where
    fromAttrVal t = first RefId <$> fromAttrVal t
