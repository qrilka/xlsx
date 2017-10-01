{-# LANGUAGE DeriveGeneric #-}
module Codec.Xlsx.Types.Comment where

import Control.DeepSeq (NFData)
import Data.Text (Text)
import GHC.Generics (Generic)

import Codec.Xlsx.Types.Common

-- | User comment for a cell
--
-- TODO: the following child elements:
-- guid, shapeId, commentPr
--
-- Section 18.7.3 "comment (Comment)" (p. 1749)
data Comment = Comment
    { _commentText    :: XlsxText
    -- ^ cell comment text, maybe formatted
    -- Section 18.7.7 "text (Comment Text)" (p. 1754)
    , _commentAuthor  :: Text
    -- ^ comment author
    , _commentVisible :: Bool
    } deriving (Eq, Show, Generic)
instance NFData Comment
