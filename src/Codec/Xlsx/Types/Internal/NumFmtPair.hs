{-# LANGUAGE OverloadedStrings #-}
module Codec.Xlsx.Types.Internal.NumFmtPair where

import           Data.Text                  (Text)

import           Codec.Xlsx.Parser.Internal
import           Codec.Xlsx.Writer.Internal

-- | Internal helper type for parsing "numFmt" recods
--
-- See 18.8.30 "numFmt (Number Format)" (p. 1777)
newtype NumFmtPair = NumFmtPair
    { unNumFmtPair :: (Int, Text)
    } deriving (Eq, Show)

instance FromCursor NumFmtPair where
    fromCursor cur = do
        fId <- fromAttribute "numFmtId" cur
        fCode <- fromAttribute "formatCode" cur
        return $ NumFmtPair (fId, fCode)

instance ToElement NumFmtPair where
    toElement nm (NumFmtPair (fId, fCode)) =
        leafElement nm [ "numFmtId"   .= toAttrVal fId
                       , "formatCode" .= toAttrVal fCode ]
