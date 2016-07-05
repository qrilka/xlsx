{-# LANGUAGE OverloadedStrings #-}
module Codec.Xlsx.Types.Internal.CfPair where

import           Text.XML.Cursor

import           Codec.Xlsx.Parser.Internal
import           Codec.Xlsx.Types.Common
import           Codec.Xlsx.Types.ConditionalFormatting
import           Codec.Xlsx.Writer.Internal


-- | Internal helper type for parsing "conditionalFormatting recods
-- TODO: pivot, extList
-- Implementing those will need this implementation to be changed
--
-- See 18.3.1.18 "conditionalFormatting (Conditional Formatting)" (p. 1610)
newtype CfPair = CfPair
    { unCfPair :: (SqRef, ConditionalFormatting)
    } deriving (Eq, Show)

instance FromCursor CfPair where
    fromCursor cur = do
        sqref <- fromAttribute "sqref" cur
        let cfRules = cur $/ element (n"cfRule") >=> fromCursor
        return $ CfPair (sqref, cfRules)

instance ToElement CfPair where
    toElement nm (CfPair (sqRef, cfRules)) =
        elementList nm [ "sqref" .= toAttrVal sqRef ]
                    (map (toElement "cfRule") cfRules)
