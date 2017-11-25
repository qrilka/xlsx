{-# LANGUAGE OverloadedStrings #-}
{-# LANGUAGE DeriveGeneric #-}
module Codec.Xlsx.Types.Internal.DvPair where

import qualified Data.Map as M
import GHC.Generics (Generic)
import Text.XML (Element(..))

import Codec.Xlsx.Parser.Internal
import Codec.Xlsx.Types.Common
import Codec.Xlsx.Types.DataValidation
import Codec.Xlsx.Writer.Internal

-- | Internal helper type for parsing data validation records
--
-- See 18.3.1.32 "dataValidation (Data Validation)" (p. 1614/1624)
newtype DvPair = DvPair
    { unDvPair :: (SqRef, DataValidation)
    } deriving (Eq, Show, Generic)

instance FromCursor DvPair where
    fromCursor cur = do
        sqref <- fromAttribute "sqref" cur
        dv    <- fromCursor cur
        return $ DvPair (sqref, dv)

instance FromXenoNode DvPair where
  fromXenoNode root = do
    sqref <- parseAttributes root $ fromAttr "sqref"
    dv <- fromXenoNode root
    return $ DvPair (sqref, dv)

instance ToElement DvPair where
    toElement nm (DvPair (sqRef,dv)) = e
        {elementAttributes = M.insert "sqref" (toAttrVal sqRef) $ elementAttributes e}
      where
        e = toElement nm dv
