{-# LANGUAGE OverloadedStrings #-}
module Codec.Xlsx.Parser.Internal
    ( n
    , parseSharedStrings
    ) where

import qualified Data.IntMap as IM
import Data.Text (Text)
import qualified Data.Text as T
import Data.XML.Types
import Text.XML.Cursor


-- | Add sml namespace to name
n :: Text -> Name
n x = Name
  { nameLocalName = x
  , nameNamespace = Just "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  , namePrefix = Nothing
  }

parseSharedStrings :: Cursor -> IM.IntMap Text
parseSharedStrings c = IM.fromAscList $ zip [0..] (c $/ element (n"si") >=> parseT)
    where
      -- it's  either <t> or <r>s with <t> inside
      parseT c' = [T.concat $ c' $| orSelf (child >=> (element (n"r"))) &/ element (n"t") &/ content]
