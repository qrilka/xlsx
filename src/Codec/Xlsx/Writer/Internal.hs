{-# LANGUAGE FlexibleInstances   #-}
{-# LANGUAGE OverloadedStrings   #-}
{-# LANGUAGE RecordWildCards     #-}
{-# LANGUAGE ScopedTypeVariables #-}
{-# OPTIONS_GHC -Wall #-}
module Codec.Xlsx.Writer.Internal (
    -- * Rendering documents
    ToDocument(..)
  , documentFromElement
    -- * Rendering elements
  , ToElement(..)
  , elementList
  , elementValue
    -- * Rendering attributes
  , ToAttrVal(..)
  , (.=)
  , (.=?)
    -- * Dealing with namespaces
  , addNS
  , mainNamespace
  ) where

import Data.Text (Text)
import Data.String (fromString)
import Text.XML
import qualified Data.Map as Map

{-------------------------------------------------------------------------------
  Rendering documents
-------------------------------------------------------------------------------}

class ToDocument a where
  toDocument :: a -> Document

documentFromElement :: Text -> Element -> Document
documentFromElement comment e = Document {
      documentRoot     = addNS mainNamespace e
    , documentEpilogue = []
    , documentPrologue = Prologue {
          prologueBefore  = [MiscComment comment]
        , prologueDoctype = Nothing
        , prologueAfter   = []
        }
    }

{-------------------------------------------------------------------------------
  Rendering elements
-------------------------------------------------------------------------------}

class ToElement a where
  toElement :: Name -> a -> Element

elementList :: Name -> [Element] -> Element
elementList nm as = Element {
      elementName       = nm
    , elementNodes      = map NodeElement as
    , elementAttributes = Map.fromList [ "count" .= length as ]
    }

{-------------------------------------------------------------------------------
  Rendering attributes
-------------------------------------------------------------------------------}

class ToAttrVal a where
  toAttrVal :: a -> Text

instance ToAttrVal Text   where toAttrVal = id
instance ToAttrVal String where toAttrVal = fromString
instance ToAttrVal Int    where toAttrVal = fromString . show
instance ToAttrVal Double where toAttrVal = fromString . show

instance ToAttrVal Bool where
  toAttrVal True  = "true"
  toAttrVal False = "false"

elementValue :: ToAttrVal a => Name -> a -> Element
elementValue nm a = Element {
      elementName       = nm
    , elementAttributes = Map.fromList [ "val" .= a ]
    , elementNodes      = []
    }

(.=) :: ToAttrVal a => Name -> a -> (Name, Text)
nm .= a = (nm, toAttrVal a)

(.=?) :: ToAttrVal a => Name -> Maybe a -> Maybe (Name, Text)
_  .=? Nothing  = Nothing
nm .=? (Just a) = Just (nm .= a)

{-------------------------------------------------------------------------------
  Dealing with namespaces
-------------------------------------------------------------------------------}

-- | Set the namespace for the entire document
--
-- This follows the same policy that the rest of the xlsx package uses.
addNS :: Text -> Element -> Element
addNS ns Element{..} = Element{
      elementName       = goName elementName
    , elementAttributes = elementAttributes
    , elementNodes      = map goNode elementNodes
    }
  where
    goName :: Name -> Name
    goName Name{..} = Name{
        nameLocalName = nameLocalName
      , nameNamespace = Just ns
      , namePrefix    = Nothing
      }

    goNode :: Node -> Node
    goNode (NodeElement e) = NodeElement $ addNS ns e
    goNode n               = n

-- | The main namespace for Excel
mainNamespace :: Text
mainNamespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
