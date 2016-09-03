{-# LANGUAGE FlexibleInstances   #-}
{-# LANGUAGE OverloadedStrings   #-}
{-# LANGUAGE RecordWildCards     #-}
{-# LANGUAGE ScopedTypeVariables #-}
module Codec.Xlsx.Writer.Internal (
    -- * Rendering documents
    ToDocument(..)
  , documentFromElement
  , documentFromNsElement
    -- * Rendering elements
  , ToElement(..)
  , countedElementList
  , elementList
  , elementListSimple
  , nonEmptyElListSimple
  , leafElement
  , emptyElement
  , elementContent
  , elementContentPreserved
  , elementValue
    -- * Rendering attributes
  , ToAttrVal(..)
  , (.=)
  , (.=?)
  , setAttr
    -- * Dealing with namespaces
  , addNS
  , mainNamespace
    -- * Misc
  , txti
  , txtb
  , txtd
  , justNonDef
  , justTrue
  , justFalse
  ) where

import           Data.Text                        (Text)
import qualified Data.Text                        as T
import           Data.Text.Lazy                   (toStrict)
import           Data.Text.Lazy.Builder           (toLazyText)
import           Data.Text.Lazy.Builder.Int
import           Data.Text.Lazy.Builder.RealFloat

import qualified Data.Map                         as Map
import           Data.String                      (fromString)
import           Text.XML

{-------------------------------------------------------------------------------
  Rendering documents
-------------------------------------------------------------------------------}

class ToDocument a where
  toDocument :: a -> Document

documentFromElement :: Text -> Element -> Document
documentFromElement comment e = documentFromNsElement comment mainNamespace e

documentFromNsElement :: Text -> Text -> Element -> Document
documentFromNsElement comment ns e = Document {
      documentRoot     = addNS ns e
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

countedElementList :: Name -> [Element] -> Element
countedElementList nm as = elementList nm [ "count" .= length as ] as

elementList :: Name -> [(Name, Text)] -> [Element] -> Element
elementList nm attrs els = Element {
      elementName       = nm
    , elementNodes      = map NodeElement els
    , elementAttributes = Map.fromList attrs
    }

elementListSimple :: Name -> [Element] -> Element
elementListSimple nm els = elementList nm [] els

nonEmptyElListSimple :: Name -> [Element] -> Maybe Element
nonEmptyElListSimple _  []  = Nothing
nonEmptyElListSimple nm els = Just $ elementListSimple nm els

leafElement :: Name -> [(Name, Text)] -> Element
leafElement nm attrs = elementList nm attrs []

emptyElement :: Name -> Element
emptyElement nm = elementList nm [] []

elementContent0 :: Name -> [(Name, Text)] -> Text -> Element
elementContent0 nm attrs txt = Element {
      elementName       = nm
    , elementAttributes = Map.fromList attrs
    , elementNodes      = [NodeContent txt]
    }

elementContent :: Name -> Text -> Element
elementContent nm txt = elementContent0 nm [] txt

elementContentPreserved :: Name -> Text -> Element
elementContentPreserved nm txt = elementContent0 nm [ preserveSpace ] txt
  where
    preserveSpace = (
        Name { nameLocalName = "space"
             , nameNamespace = Just "http://www.w3.org/XML/1998/namespace"
             , namePrefix    = Nothing
             }
      , "preserve"
      )

{-------------------------------------------------------------------------------
  Rendering attributes
-------------------------------------------------------------------------------}

class ToAttrVal a where
  toAttrVal :: a -> Text

instance ToAttrVal Text    where toAttrVal = id
instance ToAttrVal String  where toAttrVal = fromString
instance ToAttrVal Int     where toAttrVal = txti
instance ToAttrVal Integer where toAttrVal = txti
instance ToAttrVal Double  where toAttrVal = fromString . show

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

setAttr :: ToAttrVal a => Name -> a -> Element -> Element
setAttr nm a el@Element{..} = el{ elementAttributes = attrs' }
  where
    attrs' = Map.insert nm (toAttrVal a) elementAttributes

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
    goName n@Name{..} =
      case nameNamespace of
        Just _  -> n -- If a namespace was explicitly set, leave it
        Nothing -> Name{
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


txtd :: Double -> Text
txtd = toStrict . toLazyText . realFloat

txtb :: Bool -> Text
txtb = T.toLower . T.pack . show

txti :: (Integral a) => a -> Text
txti = toStrict . toLazyText . decimal

justNonDef :: (Eq a) => a -> a -> Maybe a
justNonDef defVal a | a == defVal = Nothing
                    | otherwise   = Just a

justFalse :: Bool -> Maybe Bool
justFalse = justNonDef True

justTrue :: Bool -> Maybe Bool
justTrue = justNonDef False
