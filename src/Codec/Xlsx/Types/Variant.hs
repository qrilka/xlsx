{-# LANGUAGE OverloadedStrings #-}
{-# LANGUAGE DeriveGeneric #-}
module Codec.Xlsx.Types.Variant where

import Control.DeepSeq (NFData)
import Control.Monad.Fail (MonadFail)
import Data.ByteString (ByteString)
import Data.ByteString.Base64 as B64
import Data.Text (Text)
import qualified Data.Text as T
import qualified Data.Text.Encoding as T
import GHC.Generics (Generic)
import Text.XML
import Text.XML.Cursor

import Codec.Xlsx.Parser.Internal
import Codec.Xlsx.Writer.Internal

data Variant
    = VtBlob ByteString
    | VtBool Bool
    | VtDecimal Double
    | VtLpwstr Text
    | VtInt Int
    -- TODO: vt_vector, vt_array, vt_oblob, vt_empty, vt_null, vt_i1, vt_i2,
    -- vt_i4, vt_i8, vt_ui1, vt_ui2, vt_ui4, vt_ui8, vt_uint, vt_r4, vt_r8,
    -- vt_lpstr, vt_bstr, vt_date, vt_filetime, vt_cy, vt_error, vt_stream,
    -- vt_ostream, vt_storage, vt_ostorage, vt_vstream, vt_clsid
    deriving (Eq, Show, Generic)

instance NFData Variant

{-------------------------------------------------------------------------------
  Parsing
-------------------------------------------------------------------------------}

instance FromCursor Variant where
  fromCursor = variantFromNode . node

variantFromNode :: Node -> [Variant]
variantFromNode n@(NodeElement el) | elementName el == vt "lpwstr" =
                                         fromNode n $/ content &| VtLpwstr
                                   | elementName el == vt "bool" =
                                         fromNode n $/ content >=> fmap VtBool . boolean
                                   | elementName el == vt "int" =
                                         fromNode n $/ content >=> fmap VtInt . decimal
                                   | elementName el == vt "decimal" =
                                         fromNode n $/ content >=> fmap VtDecimal . rational
                                   | elementName el == vt "blob" =
                                         fromNode n $/ content >=> fmap VtBlob . decodeBase64 . killWhitespace
variantFromNode  _ = fail "no matching nodes"

killWhitespace :: Text -> Text
killWhitespace = T.filter (/=' ')

decodeBase64 :: MonadFail m => Text -> m ByteString
decodeBase64 t = case B64.decode (T.encodeUtf8 t) of
  Right bs -> return bs
  Left err -> fail $ "invalid base64 value: " ++ err

-- | Add doc props variant types namespace to name
vt :: Text -> Name
vt x = Name
  { nameLocalName = x
  , nameNamespace = Just docPropsVtNs
  , namePrefix = Nothing
  }

docPropsVtNs :: Text
docPropsVtNs = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"

{-------------------------------------------------------------------------------
  Rendering
-------------------------------------------------------------------------------}

variantToElement :: Variant -> Element
variantToElement (VtLpwstr t)  = elementContent (vt"lpwstr")  t
variantToElement (VtBlob bs)   = elementContent (vt"blob")    (T.decodeLatin1 $ B64.encode bs)
variantToElement (VtBool b)    = elementContent (vt"bool")    (txtb b)
variantToElement (VtDecimal d) = elementContent (vt"decimal") (txtd d)
variantToElement (VtInt i)     = elementContent (vt"int")     (txti i)
