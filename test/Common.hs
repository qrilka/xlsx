module Common
  ( parseBS
  ) where
import Data.ByteString.Lazy (ByteString)
import Text.XML
import Text.XML.Cursor

import Codec.Xlsx.Parser.Internal

parseBS :: FromCursor a => ByteString -> [a]
parseBS = fromCursor . fromDocument . parseLBS_ def
