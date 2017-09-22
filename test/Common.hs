module Common
  ( parseBS
  , cursorFromElement
  ) where
import Data.ByteString.Lazy (ByteString)
import Text.XML
import Text.XML.Cursor

import Codec.Xlsx.Parser.Internal
import Codec.Xlsx.Writer.Internal

parseBS :: FromCursor a => ByteString -> [a]
parseBS = fromCursor . fromDocument . parseLBS_ def

cursorFromElement :: Element -> Cursor
cursorFromElement = fromNode . NodeElement . addNS mainNamespace Nothing
