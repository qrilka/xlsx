{-# LANGUAGE FlexibleInstances     #-}
{-# LANGUAGE MultiParamTypeClasses #-}
{-# OPTIONS_GHC -fno-warn-orphans #-}

module Common
  ( parseBS
  , cursorFromElement
  ) where
import Data.ByteString.Lazy (ByteString)
import Text.XML
import Text.XML.Cursor

import Codec.Xlsx.Parser.Internal
import Codec.Xlsx.Types (SheetState (..))
import Codec.Xlsx.Writer.Internal

import Test.SmallCheck.Series (Serial)

parseBS :: FromCursor a => ByteString -> [a]
parseBS = fromCursor . fromDocument . parseLBS_ def

cursorFromElement :: Element -> Cursor
cursorFromElement = fromNode . NodeElement . addNS mainNamespace Nothing

instance Monad m  => Serial m SheetState
