module Codec.Xlsx.Types.Common
       ( CellRef
       ) where

import Data.Text (Text)

-- | Excel cell reference (e.g. @E3@)
-- see 18.18.62 @ST_Ref@ (p. 2482)
type CellRef = Text
