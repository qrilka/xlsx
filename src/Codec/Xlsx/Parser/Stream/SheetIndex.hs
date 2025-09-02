
{-# LANGUAGE DerivingStrategies #-}
{-# LANGUAGE GeneralisedNewtypeDeriving #-}

module Codec.Xlsx.Parser.Stream.SheetIndex (SheetIndex (..), makeIndex) where

import Control.DeepSeq

-- | datatype representing a sheet index, looking it up by name
--   can be done with 'makeIndexFromName', which is the preferred approach.
--   although 'makeIndex' is available in case it's already known.
newtype SheetIndex = MkSheetIndex Int
 deriving newtype (NFData, Show)

-- | This does *no* checking if the index exists or not.
--   you could have index out of bounds issues because of this.
makeIndex :: Int -> SheetIndex
makeIndex = MkSheetIndex

{-# DEPRECATED SheetIndex, MkSheetIndex, makeIndex "SheetIndex will not be supported in future, see issue #193. Use SheetIdentifier instead" #-}
