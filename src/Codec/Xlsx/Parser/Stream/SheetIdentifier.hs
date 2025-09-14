{-# LANGUAGE PatternSynonyms #-}

module Codec.Xlsx.Parser.Stream.SheetIdentifier
  ( SheetIdentifier (SheetIdentifier, siRefId, siSheetId)
  , unsafeGetSheetIdentifier
  ) where

import Codec.Xlsx.Types.RefId ( RefId )

-- | Opaque identifier for a sheet.
data SheetIdentifier = MkSheetIdentifier RefId Int
  deriving (Eq, Ord)

pattern SheetIdentifier :: RefId -> Int -> SheetIdentifier
pattern SheetIdentifier { siRefId, siSheetId } <- MkSheetIdentifier siRefId siSheetId

{-# COMPLETE SheetIdentifier #-}

unsafeGetSheetIdentifier :: RefId -> Int -> SheetIdentifier
unsafeGetSheetIdentifier = MkSheetIdentifier
