
module Codec.Xlsx.Parser.Stream.SheetInfo where

import Data.Text (Text)
import Codec.Xlsx.Types.RefId (RefId)
import Codec.Xlsx.Parser.Stream.SheetIdentifier
    (SheetIdentifier, unsafeGetSheetIdentifier)
import Codec.Xlsx.Types (SheetState)

-- | Represents sheets from the workbook.xml file. E.g.
-- <sheet name="Data" sheetId="1" state="hidden" r:id="rId2" /
data SheetInfo = MkSheetInfo
  { sheetInfoName    :: Text,
    -- | The r:id attribute value
    sheetInfoRelId   :: RefId,
    -- | The sheetId attribute value
    sheetInfoSheetId :: Int,
    -- | The index into the sheets in the workbook
    sheetInfoSheetIndex :: Int,
    -- | The sheet visibility state
    sheetInfoState    :: SheetState
  } deriving (Show, Eq)

{-# DEPRECATED sheetInfoRelId, sheetInfoSheetId "this field will be removed in future, see issue #193" #-} 

getSheetIdentifier :: SheetInfo -> SheetIdentifier
getSheetIdentifier sheetInfo = unsafeGetSheetIdentifier (sheetInfoRelId sheetInfo) (sheetInfoSheetId sheetInfo)
