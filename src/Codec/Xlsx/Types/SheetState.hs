{-# LANGUAGE OverloadedStrings #-}
{-# LANGUAGE DeriveGeneric #-}

module Codec.Xlsx.Types.SheetState where

import GHC.Generics (Generic)
import Control.DeepSeq (NFData)
import Codec.Xlsx.Parser.Internal
import Codec.Xlsx.Writer.Internal

type SheetVisibity = SheetState

-- | Sheet visibility state
-- cf. Ecma Office Open XML Part 1:
-- 18.18.68 ST_SheetState (Sheet Visibility Types)
-- * “visible”
--     Indicates the sheet is visible (default)
-- * “hidden”
--     Indicates the workbook window is hidden, but can be shown by the user via the user interface.
-- * “veryHidden”
--     Indicates the sheet is hidden and cannot be shown in the user interface (UI). This state is only available programmatically.
data SheetState =
  Visible -- ^ state="visible"
  | Hidden -- ^ state="hidden"
  | VeryHidden -- ^ state="veryHidden"
  deriving (Eq, Show, Generic)

instance NFData SheetState

instance FromAttrVal SheetState where
    fromAttrVal "visible"       = readSuccess Visible
    fromAttrVal "hidden"        = readSuccess Hidden
    fromAttrVal "veryHidden"    = readSuccess VeryHidden
    fromAttrVal t               = invalidText "SheetState" t

instance FromAttrBs SheetState where
    fromAttrBs "visible"    = return Visible
    fromAttrBs "hidden"     = return Hidden
    fromAttrBs "veryHidden" = return VeryHidden
    fromAttrBs t            = unexpectedAttrBs "SheetState" t

instance ToAttrVal SheetState where
    toAttrVal Visible = "visible"
    toAttrVal Hidden = "hidden"
    toAttrVal VeryHidden = "veryHidden"
