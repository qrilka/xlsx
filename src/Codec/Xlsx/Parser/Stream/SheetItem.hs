{-# LANGUAGE TemplateHaskell            #-}
{-# LANGUAGE CPP            #-}
{-# LANGUAGE DerivingStrategies            #-}
{-# LANGUAGE DeriveGeneric              #-}
{-# LANGUAGE DeriveAnyClass              #-}

module Codec.Xlsx.Parser.Stream.SheetItem where

import Codec.Xlsx.Parser.Stream.Row (Row)

#ifdef USE_MICROLENS
import Lens.Micro
import Lens.Micro.GHC ()
import Lens.Micro.Mtl
import Lens.Micro.Platform
import Lens.Micro.TH
#else
import Control.Lens (makeLenses)
#endif
import Control.DeepSeq (NFData)
import GHC.Generics (Generic)


-- | Sheet item
--
-- The current sheet at a time, every sheet is constructed of these items.
data SheetItem = MkSheetItem
  { _si_sheet_index :: !Int       -- ^ The sheet number
  , _si_row         :: Row
  } deriving stock (Generic, Show)
    deriving anyclass NFData

{-# WARNING SheetItem, MkSheetItem "SheetItem will not be supported in future, see issue #193. Use Row directly instead" #-}
makeLenses 'MkSheetItem
{-# DEPRECATED si_sheet_index, si_row "SheetItem will not be supported in future, see issue #193" #-}