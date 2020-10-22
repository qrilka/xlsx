{-# LANGUAGE CPP #-}
{-# LANGUAGE OverloadedStrings #-}
{-# LANGUAGE RecordWildCards #-}
{-# LANGUAGE TemplateHaskell #-}
{-# LANGUAGE DeriveGeneric #-}
module Codec.Xlsx.Types.Protection
  ( SheetProtection(..)
  , fullSheetProtection
  , noSheetProtection
  , LegacyPassword
  , legacyPassword
  -- * Lenses
  , sprLegacyPassword
  , sprSheet              
  , sprObjects            
  , sprScenarios          
  , sprFormatCells        
  , sprFormatColumns      
  , sprFormatRows         
  , sprInsertColumns      
  , sprInsertRows         
  , sprInsertHyperlinks   
  , sprDeleteColumns      
  , sprDeleteRows         
  , sprSelectLockedCells  
  , sprSort               
  , sprAutoFilter         
  , sprPivotTables        
  , sprSelectUnlockedCells
  ) where

import GHC.Generics (Generic)

import Control.Arrow (first)
#ifdef USE_MICROLENS
import Lens.Micro.TH (makeLenses)
#else
import Control.Lens (makeLenses)
#endif
import Control.DeepSeq (NFData)
import Data.Bits
import Data.Char
import Data.Maybe (catMaybes)
import Data.Text (Text)
import qualified Data.Text as T
import Data.Text.Lazy (toStrict)
import Data.Text.Lazy.Builder (toLazyText)
import Data.Text.Lazy.Builder.Int (hexadecimal)

import Codec.Xlsx.Parser.Internal
import Codec.Xlsx.Writer.Internal

newtype LegacyPassword =
  LegacyPassword Text
  deriving (Eq, Show, Generic)
instance NFData LegacyPassword

-- | Creates legacy @XOR@ hashed password.
--
-- /Note:/ The implementation is known to work only for ASCII symbols,
-- if you know how to encode properly others - an email or a PR will
-- be highly apperciated
--
-- See Part 4, 14.7.1 "Legacy Password Hash Algorithm" (p. 73) and
-- Part 4, 15.2.3 "Additional attributes for workbookProtection
-- element (Part 1, §18.2.29)" (p. 220) and Par 4, 15.3.1.6
-- "Additional attribute for sheetProtection element (Part 1,
-- §18.3.1.85)" (p. 229)
legacyPassword :: Text -> LegacyPassword
legacyPassword = LegacyPassword . hex . legacyHash . map ord . T.unpack
  where
    hex = toStrict . toLazyText . hexadecimal
    legacyHash bs =
      mutHash (foldr (\b hash -> b `xor` mutHash hash) 0 bs) `xor` (length bs) `xor`
      0xCE4B
    mutHash ph = ((ph `shiftR` 14) .&. 1) .|. ((ph `shiftL` 1) .&. 0x7fff)

-- | Sheet protection options to enforce and specify that it needs to
-- be protected
--
-- TODO: algorithms specified in the spec with hashes, salts and spin
-- counts
--
-- See 18.3.1.85 "sheetProtection (Sheet Protection Options)" (p. 1694)
data SheetProtection = SheetProtection
  { _sprLegacyPassword :: Maybe LegacyPassword
    -- ^ Specifies the legacy hash of the password required for editing
    -- this worksheet.
    --
    -- See Part 4, 15.3.1.6 "Additional attribute for sheetProtection
    -- element (Part 1, §18.3.1.85)" (p. 229)
  , _sprSheet :: Bool
    -- ^ the value of this attribute dictates whether the other
    -- attributes of 'SheetProtection' should be applied
  , _sprAutoFilter :: Bool
    -- ^ AutoFilters should not be allowed to operate when the sheet
    -- is protected
  , _sprDeleteColumns :: Bool
    -- ^ deleting columns should not be allowed when the sheet is
    -- protected
  , _sprDeleteRows :: Bool
    -- ^ deleting rows should not be allowed when the sheet is
    -- protected
  , _sprFormatCells :: Bool
    -- ^ formatting cells should not be allowed when the sheet is
    -- protected
  , _sprFormatColumns :: Bool
    -- ^ formatting columns should not be allowed when the sheet is
    -- protected
  , _sprFormatRows :: Bool
    -- ^ formatting rows should not be allowed when the sheet is
    -- protected
  , _sprInsertColumns :: Bool
    -- ^ inserting columns should not be allowed when the sheet is
    -- protected
  , _sprInsertHyperlinks :: Bool
    -- ^ inserting hyperlinks should not be allowed when the sheet is
    -- protected
  , _sprInsertRows :: Bool
    -- ^ inserting rows should not be allowed when the sheet is
    -- protected
  , _sprObjects :: Bool
    -- ^ editing of objects should not be allowed when the sheet is
    -- protected
  , _sprPivotTables :: Bool
    -- ^ PivotTables should not be allowed to operate when the sheet
    -- is protected
  , _sprScenarios :: Bool
    -- ^ Scenarios should not be edited when the sheet is protected
  , _sprSelectLockedCells :: Bool
    -- ^ selection of locked cells should not be allowed when the
    -- sheet is protected
  , _sprSelectUnlockedCells :: Bool
    -- ^ selection of unlocked cells should not be allowed when the
    -- sheet is protected
  , _sprSort :: Bool
    -- ^ sorting should not be allowed when the sheet is protected
  } deriving (Eq, Show, Generic)
instance NFData SheetProtection

makeLenses ''SheetProtection

{-------------------------------------------------------------------------------
  Base instances
-------------------------------------------------------------------------------}

-- | no sheet protection at all
noSheetProtection :: SheetProtection
noSheetProtection =
  SheetProtection
  { _sprLegacyPassword = Nothing
  , _sprSheet = False
  , _sprAutoFilter = False
  , _sprDeleteColumns = False
  , _sprDeleteRows = False
  , _sprFormatCells = False
  , _sprFormatColumns = False
  , _sprFormatRows = False
  , _sprInsertColumns = False
  , _sprInsertHyperlinks = False
  , _sprInsertRows = False
  , _sprObjects = False
  , _sprPivotTables = False
  , _sprScenarios = False
  , _sprSelectLockedCells = False
  , _sprSelectUnlockedCells = False
  , _sprSort = False
  }

-- | protection of all sheet features which could be protected
fullSheetProtection :: SheetProtection
fullSheetProtection =
  SheetProtection
  { _sprLegacyPassword = Nothing
  , _sprSheet = True
  , _sprAutoFilter = True
  , _sprDeleteColumns = True
  , _sprDeleteRows = True
  , _sprFormatCells = True
  , _sprFormatColumns = True
  , _sprFormatRows = True
  , _sprInsertColumns = True
  , _sprInsertHyperlinks = True
  , _sprInsertRows = True
  , _sprObjects = True
  , _sprPivotTables = True
  , _sprScenarios = True
  , _sprSelectLockedCells = True
  , _sprSelectUnlockedCells = True
  , _sprSort = True
  }

{-------------------------------------------------------------------------------
  Parsing
-------------------------------------------------------------------------------}

instance FromCursor SheetProtection where
  fromCursor cur = do
    _sprLegacyPassword <- maybeAttribute "password" cur
    _sprSheet <- fromAttributeDef "sheet" False cur
    _sprAutoFilter  <- fromAttributeDef "autoFilter" True cur
    _sprDeleteColumns  <- fromAttributeDef "deleteColumns" True cur
    _sprDeleteRows  <- fromAttributeDef "deleteRows" True cur
    _sprFormatCells  <- fromAttributeDef "formatCells" True cur
    _sprFormatColumns  <- fromAttributeDef "formatColumns" True cur
    _sprFormatRows  <- fromAttributeDef "formatRows" True cur
    _sprInsertColumns  <- fromAttributeDef "insertColumns" True cur
    _sprInsertHyperlinks  <- fromAttributeDef "insertHyperlinks" True cur
    _sprInsertRows  <- fromAttributeDef "insertRows" True cur
    _sprObjects  <- fromAttributeDef "objects" False cur
    _sprPivotTables  <- fromAttributeDef "pivotTables" True cur
    _sprScenarios  <- fromAttributeDef "scenarios" False cur
    _sprSelectLockedCells  <- fromAttributeDef "selectLockedCells" False cur
    _sprSelectUnlockedCells  <- fromAttributeDef "selectUnlockedCells" False cur
    _sprSort  <- fromAttributeDef "sort" True cur    
    return SheetProtection {..}

instance FromXenoNode SheetProtection where
  fromXenoNode root =
    parseAttributes root $ do
      _sprLegacyPassword <- maybeAttr "password"
      _sprSheet <- fromAttrDef "sheet" False
      _sprAutoFilter <- fromAttrDef "autoFilter" True
      _sprDeleteColumns <- fromAttrDef "deleteColumns" True
      _sprDeleteRows <- fromAttrDef "deleteRows" True
      _sprFormatCells <- fromAttrDef "formatCells" True
      _sprFormatColumns <- fromAttrDef "formatColumns" True
      _sprFormatRows <- fromAttrDef "formatRows" True
      _sprInsertColumns <- fromAttrDef "insertColumns" True
      _sprInsertHyperlinks <- fromAttrDef "insertHyperlinks" True
      _sprInsertRows <- fromAttrDef "insertRows" True
      _sprObjects <- fromAttrDef "objects" False
      _sprPivotTables <- fromAttrDef "pivotTables" True
      _sprScenarios <- fromAttrDef "scenarios" False
      _sprSelectLockedCells <- fromAttrDef "selectLockedCells" False
      _sprSelectUnlockedCells <- fromAttrDef "selectUnlockedCells" False
      _sprSort <- fromAttrDef "sort" True
      return SheetProtection {..}

instance FromAttrVal LegacyPassword where
  fromAttrVal = fmap (first LegacyPassword) . fromAttrVal

instance FromAttrBs LegacyPassword where
  fromAttrBs = fmap LegacyPassword . fromAttrBs

{-------------------------------------------------------------------------------
  Rendering
-------------------------------------------------------------------------------}

instance ToElement SheetProtection where
  toElement nm SheetProtection {..} =
    leafElement nm $
    catMaybes
      [ "password" .=? _sprLegacyPassword
      , "sheet" .=? justTrue _sprSheet
      , "autoFilter" .=? justFalse _sprAutoFilter
      , "deleteColumns" .=? justFalse _sprDeleteColumns
      , "deleteRows" .=? justFalse _sprDeleteRows
      , "formatCells" .=? justFalse _sprFormatCells
      , "formatColumns" .=? justFalse _sprFormatColumns
      , "formatRows" .=? justFalse _sprFormatRows
      , "insertColumns" .=? justFalse _sprInsertColumns
      , "insertHyperlinks" .=? justFalse _sprInsertHyperlinks
      , "insertRows" .=? justFalse _sprInsertRows
      , "objects" .=? justTrue _sprObjects
      , "pivotTables" .=? justFalse _sprPivotTables
      , "scenarios" .=? justTrue _sprScenarios
      , "selectLockedCells" .=? justTrue _sprSelectLockedCells
      , "selectUnlockedCells" .=? justTrue _sprSelectUnlockedCells
      , "sort" .=? justFalse _sprSort
      ]

instance ToAttrVal LegacyPassword where
  toAttrVal (LegacyPassword hash) = hash
