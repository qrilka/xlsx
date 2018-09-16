{-# LANGUAGE OverloadedStrings #-}
{-# LANGUAGE TemplateHaskell #-}
{-# LANGUAGE DeriveGeneric #-}
module Codec.Xlsx.Types.Style (
    Format(..)
  , NumberFormat(..)
    -- * Format
  , formatAlignment
  , formatBorder
  , formatFill
  , formatFont
  , formatNumberFormat
  , formatProtection
  , formatPivotButton
  , formatQuotePrefix
    -- * NumberFormat
  , fmtDecimals
  , fmtDecimalsZeroes
  ) where

import Control.DeepSeq (NFData)
import Control.Lens
import Data.Default
import Data.Monoid ((<>))
import qualified Data.Text as T
import GHC.Generics (Generic)

import Codec.Xlsx.Types.StyleSheet

-- | This type gives a high-level version of representation of number format
-- used in 'Format'.
data NumberFormat
    = StdNumberFormat ImpliedNumberFormat
    | UserNumberFormat FormatCode
    deriving (Eq, Ord, Show, Generic)

instance NFData NumberFormat

-- | Basic number format with predefined number of decimals
-- as format code of number format in xlsx should be less than 255 characters
-- number of decimals shouldn't be more than 253
fmtDecimals :: Int -> NumberFormat
fmtDecimals k = UserNumberFormat $ "0." <> T.replicate k "#"

-- | Basic number format with predefined number of decimals.
-- Works like 'fmtDecimals' with the only difference that extra zeroes are
-- displayed when number of digits after the point is less than the number
-- of digits specified in the format
fmtDecimalsZeroes :: Int -> NumberFormat
fmtDecimalsZeroes k = UserNumberFormat $ "0." <> T.replicate k "0"

-- | Formatting options used to format cells
--
-- TODOs:
--
-- * Add a number format ('_cellXfApplyNumberFormat', '_cellXfNumFmtId')
-- * Add references to the named style sheets ('_cellXfId')
data Format = Format
    { _formatAlignment    :: Maybe Alignment
    , _formatBorder       :: Maybe Border
    , _formatFill         :: Maybe Fill
    , _formatFont         :: Maybe Font
    , _formatNumberFormat :: Maybe NumberFormat
    , _formatProtection   :: Maybe Protection
    , _formatPivotButton  :: Maybe Bool
    , _formatQuotePrefix  :: Maybe Bool
    } deriving (Eq, Show, Generic)

makeLenses ''Format

instance Default Format where
  def = Format
        { _formatAlignment    = Nothing
        , _formatBorder       = Nothing
        , _formatFill         = Nothing
        , _formatFont         = Nothing
        , _formatNumberFormat = Nothing
        , _formatProtection   = Nothing
        , _formatPivotButton  = Nothing
        , _formatQuotePrefix  = Nothing
        }
