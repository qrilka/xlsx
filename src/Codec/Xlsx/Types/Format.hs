{-# LANGUAGE TemplateHaskell #-}
{-# LANGUAGE DeriveGeneric #-}
module Codec.Xlsx.Types.Format (
    Format(..)
  , formatAlignment
  , formatBorder
  , formatFill
  , formatFont
  , formatNumberFormat
  , formatProtection
  , formatPivotButton
  , formatQuotePrefix
  ) where

import Control.Lens
import Data.Default
import GHC.Generics (Generic)

import Codec.Xlsx.Types.StyleSheet

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
