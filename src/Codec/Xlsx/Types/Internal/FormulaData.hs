{-# LANGUAGE DeriveGeneric #-}
{-# LANGUAGE OverloadedStrings #-}
{-# LANGUAGE RecordWildCards #-}
module Codec.Xlsx.Types.Internal.FormulaData where

import Data.Monoid ((<>))
import Data.Text (Text)
import qualified Data.Text as T
import GHC.Generics (Generic)

import Codec.Xlsx.Parser.Internal
import Codec.Xlsx.Types.Cell
import Codec.Xlsx.Types.Common

data FormulaData = FormulaData
  { frmdFormula :: CellFormula
  , frmdShared :: Maybe (SharedFormulaIndex, SharedFormulaOptions)
  } deriving Generic

defaultFormulaType :: Text
defaultFormulaType = "normal"

instance FromXenoNode FormulaData where
  fromXenoNode n = do
    (bx, ca, t, mSi, mRef) <-
      parseAttributes n $
      (,,,,) <$> fromAttrDef "bx" False
             <*> fromAttrDef "ca" False
             <*> fromAttrDef "t" defaultFormulaType
             <*> maybeAttr "si"
             <*> maybeAttr "ref"
    (expr, shared) <-
      case t of
        d | d == defaultFormulaType -> do
            formula <- contentX n
            return (NormalFormula $ Formula formula, Nothing)
        "shared" -> do
          si <-
            maybe
              (Left "missing si attribute for shared formula")
              return
              mSi
          formula <- Formula <$> contentX n
          return
            ( SharedFormula si
            , mRef >>= \ref -> return (si, (SharedFormulaOptions ref formula)))
        unexpected -> Left $ "Unexpected formula type" <> T.pack (show unexpected)
    let f =
          CellFormula
          { _cellfAssignsToName = bx
          , _cellfCalculate = ca
          , _cellfExpression = expr
          }
    return $ FormulaData f shared
