{-# LANGUAGE OverloadedStrings         #-}
{-# LANGUAGE RecordWildCards           #-}
{-# LANGUAGE TemplateHaskell           #-}
module Codec.Xlsx.Types.DataValidation where

import           Control.Lens.TH            (makeLenses)
import           Data.Default
import qualified Data.Map                   as M
import           Data.Maybe                 (catMaybes)
import           Data.Monoid                ((<>))
import           Data.Text                  (Text)
import qualified Data.Text                  as T
import           Text.XML

import           Codec.Xlsx.Types.Common
import           Codec.Xlsx.Types.ConditionalFormatting
import           Codec.Xlsx.Writer.Internal

-- See 18.18.21 "ST_DataValidationType (Data Validation Type)" (p. 2440/2450)
data ValidationType
    = ValidationTypeNone
    | ValidationTypeCustom     Formula
    | ValidationTypeDate       OperatorExpression
    | ValidationTypeDecimal    OperatorExpression
    | ValidationTypeList       [Text]
    | ValidationTypeTextLength OperatorExpression
    | ValidationTypeTime       OperatorExpression
    | ValidationTypeWhole      OperatorExpression
    deriving (Eq, Show)

-- See 18.18.18 "ST_DataValidationErrorStyle (Data Validation Error Styles)" (p. 2438/2448)
data ErrorStyle
    = ErrorStyleInformation
    | ErrorStyleStop
    | ErrorStyleWarning
    deriving (Eq, Show)

-- See 18.3.1.32 "dataValidation (Data Validation)" (p. 1614/1624)
data DataValidation = DataValidation
    { _dvAllowBlank       :: Maybe Bool
    , _dvError            :: Maybe Text
    , _dvErrorStyle       :: Maybe ErrorStyle
    , _dvErrorTitle       :: Maybe Text
    , _dvPrompt           :: Maybe Text
    , _dvPromptTitle      :: Maybe Text
    , _dvShowDropDown     :: Maybe Bool
    , _dvShowErrorMessage :: Maybe Bool
    , _dvShowInputMessage :: Maybe Bool
    , _dvSqref            :: SqRef
    , _dvValidationType   :: Maybe ValidationType
    } deriving (Eq, Show)

makeLenses ''DataValidation

instance Default DataValidation where
    def = DataValidation
      Nothing Nothing Nothing Nothing Nothing Nothing Nothing Nothing Nothing
      (SqRef []) Nothing

{-------------------------------------------------------------------------------
  Rendering
-------------------------------------------------------------------------------}

instance ToAttrVal ValidationType where
    toAttrVal ValidationTypeNone           = "none"
    toAttrVal (ValidationTypeCustom _)     = "custom"
    toAttrVal (ValidationTypeDate _)       = "date"
    toAttrVal (ValidationTypeDecimal _)    = "decimal"
    toAttrVal (ValidationTypeList _)       = "list"
    toAttrVal (ValidationTypeTextLength _) = "textLength"
    toAttrVal (ValidationTypeTime _)       = "time"
    toAttrVal (ValidationTypeWhole _)      = "whole"

instance ToAttrVal ErrorStyle where
    toAttrVal ErrorStyleInformation = "information"
    toAttrVal ErrorStyleStop        = "stop"
    toAttrVal ErrorStyleWarning     = "warning"

instance ToElement DataValidation where
    toElement nm DataValidation{..} = Element
        { elementName       = nm
        , elementAttributes = M.fromList . catMaybes $
            [        "allowBlank"       .=? _dvAllowBlank
            ,        "error"            .=? _dvError
            ,        "errorStyle"       .=? _dvErrorStyle
            ,        "errorTitle"       .=? _dvErrorTitle
            ,        "operator"         .=? op
            ,        "prompt"           .=? _dvPrompt
            ,        "promptTitle"      .=? _dvPromptTitle
            ,        "showDropDown"     .=? _dvShowDropDown
            ,        "showErrorMessage" .=? _dvShowErrorMessage
            ,        "showInputMessage" .=? _dvShowInputMessage
            , Just $ "sqref"            .=  _dvSqref
            ,        "type"             .=? _dvValidationType
            ]
        , elementNodes      = catMaybes
            [ fmap (NodeElement . toElement "formula1") f1
            , fmap (NodeElement . toElement "formula2") f2
            ]
        }
      where
        opExp (o,f1',f2') = (Just o, f1', f2')

        op    :: Maybe Text
        f1,f2 :: Maybe Formula
        (op,f1,f2) = case _dvValidationType of
            Nothing -> (Nothing, Nothing, Nothing)
            Just t  -> case t of
              ValidationTypeNone         -> (Nothing, Nothing, Nothing)
              ValidationTypeCustom f     -> (Nothing, Just f, Nothing)
              ValidationTypeDate f       -> opExp $ viewOperatorExpression f
              ValidationTypeDecimal f    -> opExp $ viewOperatorExpression f
              ValidationTypeTextLength f -> opExp $ viewOperatorExpression f
              ValidationTypeTime f       -> opExp $ viewOperatorExpression f
              ValidationTypeWhole f      -> opExp $ viewOperatorExpression f
              ValidationTypeList as      ->
                let f = Formula $ "\"" <> T.intercalate "," as <> "\""
                in  (Nothing, Just f, Nothing)
