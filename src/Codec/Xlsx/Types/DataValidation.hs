{-# LANGUAGE OverloadedStrings #-}
{-# LANGUAGE RecordWildCards   #-}
{-# LANGUAGE TemplateHaskell   #-}
{-# LANGUAGE DeriveGeneric #-}
module Codec.Xlsx.Types.DataValidation where

import Control.Lens.TH (makeLenses)
import Control.Monad ((>=>))
import Data.Char (isSpace)
import Data.Default
import qualified Data.Map as M
import Data.Maybe (catMaybes)
import Data.Monoid ((<>))
import Data.Text (Text)
import qualified Data.Text as T
import GHC.Generics (Generic)
import Text.XML (Node(..), Element(..))
import Text.XML.Cursor (Cursor, ($/), element)

import Codec.Xlsx.Parser.Internal
import Codec.Xlsx.Types.Common
import Codec.Xlsx.Writer.Internal

-- See 18.18.20 "ST_DataValidationOperator (Data Validation Operator)" (p. 2439/2449)
data ValidationExpression
    = ValBetween Formula Formula    -- ^ "Between" operator
    | ValEqual Formula              -- ^ "Equal to" operator
    | ValGreaterThan Formula        -- ^ "Greater than" operator
    | ValGreaterThanOrEqual Formula -- ^ "Greater than or equal to" operator
    | ValLessThan Formula           -- ^ "Less than" operator
    | ValLessThanOrEqual Formula    -- ^ "Less than or equal to" operator
    | ValNotBetween Formula Formula -- ^ "Not between" operator
    | ValNotEqual Formula           -- ^ "Not equal to" operator
    deriving (Eq, Show, Generic)

-- See 18.18.21 "ST_DataValidationType (Data Validation Type)" (p. 2440/2450)
data ValidationType
    = ValidationTypeNone
    | ValidationTypeCustom     Formula
    | ValidationTypeDate       ValidationExpression
    | ValidationTypeDecimal    ValidationExpression
    | ValidationTypeList       [Text]
    | ValidationTypeTextLength ValidationExpression
    | ValidationTypeTime       ValidationExpression
    | ValidationTypeWhole      ValidationExpression
    deriving (Eq, Show, Generic)

-- See 18.18.18 "ST_DataValidationErrorStyle (Data Validation Error Styles)" (p. 2438/2448)
data ErrorStyle
    = ErrorStyleInformation
    | ErrorStyleStop
    | ErrorStyleWarning
    deriving (Eq, Show, Generic)

-- See 18.3.1.32 "dataValidation (Data Validation)" (p. 1614/1624)
data DataValidation = DataValidation
    { _dvAllowBlank       :: Bool
    , _dvError            :: Maybe Text
    , _dvErrorStyle       :: ErrorStyle
    , _dvErrorTitle       :: Maybe Text
    , _dvPrompt           :: Maybe Text
    , _dvPromptTitle      :: Maybe Text
    , _dvShowDropDown     :: Bool
    , _dvShowErrorMessage :: Bool
    , _dvShowInputMessage :: Bool
    , _dvValidationType   :: ValidationType
    } deriving (Eq, Show, Generic)

makeLenses ''DataValidation

instance Default DataValidation where
    def = DataValidation
      False Nothing ErrorStyleStop Nothing Nothing Nothing False False False ValidationTypeNone

{-------------------------------------------------------------------------------
  Parsing
-------------------------------------------------------------------------------}

instance FromAttrVal ErrorStyle where
    fromAttrVal "information" = readSuccess ErrorStyleInformation
    fromAttrVal "stop"        = readSuccess ErrorStyleStop
    fromAttrVal "warning"     = readSuccess ErrorStyleWarning
    fromAttrVal t             = invalidText "ErrorStyle" t

instance FromCursor DataValidation where
    fromCursor cur = do
        _dvAllowBlank       <- fromAttributeDef "allowBlank"       False          cur
        _dvError            <- maybeAttribute   "error"                           cur
        _dvErrorStyle       <- fromAttributeDef "errorStyle"       ErrorStyleStop cur
        _dvErrorTitle       <- maybeAttribute   "errorTitle"                      cur
        mop                 <- fromAttributeDef "operator"         "between"      cur
        _dvPrompt           <- maybeAttribute   "prompt"                          cur
        _dvPromptTitle      <- maybeAttribute   "promptTitle"                     cur
        _dvShowDropDown     <- fromAttributeDef "showDropDown"     False          cur
        _dvShowErrorMessage <- fromAttributeDef "showErrorMessage" False          cur
        _dvShowInputMessage <- fromAttributeDef "showInputMessage" False          cur
        mtype               <- fromAttributeDef "type"             "none"         cur
        _dvValidationType   <- readValidationType mop mtype                       cur
        return DataValidation{..}

readValidationType :: Text -> Text -> Cursor -> [ValidationType]
readValidationType _ "none"   _   = return ValidationTypeNone
readValidationType _ "custom" cur = do
    f <- fromCursor cur
    return $ ValidationTypeCustom f
readValidationType _ "list" cur = do
    f  <- cur $/ element (n_ "formula1") >=> fromCursor
    as <- readListFormula f
    return $ ValidationTypeList as
readValidationType op ty cur = do
    opExp <- readOpExpression2 op cur
    readValidationTypeOpExp ty opExp

readListFormula :: Formula -> [[Text]]
readListFormula (Formula f) = catMaybes [readQuotedList f]
  where
    readQuotedList t
        | Just t'  <- T.stripPrefix "\"" (T.dropAround isSpace t)
        , Just t'' <- T.stripSuffix "\"" t'
        = Just $ map (T.dropAround isSpace) $ T.splitOn "," t''
        | otherwise = Nothing
  -- This parser expects a comma-separated list surrounded by quotation marks.
  -- Spaces around the quotation marks and commas are removed, but inner spaces
  -- are kept.
  --
  -- The parser seems to be consistent with how Excel treats list formulas, but
  -- I wasn't able to find a specification of the format.

readOpExpression2 :: Text -> Cursor -> [ValidationExpression]
readOpExpression2 op cur
    | op `elem` ["between", "notBetween"] = do
        f1 <- cur $/ element (n_ "formula1") >=> fromCursor
        f2 <- cur $/ element (n_ "formula2") >=> fromCursor
        readValExpression op [f1,f2]
readOpExpression2 op cur = do
    f <- cur $/ element (n_ "formula1") >=> fromCursor
    readValExpression op [f]

readValidationTypeOpExp :: Text -> ValidationExpression -> [ValidationType]
readValidationTypeOpExp "date"       oe = [ValidationTypeDate       oe]
readValidationTypeOpExp "decimal"    oe = [ValidationTypeDecimal    oe]
readValidationTypeOpExp "textLength" oe = [ValidationTypeTextLength oe]
readValidationTypeOpExp "time"       oe = [ValidationTypeTime       oe]
readValidationTypeOpExp "whole"      oe = [ValidationTypeWhole      oe]
readValidationTypeOpExp _ _             = []

readValExpression :: Text -> [Formula] -> [ValidationExpression]
readValExpression "between" [f1, f2]       = [ValBetween f1 f2]
readValExpression "equal" [f]              = [ValEqual f]
readValExpression "greaterThan" [f]        = [ValGreaterThan f]
readValExpression "greaterThanOrEqual" [f] = [ValGreaterThanOrEqual f]
readValExpression "lessThan" [f]           = [ValLessThan f]
readValExpression "lessThanOrEqual" [f]    = [ValLessThanOrEqual f]
readValExpression "notBetween" [f1, f2]    = [ValNotBetween f1 f2]
readValExpression "notEqual" [f]           = [ValNotEqual f]
readValExpression _ _                      = []

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
            [ Just $ "allowBlank"       .=  _dvAllowBlank
            ,        "error"            .=? _dvError
            , Just $ "errorStyle"       .=  _dvErrorStyle
            ,        "errorTitle"       .=? _dvErrorTitle
            ,        "operator"         .=? op
            ,        "prompt"           .=? _dvPrompt
            ,        "promptTitle"      .=? _dvPromptTitle
            , Just $ "showDropDown"     .=  _dvShowDropDown
            , Just $ "showErrorMessage" .=  _dvShowErrorMessage
            , Just $ "showInputMessage" .=  _dvShowInputMessage
            , Just $ "type"             .=  _dvValidationType
            ]
        , elementNodes      = catMaybes
            [ fmap (NodeElement . toElement "formula1") f1
            , fmap (NodeElement . toElement "formula2") f2
            ]
        }
      where
        opExp (o,f1',f2') = (Just o, Just f1', f2')

        op    :: Maybe Text
        f1,f2 :: Maybe Formula
        (op,f1,f2) = case _dvValidationType of
          ValidationTypeNone         -> (Nothing, Nothing, Nothing)
          ValidationTypeCustom f     -> (Nothing, Just f, Nothing)
          ValidationTypeDate f       -> opExp $ viewValidationExpression f
          ValidationTypeDecimal f    -> opExp $ viewValidationExpression f
          ValidationTypeTextLength f -> opExp $ viewValidationExpression f
          ValidationTypeTime f       -> opExp $ viewValidationExpression f
          ValidationTypeWhole f      -> opExp $ viewValidationExpression f
          ValidationTypeList as      ->
            let f = Formula $ "\"" <> T.intercalate "," as <> "\""
            in  (Nothing, Just f, Nothing)

viewValidationExpression :: ValidationExpression -> (Text, Formula, Maybe Formula)
viewValidationExpression (ValBetween f1 f2)         = ("between",            f1, Just f2)
viewValidationExpression (ValEqual f)               = ("equal",              f,  Nothing)
viewValidationExpression (ValGreaterThan f)         = ("greaterThan",        f,  Nothing)
viewValidationExpression (ValGreaterThanOrEqual f)  = ("greaterThanOrEqual", f,  Nothing)
viewValidationExpression (ValLessThan f)            = ("lessThan",           f,  Nothing)
viewValidationExpression (ValLessThanOrEqual f)     = ("lessThanOrEqual",    f,  Nothing)
viewValidationExpression (ValNotBetween f1 f2)      = ("notBetween",         f1, Just f2)
viewValidationExpression (ValNotEqual f)            = ("notEqual",           f,  Nothing)

