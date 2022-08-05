{-# LANGUAGE CPP #-}
{-# LANGUAGE DeriveGeneric #-}
{-# LANGUAGE MultiParamTypeClasses #-}
{-# LANGUAGE OverloadedStrings #-}
{-# LANGUAGE RecordWildCards #-}
{-# LANGUAGE TemplateHaskell #-}
module Codec.Xlsx.Types.DataValidation
  ( ValidationExpression(..)
    , ValidationType(..)
    , dvAllowBlank
    , dvError
    , dvErrorStyle
    , dvErrorTitle
    , dvPrompt
    , dvPromptTitle
    , dvShowDropDown
    , dvShowErrorMessage
    , dvShowInputMessage
    , dvValidationType
    , ErrorStyle(..)
    , DataValidation(..)
    , ListOrRangeExpression(..)
    , ValidationList
    , getPlainListValidator
    , getCellRangeValidator
    , readValidationType
    , readListFormulas
    , readOpExpression2
    , readValidationTypeOpExp
    , readValExpression
    , viewValidationExpression
  ) where

import Control.Applicative ((<|>))
import Control.DeepSeq (NFData)
#ifdef USE_MICROLENS
import Lens.Micro.TH (makeLenses)
#else
import Control.Lens.TH (makeLenses)
#endif
import Control.Monad ((>=>), guard)
import Data.ByteString (ByteString)
import Data.Char (isSpace)
import Data.Default
import qualified Data.Map as M
import Data.Maybe (catMaybes, maybeToList)
import Data.Monoid ((<>))
import Data.Text (Text)
import qualified Data.Text as T
import GHC.Generics (Generic)
import Text.XML (Element(..), Node(..))
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
instance NFData ValidationExpression

-- See 18.18.21 "ST_DataValidationType (Data Validation Type)" (p. 2440/2450)
data ValidationType
    = ValidationTypeNone
    | ValidationTypeCustom     Formula
    | ValidationTypeDate       ValidationExpression
    | ValidationTypeDecimal    ValidationExpression
    | ValidationTypeList       ListOrRangeExpression
    | ValidationTypeTextLength ValidationExpression
    | ValidationTypeTime       ValidationExpression
    | ValidationTypeWhole      ValidationExpression
    deriving (Eq, Show, Generic)
instance NFData ValidationType

type ValidationList = [Text]

data ListOrRangeExpression = ListExpression ValidationList | RangeExpression Range
    deriving (Eq, Show, Generic)
instance NFData ListOrRangeExpression


-- See 18.18.18 "ST_DataValidationErrorStyle (Data Validation Error Styles)" (p. 2438/2448)
data ErrorStyle
    = ErrorStyleInformation
    | ErrorStyleStop
    | ErrorStyleWarning
    deriving (Eq, Show, Generic)
instance NFData ErrorStyle

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
instance NFData DataValidation

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

instance FromAttrBs ErrorStyle where
    fromAttrBs "information" = return ErrorStyleInformation
    fromAttrBs "stop"        = return ErrorStyleStop
    fromAttrBs "warning"     = return ErrorStyleWarning
    fromAttrBs x             = unexpectedAttrBs "ErrorStyle" x

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

instance FromXenoNode DataValidation where
  fromXenoNode root = do
    (op, atype, genDV) <- parseAttributes root $ do
      _dvAllowBlank <- fromAttrDef "allowBlank" False
      _dvError <- maybeAttr "error"
      _dvErrorStyle <- fromAttrDef "errorStyle" ErrorStyleStop
      _dvErrorTitle <- maybeAttr "errorTitle"
      _dvPrompt <- maybeAttr "prompt"
      _dvPromptTitle <- maybeAttr "promptTitle"
      _dvShowDropDown <- fromAttrDef "showDropDown" False
      _dvShowErrorMessage <- fromAttrDef "showErrorMessage" False
      _dvShowInputMessage <- fromAttrDef "showInputMessage" False
      op <- fromAttrDef "operator" "between"
      typ <- fromAttrDef "type" "none"
      return (op, typ, \_dvValidationType -> DataValidation {..})
    valType <- parseValidationType op atype
    return $ genDV valType
    where
      parseValidationType :: ByteString -> ByteString -> Either Text ValidationType
      parseValidationType op atype =
        case atype of
          "none" -> return ValidationTypeNone
          "custom" ->
            ValidationTypeCustom <$> formula1
          "list" -> do
            f <- formula1
            case readListFormulas f of
              Nothing -> Left "validation of type \"list\" with empty formula list"
              Just fs -> return $ ValidationTypeList fs
          "date" ->
            ValidationTypeDate <$> readOpExpression op
          "decimal"    ->
            ValidationTypeDecimal <$> readOpExpression op
          "textLength" ->
            ValidationTypeTextLength <$> readOpExpression op
          "time"       ->
            ValidationTypeTime <$> readOpExpression op
          "whole"      ->
            ValidationTypeWhole <$> readOpExpression op
          unexpected ->
            Left $ "unexpected type of data validation " <> T.pack (show unexpected)
      readOpExpression "between" = uncurry ValBetween <$> formulaPair
      readOpExpression "notBetween" = uncurry ValNotBetween <$> formulaPair
      readOpExpression "equal" = ValEqual <$> formula1
      readOpExpression "greaterThan" = ValGreaterThan <$> formula1
      readOpExpression "greaterThanOrEqual" = ValGreaterThanOrEqual <$> formula1
      readOpExpression "lessThan" = ValLessThan <$> formula1
      readOpExpression "lessThanOrEqual" = ValLessThanOrEqual <$> formula1
      readOpExpression "notEqual" = ValNotEqual <$> formula1
      readOpExpression op = Left $ "data validation, unexpected operator " <> T.pack (show op)
      formula1 = collectChildren root $ fromChild "formula1"
      formulaPair =
        collectChildren root $ (,) <$> fromChild "formula1" <*> fromChild "formula2"

readValidationType :: Text -> Text -> Cursor -> [ValidationType]
readValidationType _ "none"   _   = return ValidationTypeNone
readValidationType _ "custom" cur = do
    f <- fromCursor cur
    return $ ValidationTypeCustom f
readValidationType _ "list" cur = do
    f  <- cur $/ element (n_ "formula1") >=> fromCursor
    as <- maybeToList $ readListFormulas f
    return $ ValidationTypeList as
readValidationType op ty cur = do
    opExp <- readOpExpression2 op cur
    readValidationTypeOpExp ty opExp

-- | Attempt to obtain a range expression from the list of ValidationTypeList
getCellRangeValidator :: ListOrRangeExpression -> Maybe Range
getCellRangeValidator (RangeExpression re) = Just re
getCellRangeValidator _ = Nothing

-- | Attempt to obtain a plain list from the list of ValidationTypeList
getPlainListValidator :: ListOrRangeExpression -> Maybe ValidationList
getPlainListValidator (ListExpression le) = Just le
getPlainListValidator _ = Nothing

readListFormulas :: Formula -> Maybe ListOrRangeExpression
readListFormulas (Formula f) = readQuotedList f <|> readUnquotedCellRange f
  where
    readQuotedList t
        | Just t'  <- T.stripPrefix "\"" (T.dropAround isSpace t)
        , Just t'' <- T.stripSuffix "\"" t'
        = Just . ListExpression $ map (T.dropAround isSpace) $ T.splitOn "," t''
        | otherwise = Nothing
    readUnquotedCellRange t =
      -- a CellRef expression of a range (this is not validated beyond the absence of quotes)
      -- note that the foreign sheet name can be 'single-quoted'
      let stripped = T.dropAround isSpace t
        in RangeExpression (CellRef stripped) <$ guard (not (T.null stripped))
  -- This parser expects a comma-separated list surrounded by quotation marks.
  -- Spaces around the quotation marks and commas are removed, but inner spaces
  -- are kept.
  --
  -- The parser seems to be consistent with how Excel treats list formulas, but
  -- I wasn't able to find a specification of the format.
  --
  -- Addendum: <dataValidation type="list" ...> undescriminately designates an actual list or a cell range.
  -- For a cell range validation, instead of a quoted list, it's an unquoted CellRef-like contents of the form:
  -- ActualSheetName!$C$2:$C$18

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
            let renderPlainList l =
                  let csvFy xs = T.intercalate "," xs
                      reQuote x = '"' `T.cons`  x `T.snoc` '"'
                    in reQuote (csvFy l)
                f = Formula $
                      case as of
                        RangeExpression re -> unCellRef re
                        ListExpression le -> renderPlainList le
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
