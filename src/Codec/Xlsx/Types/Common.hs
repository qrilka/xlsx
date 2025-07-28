{-# LANGUAGE OverloadedStrings #-}
{-# LANGUAGE DeriveGeneric #-}
{-# LANGUAGE TupleSections #-}
{-# LANGUAGE ScopedTypeVariables #-}
{-# LANGUAGE TemplateHaskell #-}
{-# LANGUAGE CPP #-}
{-# LANGUAGE RankNTypes #-}
{-# LANGUAGE GeneralizedNewtypeDeriving #-}
{-# LANGUAGE ViewPatterns #-}
{-# LANGUAGE PatternSynonyms #-}

module Codec.Xlsx.Types.Common
  ( CellRef(..)
  , RowCoord(..)
  , ColumnCoord(..)
  , CellCoord
  , RangeCoord
  , mapBoth
  , col2coord
  , coord2col
  , row2coord
  , coord2row
  , singleCellRef
  , singleCellRef'
  , fromSingleCellRef
  , fromSingleCellRef'
  , fromSingleCellRefNoting
  , escapeRefSheetName
  , unEscapeRefSheetName
  , mkForeignSingleCellRef
  , fromForeignSingleCellRef
  , Range
  , mkRange
  , mkRange'
  , mkForeignRange
  , fromRange
  , fromRange'
  , fromForeignRange
  , SqRef(..)
  , XlsxText(..)
  , xlsxTextToCellValue
  , Formula(..)
  , CellValue(..)
  , pattern CellDouble
  , ErrorType(..)
  , DateBase(..)
  , dateFromNumber
  , dateToNumber
  , int2col
  , col2int
  , columnIndexToText
  , textToColumnIndex
  -- ** prisms
  , _XlsxText
  , _XlsxRichText
  , _CellText
  , _CellDecimal
  , _CellBool
  , _CellRich
  , _CellError
  , RowIndex(..)
  , ColumnIndex(..)
  ) where

import GHC.Generics (Generic)

import Control.Applicative (liftA2)
import Control.Arrow
import Control.DeepSeq (NFData)
import Control.Monad (forM, guard)
import Data.Bifunctor (bimap)
import qualified Data.ByteString as BS
import Data.Char
import Data.Maybe (isJust, fromMaybe)
import Data.Function ((&))
import Data.Ix (inRange)
import qualified Data.Map as Map
import Data.Scientific (Scientific,toRealFloat,fromFloatDigits)
import Data.Text (Text)
import qualified Data.Text as T
import qualified Data.Text.Encoding as T
import Data.Time.Calendar (Day, addDays, diffDays, fromGregorian)
import Data.Time.Clock (UTCTime(UTCTime), picosecondsToDiffTime)
import Safe
import Text.XML
import Text.XML.Cursor

import Codec.Xlsx.Parser.Internal
import Codec.Xlsx.Types.RichText
import Codec.Xlsx.Writer.Internal
#ifdef USE_MICROLENS
import Lens.Micro
import Lens.Micro.Internal
import Lens.Micro.GHC ()
import Data.Profunctor.Choice
import Data.Profunctor(dimap)
#else
import Control.Lens(makePrisms)
#endif

newtype RowIndex = RowIndex {unRowIndex :: Int}
  deriving (Eq, Ord, Show, Read, Generic, Num, Real, Enum, Integral)
newtype ColumnIndex = ColumnIndex {unColumnIndex :: Int}
  deriving (Eq, Ord, Show, Read, Generic, Num, Real, Enum, Integral)
instance NFData RowIndex
instance NFData ColumnIndex

instance ToAttrVal RowIndex where
  toAttrVal = toAttrVal . unRowIndex

{-# DEPRECATED int2col
    "this function will be removed in an upcoming release, use columnIndexToText instead." #-}
int2col :: ColumnIndex -> Text
int2col = columnIndexToText

{-# DEPRECATED col2int
    "this function will be removed in an upcoming release, use textToColumnIndex instead." #-}
col2int :: Text -> ColumnIndex
col2int = textToColumnIndex

-- | convert column number (starting from 1) to its textual form (e.g. 3 -> \"C\")
columnIndexToText :: ColumnIndex -> Text
columnIndexToText = T.pack . reverse . map int2let . base26 . unColumnIndex
    where
        int2let 0 = 'Z'
        int2let x = chr $ (x - 1) + ord 'A'
        base26  0 = []
        base26  i = let i' = (i `mod` 26)
                        i'' = if i' == 0 then 26 else i'
                    in seq i' (i' : base26 ((i - i'') `div` 26))

rowIndexToText :: RowIndex -> Text
rowIndexToText = T.pack . show . unRowIndex

-- | reverse of 'columnIndexToText'
textToColumnIndex :: Text -> ColumnIndex
textToColumnIndex = ColumnIndex . T.foldl' (\i c -> i * 26 + let2int c) 0
    where
        let2int c = 1 + ord c - ord 'A'

textToRowIndex :: Text -> RowIndex
textToRowIndex = RowIndex . read . T.unpack

-- | Excel cell or cell range reference (e.g. @E3@), possibly absolute.
-- See 18.18.62 @ST_Ref@ (p. 2482)
--
-- Note: The @ST_Ref@ type can point to another sheet (supported)
-- or a sheet in another workbook (separate .xlsx file, not implemented).
newtype CellRef = CellRef
  { unCellRef :: Text
  } deriving (Eq, Ord, Show, Generic)
instance NFData CellRef

-- | A helper type for coordinates to carry the intent of them being relative or absolute (preceded by '$'):
--
-- > singleCellRefRaw' (RowRel 5, ColumnAbs 1) == "$A5"
data RowCoord
  = RowAbs !RowIndex
  | RowRel !RowIndex
  deriving (Eq, Ord, Show, Read, Generic)
instance NFData RowCoord

data ColumnCoord
  = ColumnAbs !ColumnIndex
  | ColumnRel !ColumnIndex
  deriving (Eq, Ord, Show, Read, Generic)
instance NFData ColumnCoord

type CellCoord = (RowCoord, ColumnCoord)

type RangeCoord = (CellCoord, CellCoord)

mkColumnCoord :: Bool -> ColumnIndex -> ColumnCoord
mkColumnCoord isAbs = if isAbs then ColumnAbs else ColumnRel

mkRowCoord :: Bool -> RowIndex -> RowCoord
mkRowCoord isAbs = if isAbs then RowAbs else RowRel

coord2col :: ColumnCoord -> Text
coord2col (ColumnAbs c) = "$" <> coord2col (ColumnRel c)
coord2col (ColumnRel c) = columnIndexToText c

col2coord :: Text -> ColumnCoord
col2coord t =
  let t' = T.stripPrefix "$" t
    in mkColumnCoord (isJust t') (textToColumnIndex (fromMaybe t t'))

coord2row :: RowCoord -> Text
coord2row (RowAbs c) = "$" <> coord2row (RowRel c)
coord2row (RowRel c) = rowIndexToText c

row2coord :: Text -> RowCoord
row2coord t =
  let t' = T.stripPrefix "$" t
    in mkRowCoord (isJust t') (textToRowIndex (fromMaybe t t'))

-- | Unwrap a Coord into an abstract Int coordinate
unRowCoord :: RowCoord -> RowIndex
unRowCoord (RowAbs i) = i
unRowCoord (RowRel i) = i

-- | Unwrap a Coord into an abstract Int coordinate
unColumnCoord :: ColumnCoord -> ColumnIndex
unColumnCoord (ColumnAbs i) = i
unColumnCoord (ColumnRel i) = i

-- | Helper function to apply the same transformation to both members of a tuple
--
-- > mapBoth Abs (1, 2) == (Abs 1, Abs 2s)
mapBoth :: (a -> b) -> (a, a) -> (b, b)
mapBoth f = bimap f f

-- | Render position in @(row, col)@ format to an Excel reference.
--
-- > singleCellRef (RowIndex 2, ColumnIndex 4) == CellRef "D2"
singleCellRef :: (RowIndex, ColumnIndex) -> CellRef
singleCellRef = CellRef . singleCellRefRaw

-- | Allow specifying whether a coordinate parameter is relative or absolute.
--
-- > singleCellRef' (Rel 5, Abs 1) == CellRef "$A5"
singleCellRef' :: CellCoord -> CellRef
singleCellRef' = CellRef . singleCellRefRaw'

singleCellRefRaw :: (RowIndex, ColumnIndex) -> Text
singleCellRefRaw (row, col) = T.concat [columnIndexToText col, rowIndexToText row]

singleCellRefRaw' :: CellCoord -> Text
singleCellRefRaw' (row, col) =
    coord2col col <> coord2row row

-- | Converse function to 'singleCellRef'
-- Ignores a potential foreign sheet prefix.
fromSingleCellRef :: CellRef -> Maybe (RowIndex, ColumnIndex)
fromSingleCellRef = fromSingleCellRefRaw . unCellRef

-- | Converse function to 'singleCellRef\''
-- Ignores a potential foreign sheet prefix.
fromSingleCellRef' :: CellRef -> Maybe CellCoord
fromSingleCellRef' = fromSingleCellRefRaw' . unCellRef

fromSingleCellRefRaw :: Text -> Maybe (RowIndex, ColumnIndex)
fromSingleCellRefRaw =
  fmap (first unRowCoord . second unColumnCoord) . fromSingleCellRefRaw'

fromSingleCellRefRaw' :: Text -> Maybe CellCoord
fromSingleCellRefRaw' t' = ignoreRefSheetName t' >>= \t -> do
    let (isColAbsolute, remT) =
          T.stripPrefix "$" t
          & \remT' -> (isJust remT', fromMaybe t remT')
    let (colT, rowExpr) = T.span (inRange ('A', 'Z')) remT
    let (isRowAbsolute, rowT) =
          T.stripPrefix "$" rowExpr
          & \rowT' -> (isJust rowT', fromMaybe rowExpr rowT')
    guard $ not (T.null colT) && not (T.null rowT) && T.all isDigit rowT
    row <- decimal rowT
    return $
      bimap
      (mkRowCoord isRowAbsolute)
      (mkColumnCoord isColAbsolute)
      (row, textToColumnIndex colT)

-- | Converse function to 'singleCellRef' expecting valid reference and failig with
-- a standard error message like /"Bad cell reference 'XXX'"/
fromSingleCellRefNoting :: CellRef -> (RowIndex, ColumnIndex)
fromSingleCellRefNoting ref = fromJustNote errMsg $ fromSingleCellRefRaw txt
  where
    txt = unCellRef ref
    errMsg = "Bad cell reference '" ++ T.unpack txt ++ "'"

-- | Frame and escape the referenced sheet name in single quotes (apostrophe).
--
-- Sheet name in ST_Ref can be single-quoted when it contains non-alphanum class, non-ASCII range characters.
-- Intermediate squote characters are escaped in a doubled fashion:
-- "My ' Sheet" -> 'My '' Sheet'
escapeRefSheetName :: Text -> Text
escapeRefSheetName sheetName =
   T.concat ["'", escape sheetName, "'"]
  where
    escape sn = T.splitOn "'" sn & T.intercalate "''"

-- | Unframe and unescape the referenced sheet name.
unEscapeRefSheetName :: Text -> Text
unEscapeRefSheetName = unescape . unFrame
      where
        unescape  = T.intercalate "'" . T.splitOn "''"
        unFrame sn = fromMaybe sn $ T.stripPrefix "'" sn >>= T.stripSuffix "'"

ignoreRefSheetName :: Text -> Maybe Text
ignoreRefSheetName t =
  case T.split (== '!') t of
    [_, r] -> Just r
    [r] -> Just r
    _ -> Nothing

-- | Render a single cell existing in another worksheet.
-- This function always renders the sheet name single-quoted regardless the presence of spaces.
-- A sheet name shouldn't contain @"[]*:?/\"@ chars and apostrophe @"'"@ should not happen at extremities.
--
-- > mkForeignRange "MyOtherSheet" (Rel 2, Rel 4) (Abs 6, Abs 8) == "'MyOtherSheet'!D2:$H$6"
mkForeignSingleCellRef :: Text -> CellCoord -> CellRef
mkForeignSingleCellRef sheetName coord =
    let cr = singleCellRefRaw' coord
      in CellRef $ T.concat [escapeRefSheetName sheetName, "!", cr]

-- | Converse function to 'mkForeignSingleCellRef'.
-- The provided CellRef must be a foreign range.
fromForeignSingleCellRef :: CellRef -> Maybe (Text, CellCoord)
fromForeignSingleCellRef r =
    case T.split (== '!') (unCellRef r) of
      [sheetName, ref] -> (unEscapeRefSheetName sheetName,) <$> fromSingleCellRefRaw' ref
      _ -> Nothing

-- | Excel range (e.g. @D13:H14@), actually store as as 'CellRef' in
-- xlsx
type Range = CellRef

-- | Render range
--
-- > mkRange (RowIndex 2, ColumnIndex 4) (RowIndex 6, ColumnIndex 8) == CellRef "D2:H6"
mkRange :: (RowIndex, ColumnIndex) -> (RowIndex, ColumnIndex) -> Range
mkRange fr to = CellRef $ T.concat [singleCellRefRaw fr, ":", singleCellRefRaw to]

-- | Render range with possibly absolute coordinates
--
-- > mkRange' (Abs 2, Abs 4) (6, 8) == CellRef "$D$2:H6"
mkRange' :: (RowCoord,ColumnCoord) -> (RowCoord,ColumnCoord) -> Range
mkRange' fr to =
  CellRef $ T.concat [singleCellRefRaw' fr, ":", singleCellRefRaw' to]

-- | Render a cell range existing in another worksheet.
-- This function always renders the sheet name single-quoted regardless the presence of spaces.
-- A sheet name shouldn't contain @"[]*:?/\"@ chars and apostrophe @"'"@ should not happen at extremities.
--
-- > mkForeignRange "MyOtherSheet" (Rel 2, Rel 4) (Abs 6, Abs 8) == "'MyOtherSheet'!D2:$H$6"
mkForeignRange :: Text -> CellCoord -> CellCoord -> Range
mkForeignRange sheetName fr to =
    case mkRange' fr to of
      CellRef cr -> CellRef $ T.concat [escapeRefSheetName sheetName, "!", cr]

-- | Converse function to 'mkRange' ignoring absolute coordinates.
-- Ignores a potential foreign sheet prefix.
fromRange :: Range -> Maybe ((RowIndex, ColumnIndex), (RowIndex, ColumnIndex))
fromRange r =
  mapBoth (first unRowCoord . second unColumnCoord) <$> fromRange' r

-- | Converse function to 'mkRange\'' to handle possibly absolute coordinates.
-- Ignores a potential foreign sheet prefix.
fromRange' :: Range -> Maybe RangeCoord
fromRange' t' = parseRange =<< ignoreRefSheetName (unCellRef t')
  where
    parseRange t =
      case T.split (== ':') t of
        [from, to] -> liftA2 (,) (fromSingleCellRefRaw' from) (fromSingleCellRefRaw' to)
        _ -> Nothing

-- | Converse function to 'mkForeignRange'.
-- The provided Range must be a foreign range.
fromForeignRange :: Range -> Maybe (Text, RangeCoord)
fromForeignRange r =
    case T.split (== '!') (unCellRef r) of
      [sheetName, ref] -> (unEscapeRefSheetName sheetName,) <$> fromRange' (CellRef ref)
      _ -> Nothing

-- | A sequence of cell references
--
-- See 18.18.76 "ST_Sqref (Reference Sequence)" (p.2488)
newtype SqRef = SqRef [CellRef]
    deriving (Eq, Ord, Show, Generic)

instance NFData SqRef

-- | Common type containing either simple string or rich formatted text.
-- Used in @si@, @comment@ and @is@ elements
--
-- E.g. @si@ spec says: "If the string is just a simple string with formatting applied
-- at the cell level, then the String Item (si) should contain a single text
-- element used to express the string. However, if the string in the cell is
-- more complex - i.e., has formatting applied at the character level - then the
-- string item shall consist of multiple rich text runs which collectively are
-- used to express the string.". So we have either a single "Text" field, or
-- else a list of "RichTextRun"s, each of which is some "Text" with layout
-- properties.
--
-- TODO: Currently we do not support @phoneticPr@ (Phonetic Properties, 18.4.3,
-- p. 1723) or @rPh@ (Phonetic Run, 18.4.6, p. 1725).
--
-- Section 18.4.8, "si (String Item)" (p. 1725)
--
-- See @CT_Rst@, p. 3903
data XlsxText = XlsxText Text
              | XlsxRichText [RichTextRun]
              deriving (Eq, Ord, Show, Generic)

instance NFData XlsxText

xlsxTextToCellValue :: XlsxText -> CellValue
xlsxTextToCellValue (XlsxText txt) = CellText txt
xlsxTextToCellValue (XlsxRichText rich) = CellRich rich

-- | A formula
--
-- See 18.18.35 "ST_Formula (Formula)" (p. 2457)
newtype Formula = Formula {unFormula :: Text}
    deriving (Eq, Ord, Show, Generic)

instance NFData Formula

-- | Cell values include text, numbers and booleans,
-- standard includes date format also but actually dates
-- are represented by numbers with a date format assigned
-- to a cell containing it
-- Specification (ECMA-376):
-- - 18.3.1.4 c (Cell)
-- - 18.18.11 ST_CellType (Cell Type)
data CellValue
  = CellText Text
  | CellDecimal Scientific
  | CellBool Bool
  | CellRich [RichTextRun]
  | CellError ErrorType
  deriving (Eq, Ord, Show, Generic)
{-# COMPLETE CellText, CellDecimal, CellBool, CellRich, CellError #-}

viewCellDouble :: CellValue -> Maybe Double
viewCellDouble (CellDecimal s) = Just (toRealFloat s)
viewCellDouble _ = Nothing 

-- view pattern, since 'CellDecimal' has replaced the old constructor.
pattern CellDouble :: Double -> CellValue
pattern CellDouble b <- (viewCellDouble -> Just b)
{-# COMPLETE CellText, CellDouble, CellBool, CellRich, CellError #-}

instance NFData CellValue

-- | The evaluation of an expression can result in an error having one
-- of a number of error values.
--
-- See Annex L, L.2.16.8 "Error values" (p. 4764)
data ErrorType
  = ErrorDiv0
  -- ^ @#DIV/0!@ - Intended to indicate when any number, including
  -- zero, is divided by zero.
  | ErrorNA
  -- ^ @#N/A@ - Intended to indicate when a designated value is not
  -- available. For example, some functions, such as @SUMX2MY2@,
  -- perform a series of operations on corresponding elements in two
  -- arrays. If those arrays do not have the same number of elements,
  -- then for some elements in the longer array, there are no
  -- corresponding elements in the shorter one; that is, one or more
  -- values in the shorter array are not available. This error value
  -- can be produced by calling the function @NA@.
  | ErrorName
  -- ^ @#NAME?@ - Intended to indicate when what looks like a name is
  -- used, but no such name has been defined. For example, @XYZ/3@,
  -- where @XYZ@ is not a defined name. @Total is & A10@, where
  -- neither @Total@ nor @is@ is a defined name. Presumably, @"Total
  -- is " & A10@ was intended. @SUM(A1C10)@, where the range @A1:C10@
  -- was intended.
  | ErrorNull
  -- ^ @#NULL!@ - Intended to indicate when two areas are required to
  -- intersect, but do not. For example, In the case of @SUM(B1 C1)@,
  -- the space between @B1@ and @C1@ is treated as the binary
  -- intersection operator, when a comma was intended.
  | ErrorNum
  -- ^ @#NUM!@ - Intended to indicate when an argument to a function
  -- has a compatible type, but has a value that is outside the domain
  -- over which that function is defined. (This is known as a domain
  -- error.) For example, Certain calls to @ASIN@, @ATANH@, @FACT@,
  -- and @SQRT@ might result in domain errors. Intended to indicate
  -- that the result of a function cannot be represented in a value of
  -- the specified type, typically due to extreme magnitude. (This is
  -- known as a range error.) For example, @FACT(1000)@ might result
  -- in a range error.
  | ErrorRef
  -- ^ @#REF!@ - Intended to indicate when a cell reference is
  -- invalid. For example, If a formula contains a reference to a
  -- cell, and then the row or column containing that cell is deleted,
  -- a @#REF!@ error results. If a worksheet does not support 20,001
  -- columns, @OFFSET(A1,0,20000)@ results in a @#REF!@ error.
  | ErrorValue
  -- ^ @#VALUE!@ - Intended to indicate when an incompatible type
  -- argument is passed to a function, or an incompatible type operand
  -- is used with an operator. For example, In the case of a function
  -- argument, a number was expected, but text was provided. In the
  -- case of @1+"ABC"@, the binary addition operator is not defined for
  -- text.
  deriving (Eq, Ord, Show, Generic)

instance NFData ErrorType

-- | Specifies date base used for conversion of serial values to and
-- from datetime values
--
-- See Annex L, L.2.16.9.1 "Date Conversion for Serial Values" (p. 4765)
data DateBase
  = DateBase1900
  -- ^ 1900 date base system, the lower limit is January 1, -9999
  -- 00:00:00, which has serial value -4346018. The upper-limit is
  -- December 31, 9999, 23:59:59, which has serial value
  -- 2,958,465.9999884. The base date for this date base system is
  -- December 30, 1899, which has a serial value of 0.
  | DateBase1904
  -- ^ 1904 backward compatibility date-base system, the lower limit
  -- is January 1, 1904, 00:00:00, which has serial value 0. The upper
  -- limit is December 31, 9999, 23:59:59, which has serial value
  -- 2,957,003.9999884. The base date for this date base system is
  -- January 1, 1904, which has a serial value of 0.
  deriving (Eq, Show, Generic)
instance NFData DateBase

baseDate :: DateBase -> Day
baseDate DateBase1900 = fromGregorian 1899 12 30
baseDate DateBase1904 = fromGregorian 1904 1 1

-- | Converts serial value into datetime according to the specified
-- date base. Because Excel treats 1900 as a leap year even though it isn't,
-- this function converts any numbers that represent some time in /1900-02-29/
-- in Excel to `UTCTime` /1900-03-01 00:00/.
-- See https://docs.microsoft.com/en-gb/office/troubleshoot/excel/wrongly-assumes-1900-is-leap-year for details.
--
-- > show (dateFromNumber DateBase1900 42929.75) == "2017-07-13 18:00:00 UTC"
-- > show (dateFromNumber DateBase1900 60) == "1900-03-01 00:00:00 UTC"
-- > show (dateFromNumber DateBase1900 61) == "1900-03-01 00:00:00 UTC"
dateFromNumber :: forall t. RealFrac t => DateBase -> t -> UTCTime
dateFromNumber b d
  -- 60 is Excel's 2020-02-29 00:00 and 61 is Excel's 2020-03-01
  | b == DateBase1900 && d < 60            = getUTCTime (d + 1)
  | b == DateBase1900 && d >= 60 && d < 61 = getUTCTime (61 :: t)
  | otherwise                              = getUTCTime d
  where
    getUTCTime n =
      let
        (numberOfDays, fractionOfOneDay) = properFraction n
        day = addDays numberOfDays $ baseDate b
        diffTime = picosecondsToDiffTime (round (fractionOfOneDay * 24*60*60*1E12))
      in
        UTCTime day diffTime

-- | Converts datetime into serial value.
-- Because Excel treats 1900 as a leap year even though it isn't,
-- the numbers that represent times in /1900-02-29/ in Excel, in the range /[60, 61[/,
-- are never generated by this function for `DateBase1900`. This means that
-- under those conditions this is not an inverse of `dateFromNumber`.
-- See https://docs.microsoft.com/en-gb/office/troubleshoot/excel/wrongly-assumes-1900-is-leap-year for details.
dateToNumber :: Fractional a => DateBase -> UTCTime -> a
dateToNumber b (UTCTime day diffTime) = numberOfDays + fractionOfOneDay
  where
    numberOfDays = fromIntegral (diffDays excel1900CorrectedDay $ baseDate b)
    fractionOfOneDay = realToFrac diffTime / (24 * 60 * 60)
    marchFirst1900              = fromGregorian 1900 3 1
    excel1900CorrectedDay = if day < marchFirst1900
      then addDays (-1) day
      else day

{-------------------------------------------------------------------------------
  Parsing
-------------------------------------------------------------------------------}

-- | See @CT_Rst@, p. 3903
instance FromCursor XlsxText where
  fromCursor cur = do
    let
      ts = cur $/ element (n_ "t") >=> contentOrEmpty
      rs = cur $/ element (n_ "r") >=> fromCursor
    case (ts,rs) of
      ([t], []) ->
        return $ XlsxText t
      ([], _:_) ->
        return $ XlsxRichText rs
      _ ->
        fail "invalid item"

instance FromXenoNode XlsxText where
  fromXenoNode root = do
    (mCh, rs) <-
      collectChildren root $ (,) <$> maybeChild "t" <*> fromChildList "r"
    mT <- mapM contentX mCh
    case mT of
      Just t -> return $ XlsxText t
      Nothing ->
        case rs of
          [] -> Left $ "missing rich text subelements"
          _ -> return $ XlsxRichText rs

instance FromAttrVal CellRef where
  fromAttrVal = fmap (first CellRef) . fromAttrVal

instance FromAttrBs CellRef where
  -- we presume that cell references contain only latin letters,
  -- numbers and colon
  fromAttrBs = pure . CellRef . T.decodeLatin1

instance FromAttrVal SqRef where
  fromAttrVal t = do
    rs <- mapM (fmap fst . fromAttrVal) $ T.split (== ' ') t
    readSuccess $ SqRef rs

instance FromAttrBs SqRef where
  fromAttrBs bs = do
    -- split on space
    rs <- forM  (BS.split 32 bs) fromAttrBs
    return $ SqRef rs

-- | See @ST_Formula@, p. 3873
instance FromCursor Formula where
    fromCursor cur = [Formula . T.concat $ cur $/ content]

instance FromXenoNode Formula where
  fromXenoNode = fmap Formula . contentX

instance FromAttrVal Formula where
  fromAttrVal t = readSuccess $ Formula t

instance FromAttrBs Formula where
  fromAttrBs = fmap Formula . fromAttrBs

instance FromAttrVal ErrorType where
  fromAttrVal "#DIV/0!" = readSuccess ErrorDiv0
  fromAttrVal "#N/A" = readSuccess ErrorNA
  fromAttrVal "#NAME?" = readSuccess ErrorName
  fromAttrVal "#NULL!" = readSuccess ErrorNull
  fromAttrVal "#NUM!" = readSuccess ErrorNum
  fromAttrVal "#REF!" = readSuccess ErrorRef
  fromAttrVal "#VALUE!" = readSuccess ErrorValue
  fromAttrVal t = invalidText "ErrorType" t

instance FromAttrBs ErrorType where
  fromAttrBs "#DIV/0!" = return ErrorDiv0
  fromAttrBs "#N/A" = return ErrorNA
  fromAttrBs "#NAME?" = return ErrorName
  fromAttrBs "#NULL!" = return ErrorNull
  fromAttrBs "#NUM!" = return ErrorNum
  fromAttrBs "#REF!" = return ErrorRef
  fromAttrBs "#VALUE!" = return ErrorValue
  fromAttrBs x = unexpectedAttrBs "ErrorType" x

{-------------------------------------------------------------------------------
  Rendering
-------------------------------------------------------------------------------}

-- | See @CT_Rst@, p. 3903
instance ToElement XlsxText where
  toElement nm si = Element {
      elementName       = nm
    , elementAttributes = Map.empty
    , elementNodes      = map NodeElement $
        case si of
          XlsxText text     -> [elementContent "t" text]
          XlsxRichText rich -> map (toElement "r") rich
    }

instance ToAttrVal CellRef where
  toAttrVal = toAttrVal . unCellRef

-- See 18.18.76, "ST_Sqref (Reference Sequence)", p. 2488.
instance ToAttrVal SqRef where
  toAttrVal (SqRef refs) = T.intercalate " " $ map toAttrVal refs

-- | See @ST_Formula@, p. 3873
instance ToElement Formula where
    toElement nm (Formula txt) = elementContent nm txt

instance ToAttrVal ErrorType where
  toAttrVal ErrorDiv0 = "#DIV/0!"
  toAttrVal ErrorNA = "#N/A"
  toAttrVal ErrorName = "#NAME?"
  toAttrVal ErrorNull = "#NULL!"
  toAttrVal ErrorNum = "#NUM!"
  toAttrVal ErrorRef = "#REF!"
  toAttrVal ErrorValue = "#VALUE!"

#ifdef USE_MICROLENS
-- Since micro-lens denies the existence of prisms,
-- I pasted the splice that's generated from makePrisms,
-- then I copied over the definitions from Control.Lens for the prism
-- function as well.
-- Essentially this is doing the template haskell by hand.
type Prism s t a b = forall p f. (Choice p, Applicative f) => p a (f b) -> p s (f t)
type Prism' s a = Prism s s a a

prism :: (b -> t) -> (s -> Either t a) -> Prism s t a b
prism bt seta = dimap seta (either pure (fmap bt)) . right'

_CellText :: Prism' CellValue Text
_CellText
  = (prism (\ x1_a1ZQv -> CellText x1_a1ZQv))
      (\ x_a1ZQw
         -> case x_a1ZQw of
              CellText y1_a1ZQx -> Right y1_a1ZQx
              _ -> Left x_a1ZQw)
{-# INLINE _CellText #-}
_CellDecimal :: Prism' CellValue Scientific
_CellDecimal
  = (prism (\ x1_a1ZQy -> CellDecimal x1_a1ZQy))
      (\ x_a1ZQz
         -> case x_a1ZQz of
              CellDecimal y1_a1ZQA -> Right y1_a1ZQA
              _ -> Left x_a1ZQz)
{-# INLINE _CellDecimal #-}
_CellDouble :: Prism' CellValue Double
_CellDouble
  = (prism (\ x1_a1ZQy -> CellDecimal (fromFloatDigits x1_a1ZQy)))
      (\ x_a1ZQz
         -> case x_a1ZQz of
              CellDouble y1_a1ZQA -> Right y1_a1ZQA
              _ -> Left x_a1ZQz)
{-# INLINE _CellDouble #-}
_CellBool :: Prism' CellValue Bool
_CellBool
  = (prism (\ x1_a1ZQB -> CellBool x1_a1ZQB))
      (\ x_a1ZQC
         -> case x_a1ZQC of
              CellBool y1_a1ZQD -> Right y1_a1ZQD
              _ -> Left x_a1ZQC)
{-# INLINE _CellBool #-}
_CellRich :: Prism' CellValue [RichTextRun]
_CellRich
  = (prism (\ x1_a1ZQE -> CellRich x1_a1ZQE))
      (\ x_a1ZQF
         -> case x_a1ZQF of
              CellRich y1_a1ZQG -> Right y1_a1ZQG
              _ -> Left x_a1ZQF)
{-# INLINE _CellRich #-}
_CellError :: Prism' CellValue ErrorType
_CellError
  = (prism (\ x1_a1ZQH -> CellError x1_a1ZQH))
      (\ x_a1ZQI
         -> case x_a1ZQI of
              CellError y1_a1ZQJ -> Right y1_a1ZQJ
              _ -> Left x_a1ZQI)
{-# INLINE _CellError #-}

_XlsxText :: Prism' XlsxText Text
_XlsxText
  = (prism (\ x1_a1ZzU -> XlsxText x1_a1ZzU))
      (\ x_a1ZzV
         -> case x_a1ZzV of
              XlsxText y1_a1ZzW -> Right y1_a1ZzW
              _ -> Left x_a1ZzV)
{-# INLINE _XlsxText #-}
_XlsxRichText :: Prism' XlsxText [RichTextRun]
_XlsxRichText
  = (prism (\ x1_a1ZzX -> XlsxRichText x1_a1ZzX))
      (\ x_a1ZzY
         -> case x_a1ZzY of
              XlsxRichText y1_a1ZzZ -> Right y1_a1ZzZ
              _ -> Left x_a1ZzY)
{-# INLINE _XlsxRichText #-}

#else
makePrisms ''XlsxText
makePrisms ''CellValue
#endif
