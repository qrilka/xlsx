{-# LANGUAGE OverloadedStrings #-}
{-# LANGUAGE DeriveGeneric #-}
{-# LANGUAGE ScopedTypeVariables #-}
module Codec.Xlsx.Types.Common
  ( CellRef(..)
  , singleCellRef
  , fromSingleCellRef
  , fromSingleCellRefNoting
  , Range
  , mkRange
  , fromRange
  , SqRef(..)
  , XlsxText(..)
  , xlsxTextToCellValue
  , Formula(..)
  , CellValue(..)
  , ErrorType(..)
  , DateBase(..)
  , dateFromNumber
  , dateToNumber
  , int2col
  , col2int
  ) where

import GHC.Generics (Generic)

import Control.Arrow
import Control.DeepSeq (NFData)
import Control.Monad (forM, guard)
import qualified Data.ByteString as BS
import Data.Char
import Data.Ix (inRange)
import qualified Data.Map as Map
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

-- | convert column number (starting from 1) to its textual form (e.g. 3 -> \"C\")
int2col :: Int -> Text
int2col = T.pack . reverse . map int2let . base26
    where
        int2let 0 = 'Z'
        int2let x = chr $ (x - 1) + ord 'A'
        base26  0 = []
        base26  i = let i' = (i `mod` 26)
                        i'' = if i' == 0 then 26 else i'
                    in seq i' (i' : base26 ((i - i'') `div` 26))

-- | reverse to 'int2col'
col2int :: Text -> Int
col2int = T.foldl' (\i c -> i * 26 + let2int c) 0
    where
        let2int c = 1 + ord c - ord 'A'

-- | Excel cell or cell range reference (e.g. @E3@)
-- See 18.18.62 @ST_Ref@ (p. 2482)
newtype CellRef = CellRef
  { unCellRef :: Text
  } deriving (Eq, Ord, Show, Generic)

instance NFData CellRef

-- | Render position in @(row, col)@ format to an Excel reference.
--
-- > mkCellRef (2, 4) == "D2"
singleCellRef :: (Int, Int) -> CellRef
singleCellRef = CellRef . singleCellRefRaw

singleCellRefRaw :: (Int, Int) -> Text
singleCellRefRaw (row, col) = T.concat [int2col col, T.pack (show row)]

-- | reverse to 'mkCellRef'
fromSingleCellRef :: CellRef -> Maybe (Int, Int)
fromSingleCellRef = fromSingleCellRefRaw . unCellRef

fromSingleCellRefRaw :: Text -> Maybe (Int, Int)
fromSingleCellRefRaw t = do
  let (colT, rowT) = T.span (inRange ('A', 'Z')) t
  guard $ not (T.null colT) && not (T.null rowT) && T.all isDigit rowT
  row <- decimal rowT
  return (row, col2int colT)

-- | reverse to 'mkCellRef' expecting valid reference and failig with
-- a standard error message like /"Bad cell reference 'XXX'"/
fromSingleCellRefNoting :: CellRef -> (Int, Int)
fromSingleCellRefNoting ref = fromJustNote errMsg $ fromSingleCellRefRaw txt
  where
    txt = unCellRef ref
    errMsg = "Bad cell reference '" ++ T.unpack txt ++ "'"

-- | Excel range (e.g. @D13:H14@), actually store as as 'CellRef' in
-- xlsx
type Range = CellRef

-- | Render range
--
-- > mkRange (2, 4) (6, 8) == "D2:H6"
mkRange :: (Int, Int) -> (Int, Int) -> Range
mkRange fr to = CellRef $ T.concat [singleCellRefRaw fr, T.pack ":", singleCellRefRaw to]

-- | reverse to 'mkRange'
fromRange :: Range -> Maybe ((Int, Int), (Int, Int))
fromRange r =
  case T.split (== ':') (unCellRef r) of
    [from, to] -> (,) <$> fromSingleCellRefRaw from <*> fromSingleCellRefRaw to
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
data CellValue
  = CellText Text
  | CellDouble Double
  | CellBool Bool
  | CellRich [RichTextRun]
  | CellError ErrorType
  deriving (Eq, Ord, Show, Generic)

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
-- this function converts any numbers that represent some time in 1900-02-29
-- in Excel to @UTCTime@ 1900-03-01 00:00.
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

-- | Converts datetime into serial value
dateToNumber :: Fractional a => DateBase -> UTCTime -> a
dateToNumber b (UTCTime day diffTime) = numberOfDays + fractionOfOneDay
  where
    numberOfDays = fromIntegral (diffDays excel1900CorrectedDay $ baseDate b)
    fractionOfOneDay = realToFrac diffTime / (24 * 60 * 60)
    -- For historical reasons, Excel thinks 1900 is a leap year (https://docs.microsoft.com/en-gb/office/troubleshoot/excel/wrongly-assumes-1900-is-leap-year).
    -- So dates strictly before March 1st 1900 need to be corrected.
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
