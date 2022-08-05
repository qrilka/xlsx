{-# LANGUAGE OverloadedStrings #-}
{-# LANGUAGE DeriveGeneric #-}
{-# LANGUAGE TupleSections #-}
{-# LANGUAGE ViewPatterns #-}

module Codec.Xlsx.Types.Common
  ( CellRef(..)
  , Coord(..)
  , unCoord
  , both
  , singleCellRef
  , singleCellRef'
  , fromSingleCellRef
  , fromSingleCellRef'
  , fromSingleCellRefNoting
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
  , ErrorType(..)
  , DateBase(..)
  , dateFromNumber
  , dateToNumber
  , int2col
  , col2int
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

-- | A helper type for coordinates to carry the intent of them being relative or absolute (preceded by '$'):
--
-- > singleCellRefRaw' (Rel 5, Abs 1) == "$A5"
data Coord = Abs !Int | Rel !Int
  deriving (Eq, Ord, Show, Read, Generic)
instance NFData Coord

-- | Unwrap a Coord into an abstract Int coordinate
unCoord :: Coord -> Int
unCoord (Abs i) = i
unCoord (Rel i) = i

-- | Helper function to apply the same transformation to both members of a tuple
--
-- > both Abs (1, 2) == (Abs 1, Abs 2s)
both :: (a -> b) -> (a, a) -> (b, b)
both f = bimap f f

-- | Render position in @(row, col)@ format to an Excel reference.
--
-- > singleCellRef (2, 4) == CellRef "D2"
singleCellRef :: (Int, Int) -> CellRef
singleCellRef = CellRef . singleCellRefRaw

-- | Allow specifying whether a coordinate parameter is relative or absolute.
--
-- > singleCellRef' (Rel 5, Abs 1) == CellRef "$A5"
singleCellRef' :: (Coord, Coord) -> CellRef
singleCellRef' = CellRef . singleCellRefRaw'

singleCellRefRaw :: (Int, Int) -> Text
singleCellRefRaw = singleCellRefRaw' . both Rel

singleCellRefRaw' :: (Coord, Coord) -> Text
singleCellRefRaw' (row, col) =
    T.concat [viewAbs col, int2col (unCoord col), viewAbs row, T.pack (show (unCoord row))]
  where
    viewAbs Abs{} = "$"
    viewAbs Rel{} = ""

-- | Converse function to 'singleCellRef'
fromSingleCellRef :: CellRef -> Maybe (Int, Int)
fromSingleCellRef = fromSingleCellRefRaw . unCellRef

-- | Converse function to 'singleCellRef\''
fromSingleCellRef' :: CellRef -> Maybe (Coord, Coord)
fromSingleCellRef' = fromSingleCellRefRaw' . unCellRef

fromSingleCellRefRaw :: Text -> Maybe (Int, Int)
fromSingleCellRefRaw = fmap (both unCoord) . fromSingleCellRefRaw'

fromSingleCellRefRaw' :: Text -> Maybe (Coord, Coord)
fromSingleCellRefRaw' t = do
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
      (mkCoord isRowAbsolute)
      (mkCoord isColAbsolute)
      (row, col2int colT)
  where
    mkCoord isAbs = if isAbs then Abs else Rel

-- | reverse to 'singleCellRef' expecting valid reference and failig with
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
-- > mkRange (2, 4) (6, 8) == CellRef "D2:H6"
mkRange :: (Int, Int) -> (Int, Int) -> Range
mkRange fr to = mkRange' (both Rel fr) (both Rel to)

-- | Render range with possibly absolute coordinates
--
-- > mkRange' (Abs 2, Abs 4) (6, 8) == CellRef "$D$2:H6"
mkRange' :: (Coord,Coord) -> (Coord,Coord) -> Range
mkRange' fr to =
  CellRef $ T.concat [singleCellRefRaw' fr, ":", singleCellRefRaw' to]

-- | Render range existing in another worksheet.
-- This function always renders the sheet name single-quoted regardless the presence of spaces.
--
-- > mkForeignRange "MyOtherSheet" (Rel 2, Rel 4) (Abs 6, Abs 8) == "'MyOtherSheet'!D2:$H$6"
mkForeignRange :: Text -> (Coord, Coord) -> (Coord, Coord) -> Range
mkForeignRange sheetName fr to =
    case mkRange' fr to of
      CellRef cr -> CellRef $ T.concat ["'", sheetName, "'", "!", cr]

-- | Converse function to 'mkRange'.
-- Ignores a potential foreign sheet prefix.
fromRange :: Range -> Maybe ((Int, Int), (Int, Int))
fromRange r =
  both (both unCoord) <$> fromRange' r

-- | Converse function to 'mkRange\'' to handle possibly absolute coordinates.
-- Ignores a potential foreign sheet prefix.
fromRange' :: Range -> Maybe ((Coord, Coord), (Coord, Coord))
fromRange' t' = parseRange =<< ignoreForeignSheet (unCellRef t')
  where
    ignoreForeignSheet t =
        case T.split (== '!') t of
          [_, r] -> Just r
          [r] -> Just r
          _ -> Nothing
    parseRange t =
      case T.split (== ':') t of
        [from, to] -> liftA2 (,) (fromSingleCellRefRaw' from) (fromSingleCellRefRaw' to)
        _ -> Nothing

-- | Converse function to 'mkForeignRange'.
-- The provided Range must be a foreign range.
fromForeignRange :: Range -> Maybe (Text, ((Coord, Coord), (Coord, Coord)))
fromForeignRange r =
    case T.split (== '!') (unCellRef r) of
      [sheetName, ref] -> (unsquote sheetName,) <$> fromRange' (CellRef ref)
      _ -> Nothing
  where
    unsquote (T.strip -> n) = n & T.stripPrefix "'" >>= T.stripSuffix "'" & fromMaybe n

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

-- | Convertts serial value into datetime according to the specified
-- date base
--
-- > show (dateFromNumber DateBase1900 42929.75) == "2017-07-13 18:00:00 UTC"
dateFromNumber :: RealFrac t => DateBase -> t -> UTCTime
dateFromNumber b d = UTCTime day diffTime
  where
    (numberOfDays, fractionOfOneDay) = properFraction d
    day = addDays numberOfDays $ baseDate b
    diffTime = picosecondsToDiffTime (round (fractionOfOneDay * 24*60*60*1E12))

-- | Converts datetime into serial value
dateToNumber :: Fractional a => DateBase -> UTCTime -> a
dateToNumber b (UTCTime day diffTime) = numberOfDays + fractionOfOneDay
  where
    numberOfDays = fromIntegral (diffDays day $ baseDate b)
    fractionOfOneDay = realToFrac diffTime / (24 * 60 * 60)

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
