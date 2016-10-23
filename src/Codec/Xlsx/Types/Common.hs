{-# LANGUAGE OverloadedStrings #-}
module Codec.Xlsx.Types.Common
  ( CellRef(..)
  , singleCellRef
  , fromSingleCellRef
  , Range
  , mkRange
  , fromRange
  , SqRef(..)
  , XlsxText(..)
  , Formula(..)
  , CellValue(..)
  , int2col
  , col2int
  ) where

import Control.Arrow
import Data.Char
import qualified Data.Map as Map
import Data.Monoid ((<>))
import Data.Text (Text)
import qualified Data.Text as T
import Data.Tuple (swap)
import Safe (fromJustNote)
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
  } deriving (Eq, Ord, Show)

-- | Render position in @(row, col)@ format to an Excel reference.
--
-- > mkCellRef (2, 4) == "D2"
singleCellRef :: (Int, Int) -> CellRef
singleCellRef = CellRef . singleCellRefRaw

singleCellRefRaw :: (Int, Int) -> Text
singleCellRefRaw (row, col) = T.concat [int2col col, T.pack (show row)]

-- | reverse to 'mkCellRef'
--
-- /Warning:/ the function isn't total and will throw an error if
-- incorrect value will get passed
fromSingleCellRef :: CellRef -> (Int, Int)
fromSingleCellRef = fromSingleCellRefRaw . unCellRef

fromSingleCellRefRaw :: Text -> (Int, Int)
fromSingleCellRefRaw t =
  swap $ col2int *** toInt $ T.span (not . isDigit) t
  where
    toInt = fromJustNote "non-integer row in cell reference" . decimal


-- | Excel range (e.g. @D13:H14@), actually store as as 'CellRef' in
-- xlsx
type Range = CellRef

-- | Render range
--
-- > mkRange (2, 4) (6, 8) == "D2:H6"
mkRange :: (Int, Int) -> (Int, Int) -> Range
mkRange fr to = CellRef $ T.concat [singleCellRefRaw fr, T.pack ":", singleCellRefRaw to]

-- | reverse to 'mkRange'
--
-- /Warning:/ the function isn't total and will throw an error if
-- incorrect value will get passed
fromRange :: Range -> ((Int, Int), (Int, Int))
fromRange r = case T.split (==':') (unCellRef r) of
    [from, to] -> (fromSingleCellRefRaw from, fromSingleCellRefRaw to)
    _ -> error $ "invalid range " <> show r

-- | A sequence of cell references
--
-- See 18.18.76 "ST_Sqref (Reference Sequence)" (p.2488)
newtype SqRef = SqRef [CellRef]
    deriving (Eq, Ord, Show)

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
              deriving (Show, Eq, Ord)

-- | A formula
--
-- See 18.18.35 "ST_Formula (Formula)" (p. 2457)
newtype Formula = Formula {unFormula :: Text}
    deriving (Eq, Ord, Show)


-- | Cell values include text, numbers and booleans,
-- standard includes date format also but actually dates
-- are represented by numbers with a date format assigned
-- to a cell containing it
data CellValue = CellText   Text
               | CellDouble Double
               | CellBool   Bool
               | CellRich   [RichTextRun]
               deriving (Eq, Show)

{-------------------------------------------------------------------------------
  Parsing
-------------------------------------------------------------------------------}

-- | See @CT_Rst@, p. 3903
instance FromCursor XlsxText where
  fromCursor cur = do
    let
      ts = cur $/ element (n_ "t") >=> contentOrEmpty
      contentOrEmpty c = case c $/ content of
        [t] -> [t]
        []  -> [""]
        _   -> error "invalid item: more than one text nodes under <t>!"
      rs = cur $/ element (n_ "r") >=> fromCursor
    case (ts,rs) of
      ([t], []) ->
        return $ XlsxText t
      ([], _:_) ->
        return $ XlsxRichText rs
      _ ->
        fail "invalid item"

instance FromAttrVal CellRef where
  fromAttrVal = fmap (first CellRef) . fromAttrVal

instance FromAttrVal SqRef where
  fromAttrVal t = do
    rs <- mapM (fmap fst . fromAttrVal) $ T.split (== ' ') t
    readSuccess $ SqRef rs

-- | See @ST_Formula@, p. 3873
instance FromCursor Formula where
    fromCursor cur = [Formula . T.concat $ cur $/ content]

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
