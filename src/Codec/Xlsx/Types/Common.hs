{-# LANGUAGE OverloadedStrings #-}
module Codec.Xlsx.Types.Common
       ( CellRef
       , mkCellRef
       , fromCellRef
       , SqRef (..)
       , XlsxText (..)
       , Formula (..)
       , int2col
       , col2int
       ) where

import           Control.Arrow
import           Data.Char
import qualified Data.Map                   as Map
import           Data.Text                  (Text)
import qualified Data.Text                  as T
import           Data.Tuple                 (swap)
import           Safe                       (fromJustNote)
import           Text.XML
import           Text.XML.Cursor

import           Codec.Xlsx.Parser.Internal
import           Codec.Xlsx.Types.RichText
import           Codec.Xlsx.Writer.Internal

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

-- | Excel cell reference (e.g. @E3@)
-- See 18.18.62 @ST_Ref@ (p. 2482)
type CellRef = Text

-- | Render position in @(row, col)@ format to an Excel reference.
--
-- > mkCellRef (2, 4) == "D2"
mkCellRef :: (Int, Int) -> CellRef
mkCellRef (row, col) = T.concat [int2col col, T.pack (show row)]

-- | reverse to 'mkCellRef'
--
-- /Warning:/ the function isn't total and will throw an error if
-- incorrect value will get passed
fromCellRef :: CellRef -> (Int, Int)
fromCellRef t = swap $ col2int *** toInt $ T.span (not . isDigit) t
  where
    toInt = fromJustNote "non-integer row in cell reference" . decimal

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

{-------------------------------------------------------------------------------
  Parsing
-------------------------------------------------------------------------------}

-- | See @CT_Rst@, p. 3903
instance FromCursor XlsxText where
  fromCursor cur = do
    let
      ts = cur $/ element (n"t") >=> contentOrEmpty
      contentOrEmpty c = case c $/ content of
        [t] -> [t]
        []  -> [""]
        _   -> error "invalid item: more than one text nodes under <t>!"
      rs = cur $/ element (n"r") >=> fromCursor
    case (ts,rs) of
      ([t], []) ->
        return $ XlsxText t
      ([], _:_) ->
        return $ XlsxRichText rs
      _ ->
        fail "invalid item"

instance FromAttrVal SqRef where
    fromAttrVal t = readSuccess (SqRef $ T.split (== ' ') t)

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

-- | A sequence of cell references, space delimited.
-- See 18.18.76, "ST_Sqref (Reference Sequence)", p. 2488.
instance ToAttrVal SqRef where
    toAttrVal (SqRef refs) = T.intercalate " " refs

-- | See @ST_Formula@, p. 3873
instance ToElement Formula where
    toElement nm (Formula txt) = elementContent nm txt
