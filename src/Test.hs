{-# LANGUAGE OverloadedStrings #-}
module Test where

import           Codec.Xlsx
import           Codec.Xlsx.Writer
import           Control.Arrow
import           Control.Monad.Trans.Class
import           Control.Monad.Trans.List
import           Control.Monad.Trans.State
import qualified Data.IntMap as IM
import qualified Data.Map as M
import           Data.Maybe
import           Data.Text (Text)
import qualified Data.Text as T
import qualified Data.Text.Read as T
import           Data.Time.Calendar
import           Data.Time.LocalTime
import           Debug.Trace
import           Text.XML as X
import           Text.XML.Cursor

xText t = Just CellData{cdValue=Just $ CellText t, cdStyle=Just 0}
xDate d = Just CellData{cdValue=Just $ CellLocalTime d, cdStyle=Just 0}
xDouble d = Just CellData{cdValue=Just $ CellDouble d, cdStyle=Just 0}

test = writeXlsxStyles "test.xlsx" styles [fromList sheet1 cols1 rows1, fromList sheet2 cols2 rows2]
  where
    cols1 = [ColumnsWidth 1 10 15]
    rows1 = M.fromList [(1,50)]
    sheet1 = [[xText "column1", xText "column2", xText "column4"],
              [xDate $ LocalTime (fromGregorian 2012 05 06) (TimeOfDay 7 30 50), xDouble 42.12345, xText  "False"]]
    cols2 = [ColumnsWidth 1 3 15, ColumnsWidth 4 10 35]
    rows2 = M.fromList [(1,5),(2,20)]
    sheet2 = [[xText "column1", xText "column2", xText "column2"]]

styles = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\
\<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><numFmts count=\"1\"><numFmt formatCode=\"GENERAL\" numFmtId=\"164\"/></numFmts><fonts count=\"4\"><font><name val=\"Courier New\"/><charset val=\"1\"/><family val=\"2\"/><sz val=\"10\"/></font><font><name val=\"Arial\"/><family val=\"0\"/><sz val=\"10\"/></font><font><name val=\"Arial\"/><family val=\"0\"/><sz val=\"10\"/></font><font><name val=\"Arial\"/><family val=\"0\"/><sz val=\"10\"/></font></fonts><fills count=\"2\"><fill><patternFill patternType=\"none\"/></fill><fill><patternFill patternType=\"gray125\"/></fill></fills><borders count=\"1\"><border diagonalDown=\"false\" diagonalUp=\"false\"><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count=\"20\"><xf applyAlignment=\"true\" applyBorder=\"true\" applyFont=\"true\" applyProtection=\"true\" borderId=\"0\" fillId=\"0\" fontId=\"0\" numFmtId=\"164\"><alignment horizontal=\"general\" indent=\"0\" shrinkToFit=\"false\" textRotation=\"0\" vertical=\"bottom\" wrapText=\"false\"/><protection hidden=\"false\" locked=\"true\"/></xf><xf applyAlignment=\"false\" applyBorder=\"false\" applyFont=\"true\" applyProtection=\"false\" borderId=\"0\" fillId=\"0\" fontId=\"1\" numFmtId=\"0\"></xf><xf applyAlignment=\"false\" applyBorder=\"false\" applyFont=\"true\" applyProtection=\"false\" borderId=\"0\" fillId=\"0\" fontId=\"1\" numFmtId=\"0\"></xf><xf applyAlignment=\"false\" applyBorder=\"false\" applyFont=\"true\" applyProtection=\"false\" borderId=\"0\" fillId=\"0\" fontId=\"2\" numFmtId=\"0\"></xf><xf applyAlignment=\"false\" applyBorder=\"false\" applyFont=\"true\" applyProtection=\"false\" borderId=\"0\" fillId=\"0\" fontId=\"2\" numFmtId=\"0\"></xf><xf applyAlignment=\"false\" applyBorder=\"false\" applyFont=\"true\" applyProtection=\"false\" borderId=\"0\" fillId=\"0\" fontId=\"0\" numFmtId=\"0\"></xf><xf applyAlignment=\"false\" applyBorder=\"false\" applyFont=\"true\" applyProtection=\"false\" borderId=\"0\" fillId=\"0\" fontId=\"0\" numFmtId=\"0\"></xf><xf applyAlignment=\"false\" applyBorder=\"false\" applyFont=\"true\" applyProtection=\"false\" borderId=\"0\" fillId=\"0\" fontId=\"0\" numFmtId=\"0\"></xf><xf applyAlignment=\"false\" applyBorder=\"false\" applyFont=\"true\" applyProtection=\"false\" borderId=\"0\" fillId=\"0\" fontId=\"0\" numFmtId=\"0\"></xf><xf applyAlignment=\"false\" applyBorder=\"false\" applyFont=\"true\" applyProtection=\"false\" borderId=\"0\" fillId=\"0\" fontId=\"0\" numFmtId=\"0\"></xf><xf applyAlignment=\"false\" applyBorder=\"false\" applyFont=\"true\" applyProtection=\"false\" borderId=\"0\" fillId=\"0\" fontId=\"0\" numFmtId=\"0\"></xf><xf applyAlignment=\"false\" applyBorder=\"false\" applyFont=\"true\" applyProtection=\"false\" borderId=\"0\" fillId=\"0\" fontId=\"0\" numFmtId=\"0\"></xf><xf applyAlignment=\"false\" applyBorder=\"false\" applyFont=\"true\" applyProtection=\"false\" borderId=\"0\" fillId=\"0\" fontId=\"0\" numFmtId=\"0\"></xf><xf applyAlignment=\"false\" applyBorder=\"false\" applyFont=\"true\" applyProtection=\"false\" borderId=\"0\" fillId=\"0\" fontId=\"0\" numFmtId=\"0\"></xf><xf applyAlignment=\"false\" applyBorder=\"false\" applyFont=\"true\" applyProtection=\"false\" borderId=\"0\" fillId=\"0\" fontId=\"0\" numFmtId=\"0\"></xf><xf applyAlignment=\"false\" applyBorder=\"false\" applyFont=\"true\" applyProtection=\"false\" borderId=\"0\" fillId=\"0\" fontId=\"1\" numFmtId=\"43\"></xf><xf applyAlignment=\"false\" applyBorder=\"false\" applyFont=\"true\" applyProtection=\"false\" borderId=\"0\" fillId=\"0\" fontId=\"1\" numFmtId=\"41\"></xf><xf applyAlignment=\"false\" applyBorder=\"false\" applyFont=\"true\" applyProtection=\"false\" borderId=\"0\" fillId=\"0\" fontId=\"1\" numFmtId=\"44\"></xf><xf applyAlignment=\"false\" applyBorder=\"false\" applyFont=\"true\" applyProtection=\"false\" borderId=\"0\" fillId=\"0\" fontId=\"1\" numFmtId=\"42\"></xf><xf applyAlignment=\"false\" applyBorder=\"false\" applyFont=\"true\" applyProtection=\"false\" borderId=\"0\" fillId=\"0\" fontId=\"1\" numFmtId=\"9\"></xf></cellStyleXfs><cellXfs count=\"1\"><xf applyAlignment=\"false\" applyBorder=\"false\" applyFont=\"false\" applyProtection=\"false\" borderId=\"0\" fillId=\"0\" fontId=\"0\" numFmtId=\"164\" xfId=\"0\"></xf></cellXfs><cellStyles count=\"6\"><cellStyle builtinId=\"0\" customBuiltin=\"false\" name=\"Normal\" xfId=\"0\"/><cellStyle builtinId=\"3\" customBuiltin=\"false\" name=\"Comma\" xfId=\"15\"/><cellStyle builtinId=\"6\" customBuiltin=\"false\" name=\"Comma [0]\" xfId=\"16\"/><cellStyle builtinId=\"4\" customBuiltin=\"false\" name=\"Currency\" xfId=\"17\"/><cellStyle builtinId=\"7\" customBuiltin=\"false\" name=\"Currency [0]\" xfId=\"18\"/><cellStyle builtinId=\"5\" customBuiltin=\"false\" name=\"Percent\" xfId=\"19\"/></cellStyles></styleSheet>"

test2 = do
  doc <- X.parseLBS def xml1
  let c = fromDocument doc
      x = c $// element "x" &/ element "bar" >=> (\c' -> do
        foo <- c' $| attribute "foo"
        baz <- c' $| attribute "baz"
        cont <- c' $/ content
        return [(foo, baz, cont)])
      y = c $// element "x" &/ element "y" &/  element "bar" >=> attribute "foo"
  return (x,y)

xml1 = "<x>\
\   <x>\
\       <y>\
\           <x>\
\               <bar foo=\"1\" baz=\"a\">cooo<inner>ddoo</inner>foo</bar>\
\           </x>\
\           <bar foo=\"2\"/>\
\       </y>\
\       <bar foo=\"3\"/>\
\   </x>\
\   <bar foo=\"4\"/>\
\</x>"

test3 = c $// parseRows
  where
    doc = case X.parseLBS def xml2 of
      Right d -> d
      Left _ -> error "invalid xml"
    c :: Cursor
    c = fromDocument doc
    parseRows :: Cursor -> [(Text, [Text])]
    parseRows = element "row" >=> parseRow
    parseRow :: Cursor -> [(Text, [Text])]
    parseRow c = do
      h <- c $| attribute "height"
      return (h, c $/ element "cell" &/ content {-attribute "t"-})

xml2 = "<x>\
\<row height=\"42\">\
\<cell t=\"1\">content here</cell>\
\<x><cell t=\"invalid\"/></x>\
\</row>\
\<row height=\"100500\">\
\<cell t=\"s\"/>\
\<cell t=\"n\"/>\
\</row>\
\</x>"


test4 = (c $// parseRows, c $// element (n"col") >=> attribute "width")
  where
    doc = case X.parseLBS def xml3 of
      Right d -> d
      Left _ -> error "invalid xml"
    c :: Cursor
    c = fromDocument doc
    parseRows = element (n"sheetData") &/ element (n"row") >=> parseRow
    parseRow c = do
      r <- c $| attribute "r" >=> decimal
      let ht = if attribute "customHeight" c == ["1"]
               then listToMaybe $ c $| attribute "ht" >=> rational
               else Nothing
      return (r, ht, c $/ element (n"c") >=> parseCell)
    ss = IM.fromList [(0,"first"),(1,"second"),(2,"third")]
    parseCell :: Cursor -> [(Int, Int, CellData)]
    parseCell c = do
      ix <- c $| attribute "r"
      let
         (c, r) = T.span (>'9') ix
         x = either error fst $ T.decimal r
         y = col2int c
      return (x, y, CellData s d)
        where
          s = listToMaybe $ c $| attribute "s" >=> decimal
          t = fromMaybe "n" $ listToMaybe $ c $| attribute "t"
          d = listToMaybe $ c $/ element (n"v") &/ extractValue
          extractValue c = case (t, node c) of
            ("n", X.NodeContent v) ->
              case T.rational v of
                Right (d, _) -> [CellDouble d]
                _ -> []
            ("s", X.NodeContent v) ->
              case T.decimal v of
                Right (i, _) -> maybeToList $ fmap CellText $ IM.lookup i ss
                _ -> []
            _ -> []

decimal :: Monad m => Text -> m Int
decimal t = case T.decimal t of
  Right (d, _) -> return d
  _ -> fail "invalid decimal"

rational :: Monad m => Text -> m Double
rational t = case T.rational t of
  Right (r, _) -> return r
  _ -> fail "invalid rational"

n x = Name
  {nameLocalName = x
  ,nameNamespace = Just "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  ,namePrefix = Nothing}


xml3 = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\
\<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">\
\<cols><col min=\"1\" max=\"1\" width=\"75.42578125\" foo=\"xxx\"/><col min=\"2\" max=\"257\" width=\"8.85546875\"/></cols>\
\<sheetData><row r=\"1\" spans=\"1:3\" ht=\"12.75\"><c r=\"A1\" t=\"s\"><v>0</v></c><c r=\"B1\" t=\"s\"><v>1</v></c><c r=\"C1\" t=\"s\"><v>2</v></c><c r=\"D1\" t=\"n\"><v>42</v></c></row></sheetData>\
\<other/>\
\</worksheet>"