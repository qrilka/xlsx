{-# LANGUAGE NoMonomorphismRestriction #-}
{-# LANGUAGE OverloadedStrings #-}
{-# LANGUAGE TupleSections #-}
-- | This module provides a function for reading .xlsx files
module Codec.Xlsx.Parser
    ( toXlsx
    ) where

import qualified Codec.Archive.Zip as Zip
import           Control.Applicative
import           Control.Arrow ((&&&))
import           Control.Monad.IO.Class()
import qualified Data.ByteString.Lazy as L
import           Data.ByteString.Lazy.Char8()
import           Data.List
import qualified Data.Map as M
import           Data.Maybe
import           Data.Ord
import           Data.Text (Text)
import qualified Data.Text as T
import qualified Data.Text.Read as T
import           Data.XML.Types
import           Prelude hiding (sequence)
import           Text.XML as X
import           Text.XML.Cursor

import           Codec.Xlsx.Parser.Internal
import           Codec.Xlsx.Types
import           Codec.Xlsx.Types.SharedStringTable


-- | Reads `Xlsx' from raw data (lazy bytestring)
toXlsx :: L.ByteString -> Xlsx
toXlsx bs = Xlsx sheets styles names
  where
    ar = Zip.toArchive bs
    sst = getSharedStrings ar
    styles = getStyles ar
    (wfs, names) = readWorkbook ar
    sheets = M.fromList $ map (wfName &&& extractSheet ar sst) wfs

data WorksheetFile = WorksheetFile { wfName :: Text
                                   , wfPath :: FilePath
                                   }
                   deriving Show

extractSheet :: Zip.Archive
             -> SharedStringTable
             -> WorksheetFile
             -> Worksheet
extractSheet ar sst wf = Worksheet cws rowProps cells merges sheetViews pageSetup
  where
    file = fromJust $ Zip.fromEntry <$> Zip.findEntryByPath (wfPath wf) ar
    cur = case parseLBS def file of
      Left _  -> error "could not read file"
      Right d -> fromDocument d

    -- The specification says the file should contain either 0 or 1 @sheetViews@
    -- (4th edition, section 18.3.1.88, p. 1704 and definition CT_Worksheet, p. 3910)
    sheetViewList = cur $/ element (n"sheetViews") &/ element (n"sheetView") >=> fromCursor
    sheetViews = case sheetViewList of
      [] -> Nothing
      views -> Just views

    -- Likewise, @pageSetup@ also occurs either 0 or 1 times
    pageSetup = listToMaybe $ cur $/ element (n"pageSetup") >=> fromCursor

    cws = cur $/ element (n"cols") &/ element (n"col") >=> fromCursor

    (rowProps, cells) = collect $ cur $/ element (n"sheetData") &/ element (n"row") >=> parseRow
    parseRow c = do
      r <- c $| attribute "r" >=> decimal
      let ht = if attribute "customHeight" c == ["true"]
               then listToMaybe $ c $| attribute "ht" >=> rational
               else Nothing
      let s = if attribute "s" c /= []
              then listToMaybe $ c $| attribute "s" >=> decimal
              else Nothing
      let rp = if isNothing s && isNothing ht
               then  Nothing
               else  Just (RowProps ht s)
      return (r, rp, c $/ element (n"c") >=> parseCell)
    parseCell :: Cursor -> [(Int, Int, Cell)]
    parseCell cell = do
      let
        s = listToMaybe $ cell $| attribute "s" >=> decimal
        t = fromMaybe "n" $ listToMaybe $ cell $| attribute "t"
        d = listToMaybe $ cell $/ element (n"v") &/ content >=> extractCellValue sst t
      (c, r) <- T.span (>'9') <$> (cell $| attribute "r")
      return (int r, col2int c, Cell s d)
    collect = foldr collectRow (M.empty, M.empty)
    collectRow (_, Nothing, rowCells) (rowMap, cellMap) =
      (rowMap, foldr collectCell cellMap rowCells)
    collectRow (r, Just h, rowCells) (rowMap, cellMap) =
      (M.insert r h rowMap, foldr collectCell cellMap rowCells)
    collectCell (x, y, cd) = M.insert (x,y) cd

    merges = cur $/ parseMerges
    parseMerges :: Cursor -> [Text]
    parseMerges = element (n"mergeCells") &/ element (n"mergeCell") >=> parseMerge
    parseMerge c = c $| attribute "ref"

extractCellValue :: SharedStringTable -> Text -> Text -> [CellValue]
extractCellValue sst "s" v =
    case T.decimal v of
      Right (d, _) ->
        case sstItem sst d of
          StringItemText txt  -> [CellText txt]
          StringItemRich rich -> [CellRich rich]
      _ ->
        []
extractCellValue _ "str" str = [CellText str]
extractCellValue _ "n" v =
    case T.rational v of
      Right (d, _) -> [CellDouble d]
      _ -> []
extractCellValue _ "b" "1" = [CellBool True]
extractCellValue _ "b" "0" = [CellBool False]
extractCellValue _ _ _ = []

-- | Add office document relationship namespace to name
odr :: Text -> Name
odr x = Name
  { nameLocalName = x
  , nameNamespace = Just "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  , namePrefix = Nothing
  }

-- | Add package relationship namespace to name
pr :: Text -> Name
pr x = Name
  { nameLocalName = x
  , nameNamespace = Just "http://schemas.openxmlformats.org/package/2006/relationships"
  , namePrefix = Nothing
  }

-- | Get xml cursor from the specified file inside the zip archive.
xmlCursor :: Zip.Archive -> FilePath -> Maybe Cursor
xmlCursor ar fname = parse <$> Zip.findEntryByPath fname ar
  where
    parse entry = case parseLBS def (Zip.fromEntry entry) of
        Left _  -> error "could not read file"
        Right d -> fromDocument d

-- | Get shared string table
getSharedStrings  :: Zip.Archive -> SharedStringTable
getSharedStrings x = case xmlCursor x "xl/sharedStrings.xml" of
    Nothing ->
      error "invalid shared strings"
    Just c ->
      let [sst] = fromCursor c in sst

getStyles :: Zip.Archive -> Styles
getStyles ar = case Zip.fromEntry <$> Zip.findEntryByPath "xl/styles.xml" ar of
  Nothing  -> Styles L.empty
  Just xml -> Styles xml

-- | readWorkbook pulls the names of the sheets and the defined names
readWorkbook :: Zip.Archive -> ([WorksheetFile], DefinedNames)
readWorkbook ar = case xmlCursor ar "xl/workbook.xml" of
  Nothing ->
    error "invalid workbook"
  Just c ->
    let
        sheetData = c $/ element (n"sheets") &/ element (n"sheet") >=>
                    liftA2 (,) <$> attribute "name" <*> attribute (odr"id")
        wbRels = getWbRels ar
        names = c $/ element (n"definedNames") &/ element (n"definedName") >=> mkDefinedName
    in ([WorksheetFile name ("xl/" ++ T.unpack (fromJust $ lookup rId wbRels)) | (name, rId) <- sheetData]
       ,DefinedNames names)
 where
  -- Specification says the 'name' is required.
  mkDefinedName :: Cursor -> [(Text, Maybe Text, Text)]
  mkDefinedName c = return ( head $ attribute "name" c
                           , listToMaybe $ attribute "localSheetId" c
                           , T.concat $ c $/ content
                           )

getWbRels :: Zip.Archive -> [(Text, Text)]
getWbRels ar = case xmlCursor ar "xl/_rels/workbook.xml.rels" of
  Nothing -> []
  Just c  -> c $/ element (pr"Relationship") >=>
             liftA2 (,) <$> attribute "Id" <*> attribute "Target"

int :: Text -> Int
int = either error fst . T.decimal
