{-# LANGUAGE ScopedTypeVariables #-}
{-# LANGUAGE NoMonomorphismRestriction #-}
{-# LANGUAGE OverloadedStrings         #-}
{-# LANGUAGE TupleSections             #-}
-- | This module provides a function for reading .xlsx files
module Codec.Xlsx.Parser
    ( toXlsx
    , toXlsxEither
    , ParseError (..)
    ) where

import qualified Codec.Archive.Zip                           as Zip
import           Control.Applicative
import           Control.Arrow                               ((&&&))
import           Control.Monad.IO.Class                      ()
import           Data.Bifunctor                              (bimap, first)
import qualified Data.ByteString.Lazy                        as L
import           Data.ByteString.Lazy.Char8                  ()
import           Data.List
import qualified Data.Map                                    as M
import           Data.Maybe
import           Data.Ord
import           Data.Text                                   (Text)
import qualified Data.Text                                   as T
import qualified Data.Text.Read                              as T
import           Prelude                                     hiding (sequence)
import           Safe
import           System.FilePath.Posix
import           Text.XML                                    as X
import           Text.XML.Cursor

import           Codec.Xlsx.Parser.Internal
import           Codec.Xlsx.Types
import           Codec.Xlsx.Types.Internal
import           Codec.Xlsx.Types.Internal.CfPair
import           Codec.Xlsx.Types.Internal.CommentTable
import           Codec.Xlsx.Types.Internal.CustomProperties  as CustomProperties
import           Codec.Xlsx.Types.Internal.Relationships     as Relationships
import           Codec.Xlsx.Types.Internal.SharedStringTable


-- | Reads `Xlsx' from raw data (lazy bytestring)
toXlsx :: L.ByteString -> Xlsx
toXlsx = either (error . show) id . toXlsxEither

data ParseError = InvalidZipArchive
                | MissingFile FilePath
                | InvalidFile FilePath
                deriving (Show, Eq)

-- | Reads `Xlsx' from raw data (lazy bytestring), failing with Left on parse error
toXlsxEither :: L.ByteString -> Either ParseError Xlsx
toXlsxEither bs = do
  ar <- first (const InvalidZipArchive) $ Zip.toArchiveOrFail bs
  sst <- getSharedStrings ar
  (wfs, names) <- readWorkbook ar
  sheets <- sequenceA . M.fromList $ map (wfName &&& extractSheet ar sst) wfs
  CustomProperties customPropMap <- getCustomProperties ar
  return $ Xlsx sheets (getStyles ar) names customPropMap

data WorksheetFile = WorksheetFile { wfName :: Text
                                   , wfPath :: FilePath
                                   }
                   deriving Show

extractSheet :: Zip.Archive
             -> SharedStringTable
             -> WorksheetFile
             -> Either ParseError Worksheet
extractSheet ar sst wf = do
  let  file = fromJust $ Zip.fromEntry <$> Zip.findEntryByPath (wfPath wf) ar
  cur <- bimap (const $ MissingFile (wfPath wf)) fromDocument $ parseLBS def file
  sheetRels <- getRels ar (wfPath wf)

  -- The specification says the file should contain either 0 or 1 @sheetViews@
  -- (4th edition, section 18.3.1.88, p. 1704 and definition CT_Worksheet, p. 3910)
  let  sheetViewList = cur $/ element (n"sheetViews") &/ element (n"sheetView") >=> fromCursor
       sheetViews = case sheetViewList of
         [] -> Nothing
         views -> Just views

  let commentsType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments"
      commentTarget :: Maybe FilePath
      commentTarget = relTarget <$> findRelByType commentsType sheetRels

  commentsMap :: Maybe CommentTable <- maybe (Right Nothing) (getComments ar) commentTarget --  getComments ar <$> commentTarget

        -- Likewise, @pageSetup@ also occurs either 0 or 1 times
  let pageSetup = listToMaybe $ cur $/ element (n"pageSetup") >=> fromCursor

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
        ref <- cell $| attribute "r"
        let
          s = listToMaybe $ cell $| attribute "s" >=> decimal
          t = fromMaybe "n" $ listToMaybe $ cell $| attribute "t"
          d = listToMaybe $ cell $/ element (n"v") &/ content >=> extractCellValue sst t
          f = listToMaybe $ cell $/ element (n"f") >=> fromCursor
          (c, r) = T.span (>'9') ref
          comment = commentsMap >>= lookupComment ref
        return (int r, col2int c, Cell s d comment f)
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

      condFormtattings = M.fromList . map unCfPair  $ cur $/ element (n"conditionalFormatting") >=> fromCursor
  return $ Worksheet cws rowProps cells merges sheetViews pageSetup condFormtattings

extractCellValue :: SharedStringTable -> Text -> Text -> [CellValue]
extractCellValue sst "s" v =
    case T.decimal v of
      Right (d, _) ->
        case sstItem sst d of
          XlsxText     txt  -> [CellText txt]
          XlsxRichText rich -> [CellRich rich]
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

-- | Get xml cursor from the specified file inside the zip archive.
xmlCursor :: Zip.Archive -> FilePath -> Either ParseError (Maybe Cursor)
xmlCursor ar fname = maybe (Right Nothing) parse $ Zip.findEntryByPath fname ar
  where
    parse entry = bimap (const $ InvalidFile fname) (return . fromDocument) $ parseLBS def (Zip.fromEntry entry)

-- | Get xml cursor from the given file, failing with MissingFile if not found.
xmlCursorRequired :: Zip.Archive -> FilePath -> Either ParseError Cursor
xmlCursorRequired ar fname = maybe (Left $ MissingFile fname) Right =<< xmlCursor ar fname

-- | Get shared string table
getSharedStrings  :: Zip.Archive -> Either ParseError SharedStringTable
getSharedStrings x = head . fromCursor <$> xmlCursorRequired x "xl/sharedStrings.xml"

getStyles :: Zip.Archive -> Styles
getStyles ar = case Zip.fromEntry <$> Zip.findEntryByPath "xl/styles.xml" ar of
  Nothing  -> Styles L.empty
  Just xml -> Styles xml

getComments :: Zip.Archive -> FilePath -> Either ParseError (Maybe CommentTable)
getComments ar fp = (listToMaybe . fromCursor =<<) <$> xmlCursor ar fp  --listToMaybe =<< fromCursor <$$> xmlCursor ar fp

getCustomProperties :: Zip.Archive -> Either ParseError CustomProperties
getCustomProperties ar = maybe CustomProperties.empty (head . fromCursor) <$> xmlCursor ar "docProps/custom.xml"

-- | readWorkbook pulls the names of the sheets and the defined names
readWorkbook :: Zip.Archive -> Either ParseError ([WorksheetFile], DefinedNames)
readWorkbook ar = do -- case xmlCursor ar wbPath of
  let wbPath = "xl/workbook.xml"
  cur <- xmlCursorRequired ar wbPath
  wbRels <- getRels ar wbPath
  let -- Specification says the 'name' is required.
      mkDefinedName :: Cursor -> [(Text, Maybe Text, Text)]
      mkDefinedName c = return ( head $ attribute "name" c
                               , listToMaybe $ attribute "localSheetId" c
                               , T.concat $ c $/ content
                               )

      sheets = cur $/ element (n"sheets") &/ element (n"sheet") >=>
                    liftA2 (worksheetFile wbRels) <$> attribute "name" <*> (attribute (odr"id") &| RefId)
      names = cur $/ element (n"definedNames") &/ element (n"definedName") >=> mkDefinedName
  return $ (sheets, DefinedNames names)


worksheetFile :: Relationships -> Text -> RefId -> WorksheetFile
worksheetFile wbRels name rId = WorksheetFile name path
  where
    path = relTarget . fromJustNote "sheet path" $ Relationships.lookup rId wbRels

getRels :: Zip.Archive -> FilePath -> Either ParseError Relationships
getRels ar fp = do
    let (dir, file) = splitFileName fp
        relsPath = dir </> "_rels" </> file <.> "rels"
    c <- xmlCursor ar relsPath
    return $ maybe Relationships.empty (setTargetsFrom fp . head . fromCursor) c

int :: Text -> Int
int = either error fst . T.decimal
