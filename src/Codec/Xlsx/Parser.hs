{-# LANGUAGE FlexibleContexts #-}
{-# LANGUAGE ScopedTypeVariables #-}
{-# LANGUAGE NoMonomorphismRestriction #-}
{-# LANGUAGE OverloadedStrings         #-}
{-# LANGUAGE TupleSections             #-}
-- | This module provides a function for reading .xlsx files
module Codec.Xlsx.Parser
    ( toXlsx
    , toXlsxOrError
    , ParseError (..)
    ) where

import qualified Codec.Archive.Zip                           as Zip
import           Control.Applicative
import           Control.Arrow                               ((&&&))
import           Control.Monad                               (liftM)
import           Control.Monad.Except                        (MonadError(..))
import           Control.Monad.IO.Class                      ()
import qualified Data.ByteString.Lazy                        as L
import           Data.ByteString.Lazy.Char8                  ()
import           Data.List
import qualified Data.Map                                    as M
import           Data.Maybe
import           Data.Ord
import           Data.Text                                   (Text)
import qualified Data.Text                                   as T
import qualified Data.Text.Read                              as T
import           Data.Traversable
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
toXlsx = either (error . show) id . toXlsxOrError

data ParseError = InvalidZipArchive
                | MissingFile FilePath
                | InvalidFile FilePath
                deriving (Show, Eq)

-- | Reads `Xlsx' from raw data (lazy bytestring), failing with Left on parse error
toXlsxOrError :: MonadError ParseError m => L.ByteString -> m Xlsx
toXlsxOrError bs = do
  ar <- Zip.toArchiveOrFail bs `whenLeftThrow` const InvalidZipArchive
  sst <- getSharedStrings ar
  (wfs, names) <- readWorkbook ar
  sheets <- sequence . M.fromList $ map (wfName &&& extractSheet ar sst) wfs
  CustomProperties customPropMap <- getCustomProperties ar
  return $ Xlsx sheets (getStyles ar) names customPropMap

data WorksheetFile = WorksheetFile { wfName :: Text
                                   , wfPath :: FilePath
                                   }
                   deriving Show

extractSheet :: MonadError ParseError m
             => Zip.Archive
             -> SharedStringTable
             -> WorksheetFile
             -> m Worksheet
extractSheet ar sst wf = do
  let filePath = wfPath wf
  file <- Zip.fromEntry `fmap` Zip.findEntryByPath filePath ar
          `whenNothingThrow` MissingFile filePath
  cur <- fromDocument `fmap` parseLBS def file
         `whenLeftThrow` const (InvalidFile filePath)
  sheetRels <- getRels ar filePath

  -- The specification says the file should contain either 0 or 1 @sheetViews@
  -- (4th edition, section 18.3.1.88, p. 1704 and definition CT_Worksheet, p. 3910)
  let  sheetViewList = cur $/ element (n"sheetViews") &/ element (n"sheetView") >=> fromCursor
       sheetViews = case sheetViewList of
         [] -> Nothing
         views -> Just views

  let commentsType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments"
      commentTarget :: Maybe FilePath
      commentTarget = relTarget <$> findRelByType commentsType sheetRels

  commentsMap :: Maybe CommentTable <- maybe (return Nothing) (getComments ar) commentTarget

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
xmlCursor :: MonadError ParseError m => Zip.Archive -> FilePath -> m (Maybe Cursor)
xmlCursor ar fname = maybe (return Nothing) parse $ Zip.findEntryByPath fname ar
  where parse entry = (return . fromDocument) `fmap` parseLBS def (Zip.fromEntry entry)
                      `whenLeftThrow` const (InvalidFile fname)

-- | Get xml cursor from the given file, failing with MissingFile if not found.
xmlCursorRequired :: MonadError ParseError m => Zip.Archive -> FilePath -> m Cursor
xmlCursorRequired ar fname = xmlCursor ar fname >>= (`whenNothingThrow` MissingFile fname)

-- | Get shared string table
getSharedStrings  :: MonadError ParseError m => Zip.Archive -> m SharedStringTable
getSharedStrings x = (head . fromCursor) `liftM` xmlCursorRequired x "xl/sharedStrings.xml"

getStyles :: Zip.Archive -> Styles
getStyles ar = Styles . fromMaybe L.empty $ Zip.fromEntry <$> Zip.findEntryByPath "xl/styles.xml" ar

getComments :: MonadError ParseError m => Zip.Archive -> FilePath -> m (Maybe CommentTable)
getComments ar fp = (listToMaybe . fromCursor =<<) `liftM` xmlCursor ar fp

getCustomProperties :: MonadError ParseError m => Zip.Archive -> m CustomProperties
getCustomProperties ar = maybe CustomProperties.empty (head . fromCursor) `liftM` xmlCursor ar "docProps/custom.xml"

-- | readWorkbook pulls the names of the sheets and the defined names
readWorkbook :: MonadError ParseError m => Zip.Archive -> m ([WorksheetFile], DefinedNames)
readWorkbook ar = do
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
  return (sheets, DefinedNames names)


worksheetFile :: Relationships -> Text -> RefId -> WorksheetFile
worksheetFile wbRels name rId = WorksheetFile name path
  where
    path = relTarget . fromJustNote "sheet path" $ Relationships.lookup rId wbRels

getRels :: MonadError ParseError m => Zip.Archive -> FilePath -> m Relationships
getRels ar fp = do
    let (dir, file) = splitFileName fp
        relsPath = dir </> "_rels" </> file <.> "rels"
    c <- xmlCursor ar relsPath
    return $ maybe Relationships.empty (setTargetsFrom fp . head . fromCursor) c

int :: Text -> Int
int = either error fst . T.decimal

-- | Lift a 'Maybe' value into any 'MonadError'.
-- Intended to be used infix; e.g.,
--
--    Just 3 `whenNothingThrow` SomeError
whenNothingThrow :: MonadError e m => Maybe a -> e -> m a
whenNothingThrow v e = maybe (throwError e) return v

-- | Lift an 'Either' value into any 'MonadError'.
whenLeftThrow :: MonadError e m => Either l r -> (l -> e) -> m r
whenLeftThrow v e = either (throwError . e) return v
