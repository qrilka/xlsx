{-# LANGUAGE NoMonomorphismRestriction #-}
{-# LANGUAGE OverloadedStrings         #-}
{-# LANGUAGE ScopedTypeVariables       #-}
{-# LANGUAGE TupleSections             #-}

-- | This module provides a function for reading .xlsx files
module Codec.Xlsx.Parser
    ( toXlsx
    , toXlsxEither
    , ParseError (..)
    , Parser
    ) where

import qualified Codec.Archive.Zip                           as Zip
import           Control.Applicative
import           Control.Arrow                               (left)
import           Control.Error.Safe                          (headErr)
import           Control.Error.Util                          (note)
import           Control.Lens                                hiding (element,
                                                              views, (<.>))
import           Control.Monad.Except                        (catchError,
                                                              throwError)
import qualified Data.ByteString.Lazy                        as L
import           Data.ByteString.Lazy.Char8                  ()
import           Data.List
import qualified Data.Map                                    as M
import           Data.Maybe
import           Data.Monoid
import           Data.Ord
import           Data.Text                                   (Text)
import qualified Data.Text                                   as T
import qualified Data.Text.Read                              as T
import           Data.Traversable
import           Prelude                                     hiding (sequence)
import           System.FilePath.Posix
import           Text.XML                                    as X
import           Text.XML.Cursor

import           Codec.Xlsx.Parser.Internal
import           Codec.Xlsx.Types
import           Codec.Xlsx.Types.Internal
import           Codec.Xlsx.Types.Internal.CfPair
import           Codec.Xlsx.Types.Internal.CommentTable
import           Codec.Xlsx.Types.Internal.ContentTypes      as ContentTypes
import           Codec.Xlsx.Types.Internal.CustomProperties  as CustomProperties
import           Codec.Xlsx.Types.Internal.Relationships     as Relationships
import           Codec.Xlsx.Types.Internal.SharedStringTable

-- | Reads `Xlsx' from raw data (lazy bytestring)
toXlsx :: L.ByteString -> Xlsx
toXlsx = either (error . show) id . toXlsxEither

data ParseError = InvalidZipArchive
                | MissingFile FilePath
                | InvalidFile FilePath
                | InvalidRef FilePath RefId
                deriving (Show, Eq)

type Parser = Either ParseError

-- | Reads `Xlsx' from raw data (lazy bytestring), failing with Left on parse error
toXlsxEither :: L.ByteString -> Parser Xlsx
toXlsxEither bs = do
  ar <- left (const InvalidZipArchive) $ Zip.toArchiveOrFail bs
  sst <- getSharedStrings ar
  contentTypes <- getContentTypes ar
  (wfs, names) <- readWorkbook ar
  sheets <- forM wfs $ \wf -> do
      sheet <- extractSheet ar sst contentTypes wf
      return (wfName wf, sheet)
  CustomProperties customPropMap <- getCustomProperties ar
  return $ Xlsx sheets (getStyles ar) names customPropMap

data WorksheetFile = WorksheetFile { wfName :: Text
                                   , wfPath :: FilePath
                                   }
                   deriving Show

extractSheet :: Zip.Archive
             -> SharedStringTable
             -> ContentTypes
             -> WorksheetFile
             -> Parser Worksheet
extractSheet ar sst contentTypes wf = do
  let filePath = wfPath wf
  file <- note (MissingFile filePath) $ Zip.fromEntry <$> Zip.findEntryByPath filePath ar
  cur <- fmap fromDocument . left (\_ -> InvalidFile filePath) $
         parseLBS def file
  sheetRels <- getRels ar filePath

  -- The specification says the file should contain either 0 or 1 @sheetViews@
  -- (4th edition, section 18.3.1.88, p. 1704 and definition CT_Worksheet, p. 3910)
  let  sheetViewList = cur $/ element (n"sheetViews") &/ element (n"sheetView") >=> fromCursor
       sheetViews = case sheetViewList of
         []    -> Nothing
         views -> Just views

  let commentsType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments"
      commentTarget :: Maybe FilePath
      commentTarget = relTarget <$> findRelByType commentsType sheetRels

  commentsMap :: Maybe CommentTable <- maybe (Right Nothing) (getComments ar) commentTarget

  -- Likewise, @pageSetup@ also occurs either 0 or 1 times
  let pageSetup = listToMaybe $ cur $/ element (n"pageSetup") >=> fromCursor

      cws = cur $/ element (n"cols") &/ element (n"col") >=> fromCursor

      (rowProps, cells) = collect $ cur $/ element (n"sheetData") &/ element (n"row") >=> parseRow
      parseRow :: Cursor -> [(Int, Maybe RowProperties, [(Int, Int, Cell)])]
      parseRow c = do
        r <- c $| attribute "r" >=> decimal
        let ht = if attribute "customHeight" c == ["true"]
                 then listToMaybe $ c $| attribute "ht" >=> rational
                 else Nothing
        let s = listToMaybe $ decimal =<< attribute "s" c :: Maybe Int
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

      mDrawingId = listToMaybe $ cur $/ element (n"drawing") >=> fromAttribute (odr"id")

      merges = cur $/ parseMerges
      parseMerges :: Cursor -> [Text]
      parseMerges = element (n"mergeCells") &/ element (n"mergeCell") >=> parseMerge
      parseMerge c = c $| attribute "ref"

      condFormtattings = M.fromList . map unCfPair  $ cur $/ element (n"conditionalFormatting") >=> fromCursor

  mDrawing <- case mDrawingId of
      Just dId -> do
          rel <- note (InvalidRef filePath dId) $ Relationships.lookup dId sheetRels
          Just <$> getDrawing ar contentTypes (relTarget rel)
      Nothing  ->
          return Nothing

  return $ Worksheet cws rowProps cells mDrawing merges sheetViews pageSetup condFormtattings

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
      _            -> []
extractCellValue _ "b" "1" = [CellBool True]
extractCellValue _ "b" "0" = [CellBool False]
extractCellValue _ _ _ = []

-- | Get xml cursor from the specified file inside the zip archive.
xmlCursorOptional :: Zip.Archive -> FilePath -> Parser (Maybe Cursor)
xmlCursorOptional ar fname =
    (Just <$> xmlCursorRequired ar fname) `catchError` missingToNothing
  where
    missingToNothing :: ParseError -> Either ParseError (Maybe a)
    missingToNothing (MissingFile _) = return Nothing
    missingToNothing other           = throwError other

-- | Get xml cursor from the given file, failing with MissingFile if not found.
xmlCursorRequired :: Zip.Archive -> FilePath -> Parser Cursor
xmlCursorRequired ar fname = do
    entry <- note (MissingFile fname) $ Zip.findEntryByPath fname ar
    cur <- left (\_ -> InvalidFile fname) $ parseLBS def (Zip.fromEntry entry)
    return $ fromDocument cur

-- | Get shared string table
getSharedStrings  :: Zip.Archive -> Parser SharedStringTable
getSharedStrings x = maybe sstEmpty (head . fromCursor) <$>
                     xmlCursorOptional x "xl/sharedStrings.xml"

getContentTypes :: Zip.Archive -> Parser ContentTypes
getContentTypes x = head . fromCursor <$> xmlCursorRequired x "[Content_Types].xml"

getStyles :: Zip.Archive -> Styles
getStyles ar = case Zip.fromEntry <$> Zip.findEntryByPath "xl/styles.xml" ar of
  Nothing  -> Styles L.empty
  Just xml -> Styles xml

getComments :: Zip.Archive -> FilePath -> Parser (Maybe CommentTable)
getComments ar fp = (listToMaybe . fromCursor =<<) <$> xmlCursorOptional ar fp

getCustomProperties :: Zip.Archive -> Parser CustomProperties
getCustomProperties ar = maybe CustomProperties.empty (head . fromCursor) <$> xmlCursorOptional ar "docProps/custom.xml"

getDrawing :: Zip.Archive -> ContentTypes ->  FilePath -> Parser Drawing
getDrawing ar contentTypes fp = do
    cur <- xmlCursorRequired ar fp
    drawingRels <- getRels ar fp
    unresolved <- headErr (InvalidFile fp) (fromCursor cur)
    anchors <- forM (unresolved ^. xdrAnchors) $ resolveFileInfo drawingRels
    return $ Drawing anchors
  where
    resolveFileInfo :: Relationships -> Anchor RefId -> Parser (Anchor FileInfo)
    resolveFileInfo rels uAnch = case uAnch ^. anchObject of
        pic@Picture{} -> do
            let mRefId = pic ^. picBlipFill . bfpImageInfo
            mFI <- lookupFI rels mRefId
            return uAnch{ _anchObject = pic & picBlipFill . bfpImageInfo .~ mFI }
    lookupFI _ Nothing = return Nothing
    lookupFI rels (Just rId) = do
        path <- relTarget <$> note (InvalidRef fp rId) (Relationships.lookup rId rels)
        -- content types use paths starting with /
        contentType <- note (InvalidFile path) $ ContentTypes.lookup ("/" <> path) contentTypes
        contents <- Zip.fromEntry <$> note (MissingFile path) (Zip.findEntryByPath path ar)
        return . Just $ FileInfo (stripMediaPrefix path) contentType contents
    stripMediaPrefix :: FilePath -> FilePath
    stripMediaPrefix p = fromMaybe p $ stripPrefix "xl/media/" p

-- | readWorkbook pulls the names of the sheets and the defined names
readWorkbook :: Zip.Archive -> Parser ([WorksheetFile], DefinedNames)
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

      names = cur $/ element (n"definedNames") &/ element (n"definedName") >=> mkDefinedName
  sheets <- sequence $
    cur $/ element (n"sheets") &/ element (n"sheet") >=>
      liftA2 (worksheetFile wbPath wbRels) <$> attribute "name" <*> (attribute (odr"id") &| RefId)
  return (sheets, DefinedNames names)


worksheetFile :: FilePath -> Relationships -> Text -> RefId -> Parser WorksheetFile
worksheetFile parentPath wbRels name rId = WorksheetFile name <$> path
  where
    path :: Parser FilePath
    path = relTarget <$> note (InvalidRef parentPath rId) (Relationships.lookup rId wbRels)

getRels :: Zip.Archive -> FilePath -> Parser Relationships
getRels ar fp = do
    let (dir, file) = splitFileName fp
        relsPath = dir </> "_rels" </> file <.> "rels"
    c <- xmlCursorOptional ar relsPath
    return $ maybe Relationships.empty (setTargetsFrom fp . head . fromCursor) c

int :: Text -> Int
int = either error fst . T.decimal
