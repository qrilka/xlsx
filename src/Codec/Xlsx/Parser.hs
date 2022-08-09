{-# LANGUAGE DeriveGeneric #-}
{-# LANGUAGE NoMonomorphismRestriction #-}
{-# LANGUAGE OverloadedStrings         #-}
{-# LANGUAGE RecordWildCards           #-}
{-# LANGUAGE ScopedTypeVariables       #-}
{-# LANGUAGE TupleSections             #-}

-- | This module provides a function for reading .xlsx files
module Codec.Xlsx.Parser
  ( toXlsx
  , toXlsxEither
  , toXlsxFast
  , toXlsxEitherFast
  , ParseError(..)
  , Parser
  ) where

import qualified Codec.Archive.Zip as Zip
import Control.Applicative
import Control.Arrow (left)
import Control.Error.Safe (headErr)
import Control.Error.Util (note)
import Control.Exception (Exception)
import Control.Lens hiding ((<.>), element, views)
import Control.Monad (forM, join, void)
import Control.Monad.Except (catchError, throwError)
import Data.Bool (bool)
import Data.ByteString (ByteString)
import qualified Data.ByteString.Lazy as L
import qualified Data.ByteString.Lazy as LB
import Data.ByteString.Lazy.Char8 ()
import Data.List
import Data.Map (Map)
import qualified Data.Map as M
import Data.Maybe
import Data.Monoid
import Data.Text (Text)
import qualified Data.Text as T
import Data.Traversable
import GHC.Generics (Generic)
import Prelude hiding (sequence)
import Safe
import System.FilePath.Posix
import Text.XML as X
import Text.XML.Cursor hiding (bool)
import qualified Xeno.DOM as Xeno

import Codec.Xlsx.Parser.Internal
import Codec.Xlsx.Parser.Internal.PivotTable
import Codec.Xlsx.Types
import Codec.Xlsx.Types.Cell (formulaDataFromCursor)
import Codec.Xlsx.Types.Common (xlsxTextToCellValue)
import Codec.Xlsx.Types.Internal
import Codec.Xlsx.Types.Internal.CfPair
import Codec.Xlsx.Types.Internal.CommentTable as CommentTable
import Codec.Xlsx.Types.Internal.ContentTypes as ContentTypes
import Codec.Xlsx.Types.Internal.CustomProperties
       as CustomProperties
import Codec.Xlsx.Types.Internal.DvPair
import Codec.Xlsx.Types.Internal.FormulaData
import Codec.Xlsx.Types.Internal.Relationships as Relationships
import Codec.Xlsx.Types.Internal.SharedStringTable
import Codec.Xlsx.Types.PivotTable.Internal
import Codec.Xlsx.Types.SheetState as SheetState

-- | Reads `Xlsx' from raw data (lazy bytestring)
toXlsx :: L.ByteString -> Xlsx
toXlsx = either (error . show) id . toXlsxEither

data ParseError = InvalidZipArchive
                | MissingFile FilePath
                | InvalidFile FilePath Text
                | InvalidRef FilePath RefId
                | InconsistentXlsx Text
                deriving (Eq, Show, Generic)

instance Exception ParseError

type Parser = Either ParseError

-- | Reads `Xlsx' from raw data (lazy bytestring) using @xeno@ library
-- using some "cheating":
--
-- * not doing 100% xml validation
-- * replacing only <https://www.w3.org/TR/REC-xml/#sec-predefined-ent predefined entities>
--   and <https://www.w3.org/TR/REC-xml/#NT-CharRef Unicode character references>
--   (without checking codepoint validity)
-- * almost not using XML namespaces
toXlsxFast :: L.ByteString -> Xlsx
toXlsxFast = either (error . show) id . toXlsxEitherFast

-- | Reads `Xlsx' from raw data (lazy bytestring), failing with 'Left' on parse error
toXlsxEither :: L.ByteString -> Parser Xlsx
toXlsxEither = toXlsxEitherBase extractSheet

-- | Fast parsing with 'Left' on parse error, see 'toXlsxFast'
toXlsxEitherFast :: L.ByteString -> Parser Xlsx
toXlsxEitherFast = toXlsxEitherBase extractSheetFast

toXlsxEitherBase ::
     (Zip.Archive -> SharedStringTable -> ContentTypes -> Caches -> WorksheetFile -> Parser Worksheet)
  -> L.ByteString
  -> Parser Xlsx
toXlsxEitherBase parseSheet bs = do
  ar <- left (const InvalidZipArchive) $ Zip.toArchiveOrFail bs
  sst <- getSharedStrings ar
  contentTypes <- getContentTypes ar
  (wfs, names, cacheSources, dateBase) <- readWorkbook ar
  sheets <- forM wfs $ \wf -> do
      sheet <- parseSheet ar sst contentTypes cacheSources wf
      return (wfName wf, wsState wf, sheet)
  CustomProperties customPropMap <- getCustomProperties ar
  return $ Xlsx sheets (getStyles ar) names customPropMap dateBase

data WorksheetFile = WorksheetFile { wfName :: Text
                                   , wsState :: SheetState
                                   , wfPath :: FilePath
                                   }
                   deriving (Show, Generic)

type Caches = [(CacheId, (Text, CellRef, [CacheField]))]

extractSheetFast :: Zip.Archive
                 -> SharedStringTable
                 -> ContentTypes
                 -> Caches
                 -> WorksheetFile
                 -> Parser Worksheet
extractSheetFast ar sst contentTypes caches wf = do
  file <-
    note (MissingFile filePath) $
    Zip.fromEntry <$> Zip.findEntryByPath filePath ar
  sheetRels <- getRels ar filePath
  root <-
    left (\ex -> InvalidFile filePath $ T.pack (show ex)) $
    Xeno.parse (LB.toStrict file)
  parseWorksheet root sheetRels
  where
    filePath = wfPath wf
    parseWorksheet :: Xeno.Node -> Relationships -> Parser Worksheet
    parseWorksheet root sheetRels = do
      let prefixes = nsPrefixes root
          odrNs =
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
          odrX = addPrefix prefixes odrNs
          skip = void . maybeChild
      (ws, tableIds, drawingRId, legacyDrRId) <-
        liftEither . collectChildren root $ do
          skip "sheetPr"
          skip "dimension"
          _wsSheetViews <- fmap justNonEmpty . maybeParse "sheetViews" $ \n ->
            collectChildren n $ fromChildList "sheetView"
          skip "sheetFormatPr"
          _wsColumnsProperties <-
            fmap (fromMaybe []) . maybeParse "cols" $ \n ->
              collectChildren n (fromChildList "col")
          (_wsRowPropertiesMap, _wsCells, _wsSharedFormulas) <-
            requireAndParse "sheetData" $ \n -> do
              rows <- collectChildren n $ childList "row"
              collectRows <$> forM rows parseRow
          skip "sheetCalcPr"
          _wsProtection <- maybeFromChild "sheetProtection"
          skip "protectedRanges"
          skip "scenarios"
          _wsAutoFilter <- maybeFromChild "autoFilter"
          skip "sortState"
          skip "dataConsolidate"
          skip "customSheetViews"
          _wsMerges <- fmap (fromMaybe []) . maybeParse "mergeCells" $ \n -> do
            mCells <- collectChildren n $ childList "mergeCell"
            forM mCells $ \mCell -> parseAttributes mCell $ fromAttr "ref"
          _wsConditionalFormattings <-
            M.fromList . map unCfPair <$> fromChildList "conditionalFormatting"
          _wsDataValidations <-
            fmap (fromMaybe mempty) . maybeParse "dataValidations" $ \n -> do
              M.fromList . map unDvPair <$>
                collectChildren n (fromChildList "dataValidation")
          skip "hyperlinks"
          skip "printOptions"
          skip "pageMargins"
          _wsPageSetup <- maybeFromChild "pageSetup"
          skip "headerFooter"
          skip "rowBreaks"
          skip "colBreaks"
          skip "customProperties"
          skip "cellWatches"
          skip "ignoredErrors"
          skip "smartTags"
          drawingRId <- maybeParse "drawing" $ \n ->
            parseAttributes n $ fromAttr (odrX "id")
          legacyDrRId <- maybeParse "legacyDrawing" $ \n ->
            parseAttributes n $ fromAttr (odrX "id")
          tableIds <- fmap (fromMaybe []) . maybeParse "tableParts" $ \n -> do
            tParts <- collectChildren n $ childList "tablePart"
            forM tParts $ \part ->
              parseAttributes part $ fromAttr (odrX "id")

          -- all explicitly assigned fields filled below
          return (
            Worksheet
            { _wsDrawing = Nothing
            , _wsPivotTables = []
            , _wsTables = []
            , ..
            }
            , tableIds
            , drawingRId
            , legacyDrRId)

      let commentsType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments"
          commentTarget :: Maybe FilePath
          commentTarget = relTarget <$> findRelByType commentsType sheetRels
          legacyDrPath = fmap relTarget . flip Relationships.lookup sheetRels =<< legacyDrRId
      commentsMap <-
        fmap join . forM commentTarget $ getComments ar legacyDrPath
      let commentCells =
            M.fromList
            [ (fromSingleCellRefNoting r, def { _cellComment = Just cmnt})
            | (r, cmnt) <- maybe [] CommentTable.toList commentsMap
            ]
          assignComment withCmnt noCmnt =
            noCmnt & cellComment .~ (withCmnt ^. cellComment)
          mergeComments = M.unionWith assignComment commentCells
      tables <- forM tableIds $ \rId -> do
        fp <- lookupRelPath filePath sheetRels rId
        getTable ar fp
      drawing <- forM drawingRId $ \dId -> do
        rel <- note (InvalidRef filePath dId) $ Relationships.lookup dId sheetRels
        getDrawing ar contentTypes (relTarget rel)
      let ptType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable"
      pivotTables <- forM (allByType ptType sheetRels) $ \rel -> do
        let ptPath = relTarget rel
        bs <- note (MissingFile ptPath) $ Zip.fromEntry <$> Zip.findEntryByPath ptPath ar
        note (InconsistentXlsx $ "Bad pivot table in " <> T.pack ptPath) $
          parsePivotTable (flip Prelude.lookup caches) bs

      return $ ws & wsTables .~ tables
                  & wsCells %~ mergeComments
                  & wsDrawing .~ drawing
                  & wsPivotTables .~ pivotTables
    liftEither :: Either Text a -> Parser a
    liftEither = left (\t -> InvalidFile filePath t)
    justNonEmpty v@(Just (_:_)) = v
    justNonEmpty _ = Nothing
    collectRows = foldr collectRow (M.empty, M.empty, M.empty)
    collectRow ::
         ( Int
         , Maybe RowProperties
         , [(Int, Int, Cell, Maybe (SharedFormulaIndex, SharedFormulaOptions))])
      -> ( Map Int RowProperties
         , CellMap
         , Map SharedFormulaIndex SharedFormulaOptions)
      -> ( Map Int RowProperties
         , CellMap
         , Map SharedFormulaIndex SharedFormulaOptions)
    collectRow (r, mRP, rowCells) (rowMap, cellMap, sharedF) =
      let (newCells0, newSharedF0) =
            unzip [(((x, y), cd), shared) | (x, y, cd, shared) <- rowCells]
          newCells = M.fromAscList newCells0
          newSharedF = M.fromAscList $ catMaybes newSharedF0
          newRowMap =
            case mRP of
              Just rp -> M.insert r rp rowMap
              Nothing -> rowMap
      in (newRowMap, cellMap <> newCells, sharedF <> newSharedF)
    parseRow ::
         Xeno.Node
      -> Either Text ( Int
                     , Maybe RowProperties
                     , [( Int
                        , Int
                        , Cell
                        , Maybe (SharedFormulaIndex, SharedFormulaOptions))])
    parseRow row = do
      (r, s, ht, cstHt, hidden) <-
        parseAttributes row $
        ((,,,,) <$> fromAttr "r" <*> maybeAttr "s" <*> maybeAttr "ht" <*>
         fromAttrDef "customHeight" False <*>
         fromAttrDef "hidden" False)
      let props =
            RowProps
            { rowHeight =
                if cstHt
                  then CustomHeight <$> ht
                  else AutomaticHeight <$> ht
            , rowStyle = s
            , rowHidden = hidden
            }
      cellNodes <- collectChildren row $ childList "c"
      cells <- forM cellNodes parseCell
      return
        ( r
        , if props == def
            then Nothing
            else Just props
        , cells)
    parseCell ::
         Xeno.Node
      -> Either Text ( Int
                     , Int
                     , Cell
                     , Maybe (SharedFormulaIndex, SharedFormulaOptions))
    parseCell cell = do
      (ref, s, t) <-
        parseAttributes cell $
        (,,) <$> fromAttr "r" <*> maybeAttr "s" <*> fromAttrDef "t" "n"
      (fNode, vNode, isNode) <-
        collectChildren cell $
        (,,) <$> maybeChild "f" <*> maybeChild "v" <*> maybeChild "is"
      let vConverted :: (FromAttrBs a) => Either Text (Maybe a)
          vConverted =
            case contentBs <$> vNode of
              Nothing -> return Nothing
              Just c -> Just <$> fromAttrBs c
      mFormulaData <- mapM fromXenoNode fNode
      d <-
        case t of
          ("s" :: ByteString) -> do
            si <- vConverted
            case sstItem sst =<< si of
              Just xlTxt -> return $ Just (xlsxTextToCellValue xlTxt)
              Nothing -> throwError "bad shared string index"
          "inlineStr" -> mapM (fmap xlsxTextToCellValue . fromXenoNode) isNode
          "str" -> fmap CellText <$> vConverted
          "n" -> fmap CellDouble <$> vConverted
          "b" -> fmap CellBool <$> vConverted
          "e" -> fmap CellError <$> vConverted
          unexpected ->
            throwError $ "unexpected cell type " <> T.pack (show unexpected)
      let (r, c) = fromSingleCellRefNoting ref
          f = frmdFormula <$> mFormulaData
          shared = frmdShared =<< mFormulaData
      return (r, c, Cell s d Nothing f, shared)

extractSheet ::
     Zip.Archive
  -> SharedStringTable
  -> ContentTypes
  -> Caches
  -> WorksheetFile
  -> Parser Worksheet
extractSheet ar sst contentTypes caches wf = do
  let filePath = wfPath wf
  file <- note (MissingFile filePath) $ Zip.fromEntry <$> Zip.findEntryByPath filePath ar
  cur <- fmap fromDocument . left (\ex -> InvalidFile filePath (T.pack $ show ex)) $
         parseLBS def file
  sheetRels <- getRels ar filePath

  -- The specification says the file should contain either 0 or 1 @sheetViews@
  -- (4th edition, section 18.3.1.88, p. 1704 and definition CT_Worksheet, p. 3910)
  let  sheetViewList = cur $/ element (n_ "sheetViews") &/ element (n_ "sheetView") >=> fromCursor
       sheetViews = case sheetViewList of
         []    -> Nothing
         views -> Just views

  let commentsType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments"
      commentTarget :: Maybe FilePath
      commentTarget = relTarget <$> findRelByType commentsType sheetRels
      legacyDrRId = cur $/ element (n_ "legacyDrawing") >=> fromAttribute (odr"id")
      legacyDrPath = fmap relTarget . flip Relationships.lookup sheetRels  =<< listToMaybe legacyDrRId

  commentsMap :: Maybe CommentTable <- maybe (Right Nothing) (getComments ar legacyDrPath) commentTarget

  -- Likewise, @pageSetup@ also occurs either 0 or 1 times
  let pageSetup = listToMaybe $ cur $/ element (n_ "pageSetup") >=> fromCursor

      cws = cur $/ element (n_ "cols") &/ element (n_ "col") >=> fromCursor

      (rowProps, cells0, sharedFormulas) =
        collect $ cur $/ element (n_ "sheetData") &/ element (n_ "row") >=> parseRow
      parseRow ::
           Cursor
        -> [( Int
            , Maybe RowProperties
            , [(Int, Int, Cell, Maybe (SharedFormulaIndex, SharedFormulaOptions))])]
      parseRow c = do
        r <- fromAttribute "r" c
        let prop = RowProps
              { rowHeight = do h <- listToMaybe $ fromAttribute "ht" c
                               case fromAttribute "customHeight" c of
                                 [True] -> return $ CustomHeight    h
                                 _      -> return $ AutomaticHeight h
              , rowStyle  = listToMaybe $ fromAttribute "s" c
              , rowHidden =
                  case fromAttribute "hidden" c of
                    []  -> False
                    f:_ -> f
              }
        return ( r
               , if prop == def then Nothing else Just prop
               , c $/ element (n_ "c") >=> parseCell
               )
      parseCell ::
           Cursor
        -> [(Int, Int, Cell, Maybe (SharedFormulaIndex, SharedFormulaOptions))]
      parseCell cell = do
        ref <- fromAttribute "r" cell
        let s = listToMaybe $ cell $| attribute "s" >=> decimal
            t = fromMaybe "n" $ listToMaybe $ cell $| attribute "t"
            d = listToMaybe $ extractCellValue sst t cell
            mFormulaData = listToMaybe $ cell $/ element (n_ "f") >=> formulaDataFromCursor
            f = fst <$> mFormulaData
            shared = snd =<< mFormulaData
            (r, c) = fromSingleCellRefNoting ref
            comment = commentsMap >>= lookupComment ref
        return (r, c, Cell s d comment f, shared)
      collect = foldr collectRow (M.empty, M.empty, M.empty)
      collectRow ::
           ( Int
           , Maybe RowProperties
           , [(Int, Int, Cell, Maybe (SharedFormulaIndex, SharedFormulaOptions))])
        -> (Map Int RowProperties, CellMap, Map SharedFormulaIndex SharedFormulaOptions)
        -> (Map Int RowProperties, CellMap, Map SharedFormulaIndex SharedFormulaOptions)
      collectRow (r, mRP, rowCells) (rowMap, cellMap, sharedF) =
        let (newCells0, newSharedF0) =
              unzip [(((x,y),cd), shared) | (x, y, cd, shared) <- rowCells]
            newCells = M.fromList newCells0
            newSharedF = M.fromList $ catMaybes newSharedF0
            newRowMap = case mRP of
              Just rp -> M.insert r rp rowMap
              Nothing -> rowMap
        in (newRowMap, cellMap <> newCells, sharedF <> newSharedF)

      commentCells =
        M.fromList
          [ (fromSingleCellRefNoting r, def {_cellComment = Just cmnt})
          | (r, cmnt) <- maybe [] CommentTable.toList commentsMap
          ]
      cells = cells0 `M.union` commentCells

      mProtection = listToMaybe $ cur $/ element (n_ "sheetProtection") >=> fromCursor

      mDrawingId = listToMaybe $ cur $/ element (n_ "drawing") >=> fromAttribute (odr"id")

      merges = cur $/ parseMerges
      parseMerges :: Cursor -> [Range]
      parseMerges = element (n_ "mergeCells") &/ element (n_ "mergeCell") >=> fromAttribute "ref"

      condFormtattings = M.fromList . map unCfPair  $ cur $/ element (n_ "conditionalFormatting") >=> fromCursor

      validations = M.fromList . map unDvPair $
          cur $/ element (n_ "dataValidations") &/ element (n_ "dataValidation") >=> fromCursor

      tableIds =
        cur $/ element (n_ "tableParts") &/ element (n_ "tablePart") >=>
        fromAttribute (odr "id")

  let mAutoFilter = listToMaybe $ cur $/ element (n_ "autoFilter") >=> fromCursor

  mDrawing <- case mDrawingId of
      Just dId -> do
          rel <- note (InvalidRef filePath dId) $ Relationships.lookup dId sheetRels
          Just <$> getDrawing ar contentTypes (relTarget rel)
      Nothing  ->
          return Nothing

  let ptType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable"
  pTables <- forM (allByType ptType sheetRels) $ \rel -> do
    let ptPath = relTarget rel
    bs <- note (MissingFile ptPath) $ Zip.fromEntry <$> Zip.findEntryByPath ptPath ar
    note (InconsistentXlsx $ "Bad pivot table in " <> T.pack ptPath) $
      parsePivotTable (flip Prelude.lookup caches) bs

  tables <- forM tableIds $ \rId -> do
    fp <- lookupRelPath filePath sheetRels rId
    getTable ar fp

  return $
    Worksheet
      cws
      rowProps
      cells
      mDrawing
      merges
      sheetViews
      pageSetup
      condFormtattings
      validations
      pTables
      mAutoFilter
      tables
      mProtection
      sharedFormulas

extractCellValue :: SharedStringTable -> Text -> Cursor -> [CellValue]
extractCellValue sst t cur
  | t == "s" = do
    si <- vConverted "shared string"
    case sstItem sst si of
      Just xlTxt -> return $ xlsxTextToCellValue xlTxt
      Nothing -> fail "bad shared string index"
  | t == "inlineStr" =
    cur $/ element (n_ "is") >=> fmap xlsxTextToCellValue . fromCursor
  | t == "str" = CellText <$> vConverted "string"
  | t == "n" = CellDouble <$> vConverted "double"
  | t == "b" = CellBool <$> vConverted "boolean"
  | t == "e" = CellError <$> vConverted "error"
  | otherwise = fail "bad cell value"
  where
    vConverted typeStr = do
      vContent <- cur $/ element (n_ "v") >=> \c ->
        return (T.concat $ c $/ content)
      case fromAttrVal vContent of
        Right (val, _) -> return $ val
        _ -> fail $ "bad " ++ typeStr ++ " cell value"

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
    cur <- left (\ex -> InvalidFile fname (T.pack $ show ex)) $ parseLBS def (Zip.fromEntry entry)
    return $ fromDocument cur

fromFileCursorDef ::
     FromCursor a => Zip.Archive -> FilePath -> Text -> a -> Parser a
fromFileCursorDef x fp contentsDescr defVal = do
  mCur <- xmlCursorOptional x fp
  case mCur of
    Just cur ->
      headErr (InvalidFile fp $ "Couldn't parse " <> contentsDescr) $ fromCursor cur
    Nothing -> return defVal

fromFileCursor :: FromCursor a => Zip.Archive -> FilePath -> Text -> Parser a
fromFileCursor x fp contentsDescr = do
  cur <- xmlCursorRequired x fp
  headErr (InvalidFile fp $ "Couldn't parse " <> contentsDescr) $ fromCursor cur

-- | Get shared string table
getSharedStrings  :: Zip.Archive -> Parser SharedStringTable
getSharedStrings x =
  fromFileCursorDef x "xl/sharedStrings.xml" "shared strings" sstEmpty

getContentTypes :: Zip.Archive -> Parser ContentTypes
getContentTypes x = fromFileCursor x "[Content_Types].xml" "content types"

getStyles :: Zip.Archive -> Styles
getStyles ar = case Zip.fromEntry <$> Zip.findEntryByPath "xl/styles.xml" ar of
  Nothing  -> Styles L.empty
  Just xml -> Styles xml

getComments :: Zip.Archive -> Maybe FilePath -> FilePath -> Parser (Maybe CommentTable)
getComments ar drp fp = do
    mCurComments <- xmlCursorOptional ar fp
    mCurDr <- maybe (return Nothing) (xmlCursorOptional ar) drp
    return (liftA2 hide (hidden <$> mCurDr) . listToMaybe . fromCursor =<< mCurComments)
  where
    hide refs (CommentTable m) = CommentTable $ foldl' hideComment m refs
    hideComment m r = M.adjust (\c->c{_commentVisible = False}) r m
    v nm = Name nm (Just "urn:schemas-microsoft-com:vml") Nothing
    x nm = Name nm (Just "urn:schemas-microsoft-com:office:excel") Nothing
    hidden :: Cursor -> [CellRef]
    hidden cur = cur $/ checkElement visibleShape &/
                 element (x"ClientData") >=> shapeCellRef
    visibleShape Element{..} = elementName ==  (v"shape") &&
        maybe False (any ("visibility:hidden"==) . T.split (==';')) (M.lookup "style" elementAttributes)
    shapeCellRef :: Cursor -> [CellRef]
    shapeCellRef c = do
        r0 <- c $/ element (x"Row") &/ content >=> decimal
        c0 <- c $/ element (x"Column") &/ content >=> decimal
        return $ singleCellRef (r0 + 1, c0 + 1)

getCustomProperties :: Zip.Archive -> Parser CustomProperties
getCustomProperties ar =
  fromFileCursorDef ar "docProps/custom.xml" "custom properties" CustomProperties.empty

getDrawing :: Zip.Archive -> ContentTypes ->  FilePath -> Parser Drawing
getDrawing ar contentTypes fp = do
    cur <- xmlCursorRequired ar fp
    drawingRels <- getRels ar fp
    unresolved <- headErr (InvalidFile fp "Couldn't parse drawing") (fromCursor cur)
    anchors <- forM (unresolved ^. xdrAnchors) $ resolveFileInfo drawingRels
    return $ Drawing anchors
  where
    resolveFileInfo :: Relationships -> Anchor RefId RefId -> Parser (Anchor FileInfo ChartSpace)
    resolveFileInfo rels uAnch =
      case uAnch ^. anchObject of
        Picture {..} -> do
          let mRefId = _picBlipFill ^. bfpImageInfo
          mFI <- lookupFI rels mRefId
          let pic' =
                Picture
                { _picMacro = _picMacro
                , _picPublished = _picPublished
                , _picNonVisual = _picNonVisual
                , _picBlipFill = (_picBlipFill & bfpImageInfo .~ mFI)
                , _picShapeProperties = _picShapeProperties
                }
          return uAnch {_anchObject = pic'}
        Graphic nv rId tr -> do
          chartPath <- lookupRelPath fp rels rId
          chart <- readChart ar chartPath
          return uAnch {_anchObject = Graphic nv chart tr}
    lookupFI _ Nothing = return Nothing
    lookupFI rels (Just rId) = do
      path <- lookupRelPath fp rels rId
        -- content types use paths starting with /
      contentType <-
        note (InvalidFile path "Missing content type") $
        ContentTypes.lookup ("/" <> path) contentTypes
      contents <-
        Zip.fromEntry <$> note (MissingFile path) (Zip.findEntryByPath path ar)
      return . Just $ FileInfo (stripMediaPrefix path) contentType contents
    stripMediaPrefix :: FilePath -> FilePath
    stripMediaPrefix p = fromMaybe p $ stripPrefix "xl/media/" p

readChart :: Zip.Archive -> FilePath -> Parser ChartSpace
readChart ar path = fromFileCursor ar path "chart"

-- | readWorkbook pulls the names of the sheets and the defined names
readWorkbook :: Zip.Archive -> Parser ([WorksheetFile], DefinedNames, Caches, DateBase)
readWorkbook ar = do
  let wbPath = "xl/workbook.xml"
  cur <- xmlCursorRequired ar wbPath
  wbRels <- getRels ar wbPath
  -- Specification says the 'name' is required.
  let mkDefinedName :: Cursor -> [(Text, Maybe Text, Text)]
      mkDefinedName c =
        return
          ( headNote "Missing name attribute" $ attribute "name" c
          , listToMaybe $ attribute "localSheetId" c
          , T.concat $ c $/ content)
      names =
        cur $/ element (n_ "definedNames") &/ element (n_ "definedName") >=>
        mkDefinedName
  sheets <-
    sequence $
    cur $/ element (n_ "sheets") &/ element (n_ "sheet") >=>
    liftA3 (worksheetFile wbPath wbRels) <$> attribute "name" <*> fromAttributeDef "state" SheetState.Visible <*>
    fromAttribute (odr "id")
  let cacheRefs =
        cur $/ element (n_ "pivotCaches") &/ element (n_ "pivotCache") >=>
        liftA2 (,) <$> fromAttribute "cacheId" <*> fromAttribute (odr "id")
  caches <-
    forM cacheRefs $ \(cacheId, rId) -> do
      path <- lookupRelPath wbPath wbRels rId
      bs <-
        note (MissingFile path) $ Zip.fromEntry <$> Zip.findEntryByPath path ar
      (sheet, ref, fields0, mRecRId) <-
        note (InconsistentXlsx $ "Bad pivot table cache in " <> T.pack path) $
        parseCache bs
      fields <- case mRecRId of
        Just recId -> do
          cacheRels <- getRels ar path
          recsPath <- lookupRelPath path cacheRels recId
          rCur <- xmlCursorRequired ar recsPath
          let recs = rCur $/ element (n_ "r") >=> \cur' ->
                return $ cur' $/ anyElement >=> recordValueFromNode . node
          return $ fillCacheFieldsFromRecords fields0 recs
        Nothing ->
          return fields0
      return $ (cacheId, (sheet, ref, fields))
  let dateBase = bool DateBase1900 DateBase1904 . fromMaybe False . listToMaybe $
                 cur $/ element (n_ "workbookPr") >=> fromAttribute "date1904"
  return (sheets, DefinedNames names, caches, dateBase)

getTable :: Zip.Archive -> FilePath -> Parser Table
getTable ar fp = do
  cur <- xmlCursorRequired ar fp
  headErr (InvalidFile fp "Couldn't parse drawing") (fromCursor cur)

worksheetFile :: FilePath -> Relationships -> Text -> SheetState -> RefId -> Parser WorksheetFile
worksheetFile parentPath wbRels name visibility rId =
  WorksheetFile name visibility <$> lookupRelPath parentPath wbRels rId

getRels :: Zip.Archive -> FilePath -> Parser Relationships
getRels ar fp = do
    let (dir, file) = splitFileName fp
        relsPath = dir </> "_rels" </> file <.> "rels"
    c <- xmlCursorOptional ar relsPath
    return $ maybe Relationships.empty (setTargetsFrom fp . headNote "Missing rels" . fromCursor) c

lookupRelPath :: FilePath
              -> Relationships
              -> RefId
              -> Either ParseError FilePath
lookupRelPath fp rels rId =
  relTarget <$> note (InvalidRef fp rId) (Relationships.lookup rId rels)
