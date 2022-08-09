{-# LANGUAGE FlexibleContexts #-}
{-# LANGUAGE OverloadedStrings #-}
{-# LANGUAGE RecordWildCards #-}
{-# LANGUAGE TupleSections #-}
{-# LANGUAGE DeriveGeneric #-}
-- | This module provides a function for serializing structured `Xlsx` into lazy bytestring
module Codec.Xlsx.Writer
  ( fromXlsx
  ) where

import qualified Codec.Archive.Zip as Zip
import Control.Arrow (second)
import Control.Lens hiding (transform, (.=))
import Control.Monad (forM)
import Control.Monad.ST
import Control.Monad.State (evalState, get, put)
import qualified Data.ByteString.Lazy as L
import Data.ByteString.Lazy.Char8 ()
import Data.List (foldl', mapAccumL)
import Data.Map (Map)
import qualified Data.Map as M
import Data.Maybe
import Data.Monoid ((<>))
import Data.STRef
import Data.Text (Text)
import qualified Data.Text as T
import Data.Time (UTCTime)
import Data.Time.Clock.POSIX (POSIXTime, posixSecondsToUTCTime)
import Data.Time.Format (formatTime)
import Data.Time.Format (defaultTimeLocale)
import Data.Tuple.Extra (fst3, snd3, thd3)
import GHC.Generics (Generic)
import Safe
import Text.XML

import Codec.Xlsx.Types
import Codec.Xlsx.Types.Cell (applySharedFormulaOpts)
import Codec.Xlsx.Types.Internal
import Codec.Xlsx.Types.Internal.CfPair
import qualified Codec.Xlsx.Types.Internal.CommentTable
       as CommentTable
import Codec.Xlsx.Types.Internal.CustomProperties
import Codec.Xlsx.Types.Internal.DvPair
import Codec.Xlsx.Types.Internal.Relationships as Relationships
       hiding (lookup)
import Codec.Xlsx.Types.Internal.SharedStringTable
import Codec.Xlsx.Types.PivotTable.Internal
import Codec.Xlsx.Writer.Internal
import Codec.Xlsx.Writer.Internal.PivotTable

-- | Writes `Xlsx' to raw data (lazy bytestring)
fromXlsx :: POSIXTime -> Xlsx -> L.ByteString
fromXlsx pt xlsx =
    Zip.fromArchive $ foldr Zip.addEntryToArchive Zip.emptyArchive entries
  where
    t = round pt
    utcTime = posixSecondsToUTCTime pt
    entries = Zip.toEntry "[Content_Types].xml" t (contentTypesXml files) :
              map (\fd -> Zip.toEntry (fdPath fd) t (fdContents fd)) files
    files = workbookFiles ++ customPropFiles ++
      [ FileData "docProps/core.xml"
        "application/vnd.openxmlformats-package.core-properties+xml"
        "metadata/core-properties" $ coreXml utcTime "xlsxwriter"
      , FileData "docProps/app.xml"
        "application/vnd.openxmlformats-officedocument.extended-properties+xml"
        "xtended-properties" $ appXml sheetNames
      , FileData "_rels/.rels" "application/vnd.openxmlformats-package.relationships+xml"
        "relationships" rootRelXml
      ]
    rootRelXml = renderLBS def . toDocument $ Relationships.fromList rootRels
    rootFiles =  customPropFileRels ++
        [ ("officeDocument", "xl/workbook.xml")
        , ("metadata/core-properties", "docProps/core.xml")
        , ("extended-properties", "docProps/app.xml") ]
    rootRels = [ relEntry (unsafeRefId i) typ trg
               | (i, (typ, trg)) <- zip [1..] rootFiles ]
    customProps = xlsx ^. xlCustomProperties
    (customPropFiles, customPropFileRels) = case M.null customProps of
        True  -> ([], [])
        False -> ([ FileData "docProps/custom.xml"
                    "application/vnd.openxmlformats-officedocument.custom-properties+xml"
                    "custom-properties"
                    (customPropsXml (CustomProperties customProps)) ],
                  [ ("custom-properties", "docProps/custom.xml") ])
    workbookFiles = bookFiles xlsx
    sheetNames = xlsx ^. xlSheets & mapped %~ view _1

singleSheetFiles :: Int
                 -> Cells
                 -> [FileData]
                 -> Worksheet
                 -> STRef s Int
                 -> ST s (FileData, [FileData])
singleSheetFiles n cells pivFileDatas ws tblIdRef = do
    ref <- newSTRef 1
    mCmntData <- genComments n cells ref
    mDrawingData <- maybe (return Nothing) (fmap Just . genDrawing n ref) (ws ^. wsDrawing)
    pivRefs <- forM pivFileDatas $ \fd -> do
      refId <- nextRefId ref
      return (refId, fd)
    refTables <- forM (_wsTables ws) $ \tbl -> do
      refId <- nextRefId ref
      tblId <- readSTRef tblIdRef
      modifySTRef' tblIdRef (+1)
      return (refId, genTable tbl tblId)
    let sheetFilePath = "xl/worksheets/sheet" <> show n <> ".xml"
        sheetFile = FileData sheetFilePath
            "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"
            "worksheet" $
            sheetXml
        nss = [ ("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships") ]
        sheetXml= renderLBS def{rsNamespaces=nss} $ Document (Prologue [] Nothing []) root []
        root = addNS "http://schemas.openxmlformats.org/spreadsheetml/2006/main" Nothing $
            elementListSimple "worksheet" rootEls
        rootEls = catMaybes $
            [ elementListSimple "sheetViews" . map (toElement "sheetView") <$> ws ^. wsSheetViews
            , nonEmptyElListSimple "cols" . map (toElement "col") $ ws ^. wsColumnsProperties
            , Just . elementListSimple "sheetData" $
              sheetDataXml cells (ws ^. wsRowPropertiesMap) (ws ^. wsSharedFormulas)
            , toElement "sheetProtection" <$> (ws ^. wsProtection)
            , toElement "autoFilter" <$> (ws ^. wsAutoFilter)
            , nonEmptyElListSimple "mergeCells" . map mergeE1 $ ws ^. wsMerges
            ] ++ map (Just . toElement "conditionalFormatting") cfPairs ++
            [ nonEmptyElListSimple "dataValidations" $ map (toElement "dataValidation") dvPairs
            , toElement "pageSetup" <$> ws ^. wsPageSetup
            , fst3 <$> mDrawingData
            , fst <$> mCmntData
            , nonEmptyElListSimple "tableParts"
                [leafElement "tablePart" [odr "id" .= rId] | (rId, _) <- refTables]
            ]
        cfPairs = map CfPair . M.toList $ ws ^. wsConditionalFormattings
        dvPairs = map DvPair . M.toList $ ws ^. wsDataValidations
        mergeE1 r = leafElement "mergeCell" [("ref" .= r)]

        sheetRels = if null referencedFiles
                    then []
                    else [ FileData ("xl/worksheets/_rels/sheet" <> show n <> ".xml.rels")
                           "application/vnd.openxmlformats-package.relationships+xml"
                           "relationships" sheetRelsXml ]
        sheetRelsXml = renderLBS def . toDocument . Relationships.fromList $
            [ relEntry i fdRelType (fdPath `relFrom` sheetFilePath)
            | (i, FileData{..}) <- referenced ]
        referenced = fromMaybe [] (snd <$> mCmntData) ++
                     catMaybes [ snd3 <$> mDrawingData ] ++
                     pivRefs ++
                     refTables
        referencedFiles = map snd referenced
        extraFiles = maybe [] thd3 mDrawingData
        otherFiles = sheetRels ++ referencedFiles ++ extraFiles

    return (sheetFile, otherFiles)

nextRefId :: STRef s Int -> ST s RefId
nextRefId r = do
  num <- readSTRef r
  modifySTRef' r (+1)
  return (unsafeRefId num)

sheetDataXml ::
     Cells
  -> Map Int RowProperties
  -> Map SharedFormulaIndex SharedFormulaOptions
  -> [Element]
sheetDataXml rows rh sharedFormulas =
  evalState (mapM rowEl rows) sharedFormulas
  where
    rowEl (r, cells) = do
      let mProps    = M.lookup r rh
          hasHeight = case rowHeight =<< mProps of
                        Just CustomHeight{} -> True
                        _                   -> False
          ht        = do Just height <- [rowHeight =<< mProps]
                         let h = case height of CustomHeight    x -> x
                                                AutomaticHeight x -> x
                         return ("ht", txtd h)
          s         = do Just st <- [rowStyle =<< mProps]
                         return ("s", txti st)
          hidden    = fromMaybe False $ rowHidden <$> mProps
          attrs = ht ++
            s ++
            [ ("r", txti r)
            , ("hidden", txtb hidden)
            , ("outlineLevel", "0")
            , ("collapsed", "false")
            , ("customFormat", "true")
            , ("customHeight", txtb hasHeight)
            ]
      cellEls <- mapM (cellEl r) cells
      return $ elementList "row" attrs cellEls
    cellEl r (icol, cell) = do
      let cellAttrs ref c =
            cellStyleAttr c ++ [("r" .= ref), ("t" .= xlsxCellType c)]
          cellStyleAttr XlsxCell{xlsxCellStyle=Nothing} = []
          cellStyleAttr XlsxCell{xlsxCellStyle=Just s} = [("s", txti s)]
          formula = xlsxCellFormula cell
          fEl0 = toElement "f" <$> formula
      fEl <- case formula of
        Just CellFormula{_cellfExpression=SharedFormula si} -> do
          shared <- get
          case M.lookup si shared of
            Just fOpts -> do
              put $ M.delete si shared
              return $ applySharedFormulaOpts fOpts <$> fEl0
            Nothing ->
              return fEl0
        _ ->
          return fEl0
      return $ elementList "c" (cellAttrs (singleCellRef (r, icol)) cell) $
        catMaybes [fEl, elementContent "v" . value <$> xlsxCellValue cell]

genComments :: Int -> Cells -> STRef s Int -> ST s (Maybe (Element, [ReferencedFileData]))
genComments n cells ref =
    if null comments
    then do
        return Nothing
    else do
        rId1 <- nextRefId ref
        rId2 <- nextRefId ref
        let el = refElement "legacyDrawing" rId2
        return $ Just (el, [(rId1, commentsFile), (rId2, vmlDrawingFile)])
  where
    comments = concatMap (\(row, rowCells) -> mapMaybe (maybeCellComment row) rowCells) cells
    maybeCellComment row (col, cell) = do
        comment <- xlsxComment cell
        return (singleCellRef (row, col), comment)
    commentTable = CommentTable.fromList comments
    commentsFile = FileData commentsPath
        "application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml"
        "comments"
        commentsBS
    commentsPath = "xl/comments" <> show n <> ".xml"
    commentsBS = renderLBS def $ toDocument commentTable
    vmlDrawingFile = FileData vmlPath
        "application/vnd.openxmlformats-officedocument.vmlDrawing"
        "vmlDrawing"
        vmlDrawingBS
    vmlPath = "xl/drawings/vmlDrawing" <> show n <> ".vml"
    vmlDrawingBS = CommentTable.renderShapes commentTable

genDrawing :: Int -> STRef s Int -> Drawing -> ST s (Element, ReferencedFileData, [FileData])
genDrawing n ref dr = do
  rId <- nextRefId ref
  let el = refElement "drawing" rId
  return (el, (rId, drawingFile), referenced)
  where
    drawingFilePath = "xl/drawings/drawing" <> show n <> ".xml"
    drawingCT = "application/vnd.openxmlformats-officedocument.drawing+xml"
    drawingFile = FileData drawingFilePath drawingCT "drawing" drawingXml
    drawingXml = renderLBS def{rsNamespaces=nss} $ toDocument dr'
    nss = [ ("xdr", "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing")
          , ("a",   "http://schemas.openxmlformats.org/drawingml/2006/main")
          , ("r",   "http://schemas.openxmlformats.org/officeDocument/2006/relationships") ]
    dr' = Drawing{ _xdrAnchors = reverse anchors' }
    (anchors', images, charts, _) = foldl' collectFile ([], [], [], 1) (dr ^. xdrAnchors)
    collectFile :: ([Anchor RefId RefId], [Maybe (Int, FileInfo)], [(Int, ChartSpace)], Int)
                -> Anchor FileInfo ChartSpace
                -> ([Anchor RefId RefId], [Maybe (Int, FileInfo)], [(Int, ChartSpace)], Int)
    collectFile (as, fis, chs, i) anch0 =
        case anch0 ^. anchObject of
          Picture {..} ->
            let fi = (i,) <$> _picBlipFill ^. bfpImageInfo
                pic' =
                  Picture
                  { _picMacro = _picMacro
                  , _picPublished = _picPublished
                  , _picNonVisual = _picNonVisual
                  , _picBlipFill =
                      (_picBlipFill & bfpImageInfo ?~ RefId ("rId" <> txti i))
                  , _picShapeProperties = _picShapeProperties
                  }
                anch = anch0 {_anchObject = pic'}
            in (anch : as, fi : fis, chs, i + 1)
          Graphic nv ch tr ->
            let gr' = Graphic nv (RefId ("rId" <> txti i)) tr
                anch = anch0 {_anchObject = gr'}
            in (anch : as, fis, (i, ch) : chs, i + 1)
    imageFiles =
      [ ( unsafeRefId i
        , FileData ("xl/media/" <> _fiFilename) _fiContentType "image" _fiContents)
      | (i, FileInfo {..}) <- reverse (catMaybes images)
      ]

    chartFiles =
      [ (unsafeRefId i, genChart n k chart)
      | (k, (i, chart)) <- zip [1 ..] (reverse charts)
      ]

    innerFiles = imageFiles ++ chartFiles

    drawingRels =
      FileData
        ("xl/drawings/_rels/drawing" <> show n <> ".xml.rels")
        relsCT
        "relationships"
        drawingRelsXml

    drawingRelsXml =
      renderLBS def . toDocument . Relationships.fromList $
      map (refFileDataToRel drawingFilePath) innerFiles

    referenced =
      case innerFiles of
        [] -> []
        _ -> drawingRels : (map snd innerFiles)

genChart :: Int -> Int -> ChartSpace -> FileData
genChart n i ch = FileData path contentType relType contents
  where
    path = "xl/charts/chart" <> show n <> "_" <> show i <> ".xml"
    contentType =
      "application/vnd.openxmlformats-officedocument.drawingml.chart+xml"
    relType = "chart"
    contents = renderLBS def {rsNamespaces = nss} $ toDocument ch
    nss =
      [ ("c", "http://schemas.openxmlformats.org/drawingml/2006/chart")
      , ("a", "http://schemas.openxmlformats.org/drawingml/2006/main")
      ]

data PvGenerated = PvGenerated
  { pvgCacheFiles :: [(CacheId, FileData)]
  , pvgSheetTableFiles :: [[FileData]]
  , pvgOthers :: [FileData]
  }

generatePivotFiles :: [(CellMap, [PivotTable])] -> PvGenerated
generatePivotFiles cmTables = PvGenerated cacheFiles shTableFiles others
  where
    cacheFiles = [cacheFile | (cacheFile, _, _) <- flatRendered]
    shTableFiles = map (map (\(_, tableFile, _) -> tableFile)) rendered
    others = concat [other | (_, _, other) <- flatRendered]
    firstCacheId = 1
    flatRendered = concat rendered
    (_, rendered) =
      mapAccumL
        (\c (cm, ts) -> mapAccumL (\c' t -> (c' + 1, render cm c' t)) c ts)
        firstCacheId
        cmTables
    render cm cacheIdRaw tbl =
      let PivotTableFiles {..} = renderPivotTableFiles cm cacheIdRaw tbl
          cacheId = CacheId cacheIdRaw
          cacheIdStr = show cacheIdRaw
          cachePath =
            "xl/pivotCache/pivotCacheDefinition" <> cacheIdStr <> ".xml"
          cacheFile =
            FileData
              cachePath
              (smlCT "pivotCacheDefinition")
              "pivotCacheDefinition"
              pvtfCacheDefinition
          recordsPath =
            "xl/pivotCache/pivotCacheRecords" <> cacheIdStr <> ".xml"
          recordsFile =
            FileData
            recordsPath
            (smlCT "pivotCacheRecords")
            "pivotCacheRecords"
            pvtfCacheRecords
          cacheRelsFile =
            FileData
            ("xl/pivotCache/_rels/pivotCacheDefinition" <> cacheIdStr <> ".xml.rels")
            relsCT
            "relationships" $
            renderRels [refFileDataToRel cachePath (unsafeRefId 1, recordsFile)]
          renderRels = renderLBS def . toDocument . Relationships.fromList
          tablePath = "xl/pivotTables/pivotTable" <> cacheIdStr <> ".xml"
          tableFile =
            FileData tablePath (smlCT "pivotTable") "pivotTable" pvtfTable
          tableRels =
            FileData
              ("xl/pivotTables/_rels/pivotTable" <> cacheIdStr <> ".xml.rels")
              relsCT
              "relationships" $
            renderRels [refFileDataToRel tablePath (unsafeRefId 1, cacheFile)]
      in ((cacheId, cacheFile), tableFile, [tableRels, cacheRelsFile, recordsFile])

genTable :: Table -> Int -> FileData
genTable tbl tblId = FileData{..}
  where
    fdPath = "xl/tables/table" <> show tblId <> ".xml"
    fdContentType = smlCT "table"
    fdRelType = "table"
    fdContents = renderLBS def $ tableToDocument tbl tblId

data FileData = FileData { fdPath        :: FilePath
                         , fdContentType :: Text
                         , fdRelType     :: Text
                         , fdContents    :: L.ByteString }

type ReferencedFileData = (RefId, FileData)

refFileDataToRel :: FilePath -> ReferencedFileData -> (RefId, Relationship)
refFileDataToRel basePath (i, FileData {..}) =
    relEntry i fdRelType (fdPath `relFrom` basePath)

type Cells = [(Int, [(Int, XlsxCell)])]

coreXml :: UTCTime -> Text -> L.ByteString
coreXml created creator =
  renderLBS def{rsNamespaces=nss} $ Document (Prologue [] Nothing []) root []
  where
    nss = [ ("cp", "http://schemas.openxmlformats.org/package/2006/metadata/core-properties")
          , ("dc", "http://purl.org/dc/elements/1.1/")
          , ("dcterms", "http://purl.org/dc/terms/")
          , ("xsi","http://www.w3.org/2001/XMLSchema-instance")
          ]
    namespaced = nsName nss
    date = T.pack $ formatTime defaultTimeLocale "%FT%T%QZ" created
    root = Element (namespaced "cp" "coreProperties") M.empty
           [ nEl (namespaced "dcterms" "created")
                     (M.fromList [(namespaced "xsi" "type", "dcterms:W3CDTF")]) [NodeContent date]
           , nEl (namespaced "dc" "creator") M.empty [NodeContent creator]
           , nEl (namespaced "cp" "lastModifiedBy") M.empty [NodeContent creator]
           ]

appXml :: [Text] -> L.ByteString
appXml sheetNames =
    renderLBS def $ Document (Prologue [] Nothing []) root []
  where
    sheetCount = length sheetNames
    root = Element (extPropNm "Properties") nsAttrs
           [ extPropEl "TotalTime" [NodeContent "0"]
           , extPropEl "HeadingPairs" [
                   vTypeEl "vector" (M.fromList [ ("size", "2")
                                                , ("baseType", "variant")])
                       [ vTypeEl0 "variant"
                           [vTypeEl0 "lpstr" [NodeContent "Worksheets"]]
                       , vTypeEl0 "variant"
                           [vTypeEl0 "i4" [NodeContent $ txti sheetCount]]
                       ]
                   ]
           , extPropEl "TitlesOfParts" [
                   vTypeEl "vector" (M.fromList [ ("size",     txti sheetCount)
                                                , ("baseType", "lpstr")]) $
                       map (vTypeEl0 "lpstr" . return . NodeContent) sheetNames
                   ]
           ]
    nsAttrs = M.fromList [("xmlns:vt", "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes")]
    extPropNm n = nm "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" n
    extPropEl n = nEl (extPropNm n) M.empty
    vTypeEl0 n = vTypeEl n M.empty
    vTypeEl = nEl . nm "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"

data XlsxCellData
  = XlsxSS Int
  | XlsxDouble Double
  | XlsxBool Bool
  | XlsxError ErrorType
  deriving (Eq, Show, Generic)

data XlsxCell = XlsxCell
    { xlsxCellStyle   :: Maybe Int
    , xlsxCellValue   :: Maybe XlsxCellData
    , xlsxComment     :: Maybe Comment
    , xlsxCellFormula :: Maybe CellFormula
    } deriving (Eq, Show, Generic)

xlsxCellType :: XlsxCell -> Text
xlsxCellType XlsxCell{xlsxCellValue=Just(XlsxSS _)} = "s"
xlsxCellType XlsxCell{xlsxCellValue=Just(XlsxBool _)} = "b"
xlsxCellType XlsxCell{xlsxCellValue=Just(XlsxError _)} = "e"
xlsxCellType _ = "n" -- default in SpreadsheetML schema, TODO: add other types

value :: XlsxCellData -> Text
value (XlsxSS i)       = txti i
value (XlsxDouble d)   = txtd d
value (XlsxBool True)  = "1"
value (XlsxBool False) = "0"
value (XlsxError eType) = toAttrVal eType

transformSheetData :: SharedStringTable -> Worksheet -> Cells
transformSheetData shared ws = map transformRow $ toRows (ws ^. wsCells)
  where
    transformRow = second (map transformCell)
    transformCell (c, Cell{..}) =
        (c, XlsxCell _cellStyle (fmap transformValue _cellValue) _cellComment _cellFormula)
    transformValue (CellText t) = XlsxSS (sstLookupText shared t)
    transformValue (CellDouble dbl) =  XlsxDouble dbl
    transformValue (CellBool b) = XlsxBool b
    transformValue (CellRich r) = XlsxSS (sstLookupRich shared r)
    transformValue (CellError e) = XlsxError e

bookFiles :: Xlsx -> [FileData]
bookFiles xlsx = runST $ do
  ref <- newSTRef 1
  ssRId <- nextRefId ref
  let sheets = xlsx ^. xlSheets & mapped %~ view _3
      shared = sstConstruct sheets
      sharedStrings =
        (ssRId, FileData "xl/sharedStrings.xml" (smlCT "sharedStrings") "sharedStrings" $
              ssXml shared)
  stRId <- nextRefId ref
  let style =
        (stRId, FileData "xl/styles.xml" (smlCT "styles") "styles" $
              unStyles (xlsx ^. xlStyles))
  let PvGenerated { pvgCacheFiles = cacheIdFiles
                  , pvgOthers = pivotOtherFiles
                  , pvgSheetTableFiles = sheetPivotTables
                  } =
        generatePivotFiles
          [ (_wsCells, _wsPivotTables)
          | (_, _, Worksheet {..}) <- xlsx ^. xlSheets
          ]
      sheetCells = map (transformSheetData shared) sheets
      sheetInputs = zip3 sheetCells sheetPivotTables sheets
  tblIdRef <- newSTRef 1
  allSheetFiles <- forM (zip [1..] sheetInputs) $ \(i, (cells, pvTables, sheet)) -> do
    rId <- nextRefId ref
    (sheetFile, others) <- singleSheetFiles i cells pvTables sheet tblIdRef
    return ((rId, sheetFile), others)
  let sheetFiles = map fst allSheetFiles
      sheetAttrsByRId = zipWith (\(rId, _) (name, state, _) -> (rId, name, state)) sheetFiles (xlsx ^. xlSheets)
      sheetOthers = concatMap snd allSheetFiles
  cacheRefFDsById <- forM cacheIdFiles $ \(cacheId, fd) -> do
      refId <- nextRefId ref
      return (cacheId, (refId, fd))
  let cacheRefsById = [ (cId, rId) | (cId, (rId, _)) <- cacheRefFDsById ]
      cacheRefs = map snd cacheRefFDsById
      bookFile = FileData "xl/workbook.xml" (smlCT "sheet.main") "officeDocument" $
                 bookXml sheetAttrsByRId (xlsx ^. xlDefinedNames) cacheRefsById (xlsx ^. xlDateBase)
      rels = FileData "xl/_rels/workbook.xml.rels"
             "application/vnd.openxmlformats-package.relationships+xml"
             "relationships" relsXml
      relsXml = renderLBS def . toDocument . Relationships.fromList $
            [ relEntry i fdRelType (fdPath `relFrom` "xl/workbook.xml")
            | (i, FileData{..}) <- referenced ]
      referenced = sharedStrings:style:sheetFiles ++ cacheRefs
      otherFiles = concat [rels:(map snd referenced), pivotOtherFiles, sheetOthers]
  return $ bookFile:otherFiles

bookXml :: [(RefId, Text, SheetState)]
        -> DefinedNames
        -> [(CacheId, RefId)]
        -> DateBase
        -> L.ByteString
bookXml rIdAttrs (DefinedNames names) cacheIdRefs dateBase =
  renderLBS def {rsNamespaces = nss} $ Document (Prologue [] Nothing []) root []
  where
    nss = [ ("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships") ]
    -- The @bookViews@ element is not required according to the schema, but its
    -- absence can cause Excel to crash when opening the print preview
    -- (see <https://phpexcel.codeplex.com/workitem/2935>). It suffices however
    -- to define a bookViews with a single empty @workbookView@ element
    -- (the @bookViews@ must contain at least one @wookbookView@).
    root =
      addNS "http://schemas.openxmlformats.org/spreadsheetml/2006/main" Nothing $
      elementListSimple
        "workbook"
        ( [ leafElement "workbookPr" (catMaybes ["date1904" .=? justTrue (dateBase == DateBase1904) ])
          , elementListSimple "bookViews" [emptyElement "workbookView"]
          , elementListSimple
            "sheets"
            [ leafElement
              "sheet"
              ["name" .= name, "sheetId" .= i, "state" .= state, (odr "id") .= rId]
            | (i, (rId, name, state)) <- zip [(1 :: Int) ..] rIdAttrs
            ]
          , elementListSimple
            "definedNames"
            [ elementContent0 "definedName" (definedName name lsId) val
            | (name, lsId, val) <- names
            ]
          ] ++
          maybeToList
          (nonEmptyElListSimple "pivotCaches" $ map pivotCacheEl cacheIdRefs)
        )

    pivotCacheEl (CacheId cId, refId) =
      leafElement "pivotCache" ["cacheId" .= cId, (odr "id") .= refId]

    definedName :: Text -> Maybe Text -> [(Name, Text)]
    definedName name Nothing = ["name" .= name]
    definedName name (Just lsId) = ["name" .= name, "localSheetId" .= lsId]

ssXml :: SharedStringTable -> L.ByteString
ssXml = renderLBS def . toDocument

customPropsXml :: CustomProperties -> L.ByteString
customPropsXml = renderLBS def . toDocument

contentTypesXml :: [FileData] -> L.ByteString
contentTypesXml fds = renderLBS def $ Document (Prologue [] Nothing []) root []
  where
    root = addNS "http://schemas.openxmlformats.org/package/2006/content-types" Nothing $
           Element "Types" M.empty $
           map (\fd -> nEl "Override" (M.fromList  [("PartName", T.concat ["/", T.pack $ fdPath fd]),
                                       ("ContentType", fdContentType fd)]) []) fds

-- | fully qualified XML name
qName :: Text -> Text -> Text -> Name
qName n ns p =
    Name
    { nameLocalName = n
    , nameNamespace = Just ns
    , namePrefix = Just p
    }

-- | fully qualified XML name from prefix to ns URL mapping
nsName :: [(Text, Text)] -> Text -> Text -> Name
nsName nss p n = qName n ns p
    where
      ns = fromJustNote "ns name lookup" $ lookup p nss

nm :: Text -> Text -> Name
nm ns n = Name
  { nameLocalName = n
  , nameNamespace = Just ns
  , namePrefix = Nothing}

nEl :: Name -> Map Name Text -> [Node] -> Node
nEl name attrs nodes = NodeElement $ Element name attrs nodes

-- | Creates element holding reference to some linked file
refElement :: Name -> RefId -> Element
refElement name rId = leafElement name [ odr "id" .= rId ]

smlCT :: Text -> Text
smlCT t =
  "application/vnd.openxmlformats-officedocument.spreadsheetml." <> t <> "+xml"

relsCT :: Text
relsCT = "application/vnd.openxmlformats-package.relationships+xml"
