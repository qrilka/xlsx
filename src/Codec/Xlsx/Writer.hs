{-# LANGUAGE TupleSections #-}
{-# LANGUAGE CPP               #-}
{-# LANGUAGE OverloadedStrings #-}
{-# LANGUAGE RecordWildCards   #-}
-- | This module provides a function for serializing structured `Xlsx` into lazy bytestring
module Codec.Xlsx.Writer
    ( fromXlsx
    ) where

import qualified Codec.Archive.Zip                           as Zip
import           Control.Arrow                               (second)
import           Control.Lens                                hiding (transform, (.=))
import qualified Data.ByteString.Lazy                        as L
import           Data.ByteString.Lazy.Char8                  ()
import           Data.List                                   (foldl')
import           Data.Map                                    (Map)
import           Data.STRef
import           Control.Monad.ST
import qualified Data.Map                                    as M
import           Data.Maybe
import           Data.Monoid                                 ((<>))
import           Data.Text                                   (Text)
import Data.Tuple.Extra (fst3, snd3, thd3)
import qualified Data.Text                                   as T
import           Data.Time                                   (UTCTime)
import           Data.Time.Clock.POSIX                       (POSIXTime, posixSecondsToUTCTime)
import           Data.Time.Format                            (formatTime)
#if MIN_VERSION_time(1,5,0)
import           Data.Time.Format                            (defaultTimeLocale)
#else
import           System.Locale                               (defaultTimeLocale)
#endif
import           Safe
import           Text.XML

#if !MIN_VERSION_base(4,8,0)
import           Control.Applicative
#endif

import           Codec.Xlsx.Types
import           Codec.Xlsx.Types.Internal
import           Codec.Xlsx.Types.Internal.CfPair
import qualified Codec.Xlsx.Types.Internal.CommentTable      as CommentTable
import           Codec.Xlsx.Types.Internal.CustomProperties
import           Codec.Xlsx.Types.Internal.DvPair
import           Codec.Xlsx.Types.Internal.Relationships     as Relationships hiding (lookup)
import           Codec.Xlsx.Types.Internal.SharedStringTable
import           Codec.Xlsx.Writer.Internal

-- | Writes `Xlsx' to raw data (lazy bytestring)
fromXlsx :: POSIXTime -> Xlsx -> L.ByteString
fromXlsx pt xlsx =
    Zip.fromArchive $ foldr Zip.addEntryToArchive Zip.emptyArchive entries
  where
    t = round pt
    utcTime = posixSecondsToUTCTime pt
    entries = Zip.toEntry "[Content_Types].xml" t (contentTypesXml files) :
              map (\fd -> Zip.toEntry (fdPath fd) t (fdContents fd)) files
    -- TODO: root files should be only core.xml, app.xml and workbook.xml
    files = sheetFiles ++ customPropFiles ++
      [ FileData "docProps/core.xml"
        "application/vnd.openxmlformats-package.core-properties+xml"
        "metadata/core-properties" $ coreXml utcTime "xlsxwriter"
      , FileData "docProps/app.xml"
        "application/vnd.openxmlformats-officedocument.extended-properties+xml"
        "xtended-properties" $ appXml sheetNames
      , FileData "xl/workbook.xml"
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"
        "officeDocument" $ bookXml sheetNames (xlsx ^. xlDefinedNames)
      , FileData "xl/styles.xml"
        "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"
        "styles" $ unStyles (xlsx ^. xlStyles)
      , FileData "xl/sharedStrings.xml"
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"
        "sharedStrings" $ ssXml shared
      , FileData "xl/_rels/workbook.xml.rels"
        "application/vnd.openxmlformats-package.relationships+xml" "relationships" bookRelsXml
      , FileData "_rels/.rels" "application/vnd.openxmlformats-package.relationships+xml"
        "relationships" rootRelXml
      ]
    rootRelXml = renderLBS def . toDocument $ Relationships.fromList rootRels
    rootFiles =  customPropFileRels ++
        [ ("officeDocument", "xl/workbook.xml")
        , ("metadata/core-properties", "docProps/core.xml")
        , ("extended-properties", "docProps/app.xml") ]
    rootRels = [ relEntry i typ trg
               | (i, (typ, trg)) <- zip [1..] rootFiles ]
    customProps = xlsx ^. xlCustomProperties
    (customPropFiles, customPropFileRels) = case M.null customProps of
        True  -> ([], [])
        False -> ([ FileData "docProps/custom.xml"
                    "application/vnd.openxmlformats-officedocument.custom-properties+xml"
                    "custom-properties"
                    (customPropsXml (CustomProperties customProps)) ],
                  [ ("custom-properties", "docProps/custom.xml") ])
    bookRelsXml = renderLBS def . toDocument $ bookRels sheetCount
    sheetFiles = concat $ zipWith3 singleSheetFiles [1..] sheetCells sheets
    sheetNames = xlsx ^. xlSheets . to (map fst)
    sheets = xlsx ^. xlSheets . to (map snd)
    sheetCount = length sheets
    shared = sstConstruct sheets
    sheetCells = map (transformSheetData shared) sheets

singleSheetFiles :: Int -> Cells -> Worksheet -> [FileData]
singleSheetFiles n cells ws = runST $ do
    ref <- newSTRef 1
    mCmntData <- genComments n cells ref
    mDrawingData <- maybe (return Nothing) (fmap Just . genDrawing n ref) (ws ^. wsDrawing)
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
            , nonEmptyElListSimple "cols" . map cwEl $ ws ^. wsColumns
            , Just . elementListSimple "sheetData" $ sheetDataXml cells (ws ^. wsRowPropertiesMap)
            ] ++
            map (Just . toElement "conditionalFormatting") cfPairs ++
            [ nonEmptyElListSimple "mergeCells" . map mergeE1 $ ws ^. wsMerges
            , Just $ countedElementList "dataValidations" $ map (toElement "dataValidation") dvPairs
            , toElement "pageSetup" <$> ws ^. wsPageSetup
            , fst3 <$> mDrawingData
            , fst <$> mCmntData
            ]
        cfPairs = map CfPair . M.toList $ ws ^. wsConditionalFormattings
        dvPairs = map DvPair . M.toList $ ws ^. wsDataValidations
        cwEl cw = leafElement "col" [ ("min", txti $ cwMin cw)
                                    , ("max", txti $ cwMax cw)
                                    , ("width", txtd $ cwWidth cw)
                                    , ("style", txti $ cwStyle cw)]
        mergeE1 t = leafElement "mergeCell" [("ref", t)]

        sheetRels = if null referencedFiles
                    then []
                    else [ FileData ("xl/worksheets/_rels/sheet" <> show n <> ".xml.rels")
                           "application/vnd.openxmlformats-package.relationships+xml"
                           "relationships" sheetRelsXml ]
        sheetRelsXml = renderLBS def . toDocument . Relationships.fromList $
            [ relEntry i fdRelType (fdPath `relFrom` sheetFilePath)
            | (i, FileData{..}) <- referenced ]
        referenced = fromMaybe [] (snd <$> mCmntData) ++
                     catMaybes [ snd3 <$> mDrawingData ]
        referencedFiles = map snd referenced
        extraFiles = maybe [] thd3 mDrawingData
        otherFiles = sheetRels ++ referencedFiles ++ extraFiles

    return (sheetFile:otherFiles)

sheetDataXml :: Cells -> Map Int RowProperties -> [Element]
sheetDataXml rows rh = map rowEl rows
  where
    rowEl (r, cells) = elementList "row"
                       (ht ++ s ++ [("r", txti r) ,("hidden", "false"), ("outlineLevel", "0"),
                                    ("collapsed", "false"), ("customFormat", "true"),
                                    ("customHeight", txtb hasHeight)])
                       $ map (cellEl r) cells
      where
        (ht, hasHeight, s) = case M.lookup r rh of
          Just (RowProps (Just h) (Just st)) -> ([("ht", txtd h)], True,[("s", txti st)])
          Just (RowProps Nothing  (Just st)) -> ([], True, [("s", txti st)])
          Just (RowProps (Just h) Nothing ) -> ([("ht", txtd h)], True,[])
          _ -> ([], False,[])
    cellEl r (icol, cell) =
      elementList "c" (cellAttrs (mkCellRef (r, icol)) cell)
              (catMaybes [ elementContent "v" . value <$> xlsxCellValue cell
                         , toElement "f" <$> xlsxCellFormula cell
                         ])
    cellAttrs ref cell = cellStyleAttr cell ++ [("r", ref), ("t", xlsxCellType cell)]
    cellStyleAttr XlsxCell{xlsxCellStyle=Nothing} = []
    cellStyleAttr XlsxCell{xlsxCellStyle=Just s} = [("s", txti s)]

genComments :: Int -> Cells -> STRef s Int -> ST s (Maybe (Element, [ReferencedFileData]))
genComments n cells ref =
    if null comments
    then do
        return Nothing
    else do
        rId1 <- readSTRef ref
        let rId2 = rId1 + 1
        modifySTRef' ref (+2)
        let el = refElement "legacyDrawing" rId2
        return $ Just (el, [(rId1, commentsFile), (rId2, vmlDrawingFile)])
  where
    comments = concatMap (\(row, rowCells) -> mapMaybe (maybeCellComment row) rowCells) cells
    maybeCellComment row (col, cell) = do
        comment <- xlsxComment cell
        return (mkCellRef (row, col), comment)
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
    rId <- readSTRef ref
    modifySTRef' ref (+1)
    let el = refElement "drawing" rId
    return (el, (rId, drawingFile), referenced)
  where
     drawingFilePath = "xl/drawings/drawing" <> show n <> ".xml"
     drawingFile = FileData drawingFilePath
                            "application/vnd.openxmlformats-officedocument.drawing+xml"
                            "drawing" drawingXml
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
       [ ( i
         , FileData ("xl/media/" <> _fiFilename) _fiContentType "image" _fiContents)
       | (i, FileInfo {..}) <- reverse (catMaybes images)
       ]

     chartFiles =
       [(i, genChart n k chart) | (k, (i, chart)) <- zip [1 ..] (reverse charts)]

     innerFiles = imageFiles ++ chartFiles

     drawingRels =
       FileData
         ("xl/drawings/_rels/drawing" <> show n <> ".xml.rels")
         "application/vnd.openxmlformats-package.relationships+xml"
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

data FileData = FileData { fdPath        :: FilePath
                         , fdContentType :: Text
                         , fdRelType     :: Text
                         , fdContents    :: L.ByteString }

type ReferencedFileData = (Int, FileData)

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

data XlsxCellData = XlsxSS Int
                  | XlsxDouble Double
                  | XlsxBool Bool
                    deriving (Show, Eq)
data XlsxCell = XlsxCell
    { xlsxCellStyle   :: Maybe Int
    , xlsxCellValue   :: Maybe XlsxCellData
    , xlsxComment     :: Maybe Comment
    , xlsxCellFormula :: Maybe CellFormula
    } deriving (Show, Eq)

xlsxCellType :: XlsxCell -> Text
xlsxCellType XlsxCell{xlsxCellValue=Just(XlsxSS _)} = "s"
xlsxCellType XlsxCell{xlsxCellValue=Just(XlsxBool _)} = "b"
xlsxCellType _ = "n" -- default in SpreadsheetML schema, TODO: add other types

value :: XlsxCellData -> Text
value (XlsxSS i)       = txti i
value (XlsxDouble d)   = txtd d
value (XlsxBool True)  = "1"
value (XlsxBool False) = "0"

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

bookXml :: [Text] -> DefinedNames -> L.ByteString
bookXml sheetNames (DefinedNames names) = renderLBS def $ Document (Prologue [] Nothing []) root []
  where
    numNames = [(txti i, name) | (i, name) <- zip [(1::Int)..] sheetNames]

    -- The @bookViews@ element is not required according to the schema, but its
    -- absence can cause Excel to crash when opening the print preview
    -- (see <https://phpexcel.codeplex.com/workitem/2935>). It suffices however
    -- to define a bookViews with a single empty @workbookView@ element
    -- (the @bookViews@ must contain at least one @wookbookView@).
    root = addNS "http://schemas.openxmlformats.org/spreadsheetml/2006/main" Nothing $
           Element "workbook" M.empty
             [ nEl "bookViews" M.empty [nEl "workbookView" M.empty []]
             , nEl "sheets" M.empty $
                 map (\(n, name) -> nEl "sheet"
                         (M.fromList [ ("name", name)
                                     , ("sheetId", n)
                                     , ("state", "visible")
                                     , (rId, "rId" <> n)]) [])
                 numNames
             , nEl "definedNames" M.empty $ map (\(name, lsId, val) ->
                 nEl "definedName" (definedName name lsId) [NodeContent val]) names
             ]

    rId = nm "http://schemas.openxmlformats.org/officeDocument/2006/relationships" "id"

    definedName :: Text -> Maybe Text -> Map Name Text
    definedName name Nothing     = M.fromList [("name", name)]
    definedName name (Just lsId) = M.fromList [("name", name), ("localSheetId", lsId)]

ssXml :: SharedStringTable -> L.ByteString
ssXml = renderLBS def . toDocument

customPropsXml :: CustomProperties -> L.ByteString
customPropsXml = renderLBS def . toDocument

bookRels :: Int -> Relationships
bookRels n =  Relationships.fromList (sheetRels ++ [stylesRel, ssRel])
  where
    sheetRels = [relEntry i "worksheet" ("worksheets/sheet" <> show i <> ".xml") | i <- [1..n]]
    stylesRel = relEntry (n + 1) "styles" "styles.xml"
    ssRel = relEntry (n + 2) "sharedStrings" "sharedStrings.xml"

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
refElement :: Name -> Int -> Element
refElement name rId = leafElement name [ odr "id" .= ("rId" <> txti rId) ]
