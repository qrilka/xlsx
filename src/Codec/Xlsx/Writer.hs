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
import qualified Data.Map                                    as M
import           Data.Maybe
import           Data.Monoid                                 ((<>))
import           Data.Text                                   (Text)
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
        "xtended-properties" $ appXml (xlsx ^. xlSheets)
      , FileData "xl/workbook.xml"
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"
        "officeDocument" $ bookXml (xlsx ^. xlSheets) (xlsx ^. xlDefinedNames)
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
    sheets = xlsx ^. xlSheets . to M.elems
    sheetCount = length sheets
    shared = sstConstruct sheets
    sheetCells = map (transformSheetData shared) sheets

singleSheetFiles :: Int -> Cells -> Worksheet -> [FileData]
singleSheetFiles n cells ws = sheetFile:otherFiles
  where
    sheetFilePath = "xl/worksheets/sheet" <> show n <> ".xml"
    sheetFile = FileData sheetFilePath
        "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"
        "worksheet" $
        sheetXml ws cells drawingRId
    drawingRId = if null commentsReferenced then 1 else 2
    otherFiles = sheetRels ++ referencedFiles ++ extraFiles
    referencedFiles = commentsReferenced ++ drawingReferenced
    extraFiles = drawingExtra
    commentsReferenced = commentsFiles n cells
    sheetRels = if null referencedFiles
                then []
                else [ FileData ("xl/worksheets/_rels/sheet" <> show n <> ".xml.rels")
                      "application/vnd.openxmlformats-package.relationships+xml"
                       "relationships" sheetRelsXml ]
    sheetRelsXml = renderLBS def . toDocument . Relationships.fromList $
        [ relEntry i fdRelType (fdPath `relFrom` sheetFilePath)
        | (i, FileData{..}) <- zip [1..] referencedFiles ]
    (drawingReferenced, drawingExtra) = drawingFiles n (ws ^. wsDrawing)

commentsFiles :: Int -> Cells -> [FileData]
commentsFiles n cells =
    if null comments then [] else [commentsFile]
  where
    comments = concatMap (\(row, rowCells) -> mapMaybe (maybeCellComment row) rowCells) cells
    maybeCellComment row (col, cell) = do
        comment <- xlsxComment cell
        return (mkCellRef (row, col), comment)
    commentsFile = FileData commentsPath
        "application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml"
        "comments"
        commentsBS
    commentsPath = "xl/comments" <> show n <> ".xml"
    commentsBS = renderLBS def . toDocument $ CommentTable.fromList comments

drawingFiles :: Int -> Maybe Drawing -> ([FileData], [FileData])
drawingFiles _ Nothing = ([], [])
drawingFiles n(Just dr) = ([drawingFile], referenced)
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
     (anchors', images, _) = foldl' collectImage ([], [], 1) (dr ^. xdrAnchors)
     collectImage :: ([Anchor RefId], [Maybe FileInfo], Int) -> Anchor FileInfo
                  -> ([Anchor RefId], [Maybe FileInfo], Int)
     collectImage (as, fis, i) anch0 =
         case anch0 ^. anchObject of
             pic@Picture{} ->
                 let anch = anch0{_anchObject = pic & picBlipFill . bfpImageInfo ?~ RefId ("rId" <> txti i)}
                     fi = pic ^. picBlipFill . bfpImageInfo
                 in (anch:as, fi:fis, i + 1)
     imageFiles = [ FileData ("xl/media/" <> _fiFilename)
                    _fiContentType
                    "image" _fiContents
                  | FileInfo{..} <- reverse (catMaybes images) ]
     drawingRels = FileData ("xl/drawings/_rels/drawing" <> show n <> ".xml.rels")
                   "application/vnd.openxmlformats-package.relationships+xml"
                   "relationships" drawingRelsXml
     drawingRelsXml = renderLBS def . toDocument . Relationships.fromList $
        [ relEntry i fdRelType (fdPath `relFrom` drawingFilePath)
        | (i, FileData{..}) <- zip [1..] imageFiles ]
     referenced = case images of
         [] -> []
         _  -> drawingRels:imageFiles


data FileData = FileData { fdPath        :: FilePath
                         , fdContentType :: Text
                         , fdRelType     :: Text
                         , fdContents    :: L.ByteString }

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

appXml :: Map Text Worksheet -> L.ByteString
appXml s = renderLBS def $ Document (Prologue [] Nothing []) root []
  where
    nsAttrs = M.fromList [("xmlns:vt", "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes")]
    root = Element (extPropNm "Properties") nsAttrs
           [ extPropEl "TotalTime" [NodeContent "0"]
           , extPropEl "HeadingPairs" [
                            vTypeEl "vector" (M.fromList [("size", "2"), ("baseType", "variant")])
                                        [ vTypeEl0 "variant"
                                                       [vTypeEl0 "lpstr" [NodeContent "Worksheets"]]
                                        , vTypeEl0 "variant"
                                                       [vTypeEl0 "i4" [NodeContent $ txti $ M.size s]]
                                        ]
                           ]
           , extPropEl "TitlesOfParts" [
                            vTypeEl "vector" (M.fromList [("size", txti $ M.size s),("baseType","lpstr")]) $
                                    map (vTypeEl0 "lpstr" . return . NodeContent) $ M.keys s
                           ]
           ]
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

sheetXml :: Worksheet -> Cells -> Int -> L.ByteString
sheetXml ws rows nextRId = renderLBS def $ Document (Prologue [] Nothing []) root []
  where
    cws = ws ^. wsColumns
    rh = ws ^. wsRowPropertiesMap
    merges = ws ^. wsMerges
    sheetViews = ws ^. wsSheetViews
    pageSetup = ws ^. wsPageSetup
    drawing = ws ^. wsDrawing
    cfPairs = map CfPair . M.toList $ ws ^. wsConditionalFormattings
    root = addNS "http://schemas.openxmlformats.org/spreadsheetml/2006/main" $
           Element "worksheet" M.empty $ catMaybes $
           [ renderSheetViews <$> sheetViews,
             nonEmptyNmEl "cols" M.empty $  map cwEl cws,
             justNmEl "sheetData" M.empty $ map rowEl rows
           ] ++
           map (Just . NodeElement . toElement "conditionalFormatting") cfPairs ++
           [ nonEmptyNmEl "mergeCells" M.empty $ map mergeE1 merges,
             NodeElement . toElement "pageSetup" <$> pageSetup,
             NodeElement . (\_ -> leafElement  "drawing" [ odr "id" .= ("rId" <> txti nextRId) ]) <$> drawing
           ]
    cType = xlsxCellType
    cwEl cw = NodeElement $! Element "col" (M.fromList
              [("min", txti $ cwMin cw), ("max", txti $ cwMax cw), ("width", txtd $ cwWidth cw), ("style", txti $ cwStyle cw)]) []
    rowEl (r, cells) = nEl "row"
                       (M.fromList (ht ++ s ++ [("r", txti r) ,("hidden", "false"), ("outlineLevel", "0"),
                               ("collapsed", "false"), ("customFormat", "true"),
                               ("customHeight", txtb hasHeight)]))
                       $ map (cellEl r) cells
      where
        (ht, hasHeight, s) = case M.lookup r rh of
          Just (RowProps (Just h) (Just st)) -> ([("ht", txtd h)], True,[("s", txti st)])
          Just (RowProps Nothing  (Just st)) -> ([], True, [("s", txti st)])
          Just (RowProps (Just h) Nothing ) -> ([("ht", txtd h)], True,[])
          _ -> ([], False,[])
    mergeE1 t = NodeElement $! Element "mergeCell" (M.fromList [("ref",t)]) []
    cellEl r (icol, cell) =
      nEl "c" (M.fromList (cellAttrs (mkCellRef (r, icol)) cell))
              (catMaybes [ (\v -> nEl "v" M.empty [NodeContent v]) .  value <$> xlsxCellValue cell
                         , NodeElement . toElement "f" <$> xlsxCellFormula cell
                         ])
    cellAttrs ref cell = cellStyleAttr cell ++ [("r", ref), ("t", cType cell)]
    cellStyleAttr XlsxCell{xlsxCellStyle=Nothing} = []
    cellStyleAttr XlsxCell{xlsxCellStyle=Just s} = [("s", txti s)]

bookXml :: Map Text Worksheet -> DefinedNames -> L.ByteString
bookXml wss (DefinedNames names) = renderLBS def $ Document (Prologue [] Nothing []) root []
  where
    numNames = [(txti i, name) | (i, name) <- zip [(1::Int)..] (M.keys wss)]

    -- The @bookViews@ element is not required according to the schema, but its
    -- absence can cause Excel to crash when opening the print preview
    -- (see <https://phpexcel.codeplex.com/workitem/2935>). It suffices however
    -- to define a bookViews with a single empty @workbookView@ element
    -- (the @bookViews@ must contain at least one @wookbookView@).
    root = addNS "http://schemas.openxmlformats.org/spreadsheetml/2006/main" $ Element "workbook" M.empty
           [nEl "bookViews" M.empty [nEl "workbookView" M.empty []]
           ,nEl "sheets" M.empty $
            map (\(n, name) -> nEl "sheet"
                               (M.fromList [("name", name), ("sheetId", n), ("state", "visible"),
                                            (rId, T.concat ["rId", n])]) []) numNames
           ,nEl "definedNames" M.empty $ map (\(name, lsId, val) ->
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
    root = addNS "http://schemas.openxmlformats.org/package/2006/content-types" $
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

-- | Creates an element with the given name, attributes and children,
-- if there is at least one child. Otherwise returns `Nothing`.
nonEmptyNmEl :: Name -> Map Name Text -> [Node] -> Maybe Node
nonEmptyNmEl _ _ [] = Nothing
nonEmptyNmEl name attrs nodes = justNmEl name attrs nodes

-- | Creates an element with the given name, attributes and children.
-- Always returns a node/`Just`.
justNmEl :: Name -> Map Name Text -> [Node] -> Maybe Node
justNmEl name attrs nodes = Just $ nEl name attrs nodes

nEl :: Name -> Map Name Text -> [Node] -> Node
nEl name attrs nodes = NodeElement $ Element name attrs nodes
