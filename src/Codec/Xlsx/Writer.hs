{-# LANGUAGE OverloadedStrings #-}
-- | This module provides a function for serializing structured `Xlsx` into lazy bytestring
module Codec.Xlsx.Writer
    ( fromXlsx
    ) where

import           Control.Applicative
import qualified Codec.Archive.Zip as Zip
import           Control.Arrow (second)
import           Control.Lens hiding (transform)
import qualified Data.ByteString.Lazy as L
import           Data.ByteString.Lazy.Char8 ()
import           Data.Map (Map)
import qualified Data.Map as M
import           Data.Maybe
import           Data.Monoid ((<>))
import           Data.Text (Text, )
import qualified Data.Text as T
import           Data.Text.Lazy (toStrict)
import           Data.Text.Lazy.Builder (toLazyText)
import           Data.Text.Lazy.Builder.Int
import           Data.Text.Lazy.Builder.RealFloat
import           System.Locale
import           System.Time
import           Text.XML

import           Codec.Xlsx.Types
import           Codec.Xlsx.Types.SharedStringTable
import           Codec.Xlsx.Writer.Internal

-- | Writes `Xlsx' to raw data (lazy bytestring)
fromXlsx :: ClockTime -> Xlsx -> L.ByteString
fromXlsx ct xlsx =
    Zip.fromArchive $ foldr Zip.addEntryToArchive Zip.emptyArchive entries
  where
    TOD t _ = ct
    entries = Zip.toEntry "[Content_Types].xml" t (contentTypesXml files) :
              map (\fd -> Zip.toEntry (T.unpack $ fdName fd) t (fdContents fd)) files
    files = sheetFiles ++
      [ FileData "docProps/core.xml"
        "application/vnd.openxmlformats-package.core-properties+xml" $ coreXml (toUTCTime ct) "xlsxwriter"
      , FileData "docProps/app.xml"
        "application/vnd.openxmlformats-officedocument.extended-properties+xml" $ appXml (xlsx ^. xlSheets)
      , FileData "xl/workbook.xml"
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml" $ bookXml (xlsx ^. xlSheets) (xlsx ^. xlDefinedNames)
      , FileData "xl/styles.xml"
        "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml" $ unStyles (xlsx ^. xlStyles)
      , FileData "xl/sharedStrings.xml"
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml" $ ssXml shared
      , FileData "xl/_rels/workbook.xml.rels"
        "application/vnd.openxmlformats-package.relationships+xml" $ bookRelXml sheetCount
      , FileData "_rels/.rels" "application/vnd.openxmlformats-package.relationships+xml" rootRelXml
      ]
    sheetFiles =
      [ FileData ("xl/worksheets/sheet" <> txti n <> ".xml")
        "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml" $
        sheetXml (w ^. wsColumns) (w ^. wsRowPropertiesMap) cells (w ^. wsMerges) (w ^. wsSheetViews) (w ^. wsPageSetup) |
        (n, cells, w) <- zip3 [1..] sheetCells sheets]
    sheets = xlsx ^. xlSheets . to M.elems
    sheetCount = length sheets
    shared = sstConstruct sheets
    sheetCells = map (transformSheetData shared) sheets

data FileData = FileData { fdName :: Text
                         , fdContentType :: Text
                         , fdContents :: L.ByteString}

coreXml :: CalendarTime -> Text -> L.ByteString
coreXml created creator =
  renderLBS def{rsNamespaces=nss} $ Document (Prologue [] Nothing []) root []
  where
    nss = [ ("cp", "http://schemas.openxmlformats.org/package/2006/metadata/core-properties")
          , ("dc", "http://purl.org/dc/elements/1.1/")
          , ("dcterms", "http://purl.org/dc/terms/")
          , ("xsi","http://www.w3.org/2001/XMLSchema-instance")
          ]
    namespaced = nsName nss
    date = T.pack $ formatCalendarTime defaultTimeLocale "%Y-%m-%dT%H:%M:%SZ" created
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
    { xlsxCellStyle  :: Maybe Int
    , xlsxCellValue  :: Maybe XlsxCellData
    } deriving (Show, Eq)

xlsxCellType :: XlsxCell -> Text
xlsxCellType XlsxCell{xlsxCellValue=Just(XlsxSS _)} = "s"
xlsxCellType XlsxCell{xlsxCellValue=Just(XlsxBool _)} = "b"
xlsxCellType _ = "n" -- default in SpreadsheetML schema, TODO: add other types

value :: XlsxCell -> Text
value XlsxCell{xlsxCellValue=Just(XlsxSS i)} = txti i
value XlsxCell{xlsxCellValue=Just(XlsxDouble d)} = txtd d
value XlsxCell{xlsxCellValue=Just(XlsxBool True)} = "1"
value XlsxCell{xlsxCellValue=Just(XlsxBool False)} = "0"
value _ = error "value undefined"

transformSheetData :: SharedStringTable -> Worksheet -> [(Int, [(Int, XlsxCell)])]
transformSheetData shared ws = map transformRow $ toRows (ws ^. wsCells)
  where
    transformRow = second (map transformCell)
    transformCell (c, Cell{_cellValue=v, _cellStyle=s}) =
        (c, XlsxCell s (fmap transformValue v))
    transformValue (CellText t) = XlsxSS (sstLookupText shared t)
    transformValue (CellDouble dbl) =  XlsxDouble dbl
    transformValue (CellBool b) = XlsxBool b
    transformValue (CellRich r) = XlsxSS (sstLookupRich shared r)

sheetXml :: [ColumnsWidth] -> Map Int RowProperties -> [(Int, [(Int, XlsxCell)])] -> [Text]-> Maybe [SheetView] -> Maybe PageSetup -> L.ByteString
sheetXml cws rh rows merges sheetViews pageSetup = renderLBS def $ Document (Prologue [] Nothing []) root []
  where
    cType = xlsxCellType
    root = addNS "http://schemas.openxmlformats.org/spreadsheetml/2006/main" $
           Element "worksheet" M.empty $ catMaybes
           [renderSheetViews <$> sheetViews,
            nonEmptyNmEl "cols" M.empty $  map cwEl cws,
            justNmEl "sheetData" M.empty $ map rowEl rows,
            nonEmptyNmEl "mergeCells" M.empty $ map mergeE1 merges,
            NodeElement . toElement "pageSetup" <$> pageSetup]
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
              [nEl "v" M.empty [NodeContent $ value cell] | isJust $ xlsxCellValue cell]
    cellAttrs ref cell = cellStyleAttr cell ++ [("r", ref), ("t", cType cell)]
    cellStyleAttr XlsxCell{xlsxCellStyle=Nothing} = []
    cellStyleAttr XlsxCell{xlsxCellStyle=Just s} = [("s", txti s)]

bookXml :: Map Text Worksheet -> DefinedNames -> L.ByteString
bookXml wss (DefinedNames names) = renderLBS def $ Document (Prologue [] Nothing []) root []
  where
    numNames = [(txti i, name) | (i, name) <- zip [1..] (M.keys wss)]

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

bookRelXml :: Int -> L.ByteString
bookRelXml n = renderLBS def $ Document (Prologue [] Nothing []) root []
  where
    root = addNS "http://schemas.openxmlformats.org/package/2006/relationships" $
           Element "Relationships" M.empty $
           map (\sn -> relEl sn (T.concat ["worksheets/sheet", txti sn, ".xml"]) "worksheet") [1..n]
           ++
           [relEl (n + 1) "styles.xml" "styles", relEl (n + 2) "sharedStrings.xml" "sharedStrings"]
    relEl i target typ =
      nEl "Relationship"
      (M.fromList [("Id", T.concat ["rId", txti i]), ("Target", target),
       ("Type", T.concat ["http://schemas.openxmlformats.org/officeDocument/2006/relationships/", typ])]) []

rootRelXml :: L.ByteString
rootRelXml = "<?xml version=\"1.0\" encoding=\"UTF-8\"?>\
\<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"xl/workbook.xml\"/><Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties\" Target=\"docProps/core.xml\"/><Relationship Id=\"rId3\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties\" Target=\"docProps/app.xml\"/></Relationships>"

contentTypesXml :: [FileData] -> L.ByteString
contentTypesXml fds = renderLBS def $ Document (Prologue [] Nothing []) root []
  where
    root = addNS "http://schemas.openxmlformats.org/package/2006/content-types" $
           Element "Types" M.empty $
           map (\fd -> nEl "Override" (M.fromList  [("PartName", T.concat ["/", fdName fd]),
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
      ns = fromJust $ lookup p nss

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

txti :: Int -> Text
txti = toStrict . toLazyText . decimal

txtd :: Double -> Text
txtd = toStrict . toLazyText . realFloat

txtb :: Bool -> Text
txtb = T.toLower . T.pack . show
