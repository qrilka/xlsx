{-# LANGUAGE OverloadedStrings #-}
-- | This module provides a function for serializing structured `Xlsx` into lazy bytestring
module Codec.Xlsx.Writer
    ( fromXlsx
    ) where

import qualified Codec.Archive.Zip as Zip
import           Control.Lens hiding (transform)
import           Control.Monad.Trans.State
import qualified Data.ByteString.Lazy as L
import           Data.ByteString.Lazy.Char8()
import           Data.List
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
        "application/vnd.openxmlformats-officedocument.extended-properties+xml" appXml
      , FileData "xl/workbook.xml"
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml" $ bookXml (xlsx ^. xlSheets)
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
        sheetXml (w ^. wsColumns) (w ^. wsRowPropertiesMap) cells (w ^. wsMerges) |
        (n, cells, w) <- zip3 [1..] sheetCells sheets]
    sheets = xlsx ^. xlSheets . to M.elems
    sheetCount = length sheets
    (sheetCells, shared) = runState (mapM collectSharedTransform sheets) []

data FileData = FileData { fdName :: Text
                         , fdContentType :: Text
                         , fdContents :: L.ByteString}

coreXml :: CalendarTime -> Text -> L.ByteString
coreXml created creator =
  renderLBS def $ Document (Prologue [] Nothing []) root []
  where
    date = T.pack $ formatCalendarTime defaultTimeLocale "%Y-%m-%dT%H:%M:%SZ" created
    nsAttrs = M.fromList [("xmlns:dcterms", "http://purl.org/dc/terms/")]
    root = Element (nm "http://schemas.openxmlformats.org/package/2006/metadata/core-properties" "coreProperties") nsAttrs
           [nEl (nm "http://purl.org/dc/terms/" "created")
                                 (M.fromList [(nm "http://www.w3.org/2001/XMLSchema-instance" "type", "dcterms:W3CDTF")]) [NodeContent date],
            nEl (nm "http://purl.org/dc/elements/1.1/" "creator") M.empty [NodeContent creator],
            nEl (nm "http://schemas.openxmlformats.org/package/2006/metadata/core-properties" "version") M.empty [NodeContent "0"]]

appXml :: L.ByteString
appXml = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\
\<Properties xmlns=\"http://schemas.openxmlformats.org/officeDocument/2006/extended-properties\" xmlns:vt=\"http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes\"><TotalTime>0</TotalTime></Properties>"

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

collectSharedTransform :: Worksheet -> State [Text] [(Int, [(Int, XlsxCell)])]
collectSharedTransform ws = transformed
  where
    transformed = mapM transformRow $ toRows (ws ^. wsCells)
    transformRow (r, cells) = do
      cells' <- mapM transform cells
      return (r, cells')
    transform (c, Cell{_cellValue=v, _cellStyle=s}) =
      case v of
        Just(CellText t) -> do
          shared <- get
          case t `elemIndex` shared of
            Just i ->
              return (c, XlsxCell s (Just $ XlsxSS i))
            Nothing -> do
              put $ shared ++ [t]
              return (c, XlsxCell s (Just $ XlsxSS (length shared)))
        Just(CellDouble dbl) ->
          return (c, XlsxCell s (Just $ XlsxDouble dbl))
        Just(CellBool b) ->
          return (c, XlsxCell s (Just $ XlsxBool b))
        Nothing ->
          return (c, XlsxCell s Nothing)

sheetXml :: [ColumnsWidth] -> Map Int RowProperties -> [(Int, [(Int, XlsxCell)])] -> [Text]-> L.ByteString
sheetXml cws rh rows merges = renderLBS def $ Document (Prologue [] Nothing []) root []
  where
    cType = xlsxCellType
    root = addNS "http://schemas.openxmlformats.org/spreadsheetml/2006/main" $
           Element "worksheet" M.empty
           [nEl "cols" M.empty $  map cwEl cws,
            nEl "sheetData" M.empty $ map rowEl rows,
            nEl "mergeCells" M.empty $ map mergeE1 merges]
    cwEl cw = NodeElement $! Element "col" (M.fromList
              [("min", txti $ cwMin cw), ("max", txti $ cwMax cw), ("width", txtd $ cwWidth cw), ("style", txti $ cwStyle cw)]) []
    rowEl (r, cells) = nEl "row"
                       (M.fromList (ht ++ s ++ [("r", txti r) ,("hidden", "false"), ("outlineLevel", "0"),
                               ("collapsed", "false"), ("customFormat", "true"),
                               ("customHeight", txtb hasHeight)]))
                       $ map (cellEl r) cells
      where
        (ht, hasHeight, s) = case M.lookup r rh of
          Just (RowProps (Just h) (Just s)) -> ([("ht", txtd h)], True,[("s", txti s)])
          Just (RowProps Nothing  (Just s)) -> ([], True, [("s", txti s)])
          Just (RowProps (Just h) Nothing ) -> ([("ht", txtd h)], True,[])
          _ -> ([], False,[])
    mergeE1 t = NodeElement $! Element "mergeCell" (M.fromList [("ref",t)]) []
    cellEl r (icol, cell) =
      nEl "c" (M.fromList (cellAttrs r (int2col icol) cell))
              [nEl "v" M.empty [NodeContent $ value cell] | isJust $ xlsxCellValue cell]
    cellAttrs r col cell = cellStyleAttr cell ++ [("r", T.concat [col, txti r]), ("t", cType cell)]
    cellStyleAttr XlsxCell{xlsxCellStyle=Nothing} = []
    cellStyleAttr XlsxCell{xlsxCellStyle=Just s} = [("s", txti s)]

bookXml :: Map Text Worksheet -> L.ByteString
bookXml wss = renderLBS def $ Document (Prologue [] Nothing []) root []
  where
    numNames = [(txti i, name) | (i, name) <- zip [1..] (M.keys wss)]
    root = addNS "http://schemas.openxmlformats.org/spreadsheetml/2006/main" $ Element "workbook" M.empty
           [nEl "sheets" M.empty $
            map (\(n, name) -> nEl "sheet"
                               (M.fromList [("name", name), ("sheetId", n), ("state", "visible"),
                                            (rId, T.concat ["rId", n])]) []) numNames]
    rId = nm "http://schemas.openxmlformats.org/officeDocument/2006/relationships" "id"

ssXml :: [Text] -> L.ByteString
ssXml ss =
  renderLBS def $ Document (Prologue [] Nothing []) root []
  where
    root = addNS "http://schemas.openxmlformats.org/spreadsheetml/2006/main" $ Element "sst" M.empty $
           map (\s -> nEl "si" M.empty [nEl "t" M.empty [NodeContent s]]) ss

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

nm :: Text -> Text -> Name
nm ns n = Name
  { nameLocalName = n
  , nameNamespace = Just ns
  , namePrefix = Nothing}

addNS :: Text -> Element -> Element
addNS namespace (Element (Name ln _ _) as ns) = Element (Name ln (Just namespace) Nothing) as (map addNS' ns)
  where
    addNS' (NodeElement e) = NodeElement $ addNS namespace e
    addNS' n = n

nEl :: Name -> Map Name Text -> [Node] -> Node
nEl name attrs nodes = NodeElement $ Element name attrs nodes

txti :: Int -> Text
txti = toStrict . toLazyText . decimal

txtd :: Double -> Text
txtd = toStrict . toLazyText . realFloat

txtb :: Bool -> Text
txtb = T.toLower . T.pack . show
