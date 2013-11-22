{-# LANGUAGE OverloadedStrings #-}
module Codec.Xlsx.Writer (
  writeXlsx,
  writeXlsxStyles
  ) where

import qualified Codec.Archive.Zip as Zip
import           Control.Monad.Trans.State
import           Data.ByteString.Lazy.Char8()
import qualified Data.ByteString.Lazy as L
import           Data.List
import qualified Data.IntMap as IM
import qualified Data.Map as M
import           Data.Maybe
import           Data.Text (Text)
import qualified Data.Text as T

import           Data.Text.Lazy (toStrict)
import           Data.Text.Lazy.Builder (toLazyText)
import           Data.Text.Lazy.Builder.Int
import           Data.Text.Lazy.Builder.RealFloat
import           Data.Time.Calendar
import           Data.Time.LocalTime
import           System.Locale
import           System.Time
import           Text.XML

import           Codec.Xlsx
import           Codec.Xlsx.Parser (sheet)


-- | writes list of worksheets
writeXlsx :: FilePath -> Xlsx -> Maybe MappedSheet -> IO ()
writeXlsx fp xl@(Xlsx xlA xlS (Styles sty) xlWkfls) Nothing = do
  xlWshts <- (sheet xl)  `mapM` (zipWith (\a b -> a) [0 ..] xlWkfls)
  print sty
  writeXlsxStyles fp sty xlWshts
writeXlsx fp xl@(Xlsx xlA xlS (Styles sty) xlWkfls) (Just mappedSheets) = do
  let
    sheetList :: [Worksheet]
    sheetList = (snd `fmap`)  (IM.toList.unMappedSheet $ mappedSheets)
  writeXlsxStyles fp sty sheetList





-- | writes list of worksheets as xlsx file
writeWorksheetList :: FilePath -> [Worksheet] -> IO ()
writeWorksheetList p = writeXlsxStyles p emptyStylesXml

-- | writes list of worksheets and their styling as xlsx file
writeXlsxStyles :: FilePath -> L.ByteString -> [Worksheet] -> IO ()
writeXlsxStyles p s d = constructXlsx s d >>= L.writeFile p

data FileData = FileData { fdName :: Text
                         , fdContentType :: Text
                         , fdContents :: L.ByteString}

constructXlsx :: L.ByteString -> [Worksheet] -> IO L.ByteString
constructXlsx s ws = do
  ct <- getClockTime
  let
    TOD t _ = ct
    utct = toUTCTime ct
    (sheetCells, shared) = runState (mapM collectSharedTransform ws) []
    sheetNumber = length ws
    sheetFiles = [FileData (T.concat ["xl/worksheets/sheet", txti n, ".xml"]) "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml" $
                  sheetXml (wsColumns w) (wsRowHeights w) cells | (n, cells, w) <- zip3 [1..] sheetCells ws]
    files = sheetFiles ++
      [ FileData "docProps/core.xml" "application/vnd.openxmlformats-package.core-properties+xml" $ coreXml utct "xlsxwriter"
      , FileData "docProps/app.xml" "application/vnd.openxmlformats-officedocument.extended-properties+xml" appXml
      , FileData "xl/workbook.xml" "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml" $ bookXml ws
      , FileData "xl/styles.xml" "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml" s
      , FileData "xl/sharedStrings.xml" "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml" $ ssXml shared
      , FileData "xl/_rels/workbook.xml.rels" "application/vnd.openxmlformats-package.relationships+xml" $ bookRelXml sheetNumber
      , FileData "_rels/.rels" "application/vnd.openxmlformats-package.relationships+xml" rootRelXml ]
    entries = Zip.toEntry "[Content_Types].xml" t (contentTypesXml files) :
              map (\fd -> Zip.toEntry (T.unpack $ fdName fd) t (fdContents fd)) files
    ar = foldr Zip.addEntryToArchive Zip.emptyArchive entries
  return $ Zip.fromArchive ar


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

data XlsxCellData = XlsxSS Int | XlsxDouble Double
data XlsxCell = XlsxCell{ xlsxCellStyle  :: Maybe Int
                        , xlsxCellValue  :: Maybe XlsxCellData
                        }

xlsxCellType :: XlsxCell -> Text
xlsxCellType XlsxCell{xlsxCellValue=Just(XlsxSS _)} = "s"
xlsxCellType _ = "n" -- default type, TODO: fix cell output?

value :: XlsxCell -> Text
value XlsxCell{xlsxCellValue=Just(XlsxSS i)} = txti i
value XlsxCell{xlsxCellValue=Just(XlsxDouble d)} = txtd d
value _ = error "value undefined"


collectSharedTransform :: Worksheet -> State [Text] [[XlsxCell]]
collectSharedTransform d = transformed
  where
    transformed = mapM (mapM transform) $ toList d
    transform Nothing = return $ XlsxCell Nothing Nothing
    transform (Just CellData{cdValue=v, cdStyle=s}) =
      case v of
        Just(CellText t) -> do
          shared <- get
          case t `elemIndex` shared of
            Just i ->
              return $ XlsxCell s (Just $ XlsxSS i)
            Nothing -> do
              put $ shared ++ [t]
              return $ XlsxCell s (Just $ XlsxSS (length shared))
        Just(CellDouble dbl) ->
          return $ XlsxCell s (Just $ XlsxDouble dbl)
        Just(CellLocalTime t) ->
          return $ XlsxCell s (Just $ XlsxDouble (xlsxDoubleTime t))
        Nothing ->
          return $ XlsxCell s Nothing

xlsxDoubleTime :: LocalTime -> Double
xlsxDoubleTime LocalTime{localDay=day,localTimeOfDay=time} =
  fromIntegral (diffDays day xlsxEpochStart) + timeFraction time
  where
    xlsxEpochStart = fromGregorian 1899 12 30
    timeFraction = fromRational . timeOfDayToDayFraction

sheetXml :: [ColumnsWidth] -> RowHeights -> [[XlsxCell]] -> L.ByteString
sheetXml cws rh d = renderLBS def $ Document (Prologue [] Nothing []) root []
  where
    rows = zip [1..] d
    numCols = zip [int2col n | n <- [1..]]
    cType = xlsxCellType
    root = addNS "http://schemas.openxmlformats.org/spreadsheetml/2006/main" $
           Element "worksheet" M.empty
           [nEl "cols" M.empty $  map cwEl cws,
            nEl "sheetData" M.empty $ map rowEl rows]
    cwEl cw = NodeElement $! Element "col" (M.fromList
              [("min", txti $ cwMin cw), ("max", txti $ cwMax cw), ("width", txtd $ cwWidth cw), ("style", txti $ cwStyle cw)]) []
    rowEl (r, cells) = nEl "row"
                       (M.fromList (ht ++ [("r", txti r), ("hidden", "false"), ("outlineLevel", "0"),
                               ("collapsed", "false"), ("customFormat", "false"),
                               ("customHeight", txtb hasHeight)]))
                       $ map (cellEl r) (numCols cells)
      where
        (ht, hasHeight) = case M.lookup r rh of
          Just h  -> ([("ht", txtd h)], True)
          Nothing -> ([], False)
    cellEl r (col, cell) =
      nEl "c" (M.fromList (cellAttrs r col cell)) [nEl "v" M.empty [NodeContent $ value cell] | isJust $ xlsxCellValue cell]
    cellAttrs r col cell = cellStyleAttr cell ++ [("r", T.concat [col, txti r]), ("t", cType cell)]
    cellStyleAttr XlsxCell{xlsxCellStyle=Nothing} = []
    cellStyleAttr XlsxCell{xlsxCellStyle=Just s} = [("s", txti s)]

bookXml :: [Worksheet] -> L.ByteString
bookXml wss = renderLBS def $ Document (Prologue [] Nothing []) root []
  where
    numNames = [(txti i, wsName ws) | (i, ws) <- zip [1..] wss]
    root = addNS "http://schemas.openxmlformats.org/spreadsheetml/2006/main" $ Element "workbook" M.empty
           [nEl "sheets" M.empty $
            map (\(n, name) -> nEl "sheet"
                               (M.fromList $ [("name", name), ("sheetId", n), ("state", "visible"),
                                (rId, T.concat ["rId", n])]) []) numNames]
    rId = nm "http://schemas.openxmlformats.org/officeDocument/2006/relationships" "id"

emptyStylesXml :: L.ByteString
emptyStylesXml = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\
\<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"></styleSheet>"

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

nEl :: Name -> (M.Map Name Text) -> [Node] -> Node
nEl name attrs nodes = NodeElement $ Element name attrs nodes



txti :: Int -> Text
txti = toStrict . toLazyText . decimal

txtd :: Double -> Text
txtd = toStrict . toLazyText . realFloat

txtb :: Bool -> Text
txtb = T.toLower . T.pack . show
