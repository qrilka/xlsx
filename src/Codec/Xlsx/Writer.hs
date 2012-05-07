{-# LANGUAGE OverloadedStrings #-}
{-# LANGUAGE QuasiQuotes #-}
module Codec.Xlsx.Writer (
  writeXlsx,
  writeXlsxStyles
  ) where

import qualified Codec.Archive.Zip as Zip
import           Control.Applicative((<$>))
import           Control.Monad.Trans.State
import qualified Data.ByteString.Lazy as L
import qualified Data.ByteString.Lazy.Char8
import           Data.List
import           Data.Maybe
import           Data.Text (Text)
import qualified Data.Text as T
import           Data.Time.Calendar
import           Data.Time.LocalTime
import           System.Locale
import           System.Time
import           Text.Hamlet.XML
import           Text.XML

import           Codec.Xlsx


type Worksheet = [[Cell]]

writeXlsx :: FilePath -> [Worksheet] -> IO ()
writeXlsx p d = writeXlsxStyles p emptyStylesXml d

writeXlsxStyles :: FilePath -> L.ByteString -> [Worksheet] -> IO ()
writeXlsxStyles p s d = constructXlsx s d >>= L.writeFile p

data FileData = FileData { fdName :: Text
                         , fdContentType :: Text
                         , fdContents :: L.ByteString}

constructXlsx :: L.ByteString -> [Worksheet] -> IO L.ByteString
constructXlsx s d = do
  ct <- getClockTime
  let
    TOD t _ = ct
    utct = toUTCTime ct
    (xlsxSheets, shared) = runState (mapM collectSharedTransform d) []
    sheetNumber = length d
    sheetFiles = [FileData (T.concat ["xl/worksheets/sheet", T.pack $ show n, ".xml"]) "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml" $ sheetXml cells | (n, cells) <- zip [1..] xlsxSheets]
    files = sheetFiles ++
      [ FileData "docProps/core.xml" "application/vnd.openxmlformats-package.core-properties+xml" $ coreXml utct "xlsxwriter"
      , FileData "docProps/app.xml" "application/vnd.openxmlformats-officedocument.extended-properties+xml" appXml
      , FileData "xl/workbook.xml" "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml" $ bookXml sheetNumber
      , FileData "xl/styles.xml" "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml" s
      , FileData "xl/sharedStrings.xml" "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml" $ ssXml shared
      , FileData "xl/_rels/workbook.xml.rels" "application/vnd.openxmlformats-package.relationships+xml" $ bookRelXml sheetNumber
      , FileData "_rels/.rels" "application/vnd.openxmlformats-package.relationships+xml" rootRelXml ]
    entries = (Zip.toEntry "[Content_Types].xml" t $ contentTypesXml files) :
              map (\fd -> Zip.toEntry (T.unpack $ fdName fd) t (fdContents fd)) files
    ar = foldr Zip.addEntryToArchive Zip.emptyArchive entries
  return $ Zip.fromArchive ar


coreXml :: CalendarTime -> Text -> L.ByteString
coreXml created creator =
  renderLBS def $ Document (Prologue [] Nothing []) root []
  where
    date = T.pack $ formatCalendarTime defaultTimeLocale "%Y-%m-%dT%H:%M:%SZ" created
    nsAttrs = [("xmlns:dc", "http://purl.org/dc/elements/1.1/")
              ,("xmlns:dcterms", "http://purl.org/dc/terms/")
              ,("xmlns:xsi", "http://www.w3.org/2001/XMLSchema-instance")]
    root = Element (nmp "http://schemas.openxmlformats.org/package/2006/metadata/core-properties" "cp" "coreProperties") nsAttrs [xml|
<dcterms:created xsi:type="dcterms:W3CDTF">#{date}
<dc:creator>#{creator}
<cp:revision>0
|]

appXml :: L.ByteString
appXml = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\
\<Properties xmlns=\"http://schemas.openxmlformats.org/officeDocument/2006/extended-properties\" xmlns:vt=\"http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes\"><TotalTime>0</TotalTime></Properties>"

data XlsxCellData = XlsxSS Int | XlsxDouble Double
data XlsxCell = XlsxCell{ xlsxCellIx :: (Text, Int)
                        , xlsxCellStyle  :: Maybe Int
                        , xlsxCellValue  :: Maybe XlsxCellData
                        }

xlsxCellType :: XlsxCell -> Text
xlsxCellType XlsxCell{xlsxCellValue=Just(XlsxSS _)} = "s"
xlsxCellType XlsxCell{xlsxCellValue=Just(XlsxDouble _)} = "n"

value :: XlsxCell -> Text
value XlsxCell{xlsxCellValue=Just(XlsxSS i)} = T.pack $ show i
value XlsxCell{xlsxCellValue=Just(XlsxDouble i)} = T.pack $ show i


collectSharedTransform :: [[Cell]] -> State [Text] [[XlsxCell]]
collectSharedTransform d = transformed
  where
    shared = catMaybes $ map maybeString $ concat d
    maybeString :: Cell -> Maybe Text
    maybeString Cell{cellValue=Just(CellText t)} = Just t
    maybeString _ = Nothing
    transformed = mapM (mapM transform) d
    transform Cell{cellValue=v, cellIx = ix, cellStyle=s} =
      case v of
        Just(CellText t) -> do
          shared <- get
          case t `elemIndex` shared of
            Just i -> 
              return XlsxCell{xlsxCellIx=ix, xlsxCellStyle=s, xlsxCellValue=Just $ XlsxSS i}
            Nothing -> do
              put $ shared ++ [t]
              return XlsxCell{xlsxCellIx=ix, xlsxCellStyle=s, xlsxCellValue=Just $ XlsxSS (length shared)}
        Just(CellDouble d) ->
          return XlsxCell{xlsxCellIx=ix, xlsxCellStyle=s, xlsxCellValue=Just $ XlsxDouble d}
        Just(CellLocalTime t) ->
          return XlsxCell{xlsxCellIx=ix, xlsxCellStyle=s, xlsxCellValue=Just $ XlsxDouble (xlsxDoubleTime t)}
        Nothing ->
          return XlsxCell{xlsxCellIx=ix, xlsxCellStyle=s, xlsxCellValue=Nothing}

xlsxDoubleTime :: LocalTime -> Double
xlsxDoubleTime LocalTime{localDay=day,localTimeOfDay=time} =
  fromIntegral (diffDays day xlsxEpochStart) + timeFraction time
  where
    xlsxEpochStart = fromGregorian 1899 12 30
    timeFraction = fromRational . timeOfDayToDayFraction


xlsColumns :: [Text]
xlsColumns = [T.pack [l] | l <- ['A'..'Z']]

sheetXml :: [[XlsxCell]] -> L.ByteString
sheetXml d =
  renderLBS def $ Document (Prologue [] Nothing []) root []
  where
    rows = zip [1..] d
    n = T.pack . show . fst
    cells = snd
    column = fst . xlsxCellIx
    cType = xlsxCellType
    cValue = value
    cStyle :: XlsxCell -> Maybe Text
    cStyle cell = fmap (T.pack . show) (xlsxCellStyle cell)
    root = addNS "http://schemas.openxmlformats.org/spreadsheetml/2006/main" $ Element "worksheet" [] [xml|
<sheetData>
  $forall row <- rows
    <row collapsed=false customFormat=false customHeight=false hidden=false outlineLevel=0 r=#{n row}>
      $forall cell <- cells row
        $if isJust (xlsxCellValue cell)
          $maybe s <- (cStyle cell)
            <c r=#{column cell}#{n row} t=#{cType cell} s=#{s}>
              <v>#{cValue cell}
          $nothing
            <c r=#{column cell}#{n row} t=#{cType cell}>
              <v>#{cValue cell}
|]

bookXml :: Int -> L.ByteString
bookXml n = renderLBS def $ Document (Prologue [] Nothing []) root []
  where
    ns = [1..n]
    root = addNS "http://schemas.openxmlformats.org/spreadsheetml/2006/main" $ Element "workbook" [] [xml|
<sheets>
  $forall n <- ns
    $with n' <- T.pack $ show n
      <sheet name=List#{n'} sheetId=#{n'} state=visible {http://schemas.openxmlformats.org/officeDocument/2006/relationships}id=rId#{n'}>
|]

emptyStylesXml :: L.ByteString
emptyStylesXml = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\
\<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"></styleSheet>"

ssXml :: [Text] -> L.ByteString
ssXml ss =
  renderLBS def $ Document (Prologue [] Nothing []) root []
  where
    root = addNS "http://schemas.openxmlformats.org/spreadsheetml/2006/main" $ Element "sst" [] [xml|
$forall s <- ss
    <si>
        <t>#{s}
|]

bookRelXml :: Int -> L.ByteString
bookRelXml n = renderLBS def $ Document (Prologue [] Nothing []) root []
  where
    ns = [1..n]
    n' = n + 1
    n'' = n + 2
    root = addNS "http://schemas.openxmlformats.org/package/2006/relationships" $ Element "Relationships" [] [xml|
$forall n <- ns
  <Relationship Id=rId#{T.pack $ show n} Type=http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet Target=worksheets/sheet#{T.pack $ show n}.xml>
<Relationship Id=rId#{T.pack $ show n'} Type=http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles Target=styles.xml>
<Relationship Id=rId#{T.pack $ show n''} Type=http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings Target=sharedStrings.xml>
|]

rootRelXml :: L.ByteString
rootRelXml = "<?xml version=\"1.0\" encoding=\"UTF-8\"?>\
\<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"xl/workbook.xml\"/><Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties\" Target=\"docProps/core.xml\"/><Relationship Id=\"rId3\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties\" Target=\"docProps/app.xml\"/></Relationships>"

contentTypesXml :: [FileData] -> L.ByteString
contentTypesXml fds = renderLBS def $ Document (Prologue [] Nothing []) root []
  where
    root = Element (nm "http://schemas.openxmlformats.org/package/2006/content-types" "Types") [] [xml|
$forall fd <- fds
   <{http://schemas.openxmlformats.org/package/2006/content-types}Override PartName=/#{fdName fd} ContentType=#{fdContentType fd}>
|]

nmp :: Text -> Text -> Text -> Name
nmp ns p n = Name
  { nameLocalName = n
  , nameNamespace = Just ns
  , namePrefix = Just p}

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