{-# LANGUAGE OverloadedStrings #-}
{-# LANGUAGE QuasiQuotes #-}
module Codec.Xlsx.Writer(
  writeXlsx,
  writeXlsxStyles
  ) where

import qualified Codec.Archive.Zip as Zip
import           Control.Applicative((<$>))
import qualified Data.ByteString.Lazy as L
import qualified Data.ByteString.Lazy.Char8
import           Data.List
import           Data.Maybe
import qualified Data.Text as T
import           Data.Time.Calendar
import           Data.Time.LocalTime
import           Data.XML.Types
import           System.Locale
import           System.Time
import           Text.Hamlet.XML
import qualified Text.XML as Res

import           Codec.Xlsx


-- data XlsxType = XlsxString T.Text | XlsxInt Int | XlsxBool Bool | XlsxEmpty

writeXlsx :: FilePath -> [[Cell]] -> IO ()
writeXlsx p d = writeXlsxStyles p emptyStylesXml d

writeXlsxStyles :: FilePath -> L.ByteString -> [[Cell]] -> IO ()
writeXlsxStyles p s d = constructXlsx s d >>= L.writeFile p

data FileData = FileData { fdName :: T.Text
                         , fdContentType :: T.Text
                         , fdContents :: L.ByteString}

constructXlsx :: L.ByteString -> [[Cell]] -> IO L.ByteString
constructXlsx s d = do
  ct <- getClockTime
  let
    TOD t _ = ct
    utct = toUTCTime ct
    (shared, cells) = collectSharedTransform d
    files =
      [ FileData "docProps/core.xml" "application/vnd.openxmlformats-package.core-properties+xml" $ coreXml utct "xlsxwriter"
      , FileData "docProps/app.xml" "application/vnd.openxmlformats-officedocument.extended-properties+xml" appXml
      , FileData "xl/worksheets/sheet1.xml" "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml" $ sheetXml cells
      , FileData "xl/workbook.xml" "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml" bookXml
      , FileData "xl/styles.xml" "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml" s
      , FileData "xl/sharedStrings.xml" "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml" $ ssXml shared
      , FileData "xl/_rels/workbook.xml.rels" "application/vnd.openxmlformats-package.relationships+xml" bookRelXml
      , FileData "_rels/.rels" "application/vnd.openxmlformats-package.relationships+xml" rootRelXml ]
    entries = (Zip.toEntry "[Content_Types].xml" t $ contentTypesXml files) :
              map (\fd -> Zip.toEntry (T.unpack $ fdName fd) t (fdContents fd)) files
    ar = foldr Zip.addEntryToArchive Zip.emptyArchive entries
  return $ Zip.fromArchive ar


coreXml :: CalendarTime -> T.Text -> L.ByteString
coreXml created creator =
  Res.renderLBS Res.def $ Res.Document (Prologue [] Nothing []) root []
  where
  -- :head xmlns:cp=http://schemas.openxmlformats.org/spreadsheetml/2006/main
    date = T.pack $ formatCalendarTime defaultTimeLocale "%Y-%m-%dT%H:%M:%SZ" created
    nsAttrs = [("xmlns:dc", "http://purl.org/dc/elements/1.1/")
              ,("xmlns:dcterms", "http://purl.org/dc/terms/")
              ,("xmlns:xsi", "http://www.w3.org/2001/XMLSchema-instance")]
    root = Res.Element (nmp "http://schemas.openxmlformats.org/package/2006/metadata/core-properties" "cp" "coreProperties") nsAttrs [xml|
<dcterms:created xsi:type="dcterms:W3CDTF">#{date}
<dc:creator>#{creator}
<cp:revision>0
|]

appXml :: L.ByteString
appXml = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\
\<Properties xmlns=\"http://schemas.openxmlformats.org/officeDocument/2006/extended-properties\" xmlns:vt=\"http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes\"><TotalTime>0</TotalTime></Properties>"

data XlsxCellData = XlsxSS Int | XlsxDouble Double
data XlsxCell = XlsxCell{ xlsxCellIx :: (T.Text, Int)
                        , xlsxCellStyle  :: Maybe Int
                        , xlsxCellValue  :: Maybe XlsxCellData
                        }

xlsxCellType :: XlsxCell -> T.Text
xlsxCellType XlsxCell{xlsxCellValue=Just(XlsxSS _)} = "s"
xlsxCellType XlsxCell{xlsxCellValue=Just(XlsxDouble _)} = "n"

value :: XlsxCell -> T.Text
value XlsxCell{xlsxCellValue=Just(XlsxSS i)} = T.pack $ show i
value XlsxCell{xlsxCellValue=Just(XlsxDouble i)} = T.pack $ show i

-- TODO remove duplicates, optimize
collectSharedTransform :: [[Cell]] -> ([T.Text], [[XlsxCell]])
collectSharedTransform d = (shared, transformed)
  where
    shared = catMaybes $ map maybeString $ concat d
    maybeString :: Cell -> Maybe T.Text
    maybeString Cell{cellValue=Just(CellText t)} = Just t
    maybeString _ = Nothing
    transformed = map (map transform) d
    transform Cell{cellValue=v, cellIx = ix, cellStyle=s} =
      case v of
        Just(CellText t) ->
          XlsxCell{xlsxCellIx=ix, xlsxCellStyle=s, xlsxCellValue=XlsxSS <$> (t `elemIndex` shared)}
        Just(CellDouble d) ->
          XlsxCell{xlsxCellIx=ix, xlsxCellStyle=s, xlsxCellValue=Just $ XlsxDouble d}
        Just(CellLocalTime t) ->
          XlsxCell{xlsxCellIx=ix, xlsxCellStyle=s, xlsxCellValue=Just $ XlsxDouble (xlsxDoubleTime t)}
        Nothing ->
          XlsxCell{xlsxCellIx=ix, xlsxCellStyle=s, xlsxCellValue=Nothing}
--    transform (XlsxInt i) = CellInt i
--    transform (XlsxBool b) = CellBool b

xlsxDoubleTime :: LocalTime -> Double
xlsxDoubleTime LocalTime{localDay=day,localTimeOfDay=time} =
  fromIntegral (diffDays day xlsxEpochStart) + timeFraction time
  where
    xlsxEpochStart = fromGregorian 1899 12 30
    timeFraction = fromRational . timeOfDayToDayFraction


xlsColumns :: [T.Text]
xlsColumns = [T.pack [l] | l <- ['A'..'Z']]

sheetXml :: [[XlsxCell]] -> L.ByteString
sheetXml d =
  Res.renderLBS Res.def $ Res.Document (Prologue [] Nothing []) root []
  where
    rows = zip [1..] d
    n = T.pack . show . fst
    cells = snd
    column = fst . xlsxCellIx
    cType = xlsxCellType
    cValue = value
    cStyle :: XlsxCell -> Maybe T.Text
    cStyle cell = fmap (T.pack . show) (xlsxCellStyle cell)
    root = Res.Element (nm "http://schemas.openxmlformats.org/spreadsheetml/2006/main" "worksheet") [] [xml|
<{http://schemas.openxmlformats.org/spreadsheetml/2006/main}sheetData>
  $forall row <- rows
    <{http://schemas.openxmlformats.org/spreadsheetml/2006/main}row collapsed=false customFormat=false customHeight=false hidden=false outlineLevel=0 r=#{n row}>
      $forall cell <- cells row
        $maybe s <- (cStyle cell)
          <{http://schemas.openxmlformats.org/spreadsheetml/2006/main}c r=#{column cell}#{n row} t=#{cType cell} s=#{s}>
              <{http://schemas.openxmlformats.org/spreadsheetml/2006/main}v>#{cValue cell}
        $nothing      
          <{http://schemas.openxmlformats.org/spreadsheetml/2006/main}c r=#{column cell}#{n row} t=#{cType cell}>
              <{http://schemas.openxmlformats.org/spreadsheetml/2006/main}v>#{cValue cell}
              
|]
--  ht=\"12.8\"

bookXml :: L.ByteString
bookXml = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\
\<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"><sheets><sheet name=\"List1\" sheetId=\"1\" state=\"visible\" r:id=\"rId2\"/></sheets></workbook>"

emptyStylesXml :: L.ByteString
emptyStylesXml = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\
\<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"></styleSheet>"

ssXml :: [T.Text] -> L.ByteString
ssXml ss =
  Res.renderLBS Res.def $ Res.Document (Prologue [] Nothing []) root []
  where
    root = Res.Element (nm "http://schemas.openxmlformats.org/spreadsheetml/2006/main" "sst") [] [xml|
$forall s <- ss
    <{http://schemas.openxmlformats.org/spreadsheetml/2006/main}si>
        <{http://schemas.openxmlformats.org/spreadsheetml/2006/main}t>#{s}
|]

bookRelXml :: L.ByteString
bookRelXml = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\
\<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\
\<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" Target=\"styles.xml\"/><Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet1.xml\"/><Relationship Id=\"rId3\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings\" Target=\"sharedStrings.xml\"/>\
\</Relationships>"

rootRelXml :: L.ByteString
rootRelXml = "<?xml version=\"1.0\" encoding=\"UTF-8\"?>\
\<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"xl/workbook.xml\"/><Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties\" Target=\"docProps/core.xml\"/><Relationship Id=\"rId3\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties\" Target=\"docProps/app.xml\"/></Relationships>"

contentTypesXml :: [FileData] -> L.ByteString
contentTypesXml fds = Res.renderLBS Res.def $ Res.Document (Prologue [] Nothing []) root []
  where
    root = Res.Element (nm "http://schemas.openxmlformats.org/package/2006/content-types" "Types") [] [xml|
$forall fd <- fds
   <{http://schemas.openxmlformats.org/package/2006/content-types}Override PartName=/#{fdName fd} ContentType=#{fdContentType fd}>
|]

nmp :: T.Text -> T.Text -> T.Text -> Name
nmp ns p n = Name
  { nameLocalName = n
  , nameNamespace = Just ns
  , namePrefix = Just p}

nm :: T.Text -> T.Text -> Name
nm ns n = Name
  { nameLocalName = n
  , nameNamespace = Just ns
  , namePrefix = Nothing}
