{-# LANGUAGE NoMonomorphismRestriction #-}
{-# LANGUAGE OverloadedStrings #-}
{-# LANGUAGE TupleSections #-}

module Codec.Xlsx.Parser where

import           Prelude hiding (sequence)
import           Control.Applicative
import           Control.Monad.IO.Class
import           Control.Monad (join)
import           Data.Char (ord)
import           Data.List
import           Data.Maybe
import           Data.Ord
import           Data.Function (on)
import qualified Data.IntMap as M
import qualified Data.IntSet as S
import qualified Data.Map as Map

import           Data.Text (Text)
import qualified Data.Text as T
import qualified Data.Text.Read as T
import qualified Data.ByteString.Lazy as L
import qualified Data.ByteString.Lazy.Char8 as C

--import Control.Monad.Trans.Resource (MonadThrow)
import           Data.Conduit
import           Data.XML.Types
import qualified Data.Conduit.List as CL
import qualified Text.XML.Stream.Parse as Xml
import qualified Codec.Archive.Zip as Zip

import           Codec.Xlsx

import           Debug.Trace

data Xlsx = Xlsx{ archive :: Zip.Archive
                , sharedStrings :: M.IntMap Text
                , styles :: Styles }

data Styles = Styles L.ByteString
            deriving Show

data Columns
  = AllColumns
  | Columns [String]

type MapRow = Map.Map Text Text


-- | Read archive and preload sharedStrings
xlsx :: FilePath -> IO Xlsx
xlsx fname = do
  ar <- Zip.toArchive <$> L.readFile fname
  ss <- getSharedStrings ar
  st <- getStyles ar
  return $ Xlsx ar ss st


-- | Get data from specified worksheet.
sheet :: MonadThrow m => Xlsx -> Int -> [Text] -> Source m [Cell]
sheet x sheetId cols
  =  getSheetCells x sheetId
  $= filterColumns (S.fromList $ map col2int cols)
  $= groupRows
  $= reverseRows


-- | Get all rows from specified worksheet.
sheetRows :: MonadThrow m => Xlsx -> Int -> Source m MapRow
sheetRows x sheetId
  =  getSheetCells x sheetId
  $= CL.map (\x -> trace ("pregroup:" ++ show x) $ x) $= groupRows
  $= CL.map (\x -> trace ("prereverse" ++ show x) $ x) $= reverseRows
  $= mkMapRows


-- | Make 'Conduit' from 'mkMapRowsSink'.
mkMapRows :: Monad m => Conduit [Cell] m MapRow
mkMapRows = sequence mkMapRowsSink =$= CL.concatMap id


-- | Make 'MapRow' from list of 'Cell's.
mkMapRowsSink :: Monad m => Sink [Cell] m [MapRow]
mkMapRowsSink = do
    header <- fromMaybe [] <$> CL.head
    rows   <- CL.consume

    return $ map (mkMapRow header) rows
  where
    mkMapRow header row = Map.fromList $ zipCells header row

    zipCells :: [Cell] -> [Cell] -> [(Text, Text)]
    zipCells []            _          = []
    zipCells header        []         = map (\h -> (txt h, "")) header
    zipCells header@(h:hs) row@(r:rs) =
        case comparing (fst . cellIx) h r of
          LT -> (txt h , ""   ) : zipCells hs     row
          EQ -> (txt h , txt r) : zipCells hs     rs
          GT -> (""    , txt r) : zipCells header rs

    txt = fromMaybe "" . cv
    cv Cell{cellValue=Just(CellText t)} = Just t
    cv Cell{cellValue=Nothing} = Nothing



reverseRows = CL.map reverse
groupRows = CL.groupBy ((==) `on` (snd.cellIx))
filterColumns cs = CL.filter ((`S.member` cs) . col2int . fst . cellIx)
col2int = T.foldl' (\n c -> n*26 + ord c - ord 'A' + 1) 0


getSheetCells
 :: MonadThrow m => Xlsx -> Int -> Source m Cell
getSheetCells (Xlsx{archive=ar,sharedStrings=ss}) sheetId
  | sheetId < 0 || sheetId >= length sheets
    = error "parseSheet: Invalid sheetId"
  | otherwise
    = case xmlSource ar (sheets !! sheetId) of
      Nothing -> error "An impossible happened"
      Just xml -> xml $= mkXmlCond (getCell ss)
  where
    sheets = sort
      $ filter (isPrefixOf "xl/worksheets")
      $ Zip.filesInArchive ar


-- | Parse single cell from xml stream.
getCell
 :: MonadThrow m => M.IntMap Text -> Sink Event m (Maybe Cell)
getCell ss = Xml.tagName (n"c") cAttrs cParser
  where
    cAttrs = do
      cellIx  <- {-trace "r attr" $ -}Xml.requireAttr  "r"
      style   <- Xml.optionalAttr "s"
      sharing <- Xml.optionalAttr "t"
      Xml.ignoreAttrs
      return $ (cellIx,style,sharing)

    cParser a@(ix,style,sharing) = do
      val <- case sharing of
          Just "inlineStr" -> tagSeq ["is", "t"]
          Just "s" -> {-trace "s type" $ -}tagSeq ["v"]
            >>= return . join . fmap ((`M.lookup` ss).int)
          Nothing  -> tagSeq ["v"]
      return $ {-trace "Cell" $ -}Cell (mkCellIx ix) (int <$> style) (fmap CellText val)

    mkCellIx ix = let (c,r) = T.span (>'9') ix
                  in {-trace ("mkCell" ++ show (c,r))$ -}(c,int r)


-- | Add namespace to element names
n x = Name
  {nameLocalName = x
  ,nameNamespace = Just "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  ,namePrefix = Nothing}


-- | Get text from several nested tags
tagSeq :: MonadThrow m => [Text] -> Sink Event m (Maybe Text)
tagSeq (x:xs)
  = Xml.tagNoAttr (n x)
  $ foldr (\x -> Xml.force "" . Xml.tagNoAttr (n x)) Xml.content xs


-- | Get xml event stream from the specified file inside the zip archive.
xmlSource
 :: MonadThrow m => Zip.Archive -> FilePath -> Maybe (Source m Event)
xmlSource ar fname
  =   Xml.parseLBS Xml.def
  .   Zip.fromEntry
  <$> Zip.findEntryByPath fname ar


-- Get shared strings (if there are some) into IntMap.
getSharedStrings
  :: (MonadThrow m, Functor m)
  => Zip.Archive -> m (M.IntMap Text)
getSharedStrings x
  = case xmlSource x "xl/sharedStrings.xml" of
    Nothing -> return M.empty
    Just xml -> (M.fromAscList . zip [0..]) <$> getText xml

-- | Fetch all text from xml stream.
getText xml = xml $= mkXmlCond Xml.contentMaybe $$ CL.consume


getStyles :: (MonadThrow m, Functor m) => Zip.Archive -> m Styles
getStyles ar = case (Zip.fromEntry <$> Zip.findEntryByPath "xl/styles.xml" ar) of
  Nothing  -> return (Styles L.empty)
  Just xml -> return (Styles xml)

---------------------------------------------------------------------


int :: Text -> Int
int = either error fst . T.decimal


-- | Create conduit from xml sink
-- Resulting conduit filters nodes that `f` can consume and skips everything
-- else.
--
-- FIXME: Some benchmarking required: maybe it's not very efficient to `peek`i
-- each element twice. It's possible to swap call to `f` and `CL.peek`.
mkXmlCond f = sequenceSink () $ const
  $ CL.peek >>= maybe          -- try get current event form the stream
    (return {- $ trace "stopped" -}$ Stop)              -- stop if stream is empty
    (\_ -> {-trace "calling f" $ -}f >>= maybe         -- try consume current event
      ({-trace "drop" $-} CL.drop 1 >> return (Emit () [])) -- skip it if can't process
      ({-trace "emit2" $ -}return . Emit () . (:[])))        -- return result otherwise



test p = do
  x <- xlsx p --"/home/qrilka/workspace/haskell/xlsx-templater/tmpl.xlsx" --
            -- "/home/qrilka/test.xlsx"
  runResourceT $ sheet x 0 ["A", "B", "C"] $= CL.map reverse $$ CL.consume
  

test2 = do
  let input = L.concat ["<?xml version=\"1.0\"?><foo xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><a></a><b t=\"some\"></b></foo>"]
  let xml = trace "x" $ (Xml.parseFile Xml.def "/home/qrilka/sheet1.xml")-- :: Source IO Event
  let tagger = trace "t" $ (Xml.tagName (n"c") (Xml.requireAttr "t" >>= (\t -> Xml.ignoreAttrs >> return t)) (\t -> tagSeq["v"] >> return t)) -- :: Sink Event IO (Maybe Text)
  runResourceT $ xml $={- (CL.map (\x -> trace "b" x)) =$= -}mkXmlCond (getCell $ M.fromList [(0,"some")]){- tagger -} $$ CL.consume


test3 = do
  let xml = Xml.parseLBS Xml.def "<?xml version=\"1.0\"?><styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><bs><b>bar</b><b>foo</b></bs></styleSheet>"
  xml $$ Xml.tagNoAttr (n "styleSheet") $ do
    a <- (Xml.tagNoAttr (n "a") Xml.content)
    bs <- Xml.tagNoAttr (n "bs") (Xml.many $ Xml.tagNoAttr (n "b") Xml.content)
    return (a,bs)

parseBs = Xml.tagNoAttr (n "bs") (Xml.many parseB)

parseB = Xml.tagNoAttr (n "b") Xml.content
--  a <- Xml.tagNoAttr "a" Xml.content
--  b <- Xml.tagNoAttr "b" Xml.content
--  return (b)

test4 = Xml.parseLBS Xml.def "<?xml version=\"1.0\"?><a>text</a>" $$ Xml.tagNoAttr "a" ((Xml.many $ Xml.tagNoAttr "b" Xml.content) >>= \x -> do z<-Xml.content; return (x,z))

test5 = Xml.parseLBS Xml.def "<?xml version=\"1.0\"?><a><bs><b>b1</b><b>b2</b></bs>text</a>" $$ Xml.tagNoAttr "a" ((Xml.many $ Xml.tagNoAttr "b" Xml.content) >>= \x -> do z<-Xml.content; return (x,z))

test5_bad = Xml.parseLBS Xml.def "<?xml version=\"1.0\"?><a>before<b>b1</b><b>b2</b>text</a>" $$ Xml.tagNoAttr "a" ((Xml.many $ Xml.tagNoAttr "b" Xml.content) >>= \x -> do z<-Xml.content; return (x,z))