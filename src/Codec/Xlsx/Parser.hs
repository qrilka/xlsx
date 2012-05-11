{-# LANGUAGE NoMonomorphismRestriction #-}
{-# LANGUAGE OverloadedStrings #-}
{-# LANGUAGE TupleSections #-}

module Codec.Xlsx.Parser where

import           Control.Applicative
import           Control.Monad (join)
import           Control.Monad.IO.Class
import           Data.Char (ord)
import           Data.Function (on)
import qualified Data.IntMap as M
import qualified Data.IntSet as S
import           Data.List
import           Data.Map (Map)
import qualified Data.Map as Map
import           Data.Maybe
import           Data.Ord
import           Prelude hiding (sequence)

import           Data.Text (Text)
import qualified Data.Text as T
import qualified Data.Text.Read as T
import qualified Data.ByteString.Lazy as L
import qualified Data.ByteString.Lazy.Char8 as C

import qualified Codec.Archive.Zip as Zip
import           Data.Conduit
import qualified Data.Conduit.List as CL
import           Data.XML.Types
import           System.FilePath
import qualified Text.XML.Stream.Parse as Xml

import           Codec.Xlsx

import           Debug.Trace

data Worksheet = Worksheet { wsName :: Text
                           , wsPath :: FilePath
                           }

data Xlsx = Xlsx{ archive :: Zip.Archive
                , sharedStrings :: M.IntMap Text
                , styles :: Styles
                , worksheets :: [Worksheet]
                }

data Styles = Styles L.ByteString
            deriving Show

type MapRow = Map.Map Text Text


-- | Read archive and preload sharedStrings
xlsx :: FilePath -> IO Xlsx
xlsx fname = do
  ar <- Zip.toArchive <$> L.readFile fname
  ss <- getSharedStrings ar
  st <- getStyles ar
  ws <- getWorksheets ar
  return $ Xlsx ar ss st ws


t' x sheetN = getSheetCells x sheetN $$
              sinkState Map.empty collect return
  where
    collect m c = return $ StateProcessing $ Map.insert (cellIx c) 0 m
      

-- | Get data from specified worksheet.
sheet :: MonadThrow m => Xlsx -> Int -> [Text] -> Source m [Cell]
sheet x sheetN cols  =  getSheetCells x sheetN
                        $= filterColumns (S.fromList $ map col2int cols)
                        $= groupRows
                        $= reverseRows

sheetMap :: MonadThrow m => Xlsx -> Int -> m (Maybe ((Int,Int), (Int,Int), Map (Int,Int) CellData))
sheetMap x n = getSheetCells x n $$ sinkState init collect' return
  where
    init = Nothing
    collect' m c = return $ StateProcessing $ collect m c
    collect p cell = case p of
      Nothing -> 
        Just ((x,x), (y,y), Map.singleton (x,y) cd)
      Just ((xmin,xmax), (ymin,ymax), m) ->
        Just ((min xmin x, max xmax x),
              (min ymin y, max ymax y),
              Map.insert (x,y) cd m)
      where
        x = xlsxCol2int $ fst $ cellIx cell
        y = (snd $ cellIx cell) - 1
        cd = cell2cd cell


-- | Get all rows from specified worksheet.
sheetRows :: MonadThrow m => Xlsx -> Int -> Source m MapRow
sheetRows x sheetN
  =  getSheetCells x sheetN
  $= groupRows
  $= reverseRows
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
getSheetCells (Xlsx{archive=ar, sharedStrings=ss, worksheets=sheets}) sheetN
  | sheetN < 0 || sheetN >= length sheets
    = error "parseSheet: Invalid sheet number"
  | otherwise
    = case xmlSource ar (wsPath $ sheets !! sheetN) of
      Nothing -> error "An impossible happened"
      Just xml -> xml $= mkXmlCond (getCell ss)


-- | Parse single cell from xml stream.
getCell
 :: MonadThrow m => M.IntMap Text -> Sink Event m (Maybe Cell)
getCell ss = Xml.tagName (n"c") cAttrs cParser
  where
    cAttrs = do
      cellIx  <- {-trace "r attr" $ -}Xml.requireAttr  "r"
      style   <- Xml.optionalAttr "s"
      typ <- Xml.optionalAttr "t"
      Xml.ignoreAttrs
      return $ (cellIx,style,typ)

    maybeCellDouble Nothing = Nothing
    maybeCellDouble (Just t) = either (const Nothing) (\(d,_) -> Just (CellDouble d)) $ T.rational t

    cParser a@(ix,style,typ) = do
      val <- case typ of
          Just "inlineStr" -> liftA (fmap CellText) (tagSeq ["is", "t"])
          Just "s" -> liftA (fmap CellText) (tagSeq ["v"] >>=
                                             return . join . fmap ((`M.lookup` ss).int))
          Just "n" -> liftA maybeCellDouble $ tagSeq ["v"]
          Nothing  -> liftA maybeCellDouble $ tagSeq ["v"]
      return $ {-trace "Cell" $ -}Cell (mkCellIx ix) (int <$> style) val -- (fmap CellText val)

    mkCellIx ix = let (c,r) = T.span (>'9') ix
                  in {-trace ("mkCell" ++ show (c,r))$ -}(c,int r)


-- | Add sml namespace to name
n x = Name
  {nameLocalName = x
  ,nameNamespace = Just "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  ,namePrefix = Nothing}

-- | Add office document relationship namespace to name
odr x = Name
  {nameLocalName = x
  ,nameNamespace = Just "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  ,namePrefix = Nothing}

-- | Add package relationship namespace to name
pr x = Name
  {nameLocalName = x
  ,nameNamespace = Just "http://schemas.openxmlformats.org/package/2006/relationships"
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

getWorksheets :: (MonadThrow m, Functor m) => Zip.Archive -> m [Worksheet]
getWorksheets ar = case xmlSource ar "xl/workbook.xml" of
  Nothing ->
    error "invalid workbook"
  Just xml -> do
    sheetData <- xml $= mkXmlCond getSheetData $$ CL.consume
    wbRels <- getWbRels ar
    return [Worksheet n ("xl" </> (T.unpack $ fromJust $ lookup rId wbRels)) | (n, rId) <- sheetData]

getSheetData = Xml.tagName (n"sheet") attrs return
  where
    attrs = do
      name <- Xml.requireAttr "name"
      rId  <- Xml.requireAttr (odr "id")
      Xml.ignoreAttrs
      return (name, rId)

getWbRels :: (MonadThrow m, Functor m) => Zip.Archive -> m [(Text, Text)]
getWbRels ar = case xmlSource ar "xl/_rels/workbook.xml.rels" of
  Nothing  -> return []
  Just xml -> xml $$ parseWbRels

parseWbRels = Xml.force "relationships required" $
              Xml.tagNoAttr (pr"Relationships") $
              Xml.many $ Xml.tagName (pr"Relationship") attr return
  where
    attr = do
      target <- Xml.requireAttr "Target"
      id <- Xml.requireAttr "Id"
      Xml.ignoreAttrs
      return (id, target)

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
  $ CL.peek >>= maybe            -- try get current event form the stream
    (return $ Stop)              -- stop if stream is empty
    (\_ -> f >>= maybe           -- try consume current event
           (CL.drop 1 >> return (Emit () [])) -- skip it if can't process
           (return . Emit () . (:[])))        -- return result otherwise
