{-# LANGUAGE OverloadedStrings #-}

import           Codec.Xlsx

import           Control.Applicative
import           Control.Lens
import qualified Data.ByteString.Lazy as L
import qualified Data.Map as M
import           Data.Text (Text,pack)
import           Data.Time.Clock.POSIX (getPOSIXTime)
import           Prelude


xEmpty :: Cell
xEmpty = Cell{_cellValue=Nothing, _cellStyle=Just 0}

xText :: Text -> Cell
xText t = Cell{_cellValue=Just $ CellText t, _cellStyle=Just 0}

xDouble :: Double -> Cell
xDouble d = Cell{_cellValue=Just $ CellDouble d, _cellStyle=Just 0}

styles :: Styles
styles = Styles "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\
\<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" xmlns:x14ac=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac\" mc:Ignorable=\"x14ac\">\
\  <numFmts count=\"1\">\
\    <numFmt numFmtId=\"164\" formatCode=\"yyyy/mm/dd\"/>\
\  </numFmts>\
\  <fonts count=\"2\" x14ac:knownFonts=\"1\">\
\    <font>\
\      <sz val=\"11\"/>\
\      <color theme=\"1\"/>\
\      <name val=\"Calibri\"/>\
\      <family val=\"2\"/>\
\      <scheme val=\"minor\"/>\
\    </font>\
\    <font>\
\      <b/>\
\      <sz val=\"11\"/>\
\      <color theme=\"1\"/>\
\      <name val=\"Calibri\"/>\
\      <family val=\"2\"/>\
\      <scheme val=\"minor\"/>\
\    </font>\
\  </fonts>\
\  <fills count=\"2\">\
\    <fill>\
\      <patternFill patternType=\"none\"/>\
\    </fill>\
\    <fill>\
\      <patternFill patternType=\"gray125\"/>\
\    </fill>\
\  </fills>\
\  <borders count=\"1\">\
\    <border>\
\      <left/>\
\      <right/>\
\      <top/>\
\      <bottom/>\
\      <diagonal/>\
\    </border>\
\  </borders>\
\  <cellStyleXfs count=\"1\">\
\    <xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/>\
\  </cellStyleXfs>\
\  <cellXfs count=\"5\">\
\    <xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\"/>\
\    <xf numFmtId=\"0\" fontId=\"1\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyFont=\"1\"/>\
\    <xf numFmtId=\"164\" fontId=\"1\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyNumberFormat=\"1\" applyFont=\"1\"/>\
\    <xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyFont=\"1\"/>\
\    <xf numFmtId=\"2\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyNumberFormat=\"1\" applyFont=\"1\"/>\
\  </cellXfs>\
\  <cellStyles count=\"1\">\
\    <cellStyle name=\"Normal\" xfId=\"0\" builtinId=\"0\"/>\
\  </cellStyles>\
\  <dxfs count=\"0\"/>\
\  <tableStyles count=\"0\" defaultTableStyle=\"TableStyleMedium2\" defaultPivotStyle=\"PivotStyleLight16\"/>\
\  <extLst>\
\    <ext xmlns:x14=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/main\" uri=\"{EB79DEF2-80B8-43e5-95BD-54CBDDF9020C}\">\
\      <x14:slicerStyles defaultSlicerStyle=\"SlicerStyleLight1\"/>\
\    </ext>\
\    <ext xmlns:x15=\"http://schemas.microsoft.com/office/spreadsheetml/2010/11/main\" uri=\"{9260A510-F301-46a8-8635-F512D64BE5F5}\">\
\      <x15:timelineStyles defaultTimelineStyle=\"TimeSlicerStyleLight1\"/>\
\    </ext>\
\  </extLst>\
\</styleSheet>"

main :: IO ()
main =  do
    pt <- getPOSIXTime
    L.writeFile "test.xlsx" $ fromXlsx pt $ Xlsx sheets styles def
    x <- toXlsx <$> L.readFile "test.xlsx"
    putStrLn $ "And cell (3,2) value in list 'List' is " ++
             show (x ^? xlSheets . ix "List" . wsCells . ix (3,2) . cellValue . _Just)
  where
    cols = [ColumnsWidth 1 10 15 1]
    rowProps = M.fromList [(1, RowProps (Just 50) (Just 3))]
    cells = M.fromList [((r, c), v) | r <- [1..10000], (c, v) <- zip [1..] (row r) ]
    row r = [ xText $ pack $ "column1-r" ++ show r
            , xText $ pack $ "column2-r" ++ show r
            , xEmpty
            , xText $ pack $ "column4-r" ++ show r
            , xDouble 42.12345
            , xText  "False"]
    sheets = M.fromList [("List", Worksheet cols rowProps cells [] Nothing Nothing)] -- wtf merges?
