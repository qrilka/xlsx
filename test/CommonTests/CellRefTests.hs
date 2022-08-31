
{-# LANGUAGE FlexibleInstances     #-}
{-# LANGUAGE MultiParamTypeClasses #-}
{-# LANGUAGE OverloadedStrings     #-}
{-# LANGUAGE ScopedTypeVariables   #-}
{-# LANGUAGE TypeApplications      #-}
{-# LANGUAGE ViewPatterns          #-}
{-# OPTIONS_GHC -fno-warn-orphans #-}

module CommonTests.CellRefTests
  ( tests
  ) where

import qualified Control.Applicative as Alt
import Control.Monad
import Data.Char (chr, isPrint)
import qualified Data.List as L
import Data.Text (Text)
import qualified Data.Text as T
import Data.Word (Word8)
import Test.SmallCheck.Series as Series (NonEmpty (..), Positive (..),
                                         Serial (..), cons1, cons2, generate,
                                         (\/))
import Test.Tasty (TestTree, testGroup)
import Test.Tasty.HUnit (testCase, (@?=))
import Test.Tasty.SmallCheck (testProperty)

import Codec.Xlsx.Types.Common

tests :: TestTree
tests =
  testGroup
    "Types.Common CellRef tests"
  [ testProperty "textToColumnIndex . columnIndexToText == id" $
        \(Positive i) -> i == textToColumnIndex (columnIndexToText i)

    , testProperty "row2coord . coord2row = id" $
        \(r :: RowCoord) -> r == row2coord (coord2row r)

    , testProperty "col2coord . coord2col = id" $
        \(c :: ColumnCoord) -> c == col2coord (coord2col c)

    , testProperty "fromSingleCellRef' . singleCellRef' = pure" $
        \(cellCoord :: CellCoord) -> pure cellCoord == fromSingleCellRef' (singleCellRef' cellCoord)

    , testProperty "fromRange' . mkRange' = pure" $
        \(range :: RangeCoord) -> pure range == fromRange' (uncurry mkRange' range)

    , testProperty "fromForeignSingleCellRef . mkForeignSingleCellRef = pure" $
        \(viewForeignCellParams -> params) ->
            pure params == fromForeignSingleCellRef (uncurry mkForeignSingleCellRef params)

    , testProperty "fromSingleCellRef' . mkForeignSingleCellRef = pure . snd" $
        \(viewForeignCellParams -> (nStr, cellCoord)) ->
            pure cellCoord == fromSingleCellRef' (mkForeignSingleCellRef nStr cellCoord)

    , testProperty "fromForeignSingleCellRef . singleCellRef' = const empty" $
        \(cellCoord :: CellCoord) ->
            Alt.empty == fromForeignSingleCellRef (singleCellRef' cellCoord)

    , testProperty "fromForeignRange . mkForeignRange = pure" $
        \(viewForeignRangeParams -> params@(nStr, (start, end))) ->
            pure params == fromForeignRange (mkForeignRange nStr start end)

    , testProperty "fromRange' . mkForeignRange = pure . snd" $
        \(viewForeignRangeParams -> (nStr, range@(start, end))) ->
            pure range == fromRange' (mkForeignRange nStr start end)

    , testProperty "fromForeignRange . mkRange' = const empty" $
        \(range :: RangeCoord) ->
            Alt.empty == fromForeignRange (uncurry mkRange' range)

    , testCase "building single CellRefs" $ do
        singleCellRef' (RowRel 5, ColumnRel 25) @?= CellRef "Y5"
        singleCellRef' (RowRel 5, ColumnAbs 25) @?= CellRef "$Y5"
        singleCellRef' (RowAbs 5, ColumnRel 25) @?= CellRef "Y$5"
        singleCellRef' (RowAbs 5, ColumnAbs 25) @?= CellRef "$Y$5"
        singleCellRef (5, 25) @?= CellRef "Y5"
    , testCase "parsing single CellRefs as abstract coordinates" $ do
        fromSingleCellRef (CellRef "Y5") @?= Just (5, 25)
        fromSingleCellRef (CellRef "$Y5") @?= Just (5, 25)
        fromSingleCellRef (CellRef "Y$5") @?= Just (5, 25)
        fromSingleCellRef (CellRef "$Y$5") @?= Just (5, 25)
    , testCase "parsing single CellRefs as potentially absolute coordinates" $ do
        fromSingleCellRef' (CellRef "Y5") @?= Just (RowRel 5, ColumnRel 25)
        fromSingleCellRef' (CellRef "$Y5") @?= Just (RowRel 5, ColumnAbs 25)
        fromSingleCellRef' (CellRef "Y$5") @?= Just (RowAbs 5, ColumnRel 25)
        fromSingleCellRef' (CellRef "$Y$5") @?= Just (RowAbs 5, ColumnAbs 25)
        fromSingleCellRef' (CellRef "$Y$50") @?= Just (RowAbs 50, ColumnAbs 25)
        fromSingleCellRef' (CellRef "$Y$5$0") @?= Nothing
        fromSingleCellRef' (CellRef "Y5:Z10") @?= Nothing
    , testCase "building ranges" $ do
        mkRange (5, 25) (10, 26) @?= CellRef "Y5:Z10"
        mkRange' (RowRel 5, ColumnRel 25) (RowRel 10, ColumnRel 26) @?= CellRef "Y5:Z10"
        mkRange' (RowAbs 5, ColumnAbs 25) (RowAbs 10, ColumnAbs 26) @?= CellRef "$Y$5:$Z$10"
        mkRange' (RowRel 5, ColumnAbs 25) (RowAbs 10, ColumnRel 26) @?= CellRef "$Y5:Z$10"
        mkForeignRange "myWorksheet" (RowRel 5, ColumnAbs 25) (RowAbs 10, ColumnRel 26) @?= CellRef "'myWorksheet'!$Y5:Z$10"
        mkForeignRange "my sheet" (RowRel 5, ColumnAbs 25) (RowAbs 10, ColumnRel 26) @?= CellRef "'my sheet'!$Y5:Z$10"
    , testCase "parsing ranges CellRefs as abstract coordinates" $ do
        fromRange (CellRef "Y5:Z10") @?= Just ((5, 25), (10, 26))
        fromRange (CellRef "$Y$5:$Z$10") @?= Just ((5, 25), (10, 26))
        fromRange (CellRef "myWorksheet!$Y5:Z$10") @?= Just ((5, 25), (10, 26))
    , testCase "parsing ranges CellRefs as potentially absolute coordinates" $ do
        fromRange' (CellRef "Y5:Z10") @?= Just ((RowRel 5, ColumnRel 25), (RowRel 10, ColumnRel 26))
        fromRange' (CellRef "$Y$5:$Z$10") @?= Just ((RowAbs 5, ColumnAbs 25), (RowAbs 10, ColumnAbs 26))
        fromRange' (CellRef "myWorksheet!$Y5:Z$10") @?= Just ((RowRel 5, ColumnAbs 25), (RowAbs 10, ColumnRel 26))
        fromForeignRange (CellRef "myWorksheet!$Y5:Z$10") @?= Just ("myWorksheet", ((RowRel 5, ColumnAbs 25), (RowAbs 10, ColumnRel 26)))
        fromForeignRange (CellRef "'myWorksheet'!Y5:Z10") @?= Just ("myWorksheet", ((RowRel 5, ColumnRel 25), (RowRel 10, ColumnRel 26)))
        fromForeignRange (CellRef "'my sheet'!Y5:Z10") @?= Just ("my sheet", ((RowRel 5, ColumnRel 25), (RowRel 10, ColumnRel 26)))
        fromForeignRange (CellRef "$Y5:Z$10") @?= Nothing
  ]

instance Monad m => Serial m RowIndex where
  series = cons1 (RowIndex . getPositive)

instance Monad m => Serial m ColumnIndex where
  series = cons1 (ColumnIndex . getPositive)

instance Monad m => Serial m RowCoord where
  series = cons1 (RowAbs . getPositive) \/ cons1 (RowRel . getPositive)

instance Monad m => Serial m ColumnCoord where
  series = cons1 (ColumnAbs . getPositive) \/ cons1 (ColumnRel . getPositive)

-- | Allow defining an instance to generate valid foreign range params
data MkForeignRangeRef =
  MkForeignRangeRef NameString RangeCoord
  deriving (Show)

viewForeignRangeParams :: MkForeignRangeRef -> (Text, RangeCoord)
viewForeignRangeParams (MkForeignRangeRef nameStr range) = (nameString nameStr, range)

instance Monad m => Serial m MkForeignRangeRef where
  series = cons2 MkForeignRangeRef

-- | Allow defining an instance to generate valid foreign cellref params
data MkForeignCellRef =
  MkForeignCellRef NameString CellCoord
  deriving (Show)

viewForeignCellParams :: MkForeignCellRef -> (Text, CellCoord)
viewForeignCellParams (MkForeignCellRef nameStr coord) = (nameString nameStr, coord)

instance Monad m => Serial m MkForeignCellRef where
  series = cons2 MkForeignCellRef

-- | Overload an instance for allowed sheet name chars
newtype NameChar =
  NameChar { _unNCh :: Char }
  deriving (Show, Eq)

-- | Instance for Char which broadens the pool to permitted ascii and some wchars
-- as @Serial m Char@ is only @[ 'A' .. 'Z' ]@
instance Monad m => Serial m NameChar where
  series =
    fmap NameChar $
      let wChars = ['é', '€', '愛']
          startSeq = "ab cd'" -- single quote is permitted as long as it's not 1st or last
          authorizedAscii =
            let asciiRange = L.map (chr . fromIntegral) [minBound @Word8 .. maxBound `div` 2]
                forbiddenClass = "[]*:?/\\" :: String
                isAuthorized c = isPrint c && not (c `L.elem` forbiddenClass)
                isPlanned c = isAuthorized c && not (c `L.elem` startSeq)
              in L.filter isPlanned asciiRange
        in generate $ \d -> take (d + 1) $ startSeq ++ wChars ++ authorizedAscii

-- | Allow defining an instance to generate valid sheetnames (non-empty, valid char set, squote rule)
newtype NameString =
  NameString { _unNS :: Series.NonEmpty NameChar }
  deriving (Show)

nameString :: NameString -> Text
nameString = T.pack . L.map _unNCh . getNonEmpty . _unNS

instance Monad m => Serial m NameString where
  series = fmap NameString $
    series @m @(Series.NonEmpty NameChar) >>= \t -> t <$ guard (isOk t)
        where
          -- squote isn't permitted at the beginning and the end of a sheet's name
          isOk (getNonEmpty -> s) =
            head s /= NameChar '\'' &&
            last s /= NameChar '\''
