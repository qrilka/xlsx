
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
  [ testProperty "col2int . int2col == id" $
        \(Positive i) -> i == col2int (int2col i)

    , testProperty "row2coord . coord2row = id" $
        \(r :: Coord) -> r == row2coord (coord2row r)

    , testProperty "col2coord . coord2col = id" $
        \(c :: Coord) -> c == col2coord (coord2col c)

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
        singleCellRef' (mapBoth Rel (5, 25)) @?= CellRef "Y5"
        singleCellRef' (Rel 5, Abs 25) @?= CellRef "$Y5"
        singleCellRef' (Abs 5, Rel 25) @?= CellRef "Y$5"
        singleCellRef' (mapBoth Abs (5, 25)) @?= CellRef "$Y$5"
        singleCellRef (5, 25) @?= CellRef "Y5"
    , testCase "parsing single CellRefs as abstract coordinates" $ do
        fromSingleCellRef (CellRef "Y5") @?= Just (5, 25)
        fromSingleCellRef (CellRef "$Y5") @?= Just (5, 25)
        fromSingleCellRef (CellRef "Y$5") @?= Just (5, 25)
        fromSingleCellRef (CellRef "$Y$5") @?= Just (5, 25)
    , testCase "parsing single CellRefs as potentially absolute coordinates" $ do
        fromSingleCellRef' (CellRef "Y5") @?= Just (mapBoth Rel (5, 25))
        fromSingleCellRef' (CellRef "$Y5") @?= Just (Rel 5, Abs 25)
        fromSingleCellRef' (CellRef "Y$5") @?= Just (Abs 5, Rel 25)
        fromSingleCellRef' (CellRef "$Y$5") @?= Just (mapBoth Abs (5, 25))
        fromSingleCellRef' (CellRef "$Y$50") @?= Just (mapBoth Abs (50, 25))
        fromSingleCellRef' (CellRef "$Y$5$0") @?= Nothing
        fromSingleCellRef' (CellRef "Y5:Z10") @?= Nothing
    , testCase "building ranges" $ do
        mkRange (5, 25) (10, 26) @?= CellRef "Y5:Z10"
        mkRange' (mapBoth Rel (5, 25)) (mapBoth Rel (10, 26)) @?= CellRef "Y5:Z10"
        mkRange' (mapBoth Abs (5, 25)) (mapBoth Abs (10, 26)) @?= CellRef "$Y$5:$Z$10"
        mkRange' (Rel 5, Abs 25) (Abs 10, Rel 26) @?= CellRef "$Y5:Z$10"
        mkForeignRange "myWorksheet" (Rel 5, Abs 25) (Abs 10, Rel 26) @?= CellRef "'myWorksheet'!$Y5:Z$10"
        mkForeignRange "my sheet" (Rel 5, Abs 25) (Abs 10, Rel 26) @?= CellRef "'my sheet'!$Y5:Z$10"
    , testCase "parsing ranges CellRefs as abstract coordinates" $ do
        fromRange (CellRef "Y5:Z10") @?= Just ((5, 25), (10, 26))
        fromRange (CellRef "$Y$5:$Z$10") @?= Just ((5, 25), (10, 26))
        fromRange (CellRef "myWorksheet!$Y5:Z$10") @?= Just ((5, 25), (10, 26))
    , testCase "parsing ranges CellRefs as potentially absolute coordinates" $ do
        fromRange' (CellRef "Y5:Z10") @?= Just (mapBoth (mapBoth Rel) ((5, 25), (10, 26)))
        fromRange' (CellRef "$Y$5:$Z$10") @?= Just (mapBoth (mapBoth Abs) ((5, 25), (10, 26)))
        fromRange' (CellRef "myWorksheet!$Y5:Z$10") @?= Just ((Rel 5, Abs 25), (Abs 10, Rel 26))
        fromForeignRange (CellRef "myWorksheet!$Y5:Z$10") @?= Just ("myWorksheet", ((Rel 5, Abs 25), (Abs 10, Rel 26)))
        fromForeignRange (CellRef "'myWorksheet'!Y5:Z10") @?= Just ("myWorksheet", mapBoth (mapBoth Rel) ((5, 25), (10, 26)))
        fromForeignRange (CellRef "'my sheet'!Y5:Z10") @?= Just ("my sheet", mapBoth (mapBoth Rel) ((5, 25), (10, 26)))
        fromForeignRange (CellRef "$Y5:Z$10") @?= Nothing
  ]


instance Monad m => Serial m Coord where
  series = cons1 (Abs . getPositive) \/ cons1 (Rel . getPositive)

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
