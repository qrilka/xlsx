{-# LANGUAGE FlexibleInstances #-}
{-# LANGUAGE MultiParamTypeClasses #-}
{-# LANGUAGE ScopedTypeVariables #-}
{-# LANGUAGE OverloadedStrings #-}
{-# LANGUAGE ViewPatterns #-}
{-# LANGUAGE TypeApplications #-}
{-# OPTIONS_GHC -fno-warn-orphans #-}

module CommonTests
  ( tests
  ) where

import Control.Monad
import qualified Control.Applicative as Alt
import Data.Fixed (Pico, Fixed(..), E12)
import Data.Time.Calendar (fromGregorian)
import Data.Time.Clock (UTCTime(..))
import Data.Text (Text)
import Data.Word (Word8)
import Data.Char (isPrint, chr)
import qualified Data.Text as T
import qualified Data.List as L
import Test.Tasty.SmallCheck (testProperty)
import Test.SmallCheck.Series as Series
  ( Positive(..)
  , NonEmpty(..)
  , Serial(..)
  , newtypeCons
  , cons0
  , cons1
  , cons2
  , generate
  , (\/)
  )
import Test.Tasty (testGroup, TestTree)
import Test.Tasty.HUnit (testCase, (@?=))

import Codec.Xlsx.Types.Common

tests :: TestTree
tests =
  testGroup
    "Types.Common tests"
    [ testProperty "col2int . int2col == id" $
        \(Positive i) -> i == col2int (int2col i)
    , testCase "date conversions" $ do
        dateFromNumber DateBase1900 (- 2338.0) @?= read "1893-08-06 00:00:00 UTC"
        dateFromNumber DateBase1900 2.0 @?= read "1900-01-02 00:00:00 UTC"
        dateFromNumber DateBase1900 3687.0 @?= read "1910-02-03 00:00:00 UTC"
        dateFromNumber DateBase1900 38749.0 @?= read "2006-02-01 00:00:00 UTC"
        dateFromNumber DateBase1900 2958465.0 @?= read "9999-12-31 00:00:00 UTC"
        dateFromNumber DateBase1900 59.0 @?= read "1900-02-28 00:00:00 UTC"
        dateFromNumber DateBase1900 59.5 @?= read "1900-02-28 12:00:00 UTC"
        dateFromNumber DateBase1900 60.0 @?= read "1900-03-01 00:00:00 UTC"
        dateFromNumber DateBase1900 60.5 @?= read "1900-03-01 00:00:00 UTC"
        dateFromNumber DateBase1900 61 @?= read "1900-03-01 00:00:00 UTC"
        dateFromNumber DateBase1900 61.5 @?= read "1900-03-01 12:00:00 UTC"
        dateFromNumber DateBase1900 62 @?= read "1900-03-02 00:00:00 UTC"
        dateFromNumber DateBase1904 (-3800.0) @?= read "1893-08-05 00:00:00 UTC"
        dateFromNumber DateBase1904 0.0 @?= read "1904-01-01 00:00:00 UTC"
        dateFromNumber DateBase1904 2225.0 @?= read "1910-02-03 00:00:00 UTC"
        dateFromNumber DateBase1904 37287.0 @?= read "2006-02-01 00:00:00 UTC"
        dateFromNumber DateBase1904 2957003.0 @?= read "9999-12-31 00:00:00 UTC"
     , testCase "Converting dates in the vicinity of 1900-03-01 to numbers" $ do
        -- Note how the fact that 1900-02-29 exists for Excel forces us to skip 60
        dateToNumber DateBase1900 (UTCTime (fromGregorian 1900 2 28) 0) @?= (59 :: Double)
        dateToNumber DateBase1900 (UTCTime (fromGregorian 1900 3 1) 0) @?= (61 :: Double)
     , testProperty "dateToNumber . dateFromNumber == id" $
       -- Because excel treats 1900 as a leap year, dateToNumber and dateFromNumber
       -- aren't inverses of each other in the range n E [60, 61[ for DateBase1900
       \b (n :: Pico) -> (n >= 60 && n < 61 && b == DateBase1900) || n == dateToNumber b (dateFromNumber b $ n)

    , testProperty "row2coord . coord2row = id" $ do
        \(r :: Coord) -> r == row2coord (coord2row r)

    , testProperty "col2coord . coord2col = id" $ do
        \(c :: Coord) -> c == col2coord (coord2col c)

    , testProperty "fromSingleCellRef' . singleCellRef' = pure" $ do
        \(cellCoord :: CellCoord) -> pure cellCoord == fromSingleCellRef' (singleCellRef' cellCoord)

    , testProperty "fromRange' . mkRange' = pure" $ do
        \(range :: RangeCoord) -> pure range == fromRange' (uncurry mkRange' range)

    , testProperty "fromForeignSingleCellRef . mkForeignSingleCellRef = pure" $ do
        \(viewForeignCellParams -> params) ->
            pure params == fromForeignSingleCellRef (uncurry mkForeignSingleCellRef params)

    , testProperty "fromSingleCellRef' . mkForeignSingleCellRef = pure . snd" $ do
        \(viewForeignCellParams -> (nStr, cellCoord)) ->
            pure cellCoord == fromSingleCellRef' (mkForeignSingleCellRef nStr cellCoord)

    , testProperty "fromForeignSingleCellRef . singleCellRef' = const empty" $ do
        \(cellCoord :: CellCoord) ->
            Alt.empty == fromForeignSingleCellRef (singleCellRef' cellCoord)

    , testProperty "fromForeignRange . mkForeignRange = pure" $ do
        \(viewForeignRangeParams -> params@(nStr, range@(start, end))) ->
            pure params == fromForeignRange (mkForeignRange nStr start end)

    , testProperty "fromRange' . mkForeignRange = pure . snd" $ do
        \(viewForeignRangeParams -> (nStr, range@(start, end))) ->
            pure range == fromRange' (mkForeignRange nStr start end)

    , testProperty "fromForeignRange . mkRange' = const empty" $ do
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

instance Monad m => Serial m (Fixed E12) where
  series = newtypeCons MkFixed

instance Monad m  => Serial m DateBase where
  series = cons0 DateBase1900 \/ cons0 DateBase1904

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
              in L.filter isAuthorized asciiRange
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







