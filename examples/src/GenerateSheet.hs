{-# LANGUAGE DerivingStrategies  #-}
{-# LANGUAGE LambdaCase          #-}
{-# LANGUAGE OverloadedStrings   #-}
{-# LANGUAGE ScopedTypeVariables #-}

module GenerateSheet where

import Codec.Xlsx (Cell (..), CellValue (..), DataValidation, ErrorStyle (..),
                   ErrorType (..), ListOrRangeExpression (..), SqRef (..),
                   ValidationType (..), Worksheet, atCell, atSheet, cellValueAt,
                   def, fromXlsx, mkRange, simpleCellFormula)
import Data.Text (Text)
import Data.Time.Clock.POSIX (getPOSIXTime)
import Lens.Micro ((&), (.~), (?~))

import qualified Codec.Xlsx as Xlsx
import qualified Data.ByteString.Lazy as L
import qualified Data.Map.Strict as Map

main :: IO ()
main = do
    ct <- getPOSIXTime
    let name :: String = "example"
        ws = def
                 & addHeaders
                 & addItems
                 & Xlsx.wsDataValidations .~ validations
        xlsx = def & atSheet "List1" ?~ ws

    L.writeFile (name <> ".xlsx") $ fromXlsx ct xlsx

addHeaders :: Worksheet -> Worksheet
addHeaders sheet =
    sheet
        & cellValueAt (1, 1) ?~ CellText "Make"                 -- A simple text
        & cellValueAt (1, 2) ?~ CellText "Model"                -- An Integer
        & cellValueAt (1, 3) ?~ CellText "Year"                 -- Data for Year will be a Maybe value
        & cellValueAt (1, 4) ?~ CellText "Sales (in thousands)" -- Decimal number
        & cellValueAt (1, 5) ?~ CellText "Fuel type"            -- This will be a dropdown and the allowed values are: Gasoline, Diesel, Hybrid and Electric
        & cellValueAt (1, 6) ?~ CellText "Is in production?"    -- Boolean value

-- We'll make a 6 x 6 sheet with car data
addItems :: Worksheet -> Worksheet
addItems sheet =
    sheet
        & cellValueAt (2, 1) ?~ CellText "Toyota"
        & cellValueAt (2, 2) ?~ CellText "Corolla"
        & maybe id (\year -> cellValueAt (2, 3) ?~ CellDouble year) (Just 2021)
        & cellValueAt (2, 4) ?~ CellDouble 120.5
        & cellValueAt (2, 5) ?~ (CellText $ fuelTypeToText Gasoline)
        & cellValueAt (2, 6) ?~ CellBool True

        & cellValueAt (3, 1) ?~ CellText "Ford"
        & cellValueAt (3, 2) ?~ CellText "Mustang"
        & maybe id (\year -> cellValueAt (3, 3) ?~ CellDouble year) Nothing
        & cellValueAt (3, 4) ?~ CellDouble 85.3
        & cellValueAt (3, 5) ?~ (CellText $ fuelTypeToText Gasoline)
        & cellValueAt (3, 6) ?~ CellBool True

        & cellValueAt (4, 1) ?~ CellText "Audi"
        & cellValueAt (4, 2) ?~ CellText "A4"
        & maybe id (\year -> cellValueAt (4, 3) ?~ CellDouble year) (Just 2019)
        & cellValueAt (4, 4) ?~ CellDouble 34.9
        & cellValueAt (4, 5) ?~ (CellText $ fuelTypeToText Diesel)
        & cellValueAt (4, 6) ?~ CellBool False

        & cellValueAt (5, 1) ?~ CellText "Hyundai"
        & cellValueAt (5, 2) ?~ CellText "Elantra"
        & maybe id (\year -> cellValueAt (5, 3) ?~ CellDouble year) Nothing
        & cellValueAt (5, 4) ?~ CellDouble 88.1
        & cellValueAt (5, 5) ?~ (CellText $ fuelTypeToText Hybrid)
        & cellValueAt (5, 6) ?~ CellBool True

        & cellValueAt (6, 1) ?~ CellText "Tesla"
        & cellValueAt (6, 2) ?~ CellText "Model 3"
        & maybe id (\year -> cellValueAt (6, 3) ?~ CellDouble year) (Just 2023)
        & cellValueAt (6, 4) ?~ CellDouble 210.7
        & cellValueAt (6, 5) ?~ (CellText $ fuelTypeToText Electric)
        & cellValueAt (6, 6) ?~ CellBool True

        -- This extra row is to only to demonstrate CellError in the Year column
        & cellValueAt (7, 1) ?~ CellText "Tesla"
        & cellValueAt (7, 2) ?~ CellText "Model Y"
        & maybe id (\year -> cellValueAt (7, 3) ?~ CellDouble year) (Just 2022)
        & atCell (7, 4) ?~ Cell Nothing (Just $ CellError ErrorDiv0) Nothing (Just $ simpleCellFormula "1/0")
        & cellValueAt (7, 5) ?~ (CellText $ fuelTypeToText Electric)
        & cellValueAt (7, 6) ?~ CellBool True

-- | Since values in dropdown column are bounded,
-- We can use Sum types for this ensuring safety both
-- in code and in excel.
data FuelType
    = Gasoline
    | Diesel
    | Hybrid
    | Electric
    deriving stock (Eq, Show, Enum, Bounded)

fuelTypeToText :: FuelType -> Text
fuelTypeToText = \case
    Gasoline -> "Gasoline"
    Diesel -> "Diesel"
    Hybrid -> "Hybrid"
    Electric -> "Electric"

fuelUniverse :: [FuelType]
fuelUniverse = [minBound .. maxBound]

-- | We need to add validation logic for rendering dropdown cells.
-- The core idea is to specify the allowed values using @ListExpression@
-- in a range of cells (done via mkRange).
validations :: Map.Map SqRef DataValidation
validations =
    Map.fromList
        [ ( SqRef [mkRange (2, 5) (6, 5)]
            , def
                & Xlsx.dvAllowBlank
                    .~ False
                & Xlsx.dvError
                    ?~ "Incorrect value"
                & Xlsx.dvErrorStyle
                    .~ ErrorStyleInformation
                & Xlsx.dvErrorTitle
                    ?~ "Year"
                & Xlsx.dvPrompt
                    ?~ "When was the car produced?"
                & Xlsx.dvShowErrorMessage
                    .~ True
                & Xlsx.dvShowInputMessage
                    .~ True
                & Xlsx.dvValidationType
                    .~ (ValidationTypeList . ListExpression $ fmap fuelTypeToText fuelUniverse)
            )
        ]
