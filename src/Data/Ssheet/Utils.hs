module Data.Ssheet.Utils where

import Debug.Trace
import qualified Data.Text as T
import Data.Char (chr, ord, isLetter)
import Data.Time
import qualified Data.Map.Lazy as Map

import Data.Ssheet.Types

{-
http://www.codeproject.com/Articles/2750/Excel-serial-date-to-Day-Month-Year-and-vise-versa
License: http://www.opensource.org/licenses/cddl1.php
-}

excelSerialDateToDay :: Float -> Day
excelSerialDateToDay sdf =
  excelSerialDateIToDay (truncate sdf)

excelSerialDateIToDay :: Int -> Day
{- Excel/Lotus 123 have a bug with 29-02-1900. 1900 is not a leap year, but Excel/Lotus 123 think it is... -}
excelSerialDateIToDay serialDate | serialDate == 60 = fromGregorian 1900 2 29
{- Modified Julian to DMY calculation with an addition of 2415019 -}
excelSerialDateIToDay serialDate =
  -- Because of the 29-02-1900 bug, any serial date under 60 is one off... Compensate.
  let date = if serialDate < 60 then serialDate + 1 else serialDate
      l1 = date + 68569 + 2415019
      n = ( 4 * l1 ) `quot` 146097
      l2 = l1 - (( 146097 * n + 3 ) `quot` 4)
      i = (4000 * (l2 + 1)) `quot` 1461001
      l3 = l2 - ((1461 * i) `quot` 4) + 31
      j = (80 * l3) `quot` 2447
      day = l3 - ((2447 * j) `quot` 80)
      l4 = j `quot` 11
      month = j + 2 - 12 * l4;
      year = 100 * (n - 49) + i + l4;
  in
   fromGregorian (fromIntegral year) month day

{-
void ExcelSerialDateToDMY(int serial_date, int* day, int* month, int* year)
{
    int l, n, i, j;
    // Excel/Lotus 123 have a bug with 29-02-1900. 1900 is not a
    // leap year, but Excel/Lotus 123 think it is...
    if (serial_date == 60) {
        *day = 29;
        *month = 2;
        *year = 1900;
    }
    else {
        if (serial_date < 60) {
        // Because of the 29-02-1900 bug, any serial date
        // under 60 is one off... Compensate.
        serial_date++;
        }

          // Modified Julian to DMY calculation with an addition of 2415019
        l = serial_date + 68569 + 2415019;
        n = (int)(( 4 * l ) / 146097);
        l = l - (int)(( 146097 * n + 3 ) / 4);

        i = (int)(( 4000 * ( l + 1 ) ) / 1461001);
        l = l - (int)(( 1461 * i ) / 4) + 31;
        j = (int)(( 80 * l ) / 2447);
        *day = l - (int)(( 2447 * j ) / 80);
        l = (int)(j / 11);
        *month = j + 2 - ( 12 * l );
        *year = 100 * ( n - 49 ) + i + l;
    }
}
-}

----------------------------------------------------------------------

stripTrailingEmptyCells :: [Cell] -> [Cell]
stripTrailingEmptyCells column | all (== CellEmpty) column = []
stripTrailingEmptyCells (e : column) = e : stripTrailingEmptyCells column
stripTrailingEmptyCells [] = []

----------------------------------------------------------------------

ssheetRowColToText :: RowNo -> ColNo -> T.Text
ssheetRowColToText row col = T.pack $ ssheetRowColToString row col

ssheetRowColToString :: RowNo -> ColNo -> String
ssheetRowColToString row col = ssheetColToString col ++ show (row + 1)

ssheetColToString :: ColNo -> String
ssheetColToString col =
  let first = col `div` 26
      second = col `rem` 26
      cc i = chr $ ord 'A' + i
  in
   case first of
    0 -> cc second : ""
    f | f <= 27 -> cc (first - 1) : cc second : ""
    _ -> fail $ "Column number too big: " ++ show col

----------------------------------------------------------------------

ssheetStringToRowCol :: String -> (RowNo, ColNo)
ssheetStringToRowCol s =
  case s of
   c1:c2:r | isLetter c2 -> (rowNo r, (ord c1 - ord 'A' + 1) * 26 + (ord c2 - ord 'A'))
   c:r -> (rowNo r, ord c - ord 'A')
   _ -> (-1, -1)
  where
    rowNo :: String -> RowNo
    rowNo r = (read r) - 1

ssheetStringToCol :: String -> ColNo
ssheetStringToCol s =
  case s of
   c1:c2:_ | isLetter c2 -> (ord c1 - ord 'A' + 1) * 26 + (ord c2 - ord 'A')
   c:_ -> ord c - ord 'A'
   _ -> -1

ssheetTextToRowCol :: T.Text -> (RowNo, ColNo)
ssheetTextToRowCol t = ssheetStringToRowCol (T.unpack t)

ssheetTextsToRowCol :: [T.Text] -> (RowNo, ColNo)
ssheetTextsToRowCol tt =
  case tt of
   t:_ -> ssheetTextToRowCol t
   _ -> (-2, -2)

ssheetTextToCol :: T.Text -> ColNo
ssheetTextToCol t = ssheetStringToCol (T.unpack t)

ssheetTextsToCol :: [T.Text] -> ColNo
ssheetTextsToCol tt =
  case tt of
   t:_ -> ssheetTextToCol t
   _ -> -2

----------------------------------------------------------------------

removeEmptyRow :: (RowNo, Row) -> Bool
removeEmptyRow (_, r) | Map.null r = False
removeEmptyRow _ = True

removeEmptyCol :: (ColNo, Cell) -> Bool
removeEmptyCol (_, CellEmpty) = False
removeEmptyCol _ = True

----------------------------------------------------------------------
