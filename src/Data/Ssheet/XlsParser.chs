{- -*- Haskell -*- -}

{-# LANGUAGE TypeSynonymInstances, FlexibleInstances #-}

module Data.Ssheet.XlsParser (
  xlsRead
  ) where

import Foreign.C.Types
import Foreign.C.String
import Foreign.Ptr
-- import Foreign.Storable

import qualified Data.ByteString.Lazy.Char8 as BS8
import qualified Data.ByteString.Unsafe as BS8
import qualified Data.Map.Lazy as Map

import Data.Time.Parse
import Data.Time
import Data.Maybe
import qualified Data.Text as T

import System.IO.Unsafe

import Data.Ssheet.Types
import Data.Ssheet.Utils

----------------------------------------------------------------------

#include "../xls-parser/xls-parser.h"

{# pointer *ExcelData #}
{# pointer *ExcelData as XlsData -> ExcelData #}

foreign import ccall "excel_open" _excel_open :: CString -> CULong -> CInt -> XlsData

{#fun excel_open_file as _excel_open_file { `String', `Int' } -> `XlsData' #}
{#fun pure excel_number_of_sheets as xlsNumberOfSheets { `XlsData' } -> `Int' #}
{#fun pure excel_sheet_name as _excel_sheet_name { `XlsData', `Int' } -> `String' #}
{#fun pure excel_number_of_rows as xlsNumberOfRows { `XlsData', `Int' } -> `Int' #}
{#fun pure excel_number_of_columns as xlsNumberOfColumns { `XlsData', `Int', `Int' } -> `Int' #}
{#fun pure excel_cell_as_text as xlsCellAsString { `XlsData', `Int', `Int', `Int' } -> `String' #}

----------------------------------------------------------------------

xlsRead :: SsheetOptions -> BS8.ByteString -> IO (Either Errors Ssheet)
xlsRead options excelData =
  BS8.unsafeUseAsCString (BS8.toStrict excelData) $ \excel_data_c ->
    return $ xlsExtract (_excel_open excel_data_c (fromIntegral $ BS8.length excelData) (fromIntegral $ fromEnum $ stripStrings options)) "Unable to read source"

----------------------------------------------------------------------

-- | converts data from XlsData (provided by C lib) to Ssheet
xlsExtract :: XlsData -> String -> Either Errors Ssheet
xlsExtract excelData errMsg =
  case excelData of
   d | d == nullPtr -> Left [T.pack errMsg]
   _                -> Right (map sheet [0..xlsNumberOfSheets excelData - 1])
  where
    sheet sheetNo =
      Sheet (T.pack $ _excel_sheet_name excelData sheetNo) (extractSheet sheetNo (xlsNumberOfRows excelData sheetNo))
    extractSheet sheetNo numRows =
      (Map.fromList . filter removeEmptyRow . zip [0..] . map (cellsOfRow sheetNo)) [0..numRows-1]
    cellsOfRow sheetNo rowNo =
      (Map.fromList . filter removeEmptyCol . zip [0..] . map (xlsCell excelData sheetNo rowNo)) [0..xlsNumberOfColumns excelData sheetNo rowNo - 1]

----------------------------------------------------------------------

-- | returns Cell content. Arguments: excelData sheetNo rowNo colNo
xlsCell :: XlsData -> Int -> Int -> Int -> Cell
xlsCell excelData sheetNo rowNo colNo =
  case take 3 cell of
   "" -> CellEmpty
   ":i:" -> CellInt (read $ drop 3 cell)
   ":f:" -> CellFloat (read $ drop 3 cell)
   ":d:" -> CellDate (parseDate $ drop 3 cell)
   _     -> CellString $ T.pack cell
  where
    cell = xlsCellAsString excelData sheetNo rowNo colNo
    parseDate text = localDay $ fst $ fromJust (strptime "%Y-%m-%d" text)

----------------------------------------------------------------------
