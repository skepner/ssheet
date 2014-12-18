{- |
Interface for reading spreadsheet data in Csv, Xls, Xlsx format (detected by filename extension)
and writing in Csv and JSON formats.
-}

{-# LANGUAGE FlexibleContexts, OverloadedStrings #-}

module Data.Ssheet (
  Cell(..),
  Ssheet,
  Sheet(..),
  Errors,
  SsheetOptions(..),
  ssheetDefaultOptions,
  ssheetRead,
  ssheetReadFile,
  ssheetNumberOfRows,
  ssheetNumberOfColumns,
  ssheetCell,
  ssheetJsonPrettyPrint,
  ssheetJsonPrettyPrintToString,
  ssheetJsonPrettyPutLn,
  ssheetCsv,
  ssheetCsvToFile
  ) where

import Debug.Trace
import System.FilePath
import qualified Data.Text as T
import qualified Data.Text.IO as T
import qualified Data.Text.Read as T
import qualified Data.ByteString.Lazy.Char8 as BS8
import Data.Aeson.Encode.Pretty
import Data.Ord
import Text.Printf
import qualified Data.Map.Lazy as Map

import Data.Ssheet.Types
import Data.Ssheet.XlsxParser
import Data.Ssheet.XlsParser
import Data.Ssheet.Csv

----------------------------------------------------------------------

-- | Extracts sheets from a file.
ssheetReadFile :: SsheetOptions -> FilePath -> IO (Either Errors Ssheet)
ssheetReadFile options filename =
  BS8.readFile filename >>= ssheetRead options (takeExtension filename)

----------------------------------------------------------------------

-- | Extracts sheets from a lazy ByteString.
-- | Result has to be wrapped in IO because xlsRead returns IO
ssheetRead :: SsheetOptions -> FilePath -> BS8.ByteString -> IO (Either Errors Ssheet)
ssheetRead options extension bytes =
  case extension of
   ".xlsx" -> return $ xlsxRead options bytes
   ".xls"  -> xlsRead options bytes
   ".csv"  -> return $ csvRead options bytes
   _       -> return $ Left [T.pack $ "Unrecognized filename suffix: " ++ extension]

----------------------------------------------------------------------

-- | Number of rows in the sheet, i.e. max row index plus 1
ssheetNumberOfRows :: Sheet -> RowNo
ssheetNumberOfRows = ((+1)) . fst . Map.findMax . sheetContent

----------------------------------------------------------------------

-- | Number of columns in the sheet, i.e. max column index plus 1
ssheetNumberOfColumns :: Sheet -> ColNo
ssheetNumberOfColumns =
  ((+1)) . Map.foldr maxCol 0 . sheetContent
  where
    maxCol :: Row -> ColNo -> ColNo
    maxCol row = max (fst $ Map.findMax row)

----------------------------------------------------------------------

-- | Returns cell by row and col. For invalid or too big row and col returns CellEmpty
ssheetCell :: Sheet -> RowNo -> ColNo -> Cell
ssheetCell ssheet rowNo colNo =
  case Map.lookup rowNo (sheetContent ssheet) >>= Map.lookup colNo of
   Just cell -> cell
   Nothing   -> CellEmpty

----------------------------------------------------------------------

-- | Generates JSON from Ssheet and pretty prints it into ByteString
--
--   JSON structure: [{"name": <sheet-name>, "content" : {<row-no>: {<col-no>: <cell-as-string>}}}]
ssheetJsonPrettyPrint :: Ssheet -> BS8.ByteString
ssheetJsonPrettyPrint =
  encodePretty' (defConfig {confCompare = comparing key})
  where
    key :: T.Text -> T.Text
    key a = case T.decimal a of
      Left _ -> a
      Right (d, "") -> p d
      Right _ -> a
    p :: Int -> T.Text
    p = T.pack . printf "%05d"

ssheetJsonPrettyPrintToString :: Ssheet -> String
ssheetJsonPrettyPrintToString = BS8.unpack . ssheetJsonPrettyPrint

-- | Generates JSON from Ssheet and pretty prints it using putStrLn
ssheetJsonPrettyPutLn :: Ssheet -> IO ()
ssheetJsonPrettyPutLn = BS8.putStrLn . ssheetJsonPrettyPrint

----------------------------------------------------------------------

-- | Generates CSV for single sheet
ssheetCsv :: Sheet -> T.Text
ssheetCsv sheet = csvGenerate (ssheetNumberOfRows sheet) (ssheetNumberOfColumns sheet) (ssheetCell sheet)

-- | Generates CSV for single sheet and writes it into file
ssheetCsvToFile :: Sheet -> FilePath -> IO ()
ssheetCsvToFile ssheet filename = T.writeFile filename (ssheetCsv ssheet)

----------------------------------------------------------------------
