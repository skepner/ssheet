{-# LANGUAGE OverloadedStrings #-}

module Data.Ssheet.Csv (
  csvRead,
  csvGenerate,
  ) where

----------------------------------------------------------------------

import Debug.Trace

import qualified Data.Text as T
import qualified Data.ByteString.Lazy.Char8 as BS8
import Text.ParserCombinators.Parsec (parse, Parser, (<|>), (<?>), sepEndBy, sepBy, many, noneOf, char, try, string)
import qualified Data.Map.Lazy as Map
import Data.Maybe

import Data.Ssheet.Types
import Data.Ssheet.Utils

----------------------------------------------------------------------

csvRead :: SsheetOptions -> BS8.ByteString -> Either Errors Ssheet
csvRead options source =
  csvToSsheet $ parse csvParser "bytestring" (BS8.unpack source)
  where
    csvToSsheet :: Show a => Either a [[String]] -> Either Errors Ssheet
    csvToSsheet csv =
      case csv of
       Left err -> Left [T.pack $ show err]
       Right r  -> Right [makeSheet r]
    makeSheet :: [[String]] -> Sheet
    makeSheet src = Sheet "" (makeContent src)
    makeContent :: [[String]] -> SheetContent
    makeContent = Map.fromList . filter removeEmptyRow . zip [0..] . map makeRow
    makeRow :: [String] -> Row
    makeRow = Map.fromList . filter removeEmptyCol . zip [0..] . map makeCell
    makeCell :: String -> Cell
    makeCell s = case (if stripStrings options then T.strip . T.pack else T.pack) s of
      "" -> CellEmpty
      t  -> CellString t

----------------------------------------------------------------------

csvParser :: Parser [[String]]
csvParser =
  row `sepEndBy` eol
  where
    eol = string "\n" <|> try (string "\r\n") <|> string "\r" <?> "end of line"

    row = cell `sepBy` char ','

    cell = quotedCell <|> many (noneOf ",\r\n")
    quotedCell = do
      _ <- char '"'
      content <- many quotedChar
      _ <- char '"' <?> "\" at end of cell"
      return content
    quotedChar = noneOf "\"" <|> try (string "\"\"" >> return '"')

----------------------------------------------------------------------

csvGenerate :: RowNo -> ColNo -> (RowNo -> ColNo -> Cell) -> T.Text
csvGenerate numRows numCols cellGetter =
  T.unlines $ map makeRecord [0..numRows-1]
  where
    makeRecord :: RowNo -> T.Text
    makeRecord rowNo = T.intercalate "," $ map (encloseQuotes . escape . cellAsText . cellGetter rowNo) [0..numCols-1]
    escape = T.replace "\"" "\"\""
    encloseQuotes c =
      case c of
       ""                                                    -> c
       _ | isJust (T.find (\cc -> cc == '"' || cc == ',') c) -> T.concat ["\"", c, "\""]
       _                                                     -> c

----------------------------------------------------------------------
