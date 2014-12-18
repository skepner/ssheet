-- | xlsx file reader
-- | Inspired by:
-- | http://github.com/dilshod/xlsx2csv
-- | http://github.com/staale/python-xlsx
-- | http://github.com/leegao/pyXLSX

{-# LANGUAGE OverloadedStrings #-}

module Data.Ssheet.XlsxParser (
  xlsxRead
  ) where

----------------------------------------------------------------------

import Debug.Trace
import qualified Data.ByteString.Lazy.Char8 as BS8
import Data.Monoid
import Control.Monad
import qualified Data.Text as T

import qualified Codec.Archive.Zip as Zip

import qualified Text.XML as XML
import qualified Text.XML.Cursor as XML
import Text.XML.Cursor (($//), (&|), (&//))

import qualified Data.Map.Lazy as Map

import Data.Ssheet.Types
import Data.Ssheet.Utils

----------------------------------------------------------------------

type SheetId = T.Text

-- styles.xml
type NumberFormats = [(NumberFormatId, NumberFormatCode)]
type NumberFormatId = T.Text
type NumberFormatCode = T.Text

type CellFormattingRecords = [CellFormattingRecord]
data CellFormattingRecord = CellFormattingNumber NumberFormatId
                          | CellFormattingUnknown
                          deriving (Show)

----------------------------------------------------------------------

xlsxRead :: SsheetOptions -> BS8.ByteString -> Either Errors Ssheet
xlsxRead options bytes = do
  sheetIds >>= mapM extractSheet
  where

    sheetIds :: Either Errors [(SheetName, SheetId)]
    sheetIds =
      let sheetsCursor = fmap ($// XML.laxElement "sheet") (xmlCursor "xl/workbook.xml")
          sheetName :: XML.Cursor -> Either Errors SheetName
          sheetName cursor = case XML.attribute "name" cursor of
                              [name] -> Right name
                              _ -> Left ["[Invalid xlsx]: Cannot find sheet name"]
          sheetInfo :: (SheetId, XML.Cursor) -> Either Errors (SheetName, SheetId)
          sheetInfo (sheetId, cursor) = sequenceEitherPair (sheetName cursor, Right sheetId)
          zipSheetId :: [XML.Cursor] -> [(T.Text, XML.Cursor)]
          zipSheetId = zip $ map (T.pack . show) ([1..] :: [Int])
      in
       sheetsCursor >>= mapM sheetInfo . zipSheetId

    extractSheet :: (SheetName, SheetId) -> Either Errors Sheet
    extractSheet (sheetName, sheetId) =
      worksheet >>= sheetData >>= mapM cellsOfRow >>= filterM (Right . removeEmptyRow) >>= \rows -> Right (Sheet sheetName (Map.fromList rows))
      where
        worksheet = xmlCursor $ "xl/worksheets/sheet" ++ T.unpack sheetId ++ ".xml"
        sheetData :: XML.Cursor -> Either Errors [XML.Cursor]
        sheetData cursor = case cursor $// XML.laxElement "sheetData" of
              [node] -> Right $ XML.child node
              _ -> Left ["[Invalid xlsx]: Cannot find sheetData for sheet ", sheetId]

    ----------------------------------------------------------------------

    zzip = Zip.toArchive bytes

    xmlCursor :: String -> Either Errors XML.Cursor
    xmlCursor path =
      case Zip.findEntryByPath path zzip of
       Nothing -> Left [T.concat ["[Invalid xlsx]: Cannot find zip entry: ", T.pack path]]
       Just entry ->
         case XML.parseLBS XML.def (Zip.fromEntry entry) of
          Left err -> Left [T.concat ["[Invalid xlsx]: Cannot parse zip entry ", T.pack path, ": ", T.pack $ show err]]
          Right xmlDocument -> Right $ XML.fromDocument xmlDocument


    rowNo :: XML.Cursor -> RowNo
    rowNo = (\v -> v - 1) . read . T.unpack . head . XML.attribute "r"

    cellsOfRow :: XML.Cursor -> Either Errors (RowNo, Row)
    cellsOfRow row = mapM extractCell (XML.child row) >>= filterM (Right . removeEmptyCol) >>= \cells -> return (rowNo row, Map.fromList cells)

    extractCell :: XML.Cursor -> Either Errors (ColNo, Cell)
    extractCell column = do
      sharedStringsValidated <- sharedStrings
      let sharedStringsStripped = if stripStrings options then map T.strip sharedStringsValidated else sharedStringsValidated
      case (XML.attribute "t" column, XML.attribute "s" column, XML.attribute "r" column) of
       (["s"], _, r) ->
         let sharedStringIndexS = column $// XML.child &| XML.content
             sharedStringIndex = case sharedStringIndexS of
                                 [[index]] -> Right (read $ T.unpack index :: Int)
                                 _ -> Left ["[Invalid xlsx]: Cannot find shared string index for cell"]
             mkStringCell v = if T.null v then CellEmpty else CellString v
         in
          sharedStringIndex >>= (\index -> Right (ssheetTextsToCol r, mkStringCell (sharedStringsStripped !! index)))
       ([], [cellFormatId], r) -> do
         value <- extractFormatCellValue column cellFormatId
         return (ssheetTextsToCol r, value)
       (["n"], [cellFormatId], r) -> do
         value <- extractFormatCellValue column cellFormatId
         return (ssheetTextsToCol r, value)
       (_, _, r) ->
         return (ssheetTextsToCol r, extractCellValue column)

    -- extractFormatCellValue cell formatId | trace ("extractFormatCellValue " ++ show (extractCellValue cell) ++ " formatId:" ++ show formatId) False = undefined
    extractFormatCellValue cell formatId =
      applyFormat (read (T.unpack formatId) :: Int) (extractCellValue cell)
    extractCellValue cell = case cell $// XML.laxElement "v" of
                             [columnValue] -> CellFloat (read $ T.unpack $ head $ columnValue $// XML.content)
                             _ -> CellEmpty

    sharedStrings :: Either Errors [T.Text]
    sharedStrings =      -- xl/sharedStrings.xml: <sst><si><t>STRING</t></si><si><t>STRING</t></si></sst>
      let selectSi cursor = cursor $// XML.laxElement "si" -- XML.laxElement - match element tag without namespace
          selectTChildContent cursor = cursor $// (XML.laxElement "t" >=> XML.child) &| XML.content
      in do
         cursor <- xmlCursor "xl/sharedStrings.xml"
         return $ mconcat $ map (mconcat . selectTChildContent) (selectSi cursor)

    styles :: Either Errors (NumberFormats, CellFormattingRecords)
    styles = do
      cursor <- xmlCursor "xl/styles.xml"
      let expectValue = expectSingleValue "[Invalid xlsx]: invalid xl/styles.xml"
          numFmts = cursor $// XML.laxElement "numFmts" &// XML.laxElement "numFmt"
          extractFormatAttrs node = sequenceEitherPair (expectValue (XML.attribute "numFmtId" node), expectValue (XML.attribute "formatCode" node))
          cellXfs = cursor $// XML.laxElement "cellXfs" &// XML.laxElement "xf"
          makeFormattingRecord xf = case XML.attribute "numFmtId" xf of
            [value] -> Right $ CellFormattingNumber value
            _ -> Right CellFormattingUnknown
      sequenceEitherPair (mapM extractFormatAttrs numFmts, mapM makeFormattingRecord cellXfs)

    applyFormat :: Int -> Cell -> Either Errors Cell
    -- applyFormat cellFormatId (CellFloat value) | trace ("applyFormat " ++ show cellFormatId ++ " " ++ show value) False = undefined
    applyFormat cellFormatId (CellFloat value) =
      findFormat cellFormatId >>= applyFormatCode value
      where
        findFormat cellFormatId' | cellFormatId' > 0x32 = do -- custom format id
          (numberFormats, cellFormattingRecords) <- styles
          case cellFormattingRecords !! cellFormatId' of
           CellFormattingUnknown -> return ""
           CellFormattingNumber fid ->
             case lookup fid numberFormats of
              Nothing -> return ""
              Just formatCode -> return $ T.toUpper formatCode
        findFormat cellFormatId'
          | cellFormatId' <= 0x0D && value > 39000.0 && value < 42000.0 = return "YYYY-MM-DD" -- for unclear reason Libre/Openoffice sometimes uses those formats for dates
        findFormat cellFormatId' -- standard format ids for date and time (see table at the bottom)
          | cellFormatId' >= 0x0E && cellFormatId' <= 0x16 = return "YYYY-MM-DD"
        findFormat _ = Right ""
    applyFormat _ cellValue = return cellValue

    applyFormatCode :: Float -> T.Text -> Either Errors Cell
    -- applyFormatCode value formatCode | trace ("Format [" ++ show value ++ "] with code " ++ show formatCode) False = undefined
    applyFormatCode value formatCode
      | "YY" `T.isInfixOf` formatCode =
          -- trace ("applyFormatCode " ++ show (CellDate (excelSerialDateToDay value))) $
          Right $ CellDate (excelSerialDateToDay value)
    applyFormatCode value _ = Right $ CellFloat value

    expectSingleValue :: T.Text -> [a] -> Either Errors a
    expectSingleValue _ [value] = Right value
    expectSingleValue errMsg _ = Left [errMsg]

    sequenceEitherPair :: (Either Errors a, Either Errors b) -> Either Errors (a, b)
    sequenceEitherPair pair =
      case pair of
       (Left e1, Left e2) -> Left (e1 ++ e2)
       (Left e1, Right _) -> Left e1
       (Right _, Left e2) -> Left e2
       (Right a, Right b) -> Right (a, b)

----------------------------------------------------------------------

{-
    # "std" == "standard for US English locale"
    # See e.g. gnumeric-1.x.y/src/formats.c
    0x00: "General",
    0x01: "0",
    0x02: "0.00",
    0x03: "#,##0",
    0x04: "#,##0.00",
    0x05: "$#,##0_);($#,##0)",
    0x06: "$#,##0_);[Red]($#,##0)",
    0x07: "$#,##0.00_);($#,##0.00)",
    0x08: "$#,##0.00_);[Red]($#,##0.00)",
    0x09: "0%",
    0x0a: "0.00%",
    0x0b: "0.00E+00",
    0x0c: "# ?/?",
    0x0d: "# ??/??",
    0x0e: "m/d/yy",
    0x0f: "d-mmm-yy",
    0x10: "d-mmm",
    0x11: "mmm-yy",
    0x12: "h:mm AM/PM",
    0x13: "h:mm:ss AM/PM",
    0x14: "h:mm",
    0x15: "h:mm:ss",
    0x16: "m/d/yy h:mm",
    0x25: "#,##0_);(#,##0)",
    0x26: "#,##0_);[Red](#,##0)",
    0x27: "#,##0.00_);(#,##0.00)",
    0x28: "#,##0.00_);[Red](#,##0.00)",
    0x29: "_(* #,##0_);_(* (#,##0);_(* \"-\"_);_(@_)",
    0x2a: "_($* #,##0_);_($* (#,##0);_($* \"-\"_);_(@_)",
    0x2b: "_(* #,##0.00_);_(* (#,##0.00);_(* \"-\"??_);_(@_)",
    0x2c: "_($* #,##0.00_);_($* (#,##0.00);_($* \"-\"??_);_(@_)",
    0x2d: "mm:ss",
    0x2e: "[h]:mm:ss",
    0x2f: "mm:ss.0",
    0x30: "##0.0E+0",
    0x31: "@",
-}

----------------------------------------------------------------------
