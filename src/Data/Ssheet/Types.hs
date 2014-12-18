{-# LANGUAGE OverloadedStrings, FlexibleInstances #-}

module Data.Ssheet.Types where

import Debug.Trace
import Data.Time.Calendar
import Data.Text
import Numeric

import qualified Data.Map.Lazy as Map
import Data.Aeson

----------------------------------------------------------------------

-- | Resulting data is a list of Sheets
type Ssheet = [Sheet]

-- | Sheet consists of its name and content
data Sheet = Sheet SheetName SheetContent deriving (Show, Eq)

-- | Sheet content is a map of row number to another map representing row.
type SheetContent = Map.Map RowNo Row

-- | Row is a map of column number to cell
type Row = Map.Map ColNo Cell

data Cell = CellEmpty
          | CellString Text
          | CellFloat Float
          | CellInt Int
          | CellDate Day
          deriving (Show, Eq)

-- | Errors are list of texts but there is usually just one element.
type Errors = [Text]

type SheetName = Text
type SheetNo = Int
type RowNo = Int
type ColNo = Int

----------------------------------------------------------------------

-- | Convert Cell to Text, dates are shown as YYYY-MM-DD
cellAsText :: Cell -> Text
-- cellAsText cell | trace ("cellAsText: " ++ show cell) False = undefined
cellAsText cell = case cell of
  CellEmpty -> ""
  CellString text -> text
  CellFloat f -> pack $ showFloatPerhapsAsInt f
  CellInt i -> pack $ show i
  CellDate d -> pack $ show d

-- | Convert Cell to String, dates are shown as YYYY-MM-DD
cellAsString :: Cell -> String
-- cellAsString cell | trace ("cellAsString: " ++ show cell) False = undefined
cellAsString cell = case cell of
  CellEmpty -> ""
  CellString text -> unpack text
  CellFloat f -> showFloatPerhapsAsInt f
  CellInt i -> show i
  CellDate d -> show d

-- if float has no fractional part, show it without decimal point
showFloatPerhapsAsInt :: Float -> String
showFloatPerhapsAsInt f =
  let showIt (digits, exponen) | exponen == 0 && digits == [0] = "0"
      showIt (digits, exponen) | exponen >= Prelude.length digits = show $ (floor f :: Int)
      showIt _ = show f
  in
   showIt $ floatToDigits 10 f

----------------------------------------------------------------------

-- | Convert sheet to JSON:
--   {"name": <sheet-name>, "content" : {<row-no>: {<col-no>: <cell-as-string>}}}
instance ToJSON Sheet where
  toJSON (Sheet name content) = object ["name" .= name, "content" .= toJSON (Map.mapKeys show content)]

instance ToJSON Row where
  toJSON row = toJSON (Map.mapKeys show row)

instance ToJSON Cell where
  toJSON CellEmpty = Null
  toJSON (CellString t) = toJSON t
  toJSON (CellFloat f) = toJSON f
  toJSON (CellInt i) = toJSON i
  toJSON (CellDate d) = object ["type" .= toJSON ("date" :: String), "value" .= toJSON (show d)]

----------------------------------------------------------------------

-- | Options to import xlsx, xls, csv
data SsheetOptions = SsheetOptions { stripStrings :: Bool -- ^remove leading and trailing whitespaces in CellString (default: True)
                                   }

ssheetDefaultOptions :: SsheetOptions
ssheetDefaultOptions = SsheetOptions { stripStrings = True }

----------------------------------------------------------------------

-- | Extract sheet content from sheet
sheetContent :: Sheet -> SheetContent
sheetContent (Sheet _ content) = content

----------------------------------------------------------------------
