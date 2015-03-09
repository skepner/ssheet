{-# LANGUAGE OverloadedStrings #-}

-- import Debug.Trace

import System.Exit
-- import System.Posix.Directory
import qualified Data.Text as T
-- import qualified Data.Text.IO as T
import Control.Monad
-- import qualified Data.ByteString.Lazy.Char8 as BS8
import qualified Data.Map.Lazy as Map
import Data.Time.Calendar

import Data.Ssheet

----------------------------------------------------------------------

main :: IO ()
main =
  mapM_ testIt [("tests/simple.xlsx", simpleExpected),
                ("tests/simple.xls", simpleExpected),
                ("tests/simple.csv", simpleExpectedCsv),
                ("tests/dates.xlsx", datesExpected),
                ("tests/dates1904.xlsx", dates1904Expected)
                ]
  where
    testIt (source, expected) = testSimple source expected

----------------------------------------------------------------------

showError :: [T.Text] -> IO ()
showError err = putStrLn $ "Errors: " ++ show (length err) ++ "\n" ++ T.unpack (T.unlines err)

----------------------------------------------------------------------

testSimple :: FilePath -> Ssheet -> IO ()
testSimple source expected = do
  putStrLn $ "Testing " ++ source
  either_ssheet <- ssheetReadFile ssheetDefaultOptions source
  case either_ssheet of
   Left err -> do
     showError err
     exitFailure
   Right ssheet -> do
     unless (ssheet == expected) (error $ source ++ ": unexpected data extracted:\n" ++ (ssheetJsonPrettyPrintToString ssheet))

simpleExpected :: Ssheet
simpleExpected = [Sheet "Sheet1" (Map.fromList [(0,Map.fromList [(0,CellString "A1")]),
                                                (1,Map.fromList [(0,CellFloat 42.0)]),
                                                (2,Map.fromList [(0,CellFloat 42.42)]),
                                                (3,Map.fromList [(0,CellDate (fromGregorian 2014 12 16))])])]

simpleExpectedCsv :: Ssheet
simpleExpectedCsv = [Sheet "" (Map.fromList [(0,Map.fromList [(0,CellString "A1")]),
                                             (1,Map.fromList [(0,CellString "42")]),
                                             (2,Map.fromList [(0,CellString "42.42")]),
                                             (3,Map.fromList [(0,CellString "2014-12-16")])])]

datesExpected :: Ssheet
datesExpected = [Sheet "Sheet1" (Map.fromList [(0, Map.fromList  [(0, CellDate (fromGregorian 2015 3 8))]),
                                               (1, Map.fromList  [(0, CellDate (fromGregorian 2015 3 8))]),
                                               (2, Map.fromList  [(0, CellDate (fromGregorian 2015 3 8))]),
                                               (3, Map.fromList  [(0, CellDate (fromGregorian 2015 3 8))]),
                                               (4, Map.fromList  [(0, CellDate (fromGregorian 2015 3 8))]),
                                               (5, Map.fromList  [(0, CellDate (fromGregorian 2015 3 8))]),
                                               (6, Map.fromList  [(0, CellDate (fromGregorian 2015 3 8))]),
                                               (7, Map.fromList  [(0, CellDate (fromGregorian 2015 3 8))]),
                                               (8, Map.fromList  [(0, CellDate (fromGregorian 2015 3 8))]),
                                               (9, Map.fromList  [(0, CellDate (fromGregorian 2015 3 8))]),
                                               (10, Map.fromList [(0, CellDate (fromGregorian 2015 3 8))])])]

dates1904Expected :: Ssheet
dates1904Expected = [Sheet "Sheet1" (Map.fromList [(0, Map.fromList [(0, CellDate (fromGregorian 2014 10 15))]),
                                                   (1, Map.fromList [(0, CellDate (fromGregorian 2014 10 15))])
                                                  ])]

----------------------------------------------------------------------
