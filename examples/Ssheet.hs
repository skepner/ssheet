{-# LANGUAGE DeriveDataTypeable, OverloadedStrings, RecordWildCards #-}

import Debug.Trace
import qualified System.Console.CmdArgs as CmdArgs -- command line switches
import System.Console.CmdArgs ((&=))

import System.Exit
-- import Control.Exception
import qualified Data.Text as T
import qualified Data.Text.IO as T
-- import Data.Either (lefts, rights)
-- import qualified Data.Map.Strict as Map
import qualified Data.ByteString.Lazy.Char8 as BS8

import Data.Ssheet

----------------------------------------------------------------------

main = do
  options <- parseCommandLine
  -- trace (show options) $ return ()
  case options of
   Sheet {..} -> do
     r <- ssheetReadFile ssheetDefaultOptions sheetToRead
     case r of
      Left err -> do
        showError err
        exitFailure
      Right ssheet -> do
        print $ ssheetNumberOfRows (head ssheet)
        print $ ssheetNumberOfColumns (head ssheet)
        T.putStrLn (ssheetCsv (head ssheet))
        ssheetJsonPrettyPutLn ssheet
        exitSuccess

----------------------------------------------------------------------

showError :: [T.Text] -> IO ()
showError err = putStrLn $ "Errors: " ++ show (length err) ++ "\n" ++ T.unpack (T.unlines err)

----------------------------------------------------------------------

data CommandLineOptions = Sheet { sheetToRead :: FilePath
                                }
                        deriving (Show, CmdArgs.Data, CmdArgs.Typeable)

clomSheet = Sheet { sheetToRead = CmdArgs.def &= CmdArgs.argPos 0 &= CmdArgs.typ "XLS/XLSX/CSV"
                  }

parseCommandLine = CmdArgs.cmdArgs $ CmdArgs.modes [clomSheet]
                   &= CmdArgs.program "ssheet"
                   &= CmdArgs.summary "ssheet: xlsx, xls, csv reader"

----------------------------------------------------------------------
