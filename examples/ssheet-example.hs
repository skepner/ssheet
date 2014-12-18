{-# LANGUAGE OverloadedStrings #-}

import Debug.Trace

import System.Exit
import System.Environment
import qualified Data.Text as T
-- import qualified Data.Text.IO as T

import Data.Ssheet

----------------------------------------------------------------------

main :: IO ()
main =
  getArgs >>= process
  where
    process :: [String] -> IO ()
    process args =
      if length args > 0
      then mapM_ processSheet args
      else error "No sheets provided in the command line"
    processSheet :: FilePath -> IO ()
    processSheet filename = do
      putStrLn $ "Reading " ++ filename
      ss <- ssheetReadFile ssheetDefaultOptions filename
      case ss of
       Left err -> do
         error $ "Errors: " ++ show (length err) ++ "\n" ++ T.unpack (T.unlines err)
       Right ssheet -> do
         putStrLn $ "Rows: " ++ show (ssheetNumberOfRows (head ssheet))
         putStrLn $ "Cols: " ++ show (ssheetNumberOfColumns (head ssheet))
         putStrLn $ "0:0: " ++ show (ssheetCell (head ssheet) 0 0)
         -- T.putStrLn (ssheetCsv (head ssheet))
         -- ssheetJsonPrettyPutLn ssheet

----------------------------------------------------------------------
