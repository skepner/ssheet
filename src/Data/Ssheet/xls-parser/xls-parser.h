/* Eugene Skepner 2012 */
/*  Antigenic Cartography 2012 */

/*======================================================================*/

#include <sys/types.h>

/*======================================================================*/

struct ExcelData;

struct ExcelData* excel_open_file(char* filename, int strip_strings);
struct ExcelData* excel_open(char* data, off_t size, int strip_strings);
void excel_close(struct ExcelData* data);

int excel_number_of_sheets(struct ExcelData* data);
const char* excel_sheet_name(struct ExcelData* data, int sheet_no);
int excel_number_of_rows(struct ExcelData* data, int sheet_no);
int excel_number_of_columns(struct ExcelData* data, int sheet_no, int row_no);

/*
  returns "" fo empty and non-existent cell
  returns ":d:YYYY-MM-DD" for date cell
  returns ":i:NUMBER" for integer cell
  returns ":f:NUMBER" for float cell (may look like integer)
 */
const char* excel_cell_as_text(struct ExcelData* data, int sheet_no, int row_no, int col_no);

/*======================================================================*/

/*
 * Local Variables:
 * eval: (if (fboundp 'eu-rename-buffer) (eu-rename-buffer))
 * End:
 */
