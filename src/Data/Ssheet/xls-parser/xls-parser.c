/*
  Copyright 1998-2003 Victor Wagner
  Copyright 2003 Alex Ott
  Eugene Skepner 2012
  Antigenic Cartography 2012

  This file is released under the GPL.  Details can be
  found in the file COPYING accompanying this distribution.
*/

/*======================================================================*/

#include <stdio.h>
#include <stdlib.h>
#include <string.h>             /* memset, strlen */
#include <ctype.h>
#include <errno.h>
//#include <float.h>
#include <math.h>
#include <fcntl.h>
#include <unistd.h>

#include "xls-parser.h"

/*======================================================================*/

#define MAX_NUM_SHEETS 32

struct rowdescr
{
    int last, end;
    unsigned char **cells;
};

struct Sheet
{
    struct rowdescr* rowptr;
    int lastrow;
    char* name;
};

struct ExcelData
{
    struct Sheet sheets[MAX_NUM_SHEETS];
    int current;
};

/*======================================================================*/

typedef struct {
    unsigned char* data;
    off_t pos;
    off_t size;
} Source;

struct Ole;
struct OleEntry;

static struct Ole* ole_init(Source *f);
static struct OleEntry* ole_readdir(struct Ole* ole);
static int ole_open(struct OleEntry* e);
static int ole_is_book(struct OleEntry* e);
static int ole_close(struct OleEntry* e);
static void ole_finish(struct Ole* ole);

static struct ExcelData* do_table(struct OleEntry* input, int strip_strings);

static void free_sheet(struct Sheet* sheet);

static void ExcelSerialDateToDMY(int serial_date, int* day, int* month, int* year);

/*======================================================================*/

struct ExcelData* excel_open_file(char* filename, int strip_strings)
{
    int fd;
    off_t size;
    char* buffer;
    struct ExcelData* result;

    fd = open(filename, O_RDONLY);
    if (fd < 0)
        return NULL;
    size = lseek(fd, 0, SEEK_END);
    lseek(fd, 0, SEEK_SET);
    buffer = malloc(size);
    if (buffer == NULL) {
        close(fd);
        return NULL;
    }
    if (read(fd, buffer, size) != size) {
        close(fd);
        free(buffer);
        return NULL;
    }
    result = excel_open(buffer, size, strip_strings);
    return result;
}

/*======================================================================*/

struct ExcelData* excel_open(char* data, off_t size, int strip_strings)
{
    Source source;
    struct Ole* ole;
    struct OleEntry* ole_file;
    struct ExcelData* excel_data = NULL;

    source.data = (unsigned char*)data;
    source.size = size;
    source.pos = 0;

    if ((ole = ole_init(&source)) == NULL) {
        printf("ole_init failed\n");
        return NULL;
    }

    while ((ole_file = ole_readdir(ole)) != NULL) {
        if (ole_open(ole_file) >= 0) {
            if (ole_is_book(ole_file)) {
                excel_data = do_table(ole_file, strip_strings);
                ole_close(ole_file);
            }
        }
        if (excel_data)
            break;
    }
    ole_finish(ole);
    return excel_data;
}

/*======================================================================*/

void excel_close(struct ExcelData* data)
{
    if (data != NULL) {
        int i;
        for (i = 0; i < MAX_NUM_SHEETS; ++i) {
            free_sheet(&data->sheets[i]);
        }
    }
}

/*======================================================================*/

int excel_number_of_sheets(struct ExcelData* data)
{
    return data->current;
}

/*======================================================================*/

const char* excel_sheet_name(struct ExcelData* data, int sheet_no)
{
    const char* name = "?";
    if (sheet_no >= 0 && sheet_no < data->current) {
        struct Sheet* sheet = &data->sheets[sheet_no];
        name = sheet->name;
    }
    return name;
}

//======================================================================

int excel_number_of_rows(struct ExcelData* data, int sheet_no)
{
    int rows = 0;
    if (sheet_no >= 0 && sheet_no < data->current) {
        struct Sheet* sheet = &data->sheets[sheet_no];
        rows = sheet->lastrow;
    }
    return rows;
}

/*======================================================================*/

int excel_number_of_columns(struct ExcelData* data, int sheet_no, int row_no)
{
    int columns = 0;
    if (sheet_no < data->current) {
        const struct Sheet* sheet = &data->sheets[sheet_no];
        if (row_no >= 0 && row_no < sheet->lastrow) {
            const struct rowdescr* row = sheet->rowptr + row_no;
            if (row->cells) {
                int col;
                for (col = row->last; col >= 0; --col) {
                    if (row->cells[col])
                        break;
                }
                columns = col + 1;
            }
        }
    }
    return columns;
}

/*======================================================================*/

const char* excel_cell_as_text(struct ExcelData* data, int sheet_no, int row_no, int col_no)
{
    const char* cell = NULL;
    if (sheet_no < data->current) {
        const struct Sheet* sheet = &data->sheets[sheet_no];
        if (row_no >= 0 && row_no < sheet->lastrow) {
            const struct rowdescr* row = sheet->rowptr + row_no;
            if (row->cells && col_no >= 0 && col_no <= row->last) {
                cell = (char*)row->cells[col_no];
            }
        }
    }
      // printf("%d:%d %s\n", row_no, col_no, cell);
    return cell == NULL ? "" : cell;
}

/*======================================================================*/
/*======================================================================*/

#define BBD_BLOCK_SIZE     512
#define SBD_BLOCK_SIZE      64
#define PROP_BLOCK_SIZE    128
#define OLENAMELENGHT       32
#define MSAT_ORIG_SIZE     436

typedef enum {
    oleDir=1,
    oleStream=2,
    oleRootDir=5,
    oleUnknown=3
} oleType;

struct OleEntry;

typedef struct Ole {
    Source *file;
    off_t sectorSize;
    off_t shortSectorSize;
    off_t bbdNumBlocks;
    unsigned char* BBD;
    unsigned char* SBD;
    off_t sbdNumber;
    struct OleEntry* rootEntry;
    unsigned char* properties;
    off_t propCurNumber;
    off_t propNumber;
} Ole;

typedef struct OleEntry {
    Ole* ole;
    char name[OLENAMELENGHT+1];
    long int startBlock;
    long int curBlock;
    long int length;
    off_t ole_offset;
    long int file_offset;
    unsigned char *dirPos;
    oleType type;
    long int numOfBlocks;
    long int *blocks;           /**< array of blocks numbers */
    int isBigBlock;
} oleEntry;

static int ole_eof(oleEntry* e);
static size_t ole_read(void *ptr, size_t size, size_t nmemb, oleEntry* e);

/*======================================================================*/

static struct ExcelData* excel_data_init();
static void sheet_end(struct ExcelData* data);
static int sheet_empty(struct ExcelData* data);

static unsigned char **allocate(struct ExcelData* data, int row, int col, int size);

/*======================================================================*/

static inline unsigned int getshort(unsigned char *buffer,int offset) {
    return (unsigned short int)buffer[offset]|((unsigned short int)buffer[offset+1]<<8);
}

static inline long int getlong(unsigned char *buffer,int offset) {
    return (long)buffer[offset]|((long)buffer[offset+1]<<8L)
        |((long)buffer[offset+2]<<16L)|((long)buffer[offset+3]<<24L);
}

static inline unsigned long int getulong(unsigned char *buffer,int offset) {
    return (unsigned long)buffer[offset]|((unsigned long)buffer[offset+1]<<8L)
        |((unsigned long)buffer[offset+2]<<16L)|((unsigned long)buffer[offset+3]<<24L);
}

/*======================================================================*/

#define min(a,b) ((a) < (b) ? (a) : (b))

static const char ole_sign[] = {0xD0,0xCF,0x11,0xE0,0xA1,0xB1,0x1A,0xE1,0};

/* static int source_eof(Source *f) { */
/*     return f->pos >= f->size; */
/* }     */

static void source_seek(Source* s, off_t pos)
{
    s->pos = pos;
}

static off_t source_read(Source* s, unsigned char* buffer, off_t size)
{
    memcpy(buffer, s->data + s->pos, size);
    s->pos += size;
    return size;
}

static off_t source_read_at(Source* s, off_t pos, unsigned char* buffer, off_t size)
{
    if (pos > s->size)
        pos = s->size;
    if ((pos + size) > s->size)
        size = s->size - pos;
    if (size) {
        memcpy(buffer, s->data + pos, size);
        s->pos = pos + size;
    }
    return size;
}

/*======================================================================*/

static Ole* ole_init(Source* f)
{
    Ole* ole = (Ole*)malloc(sizeof(Ole));
    unsigned char* oleBuf;
    unsigned char *tmpBuf;
    int i;
    long int sbdMaxLen, sbdCurrent, sbdStart, sbdLen, propMaxLen, propCurrent, mblock, msat_size, propLen, propStart;
    struct OleEntry* tEntry;

    memset(ole, 0, sizeof(Ole));
    ole->file = f;

    oleBuf = ole->file->data;
    if (strncmp((char*)oleBuf, ole_sign, 8) != 0) {
        ole_finish(ole);
        return NULL;
    }
    ole->sectorSize = 1 << getshort(oleBuf, 0x1e);
    ole->shortSectorSize = 1 << getshort(oleBuf, 0x20);
    mblock = getlong(oleBuf, 0x44);
    msat_size = getlong(oleBuf, 0x48);

     /* Read BBD into memory */
    ole->bbdNumBlocks = getulong(oleBuf, 0x2c);
    ole->BBD = (unsigned char*)malloc(ole->bbdNumBlocks * ole->sectorSize);

    tmpBuf = (unsigned char*)malloc(MSAT_ORIG_SIZE);
    memcpy(tmpBuf, oleBuf + 0x4c, MSAT_ORIG_SIZE);

    i=0;
    while((mblock >= 0) && (i < msat_size)) {
        tmpBuf = (unsigned char*)realloc(tmpBuf, ole->sectorSize*(i+1)+MSAT_ORIG_SIZE);
        source_read_at(ole->file, 512+mblock*ole->sectorSize, tmpBuf+MSAT_ORIG_SIZE+(ole->sectorSize-4)*i, ole->sectorSize);
        i++;
        mblock = getlong(tmpBuf, MSAT_ORIG_SIZE+(ole->sectorSize-4)*i);
    }

    for(i=0; i< ole->bbdNumBlocks; i++) {
        long int bbdSector=getlong(tmpBuf,4*i);

        if (bbdSector >= ole->file->size / ole->sectorSize || bbdSector < 0) {
            ole_finish(ole);
            return NULL;        /* Bad BBD entry */
        }
        source_read_at(ole->file, 512+bbdSector*ole->sectorSize, ole->BBD+i*ole->sectorSize, ole->sectorSize);
    }
    free(tmpBuf);

       /* Read SBD into memory */
    sbdLen = 0;
    sbdMaxLen = 10;
    sbdCurrent = sbdStart = getlong(oleBuf,0x3c);
    if (sbdStart > 0 && sbdStart < ole->file->size) {
        ole->SBD = (unsigned char*)malloc(ole->sectorSize*sbdMaxLen);
        while(1) {
            source_read_at(ole->file, 512 + sbdCurrent * ole->sectorSize, ole->SBD + sbdLen * ole->sectorSize, ole->sectorSize);
            sbdLen++;
            if (sbdLen >= sbdMaxLen) {
                sbdMaxLen += 5;
                ole->SBD = (unsigned char*)realloc(ole->SBD, ole->sectorSize*sbdMaxLen);
            }
            sbdCurrent = getlong(ole->BBD, sbdCurrent * 4);
            if (sbdCurrent < 0 || sbdCurrent >= ole->file->size / ole->sectorSize)
                break;
        }
        ole->sbdNumber = (sbdLen * ole->sectorSize) / ole->shortSectorSize;
    } else {
        ole->SBD=NULL;
    }
      /* Read property catalog into memory */
    propLen = 0;
    propMaxLen = 5;
    propCurrent = propStart = getlong(oleBuf,0x30);
    if (propStart >= 0) {
        ole->properties = (unsigned char*)malloc(propMaxLen*ole->sectorSize);
        while(1) {
            source_read_at(ole->file, 512+propCurrent*ole->sectorSize, ole->properties+propLen*ole->sectorSize, ole->sectorSize);
            propLen++;
            if (propLen >= propMaxLen) {
                propMaxLen+=5;
                ole->properties = (unsigned char*)realloc(ole->properties, propMaxLen*ole->sectorSize);
            }

            propCurrent = getlong(ole->BBD, propCurrent*4);
            if(propCurrent < 0 ||
               propCurrent >= ole->file->size / ole->sectorSize ) {
                break;
            }
        }
        ole->propNumber = (propLen*ole->sectorSize)/PROP_BLOCK_SIZE;
        ole->propCurNumber = 0;
    } else {
        ole_finish(ole);
        return NULL;
    }

      /* Find Root Entry */
    while ((tEntry = ole_readdir(ole)) != NULL) {
        if (tEntry->type == oleRootDir ) {
            ole->rootEntry=tEntry;
            break;
        }
        ole_close(tEntry);
    }
    ole->propCurNumber = 0;
    source_seek(ole->file, 0);
    if (!ole->rootEntry) {
        ole_finish(ole);
        return NULL;            /* Broken OLE structure. Cannot find root entry in this file */
    }
    return ole;
}

/*======================================================================*/

static oleEntry* ole_readdir(Ole* ole)
{
    unsigned i;
    unsigned char *oleBuf;
    oleEntry* e;
    long int chainMaxLen, chainCurrent;

    if ( ole->properties == NULL || ole->propCurNumber >= ole->propNumber || ole->file == NULL )
        return NULL;
    oleBuf = ole->properties + ole->propCurNumber * PROP_BLOCK_SIZE;
    if( !(oleBuf[0x42] == 1 || oleBuf[0x42] == 2 || oleBuf[0x42] == 3 || oleBuf[0x42] == 5))
        return NULL;
    e = (oleEntry*)malloc(sizeof(oleEntry));
    e->ole = ole;
    e->dirPos=oleBuf;
    e->type = (oleType)((unsigned char)oleBuf[0x42]);
    e->startBlock=getlong(oleBuf,0x74);
    e->blocks=NULL;

    for (i=0 ; i < getshort(oleBuf,0x40) /2; i++)
        e->name[i]=(char)oleBuf[i*2];
    e->name[i]='\0';            // e->name: Root Entry, Workbook
    ole->propCurNumber++;
    e->length=getulong(oleBuf,0x78);
        /* Read sector chain for object */
    chainMaxLen = 25;
    e->numOfBlocks = 0;
    chainCurrent = e->startBlock;
    e->isBigBlock = (e->length >= 0x1000) || !strcmp(e->name, "Root Entry");
    if (e->startBlock >= 0 &&
        /* e->length >= 0 && */
        (e->startBlock <= ole->file->size / (e->isBigBlock ? ole->sectorSize : ole->shortSectorSize))) {
        e->blocks = (long*)malloc(chainMaxLen*sizeof(long));
        while(1) {
            e->blocks[e->numOfBlocks++] = chainCurrent;
            if (e->numOfBlocks >= chainMaxLen) {
                chainMaxLen+=25;
                e->blocks = (long*)realloc(e->blocks, chainMaxLen*sizeof(long));
            }
            if ( e->isBigBlock ) {
                chainCurrent = getlong(ole->BBD, chainCurrent*4);
            } else if ( ole->SBD != NULL ) {
                chainCurrent = getlong(ole->SBD, chainCurrent*4);
            } else {
                chainCurrent=-1;
            }
            if(chainCurrent <= 0 ||
               chainCurrent >= ( e->isBigBlock ? ((ole->bbdNumBlocks * ole->sectorSize)/4) : ((ole->sbdNumber*ole->shortSectorSize)/4) ) ||
               (e->numOfBlocks >
                (long)e->length/(e->isBigBlock ? ole->sectorSize : ole->shortSectorSize))) {
                break;
            }
        }
    }

    if(e->length > (long)(e->isBigBlock ? ole->sectorSize : ole->shortSectorSize)*e->numOfBlocks)
        e->length = (e->isBigBlock ? ole->sectorSize : ole->shortSectorSize)*e->numOfBlocks;

    return e;
}

/*======================================================================*/

static int ole_open(oleEntry* e) {
    if ( e->type != oleStream)
        return -2;

    e->ole_offset=0;
    e->file_offset = e->ole->file->pos;
    return 0;
}

/*======================================================================*/

static int ole_is_book(struct OleEntry* ole_file)
{
    return strcasecmp(ole_file->name, "Workbook") == 0 || strcasecmp(ole_file->name, "Book") == 0;
}

/*======================================================================*/

static off_t calcFileBlockOffset(oleEntry *e, long int blk)
{
    off_t res;
    if (e->isBigBlock) {
        res= 512 + e->blocks[blk] * e->ole->sectorSize;
    }
    else {
        const off_t sbdPerSector = e->ole->sectorSize / e->ole->shortSectorSize;
        const off_t sbdSecNum = e->blocks[blk] / sbdPerSector;
        const off_t sbdSecMod = e->blocks[blk] % sbdPerSector;
        res = 512 + e->ole->rootEntry->blocks[sbdSecNum] * e->ole->sectorSize + sbdSecMod * e->ole->shortSectorSize;
    }
    return res;
}

/*======================================================================*/

static size_t ole_read(void *ptr, size_t size, size_t nmemb, oleEntry* e)
{
    long int llen = size*nmemb, rread=0, i;
    long int blockNumber, modBlock, toReadBlocks, toReadBytes, bytesInBlock;
    long int ssize;             /**< Size of block */
    long int newoffset;
    unsigned char *cptr = (unsigned char*)ptr;
    if( e->ole_offset+llen > e->length )
        llen= e->length - e->ole_offset;

    ssize = (e->isBigBlock ? e->ole->sectorSize : e->ole->shortSectorSize);
    blockNumber = e->ole_offset/ssize;
    if ( blockNumber >= e->numOfBlocks || llen <=0 )
        return 0;

    modBlock=e->ole_offset%ssize;
    bytesInBlock = ssize - modBlock;
    if(bytesInBlock < llen) {
        toReadBlocks = (llen-bytesInBlock)/ssize;
        toReadBytes = (llen-bytesInBlock)%ssize;
    } else {
        toReadBlocks = toReadBytes = 0;
    }
    newoffset = calcFileBlockOffset(e,blockNumber)+modBlock;
    if (e->file_offset != newoffset) {
        source_seek(e->ole->file, e->file_offset=newoffset);
    }
    rread = source_read(e->ole->file, (unsigned char*)ptr, min(llen, bytesInBlock));
    e->file_offset += rread;
    for(i=0; i<toReadBlocks; i++) {
        int readbytes;
        blockNumber++;
        newoffset = calcFileBlockOffset(e,blockNumber);
        readbytes = source_read_at(e->ole->file, e->file_offset=newoffset, cptr+rread, min(llen-rread, ssize));
        rread += readbytes;
        e->file_offset +=readbytes;
    }
    if(toReadBytes > 0) {
        int readbytes;
        blockNumber++;
        newoffset = calcFileBlockOffset(e,blockNumber);
        readbytes = source_read_at(e->ole->file, e->file_offset=newoffset, cptr+rread, toReadBytes);
        rread +=readbytes;
        e->file_offset +=readbytes;
    }
    e->ole_offset += rread;
    return rread;
}

/*======================================================================*/

static int ole_eof(oleEntry* e)
{
    return e->ole_offset >= e->length;
}

/*======================================================================*/

static void ole_finish(Ole* ole)
{
    if (ole->BBD != NULL)
        free(ole->BBD);
    if (ole->SBD != NULL)
        free(ole->SBD);
    if (ole->properties != NULL)
        free(ole->properties);
    if (ole->rootEntry != NULL)
        ole_close(ole->rootEntry);
}

/*======================================================================*/

static int ole_close(oleEntry* e)
{
    if(e == NULL)
        return -1;
    if (e->blocks != NULL)
        free(e->blocks);
    free(e);
    return 0;
}

/*======================================================================*/
/*======================================================================*/

#define MAX_NUM_DATE_FORMATS 16

typedef struct SstData
{
    unsigned char **sst;
    int sstsize; /*Number of strings in SST*/
    unsigned char *sstBuffer; /*Unparsed sst to accumulate all its parts*/
    int sstBytes; /*Size of SST Data, already accumulated in the buffer */
    int prev_rectype;
    unsigned char **saved_reference;

      /* format table to detect dates */
    size_t formatTableIndex;
    size_t formatTableSize;
    short* formatTable;
    short dateFormats[MAX_NUM_DATE_FORMATS]; /* lists codes stored in formatTable that correspond tp date formats */
    size_t numDateFormats;

} SstData;

static void sstData_init(SstData* sstData);
static void sstData_done(SstData* sstData);
static void parse_sst(SstData* sstData);
/* returns 0 on success, -1 on error */
static int process_item(SstData* sstData, struct ExcelData* data, int rectype, int reclen, unsigned char *rec, int strip_strings);
static void format_double(SstData* sstData, unsigned char* buffer, unsigned char *rec, int offset, int format_code);
static void format_rk(SstData* sstData, unsigned char* buffer, unsigned char *rec, int format_code);
static int to_utf8(char* utfbuffer, unsigned int uc);
static unsigned char *copy_unicode_string(unsigned char **src);
static void postprocess_field(unsigned char* src, int strip_strings);

/*======================================================================*/

#define MAX_MS_RECSIZE 18000

#define MS_UNIX_DATE_DIFF (70*365.2422+1)

#define DATE_FORMAT              14

#define MS1904           0x22
#define ADDIN                0x87
#define ADDMENU          0xC2
#define ARRAY                0x221
#define AUTOFILTER           0x9E
#define AUTOFILTERINFO           0x9D
#define BACKUP                   0x40
#define BLANK                0x201
#define BOF                  0x809
#define BOOKBOOL         0xDA
#define BOOLERR          0x205
#define BOTTOMMARGIN         0x29
#define BOUNDSHEET           0x85
#define CALCCOUNT        0x0C
#define CALCMODE         0x0D
#define CODEPAGE         0x42
#define COLINFO          0x7D
#define CONTINUE         0x3C
#define COORDLIST        0xA9
#define COUNTRY          0x8C
#define CRN                  0x5A
#define DBCELL                   0xD7
#define DCON                 0x50
#define DCONNAME         0x52
#define DCONREF          0x51
#define DEFAULTROWHEIGHT     0x225
#define DEFCOLWIDTH          0x55
#define DELMENU          0xC3
#define DELTA                0x10
#define DIMENSIONS           0x200
#define DOCROUTE         0xB8
#define EDG                  0x88
#define MSEOF                0x0A
#define EXTERNCOUNT          0x16
#define EXTERNNAME           0x223
#define EXTERNSHEET          0x17
#define FILEPASS         0x2F
#define FILESHARING          0x5B
#define FILESHARING2         0x1A5
#define FILTERMODE           0x9B
#define FNGROUPCOUNT         0x9C
#define FNGROUPNAME          0x9A
#define FONT                 0x231
#define FONT2                0x31
#define FOOTER                   0x15
#define FORMAT                   0x41E
#define FORMULA_RELATED          0x4BC
#define DOUBLE_STREAM_FILE   0x161
/*#define FORMULA        0x406  Microsoft docs wrong?*/
#define FORMULA          0x06
#define GCW                  0xAB
#define GRIDSET          0x82
#define PROT4REVPASS             0x1BC
#define GUTS                 0x80
#define HCENTER          0x83
#define HEADER                   0x14
#define HIDEOBJ          0x8D
#define HORIZONTALPAGEBREAKS     0x1B
#define IMDATA                   0x7F
#define INDEX                0x20B
#define INTERFACEEND         0xE2
#define INTERFACEHDR         0xE1
#define ITERATION        0x11
#define LABEL                0x204
#define LEFTMARGIN           0x26
#define LHNGRAPH         0x95
#define LHRECORD         0x94
#define LPR                  0x98
#define MMS                  0xC1
#define MULBLANK         0xBE
#define MULRK                0xBD
#define NAME                 0x218
#define NOTE                 0x1C
#define NUMBER                   0x203
#define OBJ                  0x5D
#define OBJPROTECT           0x63
#define OBPROJ                   0xD3
#define OLESIZE          0xDE
#define PALETTE          0x92
#define PANE                 0x41
#define PASSWORD         0x13
#define PLS                  0x4D
#define PRECISION        0x0E
#define PRINTGRIDLINES           0x2B
#define PRINTHEADERS         0x2A
#define PROTECT          0x12
#define PUB                  0x89
#define RECIPNAME        0xB9
#define REFMODE          0x0F
#define RIGHTMARGIN          0x27
#define RK                   0x27E
#define ROW                  0x208
#define RSTRING          0xD6
#define SAVERECALC           0x5F
#define SCENARIO         0xAF
#define SCENMAN          0xAE
#define SCENPROTECT          0xDD
#define SCL                  0xA0
#define SELECTION        0x1D
#define SETUP                0xA1
#define SHRFMLA          0xBC
#define SORT                 0x90
#define SOUND                0x96
#define STANDARDWIDTH        0x99
#define STRING                   0x207
#define STYLE                0x293
#define SUB                  0x91
#define SXDI                 0xC5
#define SXEXT                0xDC
#define SXIDSTM          0xD5
#define SXIVD                0xB4
#define SXLI                 0xB5
#define SXPI                 0xB6
#define SXSTRING         0xCD
#define SXTBL                0xD0
#define SXTBPG                   0xD2
#define SXTBRGIITM           0xD1
#define SXVD                 0xB1
#define SXVI                 0xB2
#define SXVIEW                   0xB0
#define SXVS                 0xE3
#define TABID                0x13D
#define TABIDCONF        0xEA
#define TABLE                0x236
#define TEMPLATE         0x60
#define TOPMARGIN        0x28
#define UDDESC                   0xDF
#define UNCALCED         0x5E
#define VCENTER          0x84
#define VERTICALPAGEBREAKS       0x1A
#define WINDOW1          0x3D
#define WINDOW2          0x23E
#define WINDOWPROTECT        0x19
#define WRITEACCESS          0x5C
#define WRITEPROT        0x86
#define WSBOOL                   0x81
#define XCT                  0x59
#define XF                   0xE0
#define SST              0xFC
#define CONSTANT_STRING              0xFD
#define REFRESHALL       0x1B7
#define USESELFS         0x160
#define EXTSST               0xFF
/* Vitus additions */
#define INTEGER_CELL     0x202

/*======================================================================*/

static struct ExcelData* do_table(oleEntry *input, int strip_strings)
{
    SstData sstData;
    unsigned char rec[MAX_MS_RECSIZE];
    long rectype;
    long reclen, build_year=0, build_rel=0, offset=0;
    int eof_flag=0;
    int itemsread=1;
    struct ExcelData* data = excel_data_init();

    sstData_init(&sstData);

    while (itemsread) {
        int biff_version;
        ole_read(rec, 2, 1, input);
        biff_version = getshort(rec, 0);
        ole_read(rec, 2, 1, input);
        reclen = getshort(rec, 0);
        if (biff_version == 0x0809 || biff_version == 0x0409 || biff_version == 0x0209 || biff_version == 0x0009 ) {
            if (reclen == 8 || reclen == 16) {
                if (biff_version == 0x0809 ) {
                    itemsread=ole_read(rec,4,1,input);
                    build_year=getshort((rec+2),0);
                    build_rel=getshort(rec,0);
                    if(build_year > 5 ) {
                        itemsread=ole_read(rec,8,1,input);
                        /* biff_version=8; */
                        offset=12;
                    }
                    else {
                        /* biff_version=7; */
                        offset=4;
                    }
                } else if (biff_version == 0x0209 ) {
                    /* biff_version=3; */
                    offset=2;
                } else if (biff_version == 0x0409 ) {
                    offset=2;
                    /* biff_version=4; */
                } else {
                    /* biff_version=2; */
                }
                itemsread= ole_read(rec,reclen - offset, 1,input);
                break;
            } else {
                excel_close(data);
                sstData_done(&sstData);
                return NULL;    /* Invalid BOF record */
            }
        } else {
            itemsread = ole_read(rec, 126, 1, input);
        }
    }
    if (ole_eof(input)) {
        excel_close(data);
        sstData_done(&sstData);
        return NULL;    /* No BOF record found */
    }
    while (itemsread) {
        unsigned char buffer[2];
        rectype = 0;
        itemsread = ole_read(buffer, 2, 1, input);
        if (ole_eof(input)) {
            if (process_item(&sstData, data, MSEOF, 0, NULL, strip_strings)) {
                excel_close(data);
                sstData_done(&sstData);
                return NULL;    /* Error */
            }
            sstData_done(&sstData);
            return data;
        }

        rectype = getshort(buffer, 0);
        if (itemsread == 0)
            break;
        reclen=0;

        itemsread = ole_read(buffer, 2, 1, input);
        reclen = getshort(buffer,0);
        if (reclen && reclen < MAX_MS_RECSIZE && reclen > 0) {
            itemsread = ole_read(rec, 1, reclen, input);
            rec[reclen] = '\0';
        }
        if (eof_flag && rectype != BOF) {
            break;
        }
        /* printf("item offset_end:%d rectype:0x%X reclen:%d\n", input->ole->file->pos, rectype, reclen); */
        if (process_item(&sstData, data, rectype, reclen, rec, strip_strings)) {
            excel_close(data);
            sstData_done(&sstData);
            return NULL;    /* error */
        }
        eof_flag = rectype == MSEOF;
    }
    sstData_done(&sstData);
    return data;
}

/*======================================================================*/

static void sstData_init(SstData* sstData)
{
    sstData->sst = NULL;
    sstData->sstsize = 0; /*Number of strings in SST*/
    sstData->sstBuffer = NULL; /*Unparsed sst to accumulate all its parts*/
    sstData->sstBytes = 0; /*Size of SST Data, already accumulated in the buffer */
    sstData->prev_rectype = 0;
    sstData->saved_reference = NULL; /* not owned, do not free */

    sstData->formatTableIndex = 0;
    sstData->formatTableSize = 0;
    sstData->formatTable = NULL;

      /* memset(sstData->dateFormats, 0, MAX_NUM_DATE_FORMATS * sizeof(*sstData->dateFormats)); */
    sstData->numDateFormats = 0;
}

/*======================================================================*/

static void sstData_done(SstData* sstData)
{
    if (sstData->sst)
        free(sstData->sst);
    if (sstData->sstBuffer)
        free(sstData->sstBuffer);
    if (sstData->formatTable)
        free(sstData->formatTable);
}

/*======================================================================*/

/* returns 0 on success, -1 on error */
static int process_item(SstData* sstData, struct ExcelData* data, int rectype, int reclen, unsigned char *rec, int strip_strings)
{
    if (rectype != CONTINUE && sstData->prev_rectype == SST) {
        parse_sst(sstData);
    }
    switch (rectype) {
      case FILEPASS: {
          return -1;            /* File encrypted */
      }

            /* case WRITEPROT: { */
            /*     fprintf(stderr,"File is write protected\n"); */
            /*     break; */
            /* } */

            /* case 0x42: { */
            /*     /\* if (source_charset) *\/ */
            /*     /\*     break; *\/ */
            /*     int codepage=getshort(rec,0); */
            /*     fprintf(stderr,"CODEPAGE %d\n",codepage); */
            /*     /\* if (codepage!=1200) { *\/ */
            /*     /\*  const char *cp = charset_from_codepage(codepage); *\/ */
            /*     /\*  source_charset=read_charset(cp); *\/ */
            /*     /\* } *\/ */
            /*     break; */
            /* } */

            /* case FORMAT: { */
            /*     int format_code; */
            /*     format_code = getshort(rec,0); */
            /*     break; */
            /* } */

      case SST: {
            /* Just copy SST into buffer, and wait until we get
             * all CONTINUE records
             */
            /* If exists first SST entry, then just drop it and start new*/
          if (sstData->sstBuffer != NULL)
              free(sstData->sstBuffer);
          if (sstData->sst != NULL)
              free(sstData->sst);

          sstData->sstBuffer = (unsigned char*)malloc(reclen);
          sstData->sstBytes = reclen;
          memcpy(sstData->sstBuffer,rec,reclen);
          break;
      }
      case CONTINUE: {
          if (sstData->prev_rectype != SST) {
              return 0; /* to avoid changing of sstData->prev_rectype;*/
          }
          sstData->sstBuffer = (unsigned char*)realloc(sstData->sstBuffer,sstData->sstBytes+reclen);
          memcpy(sstData->sstBuffer+sstData->sstBytes,rec,reclen);
          sstData->sstBytes+=reclen;
          return 0;
      }
      case LABEL: {
          int row,col;
          unsigned char **pcell;
          unsigned char *src=(unsigned char *)rec+6;

          sstData->saved_reference=NULL;
          row = getshort(rec,0);
          col = getshort(rec,2);
          pcell=allocate(data, row, col, 0);
          *pcell=copy_unicode_string(&src);
          postprocess_field(*pcell, strip_strings);
          break;
      }

      /* case BLANK: { */
      /*     unsigned char **pcell = allocate(data, getshort(rec,0), getshort(rec,2)); */
      /*     *pcell=NULL; */
      /*     break; */
      /* } */
      /* case MULBLANK: { */
      /*     unsigned char **pcell = allocate(data, getshort(rec,0), getshort(rec,reclen-2)); */
      /*     *pcell=NULL; */
      /*     break; */
      /* } */

      case CONSTANT_STRING: {
          int string_no = getshort(rec,6);
          if (!sstData->sst) {
              return -1;        /* CONSTANT_STRING before SST parsed */
          }

          sstData->saved_reference=NULL;
          if (string_no >= sstData->sstsize || string_no < 0 ) {
              return -1;        /* string index out of boundary */
          }
          else if (sstData->sst[string_no] != NULL) {
              const int len = strlen((char*)sstData->sst[string_no]);
              unsigned char **pcell = allocate(data, getshort(rec, 0), getshort(rec, 2), len + 1);
                // printf("CONSTANT_STRING string_no:%d %d:%d size=%d\n", string_no, getshort(rec, 0), getshort(rec, 2), len + 1);
              strcpy((char*)*pcell, (char*)sstData->sst[string_no]);
              postprocess_field(*pcell, strip_strings);
          }
          else {
              unsigned char **pcell = allocate(data, getshort(rec,0), getshort(rec,2), 1);
              **pcell = 0;
          }
          break;
      }
      case 0x03:
      case 0x103:
      case 0x303:
      case NUMBER: {
          unsigned char **pcell = allocate(data, getshort(rec,0), getshort(rec,2), 32);
          sstData->saved_reference = NULL;
          format_double(sstData, *pcell, rec, 6, getshort(rec,4));
          break;
      }
      case INTEGER_CELL: {
          unsigned char **pcell = allocate(data, getshort(rec,0), getshort(rec,2), 32);
          sprintf((char*)*pcell, ":i:%i", getshort(rec,7));
          break;
      }
      case RK: {
          unsigned char **pcell = allocate(data, getshort(rec,0), getshort(rec,2), 32);
          format_rk(sstData, *pcell, rec + 6, getshort(rec, 4));
          sstData->saved_reference=NULL;
          break;
      }
      case MULRK: {
          int offset;
          int row = getshort(rec,0);
          int col;
          int endcol = getshort(rec,reclen-2);
          sstData->saved_reference=NULL;

          for (offset = 4, col = getshort(rec,2); col <= endcol; offset += 6, col++) {
              format_rk(sstData, *(allocate(data, row, col, 32)), rec + offset + 2, getshort(rec, offset));
          }
          break;
      }
      case FORMULA: {
          sstData->saved_reference=NULL;
          if ((unsigned char)rec[12] == 0xFF && (unsigned char)rec[13] == 0xFF) {
                /* not a floating point value */
              if (rec[6]==1) {
                    /*boolean*/
                  unsigned char **pcell = allocate(data, getshort(rec,0), getshort(rec,2), 2);
                  (*pcell)[0] = '0' + rec[9];
                  (*pcell)[1] = 0;
              } else if (rec[6]==2) {
                    /*error*/
                  unsigned char **pcell = allocate(data, getshort(rec,0), getshort(rec,2), 6);
                  strcpy((char*)*pcell, "ERROR");
              } else if (rec[6]==0) {
                  sstData->saved_reference = allocate(data, getshort(rec,0), getshort(rec,2), 0);
              }
          } else {
              unsigned char **pcell = allocate(data, getshort(rec,0), getshort(rec,2), 32);
              format_double(sstData, *pcell, rec, 6, getshort(rec,4));
          }
          break;
      }
      case STRING: {
          unsigned char *src=(unsigned char *)rec;
          if (!sstData->saved_reference) {
              /* fprintf(stderr,"String record without preceeding string formula\n"); */
              break;
          }
          *sstData->saved_reference=copy_unicode_string(&src);
          postprocess_field(src, strip_strings);
          break;
      }
      case BOF: {
          if (!sheet_empty(data)) {
              /* fprintf(stderr,"BOF when current sheet is not flushed\n"); */
              sheet_end(data);
          }
          break;
      }

    case XF:
    case 0x43: /*from perl module Spreadsheet::ParseExecel */
        {
            /* we are interested only in format index here */
            if (sstData->formatTableIndex >= sstData->formatTableSize) {
                sstData->formatTableSize += 16;
                sstData->formatTable = (short*)realloc(sstData->formatTable, sstData->formatTableSize * sizeof(short));
            }
            sstData->formatTable[sstData->formatTableIndex++] = getshort(rec, 2);
            break;
        }

      case FORMAT: {
            /* need to detect which of custom formats (code >= 0xA4 in sstData->formatTable) are date */
          int len = getshort(rec, 2), pos = 5, date = 0;
          if (rec[4] & 0x08)
              pos += 2;
          if (rec[4] & 0x04)
              pos += 4;
            /* format desc stored in rec[pos:pos+len] */
            /* look for YY or yy in the format */
          if (rec[4] & 0x01) {   // utf_16_le, use odd chars
              len = len * 2 + pos;
              for (; pos < (len - 2); pos += 2) {
                  if ((rec[pos] == 'Y' || rec[pos] == 'y') && (rec[pos + 1] == 'Y' || rec[pos + 1] == 'y')) {
                      date = 1;
                      break;
                  }
              }
          }
          else {                  // latin-1
              len += pos;
              for (; pos < (len - 1); ++pos) {
                  if ((rec[pos] == 'Y' || rec[pos] == 'y') && (rec[pos + 1] == 'Y' || rec[pos + 1] == 'y')) {
                      date = 1;
                      break;
                  }
              }
          }
          if (date) {
              sstData->dateFormats[sstData->numDateFormats++] = getshort(rec, 0);
          }
            /* printf("FORMAT fmtkey:%X date:%d\n", getshort(rec, 0), date); */
          break;
      }

      case MSEOF: {
          if (sheet_empty(data))
              break;
          sheet_end(data);
          break;
      }
      case BOUNDSHEET: {
            //const int ucs = rec[7] & 0x01;
          const int name_size = rec[6];
          struct Sheet* sheet = &data->sheets[data->current];
          sheet->name = (char*)malloc(name_size + 1);
          memcpy(sheet->name, rec + 8, name_size);
          sheet->name[name_size] = 0;
          break;
      }
      case ROW:
          break;
      case INDEX:
          break;
      default:
          break;
    }
    sstData->prev_rectype = rectype;
    return 0;
}

/*======================================================================*/

static unsigned char *copy_unicode_string(unsigned char **src)
{
    int count = 0;
    int flags = 0;
    int start_offset=0;
    int to_skip = 0;
    int offset = 1;
    int charsize;
    unsigned char *dest;
    unsigned char *s;

    int i,l,len;

    flags = *((*src) + 1 + offset);
    if (! ( flags == 0 || flags == 1 || flags == 8 || flags == 9 ||
            flags == 4 || flags == 5 || flags == 0x0c || flags == 0x0d ) ) {
        count = **src;
        flags = *(*src + offset);
        offset --;
        flags = *(*src+1+offset);
        if (! ( flags == 0 || flags == 1 || flags == 8 || flags == 9 ||
                flags == 4 || flags == 5 || flags == 0x0c || flags == 0x0d ) ) {
              /*          fprintf(stderr,"Strange flags = %d, returning NULL\n", flags); */
            return NULL;
        }
    }
    else {
        count = getshort(*src, 0);
    }
    charsize = (flags & 0x01) ? 2 : 1;

    switch (flags & 12 ) {
      case 0x0c: /* Far East with RichText formating */
          to_skip = 4 * getshort((*src), 2 + offset) + getlong((*src), 4 + offset);
          start_offset = 2 + offset + 2 + 4;
          break;

      case 0x08: /* With RichText formating */
          to_skip=4*getshort(*src,2+offset);
          start_offset=2+offset+2;
          break;

      case 0x04: /* Far East */
          to_skip=getlong((*src), 2+offset);
          start_offset=2+offset+4;
          break;

      default:
          to_skip = 0;
          start_offset = 2 + offset;
    }

    dest = (unsigned char*)malloc(count + 1);
    *src += start_offset;
    len = count;
    *dest = 0;
    l = 0;
      // printf("copy_unicode_string start_offset=%d end=%d i_end=%d\n", start_offset, start_offset + len * charsize, len * charsize);
    for (s = *src, i = 0; i < count; ++i, s += charsize) {
        if ( (charsize == 1 && (*s == 1 || *s == 0)) /* Disabled by Eu on 2014-01-15 for CNIC H3 2013-10-24.xls parsing: || (charsize == 2 && (*s == 1 || *s == 0) && *(s+1) != 4) */) {
            charsize = (*s & 0x01) ? 2 : 1;
            if (charsize == 2)
                s -= 1;
            count++;
              // printf("copy_unicode_string new charsize=%d i=%d\n", charsize, i);
        }
        else {
            char c[4];
            int dl = to_utf8(c, charsize == 2 ? (unsigned short)getshort(s,0) : (unsigned short)(unsigned char)*s);
            while (l + dl >= len) {
                len += 16;
                dest = (unsigned char*)realloc(dest, len+1);
            }
            strcpy((char*)(dest + l), c);
            l += dl;
        }
    }
    *src=s+to_skip;
      // printf("copy_unicode_string dest=[%s] i_end=%d\n", dest, i);
    return dest;
}

/*======================================================================*/

static int to_utf8(char* utfbuffer, unsigned int uc)
{
    int count=0;
    if (uc< 0x80) {
        utfbuffer[0]=uc;
        count=1;
    }
    else  {
        if (uc < 0x800) {
            utfbuffer[count++]=0xC0 | (uc >> 6);
        } else {
            utfbuffer[count++]=0xE0 | (uc >>12);
            utfbuffer[count++]=0x80 | ((uc >>6) &0x3F);
        }
        utfbuffer[count++]=0x80 | (uc & 0x3F);
    }
    utfbuffer[count]=0;
    return count;
}

/*======================================================================*/

/*
 * Format code is index into format table (which is list of XF records
 * in the file
 * Second word of XF record is format type idnex
 * format index between 0x0E and 0x16 also between 0x2D and ox2F denotes
 * date if it is not used for explicitly stored formats.
 * BuiltInDateFormatIdx converts format index into index of explicit
 * built-in date formats sutable for strftime.
 */

/* Checks if format denoted by given code is date
 * Format code is index into format table (which is list of XF records
 * in the file
 * Second word of XF record is format type inex
 * format index between 0x0E and 0x16 also between 0x2D and ox2F denotes
 * date.
 * returns if it is date
 */

static int isDateFormat(SstData* sstData, size_t format_code)
{
    const int format = (format_code < sstData->formatTableIndex) ? sstData->formatTable[format_code] : 0;
    if ((format >= 0x0E && format <= 0x16) || (format >= 0x2D && format <= 0x2F))
        return 1;
    if (format >= 0xA4) {
        size_t i;
        for (i = 0; i < sstData->numDateFormats; ++i) {
            if (sstData->dateFormats[i] == format)
                return 1;
        }
    }
    return 0;
}

/*======================================================================*/

static void format_double_2(unsigned char* buffer, double value, int date)
{
    double tmp;
    if (date) {
          /* sprintf((char*)buffer, ":date:%.0f", value); */
        int day, month, year;
        ExcelSerialDateToDMY((int)value, &day, &month, &year);
        sprintf((char*)buffer, ":d:%04d-%02d-%02d", year, month, day);
    }
    else if (modf(value, &tmp) == 0.0)
        sprintf((char*)buffer, ":f:%.0f", value);
    else
        sprintf((char*)buffer, ":f:%g", value);
}

//======================================================================

static void format_double(SstData* sstData, unsigned char* buffer, unsigned char *rec,int offset, int format_code)
{
    union {
        char cc[8];
        double d;
    } dconv;
    unsigned char *d,*s;
    int i;
# ifdef WORDS_BIGENDIAN
    for(s=rec+offset+8,d=dconv.cc, i=0; i < 8; i++)
        *(d++) = *(--s);
# else
    for(s=rec+offset,d=(unsigned char*)dconv.cc, i=0; i < 8; i++)
        *(d++)=*(s++);
# endif
    format_double_2(buffer, dconv.d, isDateFormat(sstData, format_code));
}

/*======================================================================*/

static void format_rk(SstData* sstData, unsigned char* buffer, unsigned char *rec, int format_code)
{
    double value = 0.0;

    if (*(rec) & 0x02) {
        value = (double)(getlong(rec,0)>>2);
    }
    else {
        int i;
        union { char cc[8]; double d; } dconv;
        unsigned char *d, *s;
        for (i = 0; i < 8; i++)
            dconv.cc[i] = '\0';
# ifdef WORDS_BIGENDIAN
        for (s = rec+4, d = dconv.cc, i = 0; i < 4; i++)
            *(d++) = *(--s);
        dconv.cc[0] = dconv.cc[0] & 0xfc;
# else
        for(s = rec, d =(unsigned char*)dconv.cc+4, i=0; i < 4; i++)
            *(d++) = *(s++);
        dconv.cc[3] = dconv.cc[3] & 0xfc;
# endif
        value=dconv.d;
    }
    if (*(rec) & 0x01)
        value = value * 0.01;
    format_double_2(buffer, value, isDateFormat(sstData, format_code));
}

/*======================================================================*/

static void parse_sst(SstData* sstData)
{
      //(unsigned char *sstbuf,int bufsize) {
    int i; /* index into sst */
    unsigned char *curString; /* pointer into unparsed buffer*/
    unsigned char *barrier=(unsigned char *)sstData->sstBuffer + sstData->sstBytes; /*pointer to end of buffer*/
    unsigned char **parsedString;/*pointer into parsed array*/

    sstData->sstsize = getlong((sstData->sstBuffer+4),0);
    sstData->sst = (unsigned char**)malloc(sstData->sstsize*sizeof(char *));
    memset(sstData->sst, 0, sstData->sstsize*sizeof(char *));
    for (i = 0, parsedString = sstData->sst, curString = sstData->sstBuffer + 8; i < sstData->sstsize && curString < barrier; i++, parsedString++) {
          // printf("parse_sst %d\n", i);
        *parsedString = copy_unicode_string(&curString);
    }
}

/*======================================================================*/
/*======================================================================*/

/*======================================================================*/

static void clean_empty_rows_at_end(struct Sheet* sheet);

/*======================================================================*/

static struct ExcelData* excel_data_init()
{
    int i;
    struct ExcelData* data = (struct ExcelData*)malloc(sizeof(struct ExcelData));
    data->current = 0;
    for (i = 0; i < MAX_NUM_SHEETS; ++i) {
        data->sheets[i].rowptr = NULL;
        data->sheets[i].lastrow = 0;
        data->sheets[i].name = NULL;
    }
    return data;
}

/*======================================================================*/

static void sheet_end(struct ExcelData* data)
{
    clean_empty_rows_at_end(&data->sheets[data->current]);
    data->current++;
    if (data->current >= MAX_NUM_SHEETS)
        fprintf(stderr, "Too many sheets");
}

/*======================================================================*/

static int sheet_empty(struct ExcelData* data)
{
    return data->sheets[data->current].rowptr == NULL;
}

/*======================================================================*/

static void free_sheet(struct Sheet* sheet)
{
    if (sheet->rowptr != NULL) {
        int i,j;
        struct rowdescr *row;
        for (row = sheet->rowptr, i = 0; i < sheet->lastrow; i++, row++) {
            if (row->cells) {
                unsigned char **col;
                for (col = row->cells, j = 0; j< row->end; j++, col++) {
                    if (*col) {
                        free(*col);
                    }
                }
                free(row->cells);
            }
        }
        free(sheet->rowptr);
        sheet->rowptr = NULL;
        sheet->lastrow = 0;
        free(sheet->name);
    }
}

/*======================================================================*/

static void clean_empty_rows_at_end(struct Sheet* sheet)
{
    int i;
    for (i = sheet->lastrow - 1; i >= 0; --i) {
        struct rowdescr *row = &sheet->rowptr[i];
        if (row->cells) {
            unsigned char **col;
            int j, present = 0;
            for (col = row->cells, j = 0; j < row->end; j++, col++) {
                if (*col) {
                    present = 1;
                    break;
                }
            }
            if (present) {
                break;
            }
            else {
                for (col = row->cells, j = 0; j< row->end; j++, col++) {
                    if (*col) {
                        free(*col);
                    }
                }
                free(row->cells);
                row->cells = NULL;
            }
        }
    }
    sheet->lastrow = i + 1;
}

/*======================================================================*/

static unsigned char **allocate(struct ExcelData* data, int row, int col, int size)
{
    struct Sheet* sheet = &data->sheets[data->current];
    if (row >= sheet->lastrow) {
        unsigned int newrow;
        newrow = (row / 16 + 1) * 16;
        sheet->rowptr = (struct rowdescr*)realloc(sheet->rowptr, newrow * sizeof(struct rowdescr));
        memset(sheet->rowptr + sheet->lastrow, 0, (newrow - sheet->lastrow) * sizeof(struct rowdescr));
        sheet->lastrow = newrow;
    }
    if (col >= sheet->rowptr[row].end) {
        unsigned int newcol;
        newcol = (col / 16 + 1) * 16;
        sheet->rowptr[row].cells = (unsigned char**)realloc(sheet->rowptr[row].cells, newcol *sizeof(char *));
        memset(sheet->rowptr[row].cells + sheet->rowptr[row].end, 0, (newcol - sheet->rowptr[row].end) * sizeof(char *));
        sheet->rowptr[row].end = newcol;
    }
    if (col > sheet->rowptr[row].last)
        sheet->rowptr[row].last = col;
    if (size > 0)
        sheet->rowptr[row].cells[col] = (unsigned char*)malloc(size);
    return (sheet->rowptr[row].cells + col);
}

/*======================================================================*/

static void postprocess_field(unsigned char* src, int strip_strings)
{
    if (strip_strings) {
        unsigned char* begin = src;
        unsigned char* end = src + strlen((const char*)src);

        while (isspace(*begin))
            ++begin;
        while (begin < end && isspace(*(end-1)))
            --end;
        if (begin < end) {
            memmove(src, begin, end - begin);
            src[end - begin] = 0;
        }
        else {
            src[0] = 0;
        }
    }
}

/*======================================================================*/

/* http://www.codeproject.com/Articles/2750/Excel-serial-date-to-Day-Month-Year-and-vise-versa */
/* License: http://www.opensource.org/licenses/cddl1.php */
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

/*======================================================================*/

/* /\* */
/*  * prints out one value with quoting */
/*  * uses global variable quote_mode */
/*  *\/  */
/* /\* types of quoting *\/ */
/* #define QUOTE_NEVER 0 */
/* #define QUOTE_SPACES_ONLY 1 */
/* #define QUOTE_ALL_STRINGS 2 */
/* #define QUOTE_EVERYTHING 3 */

/* int quote_mode = QUOTE_ALL_STRINGS; */
/* char cell_separator = ','; */

/* static void print_value(unsigned char *value)  */
/* { */
/*     int i,len; */
/*     int quotes=0; */
/*     if (value != NULL) { */
/*         len=strlen((char *)value); */
/*     } else { */
/*         len = 0; */
/*     } */
/*     switch (quote_mode) { */
/*         case QUOTE_NEVER: */
/*             break; */
/*         case QUOTE_SPACES_ONLY:    */
/*             for (i=0;i<len;i++) { */
/*                 if (isspace(value[i]) || value[i]==cell_separator ||  */
/*                         value[i]=='"') { */
/*                     quotes=1; */
/*                     break; */
/*                 } */
/*             }     */
/*             break; */
/*         case QUOTE_ALL_STRINGS: */
/*             { char *endptr; */
/*                 strtod((char*)value,&endptr); */
/*               quotes=(*endptr != '0'); */
/*             break; */
/*             }   */
/*         case QUOTE_EVERYTHING: */
/*             quotes = 1; */
/*             break;      */
/*     }      */
/*     if (quotes) { */
/*         fputc('\"',stdout); */
/*         for (i=0;i<len;i++) { */
/*             if (value[i]=='\"') { */
/*                 fputc('\"',stdout); */
/*                 fputc('\"',stdout); */
/*             } else { */
/*                 fputc(value[i],stdout); */
/*             } */
/*         }    */
/*         fputc('\"',stdout); */
/*     } else { */
/*         fputs((char *)value,stdout); */
/*     } */
/* }     */

/* /\*======================================================================*\/ */

/* /\* */
/*  * Prints sheet to stdout. Uses global variable cell_separator */
/*  *\/  */
/* char *sheet_separator = "**** Sheet End ****\n"; */
/* static void print_sheet(struct ExcelData* data, int sheet_no) */
/* { */
/*     struct Sheet* sheet = &data->sheets[sheet_no]; */
/*     int i,j,printed=0; */
/*     struct rowdescr *row; */
/*     unsigned char **col; */
/*     int lastrow = sheet->lastrow - 1; */
/*     while (lastrow > 0 && !sheet->rowptr[lastrow].cells) */
/*         lastrow--; */
/*     for(i=0,row=sheet->rowptr;i<=lastrow;i++,row++) { */
/*         if (row->cells) { */
/*             for (j=0,col=row->cells;j<=row->last;j++,col++) { */
/*                 if (j){ */
/*                     fputc(cell_separator,stdout); */
/*                     printed=1; */
/*                 } */
/*                 if (*col) { */
/*                     print_value(*col); */
/*                     printed=1; */
/*                 } */
/*             } */
/*             if (printed) { */
/*                 fputc('\n',stdout); */
/*                 printed=0; */
/*             } */
/*         } */
/*     } */
/*     fputs(sheet_separator,stdout); */
/* } */

/*======================================================================*/

/*
 * Local Variables:
 * eval: (if (fboundp 'eu-rename-buffer) (eu-rename-buffer))
 * End:
 */
