
#include <R.h>
#include <Rinternals.h>
#include <stdlib.h> // for NULL
#include <R_ext/Rdynload.h>

/* 
generated with
tools::package_native_routine_registration_skeleton(".")
*/

/* .Call calls */
extern SEXP _openxlsx_build_cell_merges(SEXP);
extern SEXP _openxlsx_build_cell_types_integer(SEXP, SEXP);
extern SEXP _openxlsx_build_table_xml(SEXP, SEXP, SEXP, SEXP, SEXP, SEXP);
extern SEXP _openxlsx_buildCellList(SEXP, SEXP, SEXP);
extern SEXP _openxlsx_buildCellTypes(SEXP, SEXP);
extern SEXP _openxlsx_buildMatrixMixed(SEXP, SEXP, SEXP, SEXP, SEXP, SEXP, SEXP, SEXP);
extern SEXP _openxlsx_buildMatrixNumeric(SEXP, SEXP, SEXP, SEXP, SEXP, SEXP);
extern SEXP _openxlsx_calc_column_widths(SEXP, SEXP, SEXP, SEXP, SEXP, SEXP, SEXP);
extern SEXP _openxlsx_calc_number_rows(SEXP, SEXP);
extern SEXP _openxlsx_cell_ref_to_col(SEXP);
extern SEXP _openxlsx_convert_from_excel_ref(SEXP);
extern SEXP _openxlsx_convert_to_excel_ref(SEXP, SEXP);
extern SEXP _openxlsx_convert_to_excel_ref_expand(SEXP, SEXP, SEXP);
extern SEXP _openxlsx_cppReadFile(SEXP);
extern SEXP _openxlsx_get_extLst_Major(SEXP);
extern SEXP _openxlsx_get_letters();
extern SEXP _openxlsx_get_shared_strings(SEXP, SEXP);
extern SEXP _openxlsx_getAttr(SEXP, SEXP);
extern SEXP _openxlsx_getCellInfo(SEXP, SEXP, SEXP, SEXP, SEXP, SEXP);
extern SEXP _openxlsx_getChildlessNode(SEXP, SEXP);
extern SEXP _openxlsx_getChildlessNode_ss(SEXP, SEXP);
extern SEXP _openxlsx_getNodes(SEXP, SEXP);
extern SEXP _openxlsx_getOpenClosedNode(SEXP, SEXP, SEXP);
extern SEXP _openxlsx_int_2_cell_ref(SEXP);
extern SEXP _openxlsx_isInternalHyperlink(SEXP);
extern SEXP _openxlsx_loadworksheets(SEXP, SEXP, SEXP, SEXP);
extern SEXP _openxlsx_map_cell_types_to_char(SEXP);
extern SEXP _openxlsx_map_cell_types_to_integer(SEXP);
extern SEXP _openxlsx_matrixRowInds(SEXP);
extern SEXP _openxlsx_read_file_newline(SEXP);
extern SEXP _openxlsx_read_workbook(SEXP, SEXP, SEXP, SEXP, SEXP, SEXP, SEXP, SEXP, SEXP, SEXP, SEXP);
extern SEXP _openxlsx_write_worksheet_xml(SEXP, SEXP, SEXP, SEXP);
extern SEXP _openxlsx_write_worksheet_xml_2(SEXP, SEXP, SEXP, SEXP, SEXP);
extern SEXP _openxlsx_write_file(SEXP, SEXP, SEXP, SEXP);

static const R_CallMethodDef CallEntries[] = {
  {"_openxlsx_build_cell_merges",           (DL_FUNC) &_openxlsx_build_cell_merges,            1},
  {"_openxlsx_build_cell_types_integer",    (DL_FUNC) &_openxlsx_build_cell_types_integer,     2},
  {"_openxlsx_build_table_xml",             (DL_FUNC) &_openxlsx_build_table_xml,              6},
  {"_openxlsx_buildCellList",               (DL_FUNC) &_openxlsx_buildCellList,                3},
  {"_openxlsx_buildCellTypes",              (DL_FUNC) &_openxlsx_buildCellTypes,               2},
  {"_openxlsx_buildMatrixMixed",            (DL_FUNC) &_openxlsx_buildMatrixMixed,             8},
  {"_openxlsx_buildMatrixNumeric",          (DL_FUNC) &_openxlsx_buildMatrixNumeric,           6},
  {"_openxlsx_calc_column_widths",          (DL_FUNC) &_openxlsx_calc_column_widths,           7},
  {"_openxlsx_calc_number_rows",            (DL_FUNC) &_openxlsx_calc_number_rows,             2},
  {"_openxlsx_cell_ref_to_col",             (DL_FUNC) &_openxlsx_cell_ref_to_col,              1},
  {"_openxlsx_convert_from_excel_ref",      (DL_FUNC) &_openxlsx_convert_from_excel_ref,       1},
  {"_openxlsx_convert_to_excel_ref",        (DL_FUNC) &_openxlsx_convert_to_excel_ref,         2},
  {"_openxlsx_convert_to_excel_ref_expand", (DL_FUNC) &_openxlsx_convert_to_excel_ref_expand,  3},
  {"_openxlsx_cppReadFile",                 (DL_FUNC) &_openxlsx_cppReadFile,                  1},
  {"_openxlsx_get_extLst_Major",            (DL_FUNC) &_openxlsx_get_extLst_Major,             1},
  {"_openxlsx_get_letters",                 (DL_FUNC) &_openxlsx_get_letters,                  0},
  {"_openxlsx_get_shared_strings",          (DL_FUNC) &_openxlsx_get_shared_strings,           2},
  {"_openxlsx_getAttr",                     (DL_FUNC) &_openxlsx_getAttr,                      2},
  {"_openxlsx_getCellInfo",                 (DL_FUNC) &_openxlsx_getCellInfo,                  6},
  {"_openxlsx_getChildlessNode",            (DL_FUNC) &_openxlsx_getChildlessNode,             2},
  {"_openxlsx_getChildlessNode_ss",         (DL_FUNC) &_openxlsx_getChildlessNode_ss,          2},
  {"_openxlsx_getNodes",                    (DL_FUNC) &_openxlsx_getNodes,                     2},
  {"_openxlsx_getOpenClosedNode",           (DL_FUNC) &_openxlsx_getOpenClosedNode,            3},
  {"_openxlsx_int_2_cell_ref",              (DL_FUNC) &_openxlsx_int_2_cell_ref,               1},
  {"_openxlsx_isInternalHyperlink",         (DL_FUNC) &_openxlsx_isInternalHyperlink,          1},
  {"_openxlsx_loadworksheets",              (DL_FUNC) &_openxlsx_loadworksheets,               4},
  {"_openxlsx_map_cell_types_to_char",      (DL_FUNC) &_openxlsx_map_cell_types_to_char,       1},
  {"_openxlsx_map_cell_types_to_integer",   (DL_FUNC) &_openxlsx_map_cell_types_to_integer,    1},
  {"_openxlsx_matrixRowInds",               (DL_FUNC) &_openxlsx_matrixRowInds,                1},
  {"_openxlsx_read_file_newline",           (DL_FUNC) &_openxlsx_read_file_newline,            1},
  {"_openxlsx_read_workbook",               (DL_FUNC) &_openxlsx_read_workbook,               11},
  {"_openxlsx_write_worksheet_xml",         (DL_FUNC) &_openxlsx_write_worksheet_xml,          4},
  {"_openxlsx_write_worksheet_xml_2",       (DL_FUNC) &_openxlsx_write_worksheet_xml_2,        5},
  {"_openxlsx_write_file",                   (DL_FUNC) &_openxlsx_write_file,                  4},
  {NULL, NULL, 0}
};

void R_init__openxlsx(DllInfo *dll)
{
  R_registerRoutines(dll, NULL, CallEntries, NULL, NULL);
  R_useDynamicSymbols(dll, TRUE);
  R_forceSymbols(dll, FALSE);
}



