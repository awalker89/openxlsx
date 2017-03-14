
#include <R.h>
#include <Rinternals.h>
#include <stdlib.h> // for NULL
#include <R_ext/Rdynload.h>

/* 
generated with
tools::package_native_routine_registration_skeleton(".")
*/

/* .Call calls */
extern SEXP openxlsx_build_cell_merges(SEXP);
extern SEXP openxlsx_build_cell_types_integer(SEXP, SEXP);
extern SEXP openxlsx_build_table_xml(SEXP, SEXP, SEXP, SEXP, SEXP, SEXP);
extern SEXP openxlsx_buildCellList(SEXP, SEXP, SEXP);
extern SEXP openxlsx_buildCellTypes(SEXP, SEXP);
extern SEXP openxlsx_buildMatrixMixed(SEXP, SEXP, SEXP, SEXP, SEXP, SEXP, SEXP, SEXP);
extern SEXP openxlsx_buildMatrixNumeric(SEXP, SEXP, SEXP, SEXP, SEXP, SEXP);
extern SEXP openxlsx_calc_column_widths(SEXP, SEXP, SEXP, SEXP, SEXP, SEXP, SEXP);
extern SEXP openxlsx_calc_number_rows(SEXP, SEXP);
extern SEXP openxlsx_cell_ref_to_col(SEXP);
extern SEXP openxlsx_convert_from_excel_ref(SEXP);
extern SEXP openxlsx_convert_to_excel_ref(SEXP, SEXP);
extern SEXP openxlsx_convert_to_excel_ref_expand(SEXP, SEXP, SEXP);
extern SEXP openxlsx_cppReadFile(SEXP);
extern SEXP openxlsx_get_extLst_Major(SEXP);
extern SEXP openxlsx_get_letters();
extern SEXP openxlsx_get_shared_strings(SEXP, SEXP);
extern SEXP openxlsx_getAttr(SEXP, SEXP);
extern SEXP openxlsx_getCellInfo(SEXP, SEXP, SEXP, SEXP, SEXP, SEXP);
extern SEXP openxlsx_getChildlessNode(SEXP, SEXP);
extern SEXP openxlsx_getChildlessNode_ss(SEXP, SEXP);
extern SEXP openxlsx_getNodes(SEXP, SEXP);
extern SEXP openxlsx_getOpenClosedNode(SEXP, SEXP, SEXP);
extern SEXP openxlsx_int_2_cell_ref(SEXP);
extern SEXP openxlsx_isInternalHyperlink(SEXP);
extern SEXP openxlsx_loadworksheets(SEXP, SEXP, SEXP, SEXP);
extern SEXP openxlsx_map_cell_types_to_char(SEXP);
extern SEXP openxlsx_map_cell_types_to_integer(SEXP);
extern SEXP openxlsx_matrixRowInds(SEXP);
extern SEXP openxlsx_read_file_newline(SEXP);
extern SEXP openxlsx_read_workbook(SEXP, SEXP, SEXP, SEXP, SEXP, SEXP, SEXP, SEXP, SEXP, SEXP);
extern SEXP openxlsx_write_worksheet_xml(SEXP, SEXP, SEXP, SEXP);
extern SEXP openxlsx_write_worksheet_xml_2(SEXP, SEXP, SEXP, SEXP, SEXP);
extern SEXP openxlsx_writeFile(SEXP, SEXP, SEXP, SEXP);

static const R_CallMethodDef CallEntries[] = {
  {"openxlsx_build_cell_merges",           (DL_FUNC) &openxlsx_build_cell_merges,            1},
  {"openxlsx_build_cell_types_integer",    (DL_FUNC) &openxlsx_build_cell_types_integer,     2},
  {"openxlsx_build_table_xml",             (DL_FUNC) &openxlsx_build_table_xml,              6},
  {"openxlsx_buildCellList",               (DL_FUNC) &openxlsx_buildCellList,                3},
  {"openxlsx_buildCellTypes",              (DL_FUNC) &openxlsx_buildCellTypes,               2},
  {"openxlsx_buildMatrixMixed",            (DL_FUNC) &openxlsx_buildMatrixMixed,             8},
  {"openxlsx_buildMatrixNumeric",          (DL_FUNC) &openxlsx_buildMatrixNumeric,           6},
  {"openxlsx_calc_column_widths",          (DL_FUNC) &openxlsx_calc_column_widths,           7},
  {"openxlsx_calc_number_rows",            (DL_FUNC) &openxlsx_calc_number_rows,             2},
  {"openxlsx_cell_ref_to_col",             (DL_FUNC) &openxlsx_cell_ref_to_col,              1},
  {"openxlsx_convert_from_excel_ref",      (DL_FUNC) &openxlsx_convert_from_excel_ref,       1},
  {"openxlsx_convert_to_excel_ref",        (DL_FUNC) &openxlsx_convert_to_excel_ref,         2},
  {"openxlsx_convert_to_excel_ref_expand", (DL_FUNC) &openxlsx_convert_to_excel_ref_expand,  3},
  {"openxlsx_cppReadFile",                 (DL_FUNC) &openxlsx_cppReadFile,                  1},
  {"openxlsx_get_extLst_Major",            (DL_FUNC) &openxlsx_get_extLst_Major,             1},
  {"openxlsx_get_letters",                 (DL_FUNC) &openxlsx_get_letters,                  0},
  {"openxlsx_get_shared_strings",          (DL_FUNC) &openxlsx_get_shared_strings,           2},
  {"openxlsx_getAttr",                     (DL_FUNC) &openxlsx_getAttr,                      2},
  {"openxlsx_getCellInfo",                 (DL_FUNC) &openxlsx_getCellInfo,                  6},
  {"openxlsx_getChildlessNode",            (DL_FUNC) &openxlsx_getChildlessNode,             2},
  {"openxlsx_getChildlessNode_ss",         (DL_FUNC) &openxlsx_getChildlessNode_ss,          2},
  {"openxlsx_getNodes",                    (DL_FUNC) &openxlsx_getNodes,                     2},
  {"openxlsx_getOpenClosedNode",           (DL_FUNC) &openxlsx_getOpenClosedNode,            3},
  {"openxlsx_int_2_cell_ref",              (DL_FUNC) &openxlsx_int_2_cell_ref,               1},
  {"openxlsx_isInternalHyperlink",         (DL_FUNC) &openxlsx_isInternalHyperlink,          1},
  {"openxlsx_loadworksheets",              (DL_FUNC) &openxlsx_loadworksheets,               4},
  {"openxlsx_map_cell_types_to_char",      (DL_FUNC) &openxlsx_map_cell_types_to_char,       1},
  {"openxlsx_map_cell_types_to_integer",   (DL_FUNC) &openxlsx_map_cell_types_to_integer,    1},
  {"openxlsx_matrixRowInds",               (DL_FUNC) &openxlsx_matrixRowInds,                1},
  {"openxlsx_read_file_newline",           (DL_FUNC) &openxlsx_read_file_newline,            1},
  {"openxlsx_read_workbook",               (DL_FUNC) &openxlsx_read_workbook,               10},
  {"openxlsx_write_worksheet_xml",         (DL_FUNC) &openxlsx_write_worksheet_xml,          4},
  {"openxlsx_write_worksheet_xml_2",       (DL_FUNC) &openxlsx_write_worksheet_xml_2,        5},
  {"openxlsx_writeFile",                   (DL_FUNC) &openxlsx_writeFile,                    4},
  {NULL, NULL, 0}
};

void R_init_openxlsx(DllInfo *dll)
{
  R_registerRoutines(dll, NULL, CallEntries, NULL, NULL);
  R_useDynamicSymbols(dll, TRUE);
  R_forceSymbols(dll, FALSE);
}



