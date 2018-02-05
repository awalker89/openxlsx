

#include <Rcpp.h>
#include <iostream>
#include <fstream>
#include <sstream>

using namespace Rcpp;
using namespace std;



// load workbook
int cell_ref_to_col(string );
CharacterVector int_2_cell_ref(IntegerVector);

// write file
SEXP write_worksheet_xml(string, string, List, Reference, IntegerVector, CharacterVector, string );


// write_data.cpp
CharacterVector map_cell_types_to_char(IntegerVector);
IntegerVector map_cell_types_to_integer(CharacterVector);
  


std::vector<std::string> get_letters();


IntegerVector convert_from_excel_ref( CharacterVector x );

SEXP calc_column_widths(Reference sheet_data, std::vector<std::string> sharedStrings, IntegerVector autoColumns, NumericVector widths, float baseFontCharWidth, float minW, float maxW);

SEXP getOpenClosedNode(std::string xml, std::string open_tag, std::string close_tag);

std::string cppReadFile(std::string xmlFile);
std::string read_file_newline(std::string xmlFile);


SEXP getNodes(std::string xml, std::string tagIn);
std::vector<std::string> getChildlessNode_ss(std::string xml, std::string tag);
CharacterVector get_extLst_Major(std::string xml);
CharacterVector getChildlessNode(std::string xml, std::string tag);
SEXP getAttr(CharacterVector x, std::string tag);

CharacterVector get_shared_strings(std::string xmlFile, bool isFile);


List buildCellList( CharacterVector r, CharacterVector t, CharacterVector v);
SEXP openxlsx_convert_to_excel_ref(IntegerVector cols, std::vector<std::string> LETTERS);

SEXP buildMatrixNumeric(CharacterVector v, IntegerVector rowInd, IntegerVector colInd,
                        CharacterVector colNames, int nRows, int nCols);

SEXP buildMatrixMixed(CharacterVector v,
                      IntegerVector rowInd,
                      IntegerVector colInd,
                      CharacterVector colNames,
                      int nRows,
                      int nCols,
                      IntegerVector charCols,
                      IntegerVector dateCols);

List getCellInfo(std::string xmlFile,
                 CharacterVector sharedStrings,
                 bool skipEmptyRows,
                 int startRow,
                 IntegerVector rows,
                 bool getDates);


SEXP convert_to_excel_ref_expand(const std::vector<int>& cols, const std::vector<std::string>& LETTERS, const std::vector<std::string>& rows);

IntegerVector matrixRowInds(IntegerVector indices);
CharacterVector build_table_xml(std::string table, std::string ref, std::vector<std::string> colNames, bool showColNames, std::string tableStyle, bool withFilter);
int calc_number_rows(CharacterVector x, bool skipEmptyRows);
CharacterVector buildCellTypes(CharacterVector classes, int nRows);
LogicalVector isInternalHyperlink(CharacterVector x);


// helper functions
string itos(int i);
SEXP write_file(std::string parent, std::string xmlText, std::string parentEnd, std::string R_fileName);
  
  