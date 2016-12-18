
#include "openxlsx.h"






// [[Rcpp::export]]
SEXP calc_column_widths(Reference sheet_data
                          , std::vector<std::string> sharedStrings
                          , IntegerVector autoColumns
                          , NumericVector widths
                          , float baseFontCharWidth
                          , float minW
                          , float maxW){
  
  
  int n = sheet_data.field("n_elements");
  IntegerVector cell_types = sheet_data.field("t");
  StringVector cell_values(sheet_data.field("v"));
  IntegerVector cell_cols = sheet_data.field("cols");
  
  NumericVector cell_n_character(n);
  CharacterVector r(n);
  int nLen;
  
  std::string tmp;

  // get widths of all values
  for(int i = 0; i < n; i++){

    if(cell_types[i] == 1){ // "s"
      cell_n_character[i] = sharedStrings[atoi(cell_values[i])].length() - 37; //-37 for shared string tags around text
    }else{
      tmp = cell_values[i];
      nLen = tmp.length();
      cell_n_character[i] = min(nLen, 11); // For numerics - max width is 11
    }
    
  }
  
  
  // get column for each value

  // reducing to only the columns that are auto
  LogicalVector notNA = !is_na(match(cell_cols, autoColumns));
  cell_cols = cell_cols[notNA];
  cell_n_character = cell_n_character[notNA];
  widths = widths[notNA];
  IntegerVector unique_cell_cols = sort_unique(cell_cols);
  
  size_t k = unique_cell_cols.size();
  NumericVector column_widths(k);

  
  // for each unique column, get all widths for that column and take max
  for(size_t i = 0; i < k; i++){
    NumericVector wTmp = cell_n_character[cell_cols == unique_cell_cols[i]];
    NumericVector thisColWidths = widths[cell_cols == unique_cell_cols[i]];
    column_widths[i] = max(wTmp * thisColWidths / baseFontCharWidth); 
  }
  
  column_widths[column_widths < minW] = minW;
  column_widths[column_widths > maxW] = maxW;    
  
  // assign column names
  column_widths.attr("names") = unique_cell_cols;
  
  return(wrap(column_widths));
  
}





// [[Rcpp::export]]
SEXP convert_to_excel_ref(IntegerVector cols, std::vector<std::string> LETTERS){
  
  int n = cols.size();  
  CharacterVector res(n);
  
  int x;
  int modulo;
  for(int i = 0; i < n; i++){
    x = cols[i];
    string columnName;
    
    while(x > 0){  
      modulo = (x - 1) % 26;
      columnName = LETTERS[modulo] + columnName;
      x = (x - modulo) / 26;
    }
    res[i] = columnName;
  }
  
  return res ;
  
}

// [[Rcpp::export]]
IntegerVector convert_from_excel_ref( CharacterVector x ){
  
  // This function converts the Excel column letter to an integer
  
  std::vector<std::string> r = as<std::vector<std::string> >(x);
  int n = r.size();
  int k;
  
  std::string a;
  IntegerVector colNums(n);
  char A = 'A';
  int aVal = (int)A - 1;
  
  for(int i = 0; i < n; i++){
    a = r[i];
    
    // remove digits from string
    a.erase(std::remove_if(a.begin()+1, a.end(), ::isdigit), a.end());
    
    int sum = 0;
    k = a.length();
    
    for (int j = 0; j < k; j++){
      sum *= 26;
      sum += (a[j] - aVal);
      
    }
    colNums[i] = sum;
  }
  
  return colNums;
  
}





// [[Rcpp::export]]
SEXP convert_to_excel_ref_expand(const std::vector<int>& cols, const std::vector<std::string>& LETTERS, const std::vector<std::string>& rows){
  
  int n = cols.size();  
  int nRows = rows.size();
  std::vector<std::string> res(n);
  
  //Convert col number to excel col letters
  size_t x;
  size_t modulo;
  for(int i = 0; i < n; i++){
    x = cols[i];
    string columnName;
    
    while(x > 0){  
      modulo = (x - 1) % 26;
      columnName = LETTERS[modulo] + columnName;
      x = (x - modulo) / 26;
    }
    res[i] = columnName;
  }
  
  CharacterVector r(n*nRows);
  CharacterVector names(n*nRows);
  size_t c = 0;
  for(int i=0; i < nRows; i++)
    for(int j=0; j < n; j++){
      r[c] = res[j] + rows[i];
      names[c] = rows[i];
      c++;
    }
    
    r.attr("names") = names;
  return wrap(r) ;
  
}



// [[Rcpp::export]]
LogicalVector isInternalHyperlink(CharacterVector x){
  
  int n = x.size();
  std::string xml;
  std::string tag = "r:id=";
  size_t found;
  LogicalVector isInternal(n);
  
  for(int i = 0; i < n; i++){ 
    
    // find location tag  
    xml = x[i];
    found = xml.find(tag, 0);
    
    if (found != std::string::npos){
      isInternal[i] = false;
    }else{
      isInternal[i] = true;
    }
    
  }
  
  return wrap(isInternal) ;  
  
}


string itos(int i){
  
  // convert int to string
  stringstream s;
  s << i;
  return s.str();
  
}


// [[Rcpp::export]]
SEXP writeFile(std::string parent, std::string xmlText, std::string parentEnd, std::string R_fileName) {
  
  const char * s = R_fileName.c_str();
  
  std::ofstream xmlFile;
  xmlFile.open (s);
  xmlFile << "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>";
  xmlFile << parent;
  xmlFile << xmlText;
  xmlFile << parentEnd;
  xmlFile.close();
  
  return 0;
}



// [[Rcpp::export]]
std::string cppReadFile(std::string xmlFile){
  
  std::string buf;
  std::string xml;
  ifstream file;
  file.open(xmlFile.c_str());
  
  while (file >> buf)
    xml += buf + ' ';
  
  return xml;
}



// [[Rcpp::export]]
std::string read_file_newline(std::string xmlFile){
  
  ifstream file;
  file.open(xmlFile.c_str());
  std::vector<std::string> lines;
  
  std::string line;
  while ( std::getline(file, line) )
  {
    // skip empty lines:
    if (line.empty())
      continue;
    
    lines.push_back(line);
  }
  
  line = "";
  int n = lines.size();
  for(int i = 0;i < n; ++i)
    line += lines[i] + "\n";  
  
  
  return line;
}



// [[Rcpp::export]]
std::vector<std::string> get_letters(){
  
  std::vector<std::string> LETTERS(26);
  
  LETTERS[0] = "A";
  LETTERS[1] = "B";
  LETTERS[2] = "C";
  LETTERS[3] = "D";
  LETTERS[4] = "E";
  LETTERS[5] = "F";
  LETTERS[6] = "G";
  LETTERS[7] = "H";
  LETTERS[8] = "I";
  LETTERS[9] = "J";
  LETTERS[10] = "K";
  LETTERS[11] = "L";
  LETTERS[12] = "M";
  LETTERS[13] = "N";
  LETTERS[14] = "O";
  LETTERS[15] = "P";
  LETTERS[16] = "Q";
  LETTERS[17] = "R";
  LETTERS[18] = "S";
  LETTERS[19] = "T";
  LETTERS[20] = "U";
  LETTERS[21] = "V";
  LETTERS[22] = "W";
  LETTERS[23] = "X";
  LETTERS[24] = "Y";
  LETTERS[25] = "Z";
  
  return(LETTERS);
  
}