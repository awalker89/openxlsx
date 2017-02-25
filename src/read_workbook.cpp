

#include "openxlsx.h"



IntegerVector which_cpp(Rcpp::LogicalVector x) {
  IntegerVector v = seq(0, x.size() - 1);
  return v[x];
}


// [[Rcpp::export]]
CharacterVector get_shared_strings(std::string xmlFile, bool isFile){
  
  
  CharacterVector x;
  size_t pos = 0;
  std::string line;
  std::vector<std::string> lines;
  
  if(isFile){
    
    // READ IN FILE
    ifstream file;
    file.open(xmlFile.c_str());
    
    
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
    
    
  }else{
    line = xmlFile;
  }
  
  
  x = getNodes(line, "<si>");
  
  
  // define variables for sharedString part
  int n = x.size();
  CharacterVector strs(n);
  std::fill(strs.begin(), strs.end(), NA_STRING);
  
  std::string xml;
  size_t endPos = 0;
  
  std::string ttag = "<t";
  std::string tag = ">";
  std::string tagEnd = "<";
  
  
  // Now check for inline formatting
  pos = line.find("<rPr>", 0);
  if(pos == std::string::npos){
    
    // NO INLINE FORMATTING
    for(int i = 0; i < n; i++){ 
      
      // find opening tag     
      xml = x[i];
      pos = xml.find(ttag, 0); // find ttag      
      
      if(pos != std::string::npos){
        
        if(xml[pos+2] != '/'){
          pos = xml.find(tag, pos+1); // find where opening ttag ends
          endPos = xml.find(tagEnd, pos+1); // find where the node ends </t> (closing tag)
          strs[i] = xml.substr(pos+1, endPos-pos - 1).c_str();
        }
      }
    }
    
    
  }else{ // we have inline formatting
    
    
    for(int i = 0; i < n; i++){ 
      
      // find opening tag     
      xml = x[i];
      pos = xml.find(ttag, 0); // find ttag      
      
      if(xml[pos+2] != '/'){
        strs[i] = "";
        while(1){
          
          if(xml[pos+2] == '/'){
            break;
          }else{
            
            pos = xml.find(tag, pos+1); // find where opening ttag ends
            endPos = xml.find(tagEnd, pos+1); // find where the node ends </t> (closing tag)
            strs[i] += xml.substr(pos+1, endPos-pos - 1).c_str(); 
            pos = xml.find(ttag, endPos); // find ttag    
            
            if(pos == std::string::npos)
              break;
            
          }  
        }
        
      }
      
    } // end of for loop
    
    
    
  } // end of else inline formatting
  
  return wrap(strs);  
  
}



// [[Rcpp::export]]
List getCellInfo(std::string xmlFile,
                 CharacterVector sharedStrings,
                 bool skipEmptyRows,
                 int startRow,
                 IntegerVector rows,
                 bool getDates){
  
  //read in file
  std::string buf;
  
  // ifstream file;
  // file.open(xmlFile.c_str());
  // while (file >> buf)
  // xml += buf + ' ';
  std::string xml = read_file_newline(xmlFile);
  
  std::string rtag = "r=";
  std::string ttag = " t=";
  std::string stag = " s=";
  
  std::string tagEnd = "\"";
  
  std::string vtag = "<v>";
  std::string vtag2 = "<v ";
  std::string vtagEnd = "</v>";
  
  std::string cell;
  List res(6);
  
  std::size_t pos = xml.find("<sheetData>");      // find sheetData
  size_t endPos = 0;
  
  // If no data
  if(pos == std::string::npos){ 
    res = List::create(Rcpp::Named("nRows") = 0, Rcpp::Named("r") = 0);
    return res;
  }
  
  xml = xml.substr(pos + 11);     // get from "sheedData" to the end
  
  // startRow cut off
  int row_i = 0;
  if(startRow > 1){
    
    //find r and check the row number
    pos = xml.find("<row r=\"", 0);
    while(pos != std::string::npos){
      
      endPos = xml.find(tagEnd, pos + 8);
      row_i = atoi(xml.substr(pos + 8, endPos - pos - 8).c_str());
      
      if(row_i >= startRow){
        xml = xml.substr(pos);
        break;
      }else{
        pos = pos + 8;
      }
      pos = xml.find("<row r=\"", pos);
    }
    
    // no rows
    if(pos == std::string::npos){
      res = List::create(Rcpp::Named("nRows") = 0, Rcpp::Named("r") = 0);
      return res;
    }
    
  }
  
  // getting rows
  //rows cut off, loop over entire string and only take rows specified in rows vector
  if(!is_na(rows)[0]){
    
    CharacterVector xml_rows = getNodes(xml, "<row");
    int nr = xml_rows.size();
    
    // no rows
    if(nr == 0){
      res = List::create(Rcpp::Named("nRows") = 0, Rcpp::Named("r") = 0);
      return res;
    }
    
    
    // for each row pull out the row number, check the row number against 'rows'
    std::string row_xml_i;
    xml = "";
    int place = 0;
    
    // get first one and remove beginning of rows
    row_xml_i = xml_rows[0];
    pos = row_xml_i.find("<row r=\"", 0);
    endPos = row_xml_i.find(tagEnd, pos + 8);
    row_i = atoi(row_xml_i.substr(pos + 8, endPos - pos - 8).c_str());
    rows = rows[rows >= row_i];
    
    int nr_sub = rows.size();
    if(nr_sub > 0){
      for(int i = 0; i < nr; i++){
        
        row_xml_i = xml_rows[i];
        pos = row_xml_i.find("<row r=\"", 0);
        endPos = row_xml_i.find(tagEnd, pos + 8);
        row_i = atoi(row_xml_i.substr(pos + 8, endPos - pos - 8).c_str());
        
        if (std::find(rows.begin()+place, rows.end(), row_i)!= rows.end()){
          xml += row_xml_i + ' ';
          place++;
          
          if(place == nr_sub)
            break;
          
        }
        
        
      } // end of cut off unwanted xml
    }
    
  }
  
  // count cells with children
  int ocs = 0;
  string::size_type start = 0;
  while((start = xml.find("</v>", start)) != string::npos){
    ++ocs;
    start += 4;
  }
  
  if(ocs == 0){
    res = List::create(Rcpp::Named("nRows") = 0, Rcpp::Named("r") = 0);
    return res;
  }
  
  // pull out cell merges
  CharacterVector merge_cell_xml = getChildlessNode(xml, "<mergeCell ");
  
  
  CharacterVector r(ocs);
  CharacterVector t(ocs);
  CharacterVector v(ocs);
  CharacterVector string_refs(ocs);
  std::fill(string_refs.begin(), string_refs.end(), NA_STRING);
  
  int s_ocs = 0;
  if(getDates)
    s_ocs = ocs;
  
  CharacterVector s(s_ocs);
  
  std::fill(t.begin(), t.end(), "n");
  std::fill(v.begin(), v.end(), NA_STRING);
  std::fill(s.begin(), s.end(), NA_STRING);
  
  
  
  int i = 0;
  int ss_ind = 0;
  size_t nextPos = 3;
  size_t vPos = 2;
  pos = xml.find("<c ", 0);
  
  // PULL OUT CELL AND ATTRIBUTES
  while(i < ocs){
    
    if(pos != std::string::npos){
      
      nextPos = xml.find("<c ", pos + 9);
      vPos = xml.find("</v>", pos + 8); // have to atleast pass <c r="XX">
      
      if(vPos < nextPos){
        
        cell = xml.substr(pos, nextPos - pos);
        
        // Pull out ref
        pos = cell.find(rtag, 0);  // find r="
        endPos = cell.find(tagEnd, pos + 3);  // find next "
        r[i] = cell.substr(pos + 3, endPos - pos - 3).c_str();
        
        // Pull out type
        pos = cell.find(ttag, 0);  // find t="
        if(pos != std::string::npos){
          endPos = cell.find(tagEnd, pos + 4);  // find next "
          t[i] = cell.substr(pos + 4, endPos - pos - 4).c_str();
        }
        
        // Pull out style
        if(getDates){
          pos = cell.find(stag, 0);  // find s="
          if(pos != std::string::npos){
            endPos = cell.find(tagEnd, pos + 4);  // find next "
            s[i] = cell.substr(pos + 4, endPos - pos - 4).c_str();
          }
        }
        
        
        // If the value is s or shared we replace with sharedString
        // If it's b we replace with "TRUE" or "FALSE"
        // If the value is str it's already a string
        if(t[i] == "e"){
          
          v[i] = NA_STRING;
          
        }else{
          
          // find <v> tag and </v> end tag
          endPos = cell.find(vtagEnd, 0);
          if(endPos != std::string::npos){
            pos = cell.find("<v", 0);
            pos = cell.find(">", pos);
            v[i] = cell.substr(pos + 1, endPos - pos - 1);
          }
          
          
          // possible values for t are n, s, shared, b, str, e 
          
          // do replacement
          if(t[i] == "s"){
            
            ss_ind = atoi(v[i]);
            v[i] = sharedStrings[ss_ind];
            
            if(v[i] == "openxlsx_na_vlu"){
              v[i] = NA_STRING;
            }
            
            string_refs[i] = r[i];
            
          }else if(t[i] == "b"){
            if(v[i] == "1"){
              v[i] = "TRUE";
            }else{
              v[i] = "FALSE";
            }
            string_refs[i] = r[i];
            
          }else if(t[i] == "str"){
            string_refs[i] = r[i];
          }
        }
        
        i++; // INCREMENT OVER OCCURENCES
      }
      
      pos = nextPos;
      
    }
  } // end of while loop over occurences
  // END OF CELL AND ATTRIBUTION GATHERING
  
  string_refs = string_refs[!is_na(string_refs)];
  
  int nRows = calc_number_rows(r, skipEmptyRows);
  res = List::create(Rcpp::Named("r") = r,
                     Rcpp::Named("string_refs") = string_refs,
                     Rcpp::Named("v") = v,
                     Rcpp::Named("s") = s,
                     Rcpp::Named("nRows") = nRows,
                     Rcpp::Named("cellMerge") = merge_cell_xml
  );
  
  
  return wrap(res);  
  
}










// [[Rcpp::export]]
SEXP read_workbook(IntegerVector cols_in,
                   IntegerVector rows_in,
                   CharacterVector v,
                   
                   IntegerVector string_inds,
                   LogicalVector is_date,
                   bool hasColNames,
                   bool skipEmptyRows,
                   bool skipEmptyCols,
                   int nRows,
                   Function clean_names
){
  
  
  IntegerVector cols = clone(cols_in);
  IntegerVector rows = clone(rows_in);
  
  int nCells = rows.size();
  int nDates = is_date.size();
  
  /* do we have any dates */
  bool has_date;
  if(nDates == 1){
    if(is_true(any(is_na(is_date)))){
      has_date = false;
    }else{
      has_date = true;
    }
  }else if(nDates == nCells){
    has_date = true;
  }else{
    has_date = false;
  }
  
  bool has_strings = true;
  IntegerVector st_inds0 (1);
  st_inds0[0] = string_inds[0];
  if(is_true(all(is_na(st_inds0))))
    has_strings = false;
  
  
  
  IntegerVector uni_cols = sort_unique(cols);
  if(!skipEmptyCols){  // want to keep all columns - just create a sequence from 1:max(cols)
    uni_cols = seq(1, max(uni_cols));
    cols = cols - 1;
  }else{
    cols = match(cols, uni_cols) - 1;
  }
  
  // scale columns from i:j to 1:(j-i+1)
  int nCols = *std::max_element(cols.begin(), cols.end()) + 1;
  
  // scale rows from i:j to 1:(j-i+1)
  IntegerVector uni_rows = sort_unique(rows);
  
  if(skipEmptyRows){
    rows = match(rows, uni_rows) - 1;
    //int nRows = *std::max_element(rows.begin(), rows.end()) + 1;
  }else{
    rows = rows - rows[0];
  }
  
  // Check if first row are all strings
  //get first row number
  
  CharacterVector col_names(nCols);
  IntegerVector removeFlag;
  int pos = 0;
  
  // If we are told col_names exist take the first row and fill any gaps with X.i
  if(hasColNames){
    
    int row_1 = rows[0];
    char name[6];
    
    IntegerVector row_inds = which_cpp(rows == row_1);
    IntegerVector header_cols = cols[row_inds];
    IntegerVector header_inds = match(seq(0, nCols), na_omit(header_cols));
    LogicalVector missing_header = is_na(header_inds);
    
    // looping over each column
    for(int i=0; i < nCols; i++){
      
      if(missing_header[i]){  // a missing header element
        
        sprintf(&(name[0]), "X%d", i+1);
        col_names[i] = name;
        
      }else{  // this is a header elements 
        
        col_names[i] = v[pos];
        pos++;
        
      }
      
    }
    
    // tidy up column names
    col_names = clean_names(col_names);
    
    //--------------------------------------------------------------------------------
    // Remove elements from rows, cols, v that have been used as headers
    
    // I've used the first pos elements as headers
    // stringInds contains the indexes of v which are strings
    // string_inds <- string_inds[string_inds > pos]
    if(has_strings){
      string_inds = string_inds[string_inds > pos];
      string_inds = string_inds - pos;
    }
    
    
    rows.erase (rows.begin(), rows.begin() + pos);
    rows = rows - 1;
    v.erase (v.begin(), v.begin() + pos);
    
    //If nothing left return a data.frame with 0 rows
    if(rows.size() == 0){
      
      List dfList(nCols);
      IntegerVector rowNames(0);
      
      for(int i = 0; i < nCols; i++){
        dfList[i] = LogicalVector(0); // this is what read.table does (bool type)
      }
      
      dfList.attr("names") = col_names;
      dfList.attr("row.names") = rowNames;
      dfList.attr("class") = "data.frame";
      return wrap(dfList);
    }
    
    cols.erase(cols.begin(), cols.begin() + pos);
    nRows--; // decrement number of rows as first row is now being used as col_names
    nCells = nCells - pos;
    
    // End Remove elements from rows, cols, v that have been used as headers
    //--------------------------------------------------------------------------------
    
    
    
  }else{ // else col_names is FALSE
    char name[6];
    for(int i =0; i < nCols; i++){
      sprintf(&(name[0]), "X%d", i+1);
      col_names[i] = name;
    }
  }
  
  
  // ------------------ column names complete
  
  
  
  
  
  // Possible there are no string_inds to begin with and value of string_inds is 0
  // Possible we have string_inds but they have now all been used up by headers
  bool allNumeric = false;
  if((string_inds.size() == 0) | all(is_na(string_inds)))
    allNumeric = true;
  
  if(has_date){
    if(is_true(any(is_date)))
      allNumeric = false;
  }
  
  // If we have colnames some elements where used to create these -so we remove the corresponding number of elements
  if(hasColNames & has_date)
    is_date.erase(is_date.begin(), is_date.begin() + pos);
  
  
  
  //Intialise return data.frame
  SEXP m; 
  
  // for(int i = 0; i < rows.size(); i++)
  //   Rcout << "rows[i]: " << rows[i] << endl;
  // 
  // Rcout << "nRows " << nRows << endl;
  // Rcout << "nCols: " << nCols << endl;
  // Rcout << "cols.size(): " << cols.size() << endl;
  // Rcout << "rows.size(): " << rows.size() << endl;
  // Rcout << "is_date.size(): " << is_date.size() << endl;
  // Rcout << "v.size(): " << v.size() << endl;
  // Rcout << "has_date: " << has_date << endl;

  if(allNumeric){
    
    m = buildMatrixNumeric(v, rows, cols, col_names, nRows, nCols);
    
  }else{
    
    // If it contains any strings it will be a character column
    IntegerVector char_cols_unique;
    if(all(is_na(string_inds))){
      char_cols_unique = -1;
    }else{
      
      IntegerVector columns_which_are_characters = cols[string_inds - 1];
      char_cols_unique = unique(columns_which_are_characters);
      
    }
    
    //date columns
    IntegerVector date_columns(1);
    if(has_date){
      
      date_columns = cols[is_date];
      date_columns = sort_unique(date_columns);
      
    }else{
      date_columns[0] = -1;
    }
    
    // List d(10);
    // d[0] = v;
    // d[2] = rows;
    // d[3] = cols;
    // d[4] = col_names;
    // d[5] = nRows;
    // d[6] = nCols;
    // d[7] = char_cols_unique;
    // d[8] = date_columns;
    // return(wrap(d));
    // Rcout << "Running buildMatrixMixed" << endl;
    
    m = buildMatrixMixed(v, rows, cols, col_names, nRows, nCols, char_cols_unique, date_columns);
    
  }
  
  return wrap(m) ;
  
  
}







// [[Rcpp::export]]  
int calc_number_rows(CharacterVector x, bool skipEmptyRows){
  
  int n = x.size();
  if(n == 0)
    return(0);
  
  int nRows;
  
  if(skipEmptyRows){
    
    CharacterVector res(n);
    std::string r;
    for(int i = 0; i < n; i++){
      r = x[i];
      r.erase(std::remove_if(r.begin(), r.end(), ::isalpha), r.end());
      res[i] = r;
    }
    
    CharacterVector uRes = unique(res);
    nRows = uRes.size();
    
  }else{
    
    std::string fRef = as<std::string>(x[0]);
    std::string lRef = as<std::string>(x[n-1]);
    fRef.erase(std::remove_if(fRef.begin(), fRef.end(), ::isalpha), fRef.end());
    lRef.erase(std::remove_if(lRef.begin(), lRef.end(), ::isalpha), lRef.end());
    int firstRow = atoi(fRef.c_str());
    int lastRow = atoi(lRef.c_str());
    nRows = lastRow - firstRow + 1;
    
  }
  
  return(nRows);
  
}

