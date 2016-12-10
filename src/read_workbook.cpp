

#include "openxlsx.h"






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
SEXP readWorkbook(CharacterVector v,
                  CharacterVector r,
                  CharacterVector string_refs,
                  LogicalVector is_date,
                  int nRows,
                  bool hasColNames,
                  bool skipEmptyRows,
                  bool skipEmptyCols,
                  Function clean_names
){
  
  
  // Convert r to column number and shift to scale 0:nCols  (from eg. J:AA in the worksheet)
  int nCells = r.size();
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
  
  IntegerVector colNumbers = convert_from_excel_ref(r); 
  IntegerVector uCols = sort_unique(colNumbers);
  
  if(!skipEmptyCols){  // want to keep all columns - just create a sequence from 1:max(cols)
    uCols = seq(1, max(uCols));
  }
  
  colNumbers = match(colNumbers, uCols) - 1;
  int nCols = *std::max_element(colNumbers.begin(), colNumbers.end()) + 1;
  
  // Check if first row are all strings
  //get first row number
  
  CharacterVector colNames(nCols);
  IntegerVector removeFlag;
  int pos = 0;
  
  // If we are told colNames exist take the first row and fill any gaps with X.i
  if(hasColNames){
    
    char name[6];
    std::vector<std::string> firstRowNumbers(nCols);
    std::string ref;
    
    // convert first nCols r to firstRowNumbers
    // get r for which firstRowNumbers == firstRowNumbers[0];
    for(int i = 0; i < nCols; i++){
      ref = r[i];
      ref.erase(std::remove_if(ref.begin(), ref.end(), ::isalpha), ref.end());
      firstRowNumbers[i] = ref;
    }
    
    ref = firstRowNumbers[0];
    
    for(int i = 0; i < nCols; i++){
      if((i == colNumbers[pos]) & (firstRowNumbers[pos] == ref)){
        colNames[i] = v[pos];
        pos = pos + 1;
      }else{
        sprintf(&(name[0]), "X%d", i+1);
        colNames[i] = name;
      }
    }
    
    // remove elements in string_refs that appear in the first row (if there are numerics in first row this will remove less elements than pos)
    CharacterVector::iterator first = r.begin();
    CharacterVector::iterator last = r.begin() + pos;
    CharacterVector toRemove(first, last);
    removeFlag = match(string_refs, toRemove);
    string_refs.erase(string_refs.begin(), string_refs.begin() + sum(!is_na(removeFlag)));
    
    // remove elements that are now being used as colNames (there are int pos many of these)    
    //check we have some r left if not return a data.frame with zero rows
    if(pos > 0)
      r.erase(r.begin(), r.begin() +  pos); 
    
    // tidy up column names
    colNames = clean_names(colNames);
    
    //If nothing left return a data.frame with 0 rows
    if(r.size() == 0){
      List dfList(nCols);
      IntegerVector rowNames(0);
      
      for(int i =0; i < nCols; i++){
        dfList[i] = LogicalVector(0);
      }
      
      dfList.attr("names") = colNames;
      dfList.attr("row.names") = rowNames;
      dfList.attr("class") = "data.frame";
      return wrap(dfList);
    }
    
    colNumbers.erase(colNumbers.begin(), colNumbers.begin() + pos); 
    nRows--; // decrement number of rows as first row is now being used as colNames
    nCells = nCells - pos;
    
  }else{ // else colNames is FALSE
    char name[6];
    for(int i =0; i < nCols; i++){
      sprintf(&(name[0]), "X%d", i+1);
      colNames[i] = name;
    }
  }
  
  /* column names now sorted */
  
  // getRow numbers from r 
  IntegerVector rowNumbers(nCells);  
  if(nCells == nRows*nCols){
    
    IntegerVector uRows = seq(0, nRows-1);
    rowNumbers = rep_each(uRows, nCols); 
    
  }else{
    
    std::vector<std::string> rs = as<std::vector<std::string> >(r);
    for(int i = 0; i < nCells; i++){
      std::string a = rs[i];
      a.erase(std::remove_if(a.begin(), a.end(), ::isalpha), a.end());
      rowNumbers[i] = atoi(a.c_str()) - 1; 
    }
    
    if(skipEmptyRows){
      rowNumbers = matrixRowInds(rowNumbers);
    }else{
      rowNumbers = rowNumbers - rowNumbers[0];
    }
    
  }
  
  // Possible there are no stringInds to begin with and value of stringInds is 0
  // Possible we have stringInds but they have now all been used up by string_refs
  bool allNumeric = false;
  if((string_refs.size() == 0) | all(is_na(string_refs)))
    allNumeric = true;
  
  if(has_date){
    if(is_true(any(is_date)))
      allNumeric = false;
  }
  
  // If we build colnames we remove the ref & values for the pos number of elements used 
  if(hasColNames & has_date){
    is_date.erase(is_date.begin(), is_date.begin() + pos);
  }
  
  
  //Intialise return data.frame
  SEXP m; 
  v.erase(v.begin(), v.begin() + pos);
  
  if(allNumeric){
    
    m = buildMatrixNumeric(v, rowNumbers, colNumbers, colNames, nRows, nCols);
    
  }else{
    
    // If it contains any strings it will be a character column
    IntegerVector charCols;
    if(all(is_na(string_refs))){
      charCols = -1;
    }else{
      charCols = match(string_refs, r);
      IntegerVector charColNumbers = colNumbers[charCols-1];
      charCols = unique(charColNumbers);
    }
    
    //Rcout << "colNumbers.size(): " << colNumbers.size() << endl; 
    //Rcout << "is_date.size(): " << is_date.size() << endl; 
    //Rcout << "v.size(): " << v.size() << endl; 
    //Rcout << "vn.size(): " << vn.size() << endl; 
    //Rcout << "rowNumbers.size(): " << rowNumbers.size() << endl; 
    
    //date columns
    IntegerVector dateCols(1);
    if(has_date){
      
      dateCols = colNumbers[is_date];
      dateCols = sort_unique(dateCols);
      
    }else{
      dateCols[0] = -1;
    }
    
    m = buildMatrixMixed(v, rowNumbers, colNumbers, colNames, nRows, nCols, charCols, dateCols);
    
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

