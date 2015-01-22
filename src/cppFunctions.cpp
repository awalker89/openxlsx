#include <Rcpp.h>
#include <iostream>
#include <fstream>
#include <sstream>

using namespace Rcpp;
using namespace std;



// [[Rcpp::export]]
IntegerVector RcppConvertFromExcelRef( CharacterVector x ){
  
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

string itos(int i) // convert int to string
{
  stringstream s;
  s << i;
  return s.str();
}




// [[Rcpp::export]]
SEXP calcColumnWidths(List sheetData, std::vector<std::string> sharedStrings, IntegerVector autoColumns, NumericVector widths, float baseFontCharWidth, float minW, float maxW){
  
  size_t n = sheetData.size();
  
  IntegerVector v(n);
  CharacterVector r(n);
  int nLen;
  
  std::vector<std::string> tmp;
  //** tmp[2] is element "v" of sheetData
  
  // get widths of all values
  for(size_t i = 0; i < n; i++){
    
    tmp = as<std::vector<std::string> >(sheetData[i]);
    if(tmp[1] == "s"){
      v[i] = sharedStrings[atoi(tmp[2].c_str())].length() - 16; //-16 for shared string tags around text
    }else{
      nLen = tmp[2].length();
      v[i] = min(nLen, 11); // For numerics - max width is 11
      
    }
    
    r[i] = tmp[0];
      
  }
  

  // get column for each value
  IntegerVector colNumbers = RcppConvertFromExcelRef(r);
  IntegerVector colGroups = match(colNumbers, autoColumns);

  // reducing to only the columns that are auto
  LogicalVector notNA = !is_na(colGroups);
  colNumbers = colNumbers[notNA];
  v = v[notNA];
  widths = widths[notNA];
  IntegerVector uCols = sort_unique(colNumbers);
  
  size_t nUnique = uCols.size();
  NumericVector wTmp; // this will hold the widths for a specific column
  
  NumericVector columnWidths(nUnique);
  NumericVector thisColWidths;
  
  // for each unique column, get all widths for that column and take max
  for(size_t i = 0; i < nUnique; i++){
    wTmp = v[colNumbers == uCols[i]];
    thisColWidths = widths[colNumbers == uCols[i]];
    columnWidths[i] = max(wTmp * thisColWidths / baseFontCharWidth); 
  }
  
  columnWidths[columnWidths < minW] = minW;
  columnWidths[columnWidths > maxW] = maxW;    

  // assign column names
  columnWidths.attr("names") = uCols;
  
  return(wrap(columnWidths));
    
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
SEXP getVals(CharacterVector x){
  
  size_t n = x.size();
  std::string xml;
  CharacterVector v(n);
  size_t pos = 0;
  size_t endPos = 0;
  
  std::string vtag = "<v>";
  std::string vtagEnd = "</v>";
  
  for(size_t i = 0; i < n; i++){ 
    
    // find opening tag     
    xml = x[i];
    pos = xml.find(vtag, 0);
    
    if(pos == std::string::npos){
      v[i] = NA_STRING;
    }else{
      
      endPos = xml.find(vtagEnd, pos+3);
      v[i] = xml.substr(pos+3, endPos-pos-3).c_str();
      
    }
    
  }
  
  return wrap(v) ;  
  
}

// [[Rcpp::export]]
SEXP getNodes(std::string xml, std::string tagIn){
  
  if(xml.length() == 0)
  return wrap(NA_STRING);
  
  xml = " " + xml;
  std::vector<std::string> r;
  size_t pos = 0;
  size_t endPos = 0;
  std::string tag = tagIn;
  std::string tagEnd = tagIn.insert(1,"/");
  
  size_t k = tag.length();
  size_t l = tagEnd.length();
  
  while(1){
    
    pos = xml.find(tag, pos+1);
    endPos = xml.find(tagEnd, pos+k);
    
    if((pos == std::string::npos) | (endPos == std::string::npos))
    break;
    
    r.push_back(xml.substr(pos, endPos-pos+l).c_str());
    
  }  
  
  return wrap(r) ;  
  
}



// [[Rcpp::export]]
SEXP getChildlessNode(std::string xml, std::string tag){
  
  size_t k = tag.length();
  if(xml.length() == 0)
  return wrap(NA_STRING);
  
  xml = " " + xml;
  
  std::vector<std::string> r;
  size_t pos = 0;
  size_t endPos = 0;
  std::string tagEnd = "/>";
  
  while(1){
    
    pos = xml.find(tag, pos+1);    
    if(pos == std::string::npos)
    break;
    
    endPos = xml.find(tagEnd, pos+k);
    
    r.push_back(xml.substr(pos, endPos-pos+2).c_str());
    
  }
  
  return wrap(r) ;  
  
}



// [[Rcpp::export]]
SEXP getAttr(CharacterVector x, std::string tag){
  
  size_t n = x.size();
  size_t k = tag.length();
  
  if(n == 0)
  return wrap(-1);
  
  std::string xml;
  CharacterVector r(n);
  size_t pos = 0;
  size_t endPos = 0;
  std::string rtagEnd = "\"";
  
  for(size_t i = 0; i < n; i++){ 
    
    // find opening tag     
    xml = x[i];
    pos = xml.find(tag, 0);
    
    if(pos == std::string::npos){
      r[i] = NA_STRING;
    }else{  
      endPos = xml.find(rtagEnd, pos+k);
      r[i] = xml.substr(pos+k, endPos-pos-k).c_str();
    }
  }
  
  return wrap(r) ;  
  
}



// [[Rcpp::export]]
SEXP childrenCounter(CharacterVector x){
  
  size_t n = x.size();
  if(n == 0)
  return wrap(0);
  
  std::string xml;
  IntegerVector nChildren(n);
  size_t pos = 0;
  std::string tag = "<c ";
  
  for(size_t i = 0; i < n; i++){ 
    
    xml = x[i];
    nChildren[i] = 0;
    pos = xml.find(tag, 0);
    
    while(pos != std::string::npos){
      nChildren[i]++;
      pos = xml.find(tag, pos+1);
    }
    
  }
  
  return wrap(nChildren) ;  
  
}


// [[Rcpp::export]]
SEXP getCellStyles(CharacterVector x){
  
  size_t n = x.size();
  
  if(n == 0){
    CharacterVector ret(1);
    ret[0] = NA_STRING;
    return wrap(ret);
  }
  
  
  std::string xml;
  CharacterVector t(n);
  size_t pos = 0;
  size_t endPos = 0;
  
  std::string rtag = " s=";
  std::string rtagEnd = "\"";
  
  for(size_t i = 0; i < n; i++){ 
    
    xml = x[i];
    pos = xml.find(rtag, 1);
    
    if(pos == std::string::npos){
      t[i] = "n";
    }else{
      endPos = xml.find(rtagEnd, pos+4);
      t[i] = xml.substr(pos+4, endPos-pos-4).c_str();
    }
  }
  
  return wrap(t) ;  
  
}


// [[Rcpp::export]]
SEXP getCellTypes(CharacterVector x){
  
  int n = x.size();
  
  if(n == 0)
  return wrap(-1);
  
  std::string xml;
  CharacterVector t(n);
  size_t pos = 0;
  size_t endPos = 0;
  
  std::string rtag = " t=";
  std::string rtagEnd = "\"";
  
  for(int i = 0; i < n; i++){ 
    
    xml = x[i];
    pos = xml.find(rtag, 1);
    
    if(pos == std::string::npos){
      t[i] = "n";
    }else{
      endPos = xml.find(rtagEnd, pos+4);
      t[i] = xml.substr(pos+4, endPos-pos-4).c_str();
    }
  }
  
  return wrap(t) ;  
  
}


// [[Rcpp::export]]
SEXP getCells(CharacterVector x){
  
  int n = x.size();
  
  if(n == 0)
  return wrap(-1);
  
  std::string xml;
  CharacterVector r(n);
  size_t pos = 0;
  size_t endPos = 0;
  
  std::string tag = "<c ";
  std::string tagEnd = "</c>";
  std::string tagEnd2 = "/>";
  
  for(int i = 0; i < n; i++){ 
    
    // find opening tag     
    xml = x[i];
    pos = xml.find(tag, 0);
    endPos = xml.find(tagEnd, pos+3);
    
    if(endPos == std::string::npos)
    endPos = xml.find(tagEnd2, pos+3);
    
    r[i] = xml.substr(pos+3, endPos-pos-3).c_str();
    
  }
  
  return wrap(r) ;  
  
}


// [[Rcpp::export]]
SEXP getFunction(CharacterVector x){
  
  int n = x.size();
  
  if(n == 0)
  return wrap(-1);
  
  std::string xml;
  CharacterVector r(n);
  size_t pos = 0;
  size_t endPos = 0;
  
  std::string tag = "<f";
  std::string tagEnd1 = "</f>";
  std::string tagEnd2 = "/>";
  
  for(int i = 0; i < n; i++){ 
    
    // find opening tag     
    xml = x[i];
    pos = xml.find(tag, 0);
    
    if(pos == std::string::npos){
      r[i] = NA_STRING;
    }else{
      
      endPos = xml.find(tagEnd1, pos+3);
      if(endPos == std::string::npos){
        endPos = xml.find(tagEnd2, pos+3);
        r[i] = xml.substr(pos, endPos-pos+2).c_str();
      }else{
        r[i] = xml.substr(pos, endPos-pos+4).c_str();
      }
      
    }
  }
  
  return wrap(r) ;  
  
}


// [[Rcpp::export]]
SEXP getRefs(CharacterVector x, int startRow){
  
  int n = x.size();
  
  if(n == 0)
  return wrap(-1);
  
  std::string xml;
  CharacterVector r(n);
  size_t pos = 0;
  size_t endPos = 0;
  
  std::string rtag = "r=";
  std::string rtagEnd = "\"";
  
  if(startRow == 1){
    for(int i = 0; i < n; i++){ 
      
      // find opening tag     
      xml = x[i];
      pos = xml.find(rtag, 0);
      endPos = xml.find(rtagEnd, pos+3);
      r[i] = xml.substr(pos+3, endPos-pos-3).c_str();
      
    }
    
  }else{
    
    // if a startRow is specified
    // loop throught v, if row is less than startRow, remove from r and v
    int rowN;
    std::string row;
    bool below = true;
    
    for(int i = 0; i < n; i++){
      
      // find opening tag     
      xml = x[i];
      pos = xml.find(rtag, 0);
      endPos = xml.find(rtagEnd, pos+3);
      r[i] = xml.substr(pos+3, endPos-pos-3).c_str();
      
      if(below){
        row = r[i];
        // remove digits from string
        row.erase(std::remove_if(row.begin(), row.end(), ::isalpha), row.end());
        rowN = atoi(row.c_str());   
        if(rowN < startRow){
          r[i] = NA_STRING;
        }else{
          below = false;
        } 
      } 
    } // end of loop
    
    r = na_omit(r);
    
  } // end of else
  
  return wrap(r) ;  
  
}



// [[Rcpp::export]]
CharacterVector getSharedStrings(CharacterVector x){
  
  int n = x.size();
  std::string xml;
  
  size_t pos = 0;
  size_t endPos = 0;
  
  std::string ttag = "<t";
  std::string tag = ">";
  std::string tagEnd = "<";
  
  CharacterVector strs(n);
  std::fill(strs.begin(), strs.end(), NA_STRING);
  
  IntegerVector emptyStrs;
  
  for(int i = 0; i < n; i++){ 
    
    // find opening tag     
    xml = x[i];
    pos = xml.find(ttag, 0); // find ttag      
    
    if(pos != std::string::npos){
      
      if(xml[pos+2] == '/'){
        emptyStrs.push_back(i); //this should only ever grow to length 1
      }else{
        pos = xml.find(tag, pos+1); // find where opening ttag ends
        endPos = xml.find(tagEnd, pos+1); // find where the node ends </t> (closing tag)
        strs[i] = xml.substr(pos+1, endPos-pos - 1).c_str();
      }
    }
    
  }
  
  strs.attr("empty") = emptyStrs;
  
  return wrap(strs) ;  
  
}




// [[Rcpp::export]]
CharacterVector getSharedStrings2(CharacterVector x){
  
  int n = x.size();
  std::string xml;
  
  size_t pos = 0;
  size_t endPos = 0;
  
  std::string ttag = "<t";
  std::string tag = ">";
  std::string tagEnd = "<";
  
  CharacterVector strs(n);
  std::fill(strs.begin(), strs.end(), NA_STRING);
  
  IntegerVector emptyStrs;
  
  for(int i = 0; i < n; i++){ 
    
    // find opening tag     
    xml = x[i];
    pos = xml.find(ttag, 0); // find ttag      
    
    if(xml[pos+2] == '/'){
       emptyStrs.push_back(i); //this should only ever grow to length 1
    }else{
      strs[i] = "";
      while(1){

        if(xml[pos+2] == '/'){
           emptyStrs.push_back(i); //this should only ever grow to length 1
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
    
    

    
  }
  
  strs.attr("empty") = emptyStrs;
  
  return wrap(strs) ;  
  
}

// [[Rcpp::export]]
List getNumValues(CharacterVector inFile, int n, std::string tagIn) {
  
  std::string xmlFile = as<std::string>(inFile);
  std::string buf;
  std::string xml;
  ifstream file;
  file.open(xmlFile.c_str());
  
  std::vector<std::string> tokens;
  while (file >> buf)
  xml += buf;
  
  std::string tag = tagIn;
  std::string tagEnd = tagIn.insert(1,"/");
  
  size_t pos = 0;
  size_t endPos = 0;
  CharacterVector vc(n);
  NumericVector vn(n);
  
  for(int i = 0; i < n; i++){ 
    // find opening tag 
    pos = xml.find(tag, endPos+1);
    
    if (pos != std::string::npos){
      endPos = xml.find(tagEnd, pos+3);    
      vc[i] = xml.substr(pos+3, endPos-pos-3).c_str();
      vn[i] = atof(vc[i]);
    } 
  }
  
  List res(2);
  res[0] = vn;
  res[1] = vc;
  
  return res;
  
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
List buildCellList( CharacterVector r, CharacterVector t, CharacterVector v) {
  
  //Valid combinations
  //  r t v	
  //  T	F	F	
  //  T	T	T
  //  F F	F	
  
  int n = r.size();
  List cells(n);
  LogicalVector hasV = !is_na(v);
  LogicalVector hasR = !is_na(r);
  
  for(int i=0; i < n; i++){
    
    if(hasR[i]){
      
      if(hasV[i]){
        
        cells[i] = CharacterVector::create(
          Named("r") = r[i],
          Named("t") = t[i],
          Named("v") = v[i],
          Named("f") = NA_STRING); 
      }else{
        
        
        cells[i] = CharacterVector::create(
          Named("r") = r[i],
          Named("t") = NA_STRING,
          Named("v") = NA_STRING,
          Named("f") = NA_STRING); 
      }
      
    }else{
      
      cells[i] = CharacterVector::create(
        Named("r") = NA_STRING,
        Named("t") = NA_STRING,
        Named("v") = NA_STRING,
        Named("f") = NA_STRING);  
    }
    
  } // end of for loop
  
  return wrap(cells) ;
}


// [[Rcpp::export]]
SEXP buildLoadCellList( CharacterVector r, CharacterVector t, CharacterVector v, CharacterVector f) {
  
  //No cases with r is NA as only take cell with children
  
  //Valid combinations
  //  r t	v	f
  //  T	F	F	F done
  //  T	T	T	F done
  //  T	T	T	T done
  
  
  int n = r.size();
  List cells(n);
  LogicalVector hasV = !is_na(v);
  LogicalVector hasF = !is_na(f);
  
  for(int i=0; i < n; i++){
    
    //If we have a function    
    if(hasF[i]){
      
      cells[i] = CharacterVector::create(
        Named("r") = r[i],
        Named("t") = t[i],
        Named("v") = v[i],
        Named("f") = f[i]); 
        
        
    }else if(hasV[i]){
      
      cells[i] = CharacterVector::create(
        Named("r") = r[i],
        Named("t") = t[i],
        Named("v") = v[i],
        Named("f") = NA_STRING); 
        
    }else{ //only have s and r
    cells[i] = CharacterVector::create(
      Named("r") = r[i],
      Named("t") = NA_STRING,
      Named("v") = NA_STRING,
      Named("f") = NA_STRING);
    }
  } // end of for loop
  
  return cells ;
}







// [[Rcpp::export]]
SEXP constructCellData(IntegerVector cols, std::vector<std::string> LETTERS, std::vector<std::string> rows, CharacterVector t, CharacterVector v) {
  
  //All cells here will have data in them: valid r, valid t and valid v 
  
  int nCols = cols.size();  
  int nRows = rows.size();
  int n = nCols*nRows;
  
  List cells(n);
  std::vector<std::string> res(n);
  
  int x;
  int modulo;
  for(int i = 0; i < nCols; i++){
    x = cols[i];
    string columnName;
    
    while(x > 0){  
      modulo = (x - 1) % 26;
      columnName = LETTERS[modulo] + columnName;
      x = (x - modulo) / 26;
    }
    res[i] = columnName;
  }
  
  CharacterVector names(n);
  int c = 0;
  for(int i=0; i < nRows; i++){
    for(int j=0; j < nCols; j++){
      
      names[c] = rows[i];
      cells[c] = CharacterVector::create(
        Named("r") = res[j] + rows[i],
        Named("t") = t[c],
        Named("v") = v[c]);
        c++; 
    }
  }
  
  cells.attr("names") = names;
  
  return cells ;
}




// [[Rcpp::export]]
SEXP convert2ExcelRef(IntegerVector cols, std::vector<std::string> LETTERS){
  
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
SEXP ExcelConvertExpand(IntegerVector cols, std::vector<std::string> LETTERS, std::vector<std::string> rows){
  
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
SEXP buildMatrixNumeric(NumericVector v, IntegerVector rowInd, IntegerVector colInd,
CharacterVector colNames, int nRows, int nCols){
  
  int k = v.size();
  NumericMatrix m(nRows, nCols);
  std::fill(m.begin(), m.end(), NA_REAL);
  
  for(int i = 0; i < k; i++)
    m(rowInd[i], colInd[i]) = v[i];
  
  List dfList(nCols);
  for(int i=0; i < nCols; ++i)
    dfList[i] = m(_,i);
  
  std::vector<int> rowNames(nRows);
  for(int i = 0;i < nRows; ++i)
    rowNames[i] = i+1;
  
  dfList.attr("names") = colNames;
  dfList.attr("row.names") = rowNames;
  dfList.attr("class") = "data.frame";
  
  return Rcpp::wrap(dfList);
  
  
}


// [[Rcpp::export]]
SEXP buildMatrixMixed(CharacterVector v,
                      NumericVector vn,
                      IntegerVector rowInd,
                      IntegerVector colInd,
                      CharacterVector colNames,
                      int nRows,
                      int nCols,
                      IntegerVector charCols,
                      int originAdj){
  
  int k = v.size();
  CharacterMatrix m(nRows, nCols);
  std::fill(m.begin(), m.end(), NA_STRING);
  
  NumericMatrix mN(nRows, nCols);
  std::fill(mN.begin(), mN.end(), NA_REAL);
  
  // fill matrix
  for(int i = 0;i < k; i++){
    m(rowInd[i], colInd[i]) = v[i];
    mN(rowInd[i], colInd[i]) = vn[i];
  }
  
  List dfList(nCols); 

  //vectors that are assigned to data.frame
  
  for(int i = 0; i < nCols; i++){

    
    CharacterVector tmp(nRows);
  
    for(int ri = 0; ri < nRows; ri++)
      tmp[ri] = m(ri,i);
    
    LogicalVector notNAElements = !is_na(tmp);
    
    // if i in charCols
    if (std::find(charCols.begin(), charCols.end(), i) != charCols.end()){ // If column is character class
    
      // determin if column is logical or date
      bool logCol = true;
      for(int ri = 0; ri < nRows; ri++){
        if(notNAElements[ri]){
          if((m(ri, i) != "TRUE") & (m(ri, i) != "FALSE")){
            logCol = false;
            break;
          }
        }
      }
      
      
      bool dateCol = true;
      for(int ri = 0; ri < nRows; ri++){
        if(notNAElements[ri]){
          if(m(ri, i) != "openxlsxdt"){
            dateCol = false;
            break;
          }
        }
      }
      
      
      if(logCol){
        
        LogicalVector logtmp(nRows);
        for(int ri=0; ri < nRows; ri++){
          if(!notNAElements[ri]){
            logtmp[ri] = NA_LOGICAL; //IF TRUE, TRUE else FALSE
          }else{
            logtmp[ri] = (tmp[ri] == "TRUE");
          }
        }
        
        dfList[i] = logtmp;
        
      }else if(dateCol){
              
        DateVector datetmp(nRows);
        for(int ri=0; ri < nRows; ri++){
          if(!notNAElements[ri]){
            datetmp[ri] = NA_REAL; //IF TRUE, TRUE else FALSE
          }else{
            datetmp[ri] = Date(mN(ri,i) - originAdj);
          }
        }
        
        dfList[i] = datetmp;
          
      
      }else{
        
        dfList[i] = tmp;
        
      }
    
    }else{ // else if column NOT character class
    
    NumericVector ntmp(nRows);
    for(int ri = 0; ri < nRows; ri++){
      if(notNAElements[ri]){
        ntmp[ri] = mN(ri, i); 
      }else{
        ntmp[ri] = NA_REAL; 
      }
    }
    
    dfList[i] = ntmp;
    
    }
    
  }
  
  std::vector<int> rowNames(nRows);
  for(int i = 0;i < nRows; ++i)
  rowNames[i] = i+1;
  
  dfList.attr("names") = colNames;
  dfList.attr("row.names") = rowNames;
  dfList.attr("class") = "data.frame";
  
  return wrap(dfList);
  
}


// [[Rcpp::export]]
IntegerVector matrixRowInds(IntegerVector indices) {
  
  int n = indices.size();
  LogicalVector notDup = !duplicated(indices);
  IntegerVector res(n);
  
  int j = -1;
  for(int i =0; i < n; i ++){
    if(notDup[i])
    j++;
    res[i] = j;
  }
  
  return wrap(res);
  
}


// [[Rcpp::export]]
List buildCellMerges(List comps){
  
  size_t nMerges = comps.size(); 
  List res(nMerges);
  
  for(size_t i =0; i < nMerges; i++){
    IntegerVector col = RcppConvertFromExcelRef(comps[i]);  
    CharacterVector comp = comps[i];
    IntegerVector row(2);  
    
    for(size_t j = 0; j < 2; j++){
      std::string rt(comp[j]);      
      rt.erase(std::remove_if(rt.begin(), rt.end(), ::isalpha), rt.end());
      row[j] = atoi(rt.c_str());
    }
    
    size_t ca(col[0]);
    size_t ck = size_t(col[1]) - ca + 1;
    
    std::vector<int> v(ck) ;
    for(size_t j = 0; j < ck; j++)
      v[j] = j + ca;
    
    size_t ra(row[0]);
    
    size_t rk = int(row[1]) - ra + 1;
    std::vector<int> r(rk) ;
    for(size_t j = 0; j < rk; j++)
      r[j] = j + ra;
    
    CharacterVector M(ck*rk);
    int ind = 0;
    for(size_t j = 0; j < ck; j++){
      for(size_t k = 0; k < rk; k++){
        char name[30];
        sprintf(&(name[0]), "%d-%d", r[k], v[j]);
        M(ind) = name;
        ind++;
      }
    }
    
    res[i] = M;
  }
  
  return wrap(res) ;
  
}


// [[Rcpp::export]]
SEXP readWorkbook(CharacterVector v,
                  NumericVector vn,
                  IntegerVector stringInds,
                  CharacterVector r,
                  CharacterVector tR,
                  int nRows,
                  bool hasColNames,
                  bool skipEmptyRows,
                  int originAdj
                  ){
  
  // Convert r to column number and shift to scale 0:nCols  (from eg. J:AA in the worksheet)
  int nCells = r.size();
  IntegerVector colNumbers = RcppConvertFromExcelRef(r); 
  colNumbers = match(colNumbers, sort_unique(colNumbers)) - 1;
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
    
    // remove elements in tR that appear in the first row (if there are numerics in first row this will remove less elements than pos)
    CharacterVector::iterator first = r.begin();
    CharacterVector::iterator last = r.begin() + pos;
    CharacterVector toRemove(first, last);
    removeFlag = match(tR, toRemove);
    tR.erase(tR.begin(), tR.begin() + sum(!is_na(removeFlag)));

    // remove elements that are now being used as colNames (there are int pos many of these)    
    //check we have some r left if not return a data.frame with zero rows
    r.erase(r.begin(), r.begin() +  pos); 
    if(r.size() == 0){
      List dfList(nCols);
      IntegerVector rowNames(0);
      
      dfList.attr("names") = colNames;
      dfList.attr("row.names") = rowNames;
      dfList.attr("class") = "data.frame";
      return wrap(dfList);
    }
    
    colNumbers.erase(colNumbers.begin(), colNumbers.begin() + pos); 
    nRows--; // decrement number of rows as first row is now being used as colNames
    nCells = nCells - pos;
        
  }else{
    char name[6];
    for(int i =0; i < nCols; i++){
      sprintf(&(name[0]), "X%d", i+1);
      colNames[i] = name;
    }
  }
  
  bool allNumeric = false;
  
  if((tR.size() == 0) | (stringInds[0] == -1)) //If the new resized tR is length 0 there are no more strings
    allNumeric = true;
    
  // getRow numbers from r 
  IntegerVector rowNumbers(nCells);  
  IntegerVector uRows;
  if(nCells == nRows*nCols){
    
    uRows = seq(0, nRows-1);
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
    
    uRows = unique(rowNumbers);
  }
  
  if(hasColNames) // remove elements that have been "used"
   vn.erase(vn.begin(), vn.begin() + pos);
      
  SEXP m;
  if(allNumeric){
    
    m = buildMatrixNumeric(vn, rowNumbers, colNumbers, colNames, nRows, nCols);
    
  }else{
    
    if(hasColNames)// remove elements that have been "used" (we have done this for r already)
      v.erase(v.begin(), v.begin() + pos);
    
    // If it contains any strings it will be a character column
    IntegerVector charCols = match(tR, r);    
    
    int tRSize = tR.size();
    IntegerVector charColNumbers(tRSize);

    for(int i = 0; i < tRSize; i++)
      charColNumbers[i] = colNumbers[charCols[i]-1];
    
    charCols = unique(charColNumbers);

    m = buildMatrixMixed(v, vn, rowNumbers, colNumbers, colNames, nRows, nCols, charCols, originAdj);
    
   }
  
  return wrap(m) ;
  
}






// [[Rcpp::export]]
SEXP quickBuildCellXML(std::string prior, std::string post, List sheetData, IntegerVector rowNumbers, CharacterVector styleInds, std::string R_fileName){
  
  // open file and write header XML
  const char * s = R_fileName.c_str();
  std::ofstream xmlFile;
  xmlFile.open (s);
  xmlFile << "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>";
  
  // If no data write childless node and return
  if(sheetData.size() == 0){
    xmlFile << prior;
    xmlFile << "<sheetData/>";
    xmlFile << post;
    xmlFile.close();
    return Rcpp::wrap(1);
  }
  
  // write xml prior to stringData and opening tag
  xmlFile << prior;
  xmlFile << "<sheetData>";
  
  
  //NOTES ON WHY THIS WORKS
  //If no row heights every row has children
  //If data only added once, sheetData will be in order
  // Thus all rows have children (all starting row tags are open) and no need to sort/find
  
  CharacterVector rows = sheetData.attr("names");    
  CharacterVector uniqueRows(sort_unique(rowNumbers));

  size_t n = rows.size();
  size_t k = uniqueRows.size();
  std::string xml;
  
  LogicalVector hasStyle = !is_na(styleInds);
  CharacterVector cVal;
  
  // write sheetData
  size_t j = 0;
  String currentRow = uniqueRows[0];
  for(size_t i = 0; i < k; i++){
    
    std::string cellXML;
    
    while(currentRow == uniqueRows[i]){
      
      cVal = sheetData[j];
      LogicalVector attrs = !is_na(cVal);
      
      //cell XML strings
      cellXML += "<c r=\"" + cVal[0];
      
      if(hasStyle[j])
        cellXML += "\" s=\""+ styleInds[j];
      
      //If we have a t value we must have a v value
      if(attrs[1]){
        
        //If we have a c value we might have an f value
        if(attrs[3]){
          cellXML += "\" t=\"" + cVal[1] + "\">" + cVal[3] + "<v>" + cVal[2] + "</v></c>";
        }else{
          cellXML += "\" t=\"" + cVal[1] + "\"><v>" + cVal[2] + "</v></c>";
        }
        
      }else{
        cellXML += "\"/>";
      }
      
      j += 1;
      
      if(j == n)
        break;
      
      currentRow = rows[j];
    }
    
    xmlFile << "<row r=\"" + uniqueRows[i] + "\">" + cellXML + "</row>";
    
  }
  
  // write closing tag and XML post sheetData
  xmlFile << "</sheetData>";
  xmlFile << post;
  
  //close file
  xmlFile.close();
  
  return Rcpp::wrap(1);
  
}


// [[Rcpp::export]]
CharacterVector buildTableXML(std::string table, std::string ref, std::vector<std::string> colNames, bool showColNames, std::string tableStyle, bool withFilter){
  
  int n = colNames.size();
  std::string tableCols;
  std::string tableStyleXML = "<tableStyleInfo name=\"" + tableStyle + "\" showFirstColumn=\"0\" showLastColumn=\"0\" showRowStripes=\"1\" showColumnStripes=\"0\"/>";
  table += " totalsRowShown=\"0\">";
  
  if(withFilter)
    table += "<autoFilter ref=\"" + ref + "\"/>";
  
  
  for(int i = 0; i < n; i ++){
    tableCols += "<tableColumn id=\"" + itos(i+1) + "\" name=\"" + colNames[i] + "\"/>";
  }
  
  tableCols = "<tableColumns count=\"" + itos(n) + "\">" + tableCols + "</tableColumns>"; 
  
  table = table + tableCols + tableStyleXML + "</table>";
  
  
  return wrap(table);
  
}



// [[Rcpp::export]]   
List uniqueCellAppend(List sheetData, CharacterVector r, List newCells){
  
  int n = sheetData.size();
  int cn = newCells.size();
  
  // get all r values
  CharacterVector exCells(n);
  CharacterVector tmp;
  for(int i =0; i < n; i ++){
    tmp = sheetData[i];
    exCells[i] = tmp[0];
  }
  
  // find which exCells we keep
  IntegerVector overwrite = match(exCells, r);
  LogicalVector toKeep = is_na(overwrite);
  int k = sum(toKeep);
  
  //If no cells to keep
  if(k == 0){
    newCells.attr("overwrite") = true;
    return wrap(newCells); 
  }
  
  //If keeping all cells (no cells to be overwritten)
  List newSheetData(k + cn);
  CharacterVector exNames = sheetData.attr("names");
  CharacterVector newCellNames = newCells.attr("names");
  CharacterVector newNames(k + cn);
  
  // create new List of toKeep + newCells
  int j = 0;
  for(int i = 0; i < n; i ++){
    
    if(toKeep[i]){
      newSheetData[j] = sheetData[i];
      newNames[j] = exNames[i];
      j++;
    }
    
    if(j == k)
    break;
  }
  
  for(int i = 0; i < cn; i++){
    newSheetData[j] = newCells[i];
    newNames[j] = newCellNames[i];
    j++;
  }
  
  newSheetData.attr("names") = newNames;
  if(k < n){
    newSheetData.attr("overwrite") = true;
  }else{
    newSheetData.attr("overwrite") = false;
  }
  
  return wrap(newSheetData) ;
  
}



// [[Rcpp::export]]
SEXP getHyperlinkRefs(CharacterVector x){
  
  int n = x.size();
  
  std::string xml;
  CharacterVector r(n);
  size_t pos = 0;
  size_t endPos = 0;
  
  std::string rtag = "ref=";
  std::string rtagEnd = "\"";
  
  for(int i = 0; i < n; i++){ 
    
    // find opening tag     
    xml = x[i];
    pos = xml.find(rtag, 0);
    endPos = xml.find(rtagEnd, pos+5);
    r[i] = xml.substr(pos+5, endPos-pos-5).c_str();
    
  }
  
  return wrap(r) ;  
  
}




// [[Rcpp::export]]   
List writeCellStyles(List sheetData, CharacterVector rows, IntegerVector cols, String styleId, std::vector<std::string> LETTERS){
  
  
  
  int nStyleCells = rows.size();
  int n = sheetData.size();
  CharacterVector exCellNames = "";
  
  // cell names
  if(n > 0)
    exCellNames = sheetData.attr("names");
  
  //new cell names will be elements of rows where styleCell doesn't exist  
    
  // create cellRefs from rows & cols
  CharacterVector cellRefsToBeStyled(nStyleCells);
  CharacterVector colNames = convert2ExcelRef(cols, LETTERS);
  
  std::string r;
  std::string c;
  
  for(int i = 0; i < nStyleCells; i++){
    r = rows[i];
    c = colNames[i];
    cellRefsToBeStyled[i] = c + r; 
  }

  // get refs of existing cells
  CharacterVector exCells(n);
  CharacterVector tmp;
  for(int i = 0; i < n; i++){
    tmp = sheetData[i];
    exCells[i] = tmp[0];
  }
  
  // get all existing cells that need to be styled with this ID
  IntegerVector exPos = match(exCells, cellRefsToBeStyled);
  LogicalVector toApply = !is_na(exPos);
    
  IntegerVector stylePos = match(cellRefsToBeStyled, exCells);
  LogicalVector isNewCell = is_na(stylePos);
  int nNewCells = sum(isNewCell); 
  
  //List for new sheetData
  List newSheetData(n + nNewCells);
  CharacterVector newSheetDataNames(n + nNewCells);
  
  for(int i = 0; i < n; i++){
    if(toApply[i]){
      tmp = sheetData[i];
      tmp[2] = styleId;
      newSheetData[i] = tmp;
    }else{
      newSheetData[i] = sheetData[i];
    }
    newSheetDataNames[i] = exCellNames[i];
  }
  
  //return(newSheetData);
  
  int j = n;
  // append new cells with styleId
  for(int i = 0; i < nStyleCells; i++){
    
    if(isNewCell[i]){
      newSheetData[j] = CharacterVector::create(
        Named("r") = cellRefsToBeStyled[i],
        Named("t") = NA_STRING,
        Named("s") = styleId,
        Named("v") = NA_STRING,
        Named("f") = NA_STRING);
        
        newSheetDataNames[j] = rows[i];
        j++;
    }
    
  }
  
  newSheetData.attr("names") = newSheetDataNames;
  
  return wrap(newSheetData);
  
}


// [[Rcpp::export]]  
SEXP calcNRows(CharacterVector x, bool skipEmptyRows){
  
  int n = x.size();
  if(n == 0)
    return(wrap(0));
  
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
  
  return wrap(nRows);
  
}



// [[Rcpp::export]]  
CharacterVector buildCellTypes(CharacterVector classes, int nRows){
  
  
  int nCols = classes.size();
  CharacterVector colLabels(nCols);
  for(int i=0; i < nCols; i++){
    
    if((classes[i] == "numeric") | (classes[i] == "integer") | (classes[i] == "raw") ){
      colLabels[i] = "n"; 
    }else if(classes[i] == "character"){
      colLabels[i] = "s"; 
    }else if(classes[i] == "logical"){
      colLabels[i] = "b";
    }else if(classes[i] == "hyperlink"){
      colLabels[i] = "h";
    }else{
      colLabels[i] = "s";
    }
    
  }
  
  CharacterVector cellTypes = rep(colLabels, nRows); 
  
  return wrap(cellTypes);
  
  
}



// [[Rcpp::export]]
CharacterVector removeEmptyNodes(CharacterVector x, CharacterVector emptyNodes){
  
  int n = x.size();
  int nEmpty = emptyNodes.size();
  
  std::string cell;
  std::string emptyNode;
  
  
  for(int e = 0; e < nEmpty; e++){
    
    emptyNode = emptyNodes[e];
    
    for(int i =0; i<n;i++){
      cell = x[i];
      
      if(cell.find(emptyNode, 0) != std::string::npos){
        if(cell.find("t=\"s\"", 0) != std::string::npos){
          x[i] = NA_STRING;
        }
      }
    }   
  }
  
  x = na_omit(x);
  
  return wrap(x);
  
}



// [[Rcpp::export]]
CharacterVector getCellsWithChildren(std::string xmlFile, CharacterVector emptyNodes){
  
  //read in file without spaces
  std::string xml = cppReadFile(xmlFile);
  
  // std::string tag = "<c ";
  //  std::string tagEnd1 = ">";
  //  std::string tagEnd2 = "</c>";
  
  // count cells with children
  int occurrences = 0;
  string::size_type start = 0;
  while((start = xml.find("</v>", start)) != string::npos) {
    ++occurrences;
    start += 4;
  }
  
  CharacterVector cells(occurrences);
  std::fill(cells.begin(), cells.end(), NA_STRING);
  
  int i = 0;
  size_t nextPos = 3;
  size_t vPos = 2;
  std::string sub;
  size_t pos = xml.find("<c ", 0);
  
  while(i < occurrences){
    
    if(pos != std::string::npos){
      nextPos = xml.find("<c ", pos+9);
      vPos = xml.find("</v>", pos+8); // have to atleast pass <c r="XX">
      
      if(vPos < nextPos){
        cells[i] = xml.substr(pos, nextPos-pos);
        i++; 
      }
      
      pos = nextPos;
      
    }
  }
  
  
  if(emptyNodes[0] != "<v></v>")
  cells = removeEmptyNodes(cells, emptyNodes);
  
  return wrap(cells) ;  
  
}





// [[Rcpp::export]]
SEXP quickBuildCellXML2(std::string prior, std::string post, List sheetData, IntegerVector rowNumbers, CharacterVector styleInds, CharacterVector rowHeights, std::string R_fileName){
  
  // open file and write header XML
  const char * s = R_fileName.c_str();
  std::ofstream xmlFile;
  xmlFile.open (s);
  xmlFile << "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>";
  
  // If no data write childless node and return
  if(sheetData.size() == 0){
    xmlFile << prior;
    xmlFile << "<sheetData/>";
    xmlFile << post;
    xmlFile.close();
    return Rcpp::wrap(1);
  }
  
  // write xml prior to stringData and opening tag
  xmlFile << prior;
  xmlFile << "<sheetData>";
  
  // sheetData will be in order, jsut need to check for rowHeights
  
  CharacterVector rows = sheetData.attr("names");    
  CharacterVector uniqueRows(sort_unique(rowNumbers));

  CharacterVector rowHeightRows = rowHeights.attr("names");
  size_t nRowHeights = rowHeights.size();
  
  size_t n = rows.size();
  size_t k = uniqueRows.size();
  std::string xml;
  
  size_t j = 0;
  size_t h = 0;
  String currentRow = uniqueRows[0];
  
  LogicalVector hasStyle = !is_na(styleInds);
  CharacterVector cVal;
  
  for(size_t i = 0; i < k; i++){
    
    std::string cellXML;
    bool hasRowData = true;
    
    while(currentRow == uniqueRows[i]){
      
      cVal = sheetData[j];
      LogicalVector attrs = !is_na(cVal);
      hasRowData = true;
      
      j += 1;
      if(!attrs[0]){ //If r IS NA we have no row data we only have a rowHeight
        hasRowData = false;
        
        if(j == n)
         break;
              
        currentRow = rows[j];
        break;
      }
      
      //cell XML strings
      cellXML += "<c r=\"" + cVal[0];
      
      if(hasStyle[j-1]) // Style is j-1 because sheet data is set as j before j is incremented
        cellXML += "\" s=\"" + styleInds[j-1];
      
      //If we have a t value we must have a v value
      if(attrs[1]){
        //If we have a c value we might have an f value
        if(attrs[3]){
          cellXML += "\" t=\"" + cVal[1] + "\">" + cVal[3] + "<v>" + cVal[2] + "</v></c>";
        }else{
          cellXML += "\" t=\"" + cVal[1] + "\"><v>" + cVal[2] + "</v></c>";
        }
        
      }else{
        cellXML += "\"/>";
      }
      
      if(j == n)
        break;
      
      currentRow = rows[j];
      
    }
    
    
    if(h < nRowHeights){
      
      if((uniqueRows[i] == rowHeightRows[h]) & hasRowData){ // this row has a row height and cellXML data
        
        xmlFile << "<row r=\"" + uniqueRows[i] + "\" ht=\"" + rowHeights[h] + "\" customHeight=\"1\">" + cellXML + "</row>";   
        h++;
        
      }else if(hasRowData){

        xmlFile << "<row r=\"" + uniqueRows[i] + "\">" + cellXML + "</row>";
      
      }else{

        xmlFile << "<row r=\"" + uniqueRows[i] + "\" ht=\"" + rowHeights[h] + "\" customHeight=\"1\"/>"; 
        h++;
      }
      
    }else{
      
      xmlFile << "<row r=\"" + uniqueRows[i] + "\">" + cellXML + "</row>";
      
    }
        
  }
  
  // write closing tag and XML post sheetData
  xmlFile << "</sheetData>";
  xmlFile << post;
  
  //close file
  xmlFile.close();
  
  return wrap(1);
  
}






// [[Rcpp::export]]
SEXP getRefsVals(CharacterVector x, int startRow){
  
  int n = x.size();
  std::string xml;
  CharacterVector r(n);
  CharacterVector v(n);
  size_t pos = 0;
  size_t endPos = 0;
  
  std::string rtag = "r=";
  std::string rtagEnd = "\"";
  
  std::string vtag = "<v>";
  std::string vtag2 = "<v ";
  std::string vtagEnd = "</v>";
  
  if(startRow == 1){
    for(int i = 0; i < n; i++){ 
      
      // find r tag     
      xml = x[i];
      pos = xml.find(rtag, 0);  // find r=
      endPos = xml.find(rtagEnd, pos + 3);  // find r= "
      r[i] = xml.substr(pos + 3, endPos - pos - 3).c_str();
      
      // find <v> tag and </v> end tag
      pos = xml.find(vtag, endPos+1);
      
      if(pos != std::string::npos){
        endPos = xml.find(vtagEnd, pos + 3);
        v[i] = xml.substr(pos + 3, endPos - pos - 3).c_str();
      }else{
        pos = xml.find(vtag2, endPos + 1);
        pos = xml.find(">", pos + 1);
        endPos = xml.find(vtagEnd, pos + 3);
        v[i] = xml.substr(pos + 1, endPos - pos - 3).c_str();
      }
    }
    
  }else{
    // if a startRow is specified
    // loop throught v, if row is less than startRow, remove from r and v
    int rowN;
    std::string row;
    bool below = true;
    
    for(int i = 0; i < n; i++){
      
      // find opening tag     
      xml = x[i];
      pos = xml.find(rtag, 0);
      endPos = xml.find(rtagEnd, pos+3);
      r[i] = xml.substr(pos+3, endPos-pos-3).c_str();
      
      if(below){
        row = r[i];
        // remove digits from string
        row.erase(std::remove_if(row.begin(), row.end(), ::isalpha), row.end());
        rowN = atoi(row.c_str());
        
        if(rowN < startRow){
          r[i] = NA_STRING;
          v[i] = NA_STRING;
          
        }else{
          below = false;
      // find <v> tag and </v> end tag
          pos = xml.find(vtag, endPos+1);
          
          if(pos != std::string::npos){
            endPos = xml.find(vtagEnd, pos + 3);
            v[i] = xml.substr(pos + 3, endPos - pos - 3).c_str();
          }else{
            pos = xml.find(vtag2, endPos + 1);
            pos = xml.find(">", pos + 1);
            endPos = xml.find(vtagEnd, pos + 3);
            v[i] = xml.substr(pos + 1, endPos - pos - 3).c_str();
          }
        }
        
      }else{
      // find <v> tag and </v> end tag
      pos = xml.find(vtag, endPos+1);
      
        if(pos != std::string::npos){
          endPos = xml.find(vtagEnd, pos + 3);
          v[i] = xml.substr(pos + 3, endPos - pos - 3).c_str();
        }else{
          pos = xml.find(vtag2, endPos + 1);
          pos = xml.find(">", pos + 1);
          endPos = xml.find(vtagEnd, pos + 3);
          v[i] = xml.substr(pos + 1, endPos - pos - 3).c_str();
      }       
      }
      
    } // end of loop
    
    v = na_omit(v);
    r = na_omit(r);
    
  } // end of else
  
  
  
  List res(2);
  res[0] = r;
  res[1] = v;
  
  return wrap(res) ;  
  
}






// [[Rcpp::export]]
std::string createAlignmentNode(List style){
  
  std::vector<std::string> nms = style.attr("names");
  std::string alignNode = "<alignment";
  
  // textRotation
  if(std::find(nms.begin(), nms.end(), "textRotation") != nms.end())
    alignNode += " textRotation=\""+ as<std::string>(style["textRotation"]) + "\"";
  
  // horizontal
  if(std::find(nms.begin(), nms.end(), "halign") != nms.end())
    alignNode += " horizontal=\""+ as<std::string>(style["halign"]) + "\"";
  
  // vertical
  if(std::find(nms.begin(), nms.end(), "valign") != nms.end())
    alignNode += " vertical=\""+ as<std::string>(style["valign"]) + "\"";
  
  if(std::find(nms.begin(), nms.end(), "wrapText") != nms.end())
    alignNode += " wrapText=\"1\"";
  
  alignNode += "/>";
  
  return alignNode;
  
}



// [[Rcpp::export]]
std::string createFillNode(List style){
  
  std::vector<std::string> nms = style.attr("names");
  std::string fillNode = "<fill><patternFill patternType=\"solid\">";
  List tmp;
  List nm;
    
  // foreground
  if(std::find(nms.begin(), nms.end(), "fillFg") != nms.end()){
    tmp = style["fillFg"];
    nm = tmp.attr("names");
    int n = tmp.size();
    fillNode += "<fgColor";
    
    for(int i = 0; i < n; i ++)
      fillNode +=  " " +as<std::string>(nm[i]) +"=\"" +  as<std::string>(tmp[i]) + "\"";
    
    fillNode += "/>";
  }
  
  // background
  if(std::find(nms.begin(), nms.end(), "fillBg") != nms.end()){
    tmp = style["fillBg"];
    nm = tmp.attr("names");
    int n = tmp.size();
    fillNode += "<bgColor";
    
    for(int i = 0; i < n; i ++)
      fillNode +=  " " + as<std::string>(nm[i]) +"=\"" +  as<std::string>(tmp[i]) + "\"";
    
    fillNode += "/>";
  }
  
  fillNode += "</patternFill></fill>";
  
  return fillNode;
  
}







// [[Rcpp::export]]
std::string createFontNode(List style, std::string defaultFontSize, std::string defaultFontColour, std::string defaultFontName){
  
  std::vector<std::string> nms = style.attr("names");
  std::string fontNode = "<font>";
  List tmp;
  std::string nm;
    
  // size
  if(std::find(nms.begin(), nms.end(), "fontSize") != nms.end()){
    tmp = style["fontSize"];
    nm = as<std::string>(tmp.attr("names"));
    fontNode +=  "<sz "+ nm +"=\"" +  as<std::string>(tmp[0]) + "\"/>";
  }else{
    fontNode += defaultFontSize;
  }
  
  // colour
  if(std::find(nms.begin(), nms.end(), "fontColour") != nms.end()){
    tmp = style["fontColour"];
    nm = as<std::string>(tmp.attr("names"));
    fontNode +=  "<color "+ nm +"=\"" +  as<std::string>(tmp[0]) + "\"/>";
  }else{
    fontNode += defaultFontColour;
  }
  
  // name
  if(std::find(nms.begin(), nms.end(), "fontName") != nms.end()){
    tmp = style["fontName"];
    nm = as<std::string>(tmp.attr("names"));
    fontNode +=  "<name "+ nm +"=\"" +  as<std::string>(tmp[0]) + "\"/>";
  }else{
    fontNode += defaultFontName;
  }
  
  
  // fontFamily
  if(std::find(nms.begin(), nms.end(), "fontFamily") != nms.end()){
   tmp = style["fontFamily"];
   fontNode +=  "<family val=\"" +  as<std::string>(tmp[0]) + "\"/>";
  }
  
  // font scheme
  if(std::find(nms.begin(), nms.end(), "fontScheme") != nms.end()){
    tmp = style["fontScheme"];
    fontNode +=  "<scheme val=\"" +  as<std::string>(tmp[0]) + "\"/>";
  }
  
  // Font decorations
  if(std::find(nms.begin(), nms.end(), "fontDecoration") != nms.end()){
    
    tmp = style["fontDecoration"];
    int n = tmp.size();
    string d;
    
    for(int i = 0;i < n; i++){
       
      d = as<std::string>(tmp[i]);
      if(d == "BOLD")
        fontNode += "<b/>";
        
      if(d == "ITALIC")
        fontNode += "<i/>";
        
      if(d == "UNDERLINE")
        fontNode += "<u val=\"single\"/>";;
        
      if(d == "UNDERLINE2")
        fontNode += "<u val=\"double\"/>";;
        
      if(d == "STRIKEOUT")
        fontNode += "<strike/>";
      
    }
  }
  
  // Create new font and return Id  

  fontNode += "</font>";
    
  return fontNode;
  
}




// [[Rcpp::export]]
std::string createBorderNode(List style, CharacterVector borders){
  
  std::vector<std::string> nms = style.attr("names");
  std::string borderNode = "<border>";
  List tmp;
  std::string colourName;
  
  if(std::find(nms.begin(), nms.end(), "borderLeft") != nms.end()){
    tmp = style["borderLeftColour"];
    colourName = as<std::string>(tmp.attr("names"));
    borderNode +=  "<left style=\"" + as<std::string>(style["borderLeft"]) + "\"><color " + colourName + "=\"" +  as<std::string>(tmp[0]) + "\"/></left>";
  }    
  
  if(std::find(nms.begin(), nms.end(), "borderRight") != nms.end()){
    tmp = style["borderRightColour"];
    colourName = as<std::string>(tmp.attr("names"));
    borderNode +=  "<right style=\"" + as<std::string>(style["borderRight"]) + "\"><color " + colourName + "=\"" +  as<std::string>(tmp[0]) + "\"/></right>";
  }  
  
  if(std::find(nms.begin(), nms.end(), "borderTop") != nms.end()){
    tmp = style["borderTopColour"];
    colourName = as<std::string>(tmp.attr("names"));
    borderNode +=  "<top style=\"" + as<std::string>(style["borderTop"]) + "\"><color " + colourName + "=\"" +  as<std::string>(tmp[0]) + "\"/></top>";
  }  
  
  if(std::find(nms.begin(), nms.end(), "borderBottom") != nms.end()){
    tmp = style["borderBottomColour"];
    colourName = as<std::string>(tmp.attr("names"));
    borderNode +=  "<bottom style=\"" + as<std::string>(style["borderBottom"]) + "\"><color " + colourName + "=\"" +  as<std::string>(tmp[0]) + "\"/></bottom>";
  }  
    
  borderNode += "</border>";
    
  return borderNode;
  
}






// [[Rcpp::export]]
SEXP getCellStylesPossiblyMissing(CharacterVector x){
  
  size_t n = x.size();
  
  if(n == 0){
    CharacterVector ret(1);
    ret[0] = NA_STRING;
    return wrap(ret);
  }
  
  
  std::string xml;
  CharacterVector t(n);
  size_t pos = 0;
  size_t endPos = 0;
  
  std::string rtag = " s=";
  std::string rtagEnd = "\"";
  
  for(size_t i = 0; i < n; i++){ 
    
    xml = x[i];
    pos = xml.find(rtag, 1);
    
    if(pos == std::string::npos){
      t[i] = NA_STRING;
    }else{
      endPos = xml.find(rtagEnd, pos+4);
      t[i] = xml.substr(pos+4, endPos-pos-4).c_str();
    }
  }
  
  return wrap(t) ;  
  
}









