
#include <Rcpp.h>
#include <iostream>
#include <fstream>
#include <sstream>

using namespace Rcpp;
using namespace std;



// [[Rcpp::export]]
std::vector<std::string> makeLetters(){
  
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


// [[Rcpp::export]]
IntegerVector RcppConvertFromExcelRef2( std::vector<std::string>& r ){
  
  // This function converts the Excel column letter to an integer
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
      v[i] = sharedStrings[atoi(tmp[2].c_str())].length() - 37; //-37 for shared string tags around text
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
std::string cppReadFile2(std::string xmlFile){
  
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
  
  // This function loops over all characters in xml, looking for tag
  // tag should look liked <tag>
  // tagEnd is then generated to be <tag/>
  
  
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
std::vector<std::string> getChildlessNode_ss(std::string xml, std::string tag){
  
  size_t k = tag.length();
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
  
  return r ;  
  
}


// [[Rcpp::export]]
CharacterVector get_extLst_Major(std::string xml){
  
  // find page margin or pagesetup then take the extLst after that
  
  if(xml.length() == 0)
    return wrap(NA_STRING);
  
  std::vector<std::string> r;
  std::string tagEnd = "</extLst>";
  size_t endPos = 0;
  std::string node;
  
  
  size_t pos = xml.find("<pageSetup ", 0);   
  if(pos == std::string::npos)
    pos = xml.find("<pageMargins ", 0);   
  
  if(pos == std::string::npos)
    pos = xml.find("</conditionalFormatting>", 0);   
  
  if(pos == std::string::npos)
    return wrap(NA_STRING);
  
  while(1){
    
    pos = xml.find("<extLst>", pos + 1);  
    if(pos == std::string::npos)
      break;
    
    endPos = xml.find(tagEnd, pos + 8);
    
    node = xml.substr(pos + 8, endPos - pos - 8);
    pos = xml.find("conditionalFormattings", pos + 1);  
    if(pos == std::string::npos)
      break;
    
    r.push_back(node.c_str());
    
  }
  
  return wrap(r) ;  
  
}



// [[Rcpp::export]]
CharacterVector getChildlessNode(std::string xml, std::string tag){
  
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
CharacterVector getSharedStrings(std::string xmlFile, bool isFile){
  
  
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
  //  T F	T (must be a formula)	
  
  int n = r.size();
  List cells(n);
  LogicalVector hasV = !is_na(v);
  LogicalVector hasR = !is_na(r);
  LogicalVector hasT = !is_na(t);
  
  for(int i=0; i < n; i++){
    
    if(hasR[i]){
      
      if(hasV[i]){
        
        if(hasT[i]){
          
          //  r t v	
          //  T	T	T (2)
          cells[i] = CharacterVector::create(
            Named("r") = r[i],
                          Named("t") = t[i],
                                        Named("v") = v[i],
                                                      Named("f") = NA_STRING); 
          
        }else{
          
          //  r t f	
          //  T	T	T (4 - formula)
          cells[i] = CharacterVector::create(
            Named("r") = r[i],
                          Named("t") = "str",
                          Named("v") = NA_STRING,
                          Named("f") = "<f>" + v[i] + "</f>"); 
          
          
        }
        
      }else{
        
        //  r t v	
        //  T	F	F	(1)
        cells[i] = CharacterVector::create(
          Named("r") = r[i],
                        Named("t") = NA_STRING,
                        Named("v") = NA_STRING,
                        Named("f") = NA_STRING); 
      }
      
    }else{
      
      //  r t v	
      //  F F	F	(3)
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
  
  //Valid combinations
  //  r t	v	f
  //  T	F	F	F done
  //  T	T	T	F done
  //  T	T	T	T done
  //  T	F	F	T done  
  
  int n = r.size();
  List cells(n);
  LogicalVector hasV = !is_na(v);
  LogicalVector hasF = !is_na(f);
  CharacterVector nms(n);
  std::string ri;
  
  for(int i=0; i < n; i++){
    
    ri = as<std::string>(r[i]);      
    ri.erase(std::remove_if(ri.begin(), ri.end(), ::isalpha), ri.end());
    nms[i] = ri;
    
    
    //If we have a function    
    if(hasF[i] & !hasV[i]){
      
      cells[i] = CharacterVector::create(
        Named("r") = r[i],
                      Named("t") = NA_STRING,
                      Named("v") = NA_STRING,
                      Named("f") = f[i]); 
      
    }else if(hasF[i]){
      
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
  
  if(n > 0)
    cells.attr("names") = nms;
  
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
SEXP buildMatrixNumeric(CharacterVector v, IntegerVector rowInd, IntegerVector colInd,
                        CharacterVector colNames, int nRows, int nCols){
  
  //Rcout << "Running buildMatrixNumeric" << endl;
  
  int k = v.size();
  NumericMatrix m(nRows, nCols);
  std::fill(m.begin(), m.end(), NA_REAL);
  
  for(int i = 0; i < k; i++)
    m(rowInd[i], colInd[i]) = atof(v[i]);
  
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
                      IntegerVector rowInd,
                      IntegerVector colInd,
                      CharacterVector colNames,
                      int nRows,
                      int nCols,
                      IntegerVector charCols,
                      IntegerVector dateCols){
  
  
  /* List d(10);
   d[0] = v;
   d[1] = vn;
   d[2] = rowInd;
   d[3] = colInd;
   d[4] = colNames;
   d[5] = nRows;
   d[6] = nCols;
   d[7] = charCols;
   d[8] = dateCols;
   d[9] = originAdj;
   return(wrap(d));
   */
  
  int k = v.size();
  std::string dt_str;
  
  // create and fill matrix
  CharacterMatrix m(nRows, nCols);
  std::fill(m.begin(), m.end(), NA_STRING);
  
  for(int i = 0;i < k; i++)
    m(rowInd[i], colInd[i]) = v[i];
  
  
  
  // this will be the return data.frame
  List dfList(nCols); 
  
  
  // loop over each column and check type
  for(int i = 0; i < nCols; i++){
    
    CharacterVector tmp(nRows);
    
    for(int ri = 0; ri < nRows; ri++)
      tmp[ri] = m(ri,i);
    
    LogicalVector notNAElements = !is_na(tmp);
    
    
    // If column is date class and no strings exist in column
    if( (std::find(dateCols.begin(), dateCols.end(), i) != dateCols.end()) &
        (std::find(charCols.begin(), charCols.end(), i) == charCols.end()) ){
      
      // these are all dates and no characters --> safe to convert numerics
      
      DateVector datetmp(nRows);
      for(int ri=0; ri < nRows; ri++){
        if(!notNAElements[ri]){
          datetmp[ri] = NA_REAL; //IF TRUE, TRUE else FALSE
        }else{
          // dt_str = as<std::string>(m(ri,i));
          dt_str = m(ri,i);
          datetmp[ri] = Rcpp::Date(atoi(dt_str.substr(5,2).c_str()), atoi(dt_str.substr(8,2).c_str()), atoi(dt_str.substr(0,4).c_str()) );
          //datetmp[ri] = Date(atoi(m(ri,i)) - originAdj);
          //datetmp[ri] = Date(as<std::string>(m(ri,i)));
        }
      }
      
      dfList[i] = datetmp;
      
      
      // character columns
    }else if(std::find(charCols.begin(), charCols.end(), i) != charCols.end()){
      
      // determine if column is logical or date
      bool logCol = true;
      for(int ri = 0; ri < nRows; ri++){
        if(notNAElements[ri]){
          if((m(ri, i) != "TRUE") & (m(ri, i) != "FALSE")){
            logCol = false;
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
        
      }else{
        
        dfList[i] = tmp;
        
      }
      
    }else{ // else if column NOT character class (thus numeric)
      
      NumericVector ntmp(nRows);
      for(int ri = 0; ri < nRows; ri++){
        if(notNAElements[ri]){
          ntmp[ri] = atof(m(ri, i)); 
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
std::vector<int> matrixRowInds2(std::vector<int>& x) {
  
  int n = x.size();
  std::vector<int> v(n);
  v[0] = 0;
  
  int c = 0;
  for(int i = 1; i < n; ++i){
    if(x[i] != x[i-1])
      ++c;
    v[i] = c;
  }
  
  return v;
  
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
        
        //If we have a c value we might have an f value (cval[3] is f)
        if(attrs[3]){
          
          if(attrs[2]){ // If v is NOT NA
            cellXML += "\" t=\"" + cVal[1] + "\">" + cVal[3] + "<v>" + cVal[2] + "</v></c>";
          }else{
            cellXML += "\" t=\"" + cVal[1] + "\">" + cVal[3] + "</c>";
          }
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
  
  if(n == 0)
    return wrap("");
  
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
int calcNRows(CharacterVector x, bool skipEmptyRows){
  
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
    }else if(classes[i] == "openxlsx_formula"){
      colLabels[i] = NA_STRING;
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
CharacterVector getCellsWithChildrenLimited(std::string xmlFile, CharacterVector emptyNodes, int n){
  
  //read in file without spaces
  std::string xml = cppReadFile(xmlFile);
  
  // std::string tag = "<c ";
  //  std::string tagEnd1 = ">";
  //  std::string tagEnd2 = "</c>";
  
  
  // count number of rows
  int ocs = 0;
  string::size_type start = 0;
  
  while(ocs <= n) {
    
    start = xml.find("<row r", start);
    if(start == string::npos){
      break;
    }else{
      ++ocs;
      start = start + 5;
    }
    
  }
  
  xml = xml.substr(0, start);
  
  // count cells with children
  ocs = 0;
  start = 0;
  while((start = xml.find("</v>", start)) != string::npos) {
    ++ocs;
    start += 4;
  }
  
  CharacterVector cells(ocs);
  std::fill(cells.begin(), cells.end(), NA_STRING);
  
  int i = 0;
  size_t nextPos = 3;
  size_t vPos = 2;
  std::string sub;
  size_t pos = xml.find("<c ", 0);
  
  while(i < ocs){
    
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
CharacterVector getCellsWithChildren(std::string xmlFile, CharacterVector emptyNodes){
  
  //read in file without spaces
  std::string xml = cppReadFile(xmlFile);
  
  // std::string tag = "<c ";
  //  std::string tagEnd1 = ">";
  //  std::string tagEnd2 = "</c>";
  
  // count cells with children
  int ocs = 0;
  string::size_type start = 0;
  while((start = xml.find("</v>", start)) != string::npos) {
    ++ocs;
    start += 4;
  }
  
  CharacterVector cells(ocs);
  std::fill(cells.begin(), cells.end(), NA_STRING);
  
  int i = 0;
  size_t nextPos = 3;
  size_t vPos = 2;
  std::string sub;
  size_t pos = xml.find("<c ", 0);
  
  while(i < ocs){
    
    if(pos != std::string::npos){
      nextPos = xml.find("<c ", pos+9);
      vPos = xml.find("</v>", pos+8); // have to atleast pass <c r="XX">
      
      if(vPos < nextPos){
        cells[i] = xml.substr(pos, nextPos - pos);
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
          
          if(attrs[2]){ // If v is NOT NA
            cellXML += "\" t=\"" + cVal[1] + "\">" + cVal[3] + "<v>" + cVal[2] + "</v></c>";
          }else{
            cellXML += "\" t=\"" + cVal[1] + "\">" + cVal[3] + "</c>";
          }
          
        }else{
          cellXML += "\" t=\"" + cVal[1] + "\"><v>" + cVal[2] + "</v></c>";
        }
      }else if(attrs[3]){
        cellXML += "\">" + cVal[3] + "</c>";
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
  
  IntegerVector colNumbers = RcppConvertFromExcelRef(r); 
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
  std::string xml = cppReadFile2(xmlFile);
  
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
  
  int nRows = calcNRows(r, skipEmptyRows);
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
SEXP loadworksheets(Reference wb, List styleObjects, std::vector<std::string> xmlFiles, LogicalVector is_chart_sheet){
  
  List worksheets = wb.field("worksheets");
  int n_sheets = is_chart_sheet.size();
  CharacterVector sheetNames = worksheets.attr("names");
  
  // variable set up
  std::string tagEnd = "\"";
  std::string cell;
  List sheetData(n_sheets);
  List freezePane(n_sheets);
  List colWidths(n_sheets);
  List rowHeights(n_sheets);
  List dataCount(n_sheets);
  List hyperLinks(n_sheets);
  List wbstyleObjects;
  
  // loop over each worksheet file
  for(int i = 0; i < n_sheets; i++){
    
    if(is_chart_sheet[i]){
      
      colWidths[i] = List(0);
      rowHeights[i] = List(0);
      sheetData[i] = List(0);
      freezePane[i] = List(0);
      rowHeights[i] = List(0);
      dataCount[i] = 0;
      hyperLinks[i] = List(0);
      
    }else{
      
      
      
      colWidths[i] = List(0);
      rowHeights[i] = List(0);
      sheetData[i] = List(0);
      List this_worksheet = worksheets[i];
      
      
      //read in file
      std::string xmlFile = xmlFiles[i];
      
      std::string buf;
      std::string xml = cppReadFile2(xmlFile);
      // ifstream file;
      // file.open(xmlFile.c_str());
      // while (file >> buf)
      // xml += buf + ' ';
      
      std::size_t pos = xml.find("<sheetData>");  // find sheetData
      size_t endPos = 0;
      size_t tmp_pos = 0;
      
      bool has_data = true;
      if(pos == string::npos){
        has_data = false;
        pos = xml.find("<sheetData/>");
      }
      
      /* --- Everything before pos --- */
      std::string xml_pre = xml.substr(0, pos);
      
      // sheetPR
      CharacterVector sheetPr = getNodes(xml_pre, "<sheetPr>");
      
      if(sheetPr.size() == 0){
        sheetPr = getNodes(xml_pre, "<sheetPr");
        
        for(int j = 0; j < sheetPr.size(); j++){
          std::string sp = as<std::string>(sheetPr[j]);
          char ch = *sp.rbegin();  
          if(ch != '>')
            sp += ">";
          sheetPr[j] = sp;
        }
      }
      
      if(sheetPr.size() == 0)
        sheetPr = getChildlessNode(xml_pre, "<sheetPr");
      
      if(sheetPr.size() > 0)
        this_worksheet["sheetPr"] = sheetPr;
      
      
      
      // Freeze Panes
      CharacterVector node_xml = getChildlessNode(xml_pre, "<pane ");
      if(node_xml.size() > 0){
        freezePane[i] = node_xml;
      }else{
        freezePane[i] = List(0);
      }
      
      // SheetViews
      node_xml = getNodes(xml_pre, "<sheetViews>");
      if(node_xml.size() > 0)
        this_worksheet["sheetViews"] = node_xml;
      
      
      //colwidths
      std::vector<std::string> cols = getChildlessNode_ss(xml_pre, "<col ");
      if(cols.size() > 0){
        
        NumericVector widths;
        IntegerVector columns;
        
        
        for(size_t ci = 0; ci < cols.size(); ci++){
          
          double tmp = 0;
          int min_c = 0;
          int max_c = 0;
          buf = cols[ci];
          if(buf.find("customWidth", 0) != string::npos){
            
            tmp_pos = buf.find("min=\"", 0);
            endPos = buf.find(tagEnd, tmp_pos + 5);
            min_c = atoi(buf.substr(tmp_pos + 5, endPos - tmp_pos - 5).c_str());
            
            tmp_pos = buf.find("max=\"", 0);
            endPos = buf.find(tagEnd, tmp_pos + 5);
            max_c = atoi(buf.substr(tmp_pos + 5, endPos - tmp_pos - 5).c_str());
            
            tmp_pos = buf.find("width=\"", 0);
            endPos = buf.find(tagEnd, tmp_pos + 7);
            tmp = atof(buf.substr(tmp_pos + 7, endPos - tmp_pos - 7).c_str()) - 0.71;
            
            if(min_c != max_c){
              while(min_c <= max_c){
                widths.push_back(tmp);
                columns.push_back(min_c);
                min_c++;
              }
            }else{
              widths.push_back(tmp);
              columns.push_back(min_c);
            }
          }
          
        }
        
        if(widths.size() > 0){
          CharacterVector tmp_widths(widths);
          tmp_widths.attr("names") = columns;
          colWidths[i] = tmp_widths;
        }
        
      }
      
      
      
      
      /* --- Everything after sheetData --- */
      size_t pos_post = 0;
      if(has_data){
        pos_post = xml.find("</sheetData>");   
      }else{
        pos_post = pos; 
      }
      
      std::string xml_post = xml.substr(pos_post);
      
      
      node_xml = getChildlessNode(xml_post, "<autoFilter ");
      if(node_xml.size() > 0)
        this_worksheet["autoFilter"] = node_xml;
      
      node_xml = getChildlessNode(xml_post, "<hyperlink ");
      if(node_xml.size() > 0){
        hyperLinks[i] = node_xml;
      }else{
        hyperLinks[i] = List(0);
      }
      
      node_xml = getChildlessNode(xml_post, "<pageMargins ");
      if(node_xml.size() > 0)
        this_worksheet["pageMargins"] = node_xml;
      
      
      node_xml = getChildlessNode(xml_post, "<pageSetup ");
      if(node_xml.size() > 0){
        for(int j = 0; j < node_xml.size(); j++){
          
          std::string pageSetup_tmp = as<std::string>(node_xml[j]);
          size_t ps_pos = pageSetup_tmp.find("r:id=\"rId", 0);
          if(ps_pos != std::string::npos){
            
            std::string pageSetup_tmp2 = pageSetup_tmp.substr(0, ps_pos + 9) + "2";
            ps_pos = pageSetup_tmp.find("\"", ps_pos + 9);
            pageSetup_tmp  = pageSetup_tmp2 + pageSetup_tmp.substr(ps_pos);
            
          }
          
          node_xml[j] = pageSetup_tmp;
          
        }
        this_worksheet["pageSetup"] = node_xml;
      }
      
      
      
      node_xml = getChildlessNode(xml_post, "<mergeCell ");
      if(node_xml.size() > 0)
        this_worksheet["mergeCells"] = node_xml;
      
      
      node_xml = getChildlessNode(xml_post, "<drawing ");
      if(node_xml.size() == 0)
        node_xml = getChildlessNode(xml_post, "<legacyDrawing ");
      
      if(node_xml.size() > 0){
        for(int j = 0; j < node_xml.size(); j++){
          
          std::string drawingId_tmp = as<std::string>(node_xml[j]);
          size_t ps_pos = drawingId_tmp.find("r:id=\"rId", 0);
          
          std::string drawingId_tmp2 = drawingId_tmp.substr(0, ps_pos + 9) + "1";
          ps_pos = drawingId_tmp.find("\"", ps_pos + 9);
          
          drawingId_tmp  = drawingId_tmp2 + drawingId_tmp.substr(ps_pos);
          node_xml[j] = drawingId_tmp;
          
        }
      }
      
      
      //  conditionalFormatting
      CharacterVector conForm = getNodes(xml_post, "<conditionalFormatting");
      if(conForm.size() > 0){
        
        // get sqref attribute
        size_t tmp_pos = 0;
        int end_pos = 0;
        std::string sqref;
        CharacterVector cf;
        CharacterVector cf_names;
        
        
        for(int ci = 0; ci < conForm.size(); ci++){
          
          buf = conForm[ci];
          
          tmp_pos = buf.find("sqref=\"", 0);
          end_pos = buf.find("\"", tmp_pos + 7);
          
          sqref = buf.substr(tmp_pos + 7, end_pos - tmp_pos - 7);
          buf = buf.substr(0, buf.find("</conditionalFormatting"));
          buf = buf.substr(buf.find("<cfRule"));
          
          int ocs = 0;
          string::size_type start = 0;
          while((start = buf.find("<cfRule", start)) != string::npos){
            ++ocs;
            start += 7;
          }
          
          if(ocs == 1){
            cf_names.push_back(sqref);
            cf.push_back(buf);
          }else if(ocs > 1){
            
            tmp_pos = buf.find("<cfRule", 0);
            while(tmp_pos != std::string::npos){
              
              end_pos = buf.find("<cfRule", tmp_pos + 7);
              cf.push_back(buf.substr(tmp_pos, end_pos - tmp_pos));
              cf_names.push_back(sqref);
              tmp_pos = end_pos;
              
            }
          }
          
        } // end of loop through conditional formats
        
        cf.attr("names") = cf_names;
        this_worksheet["conditionalFormatting"] = cf;
        
      } // end of if(conForm.size() > 0)
      
      // extLst
      node_xml = get_extLst_Major(xml_post);
      if(node_xml.size() > 0)
        this_worksheet["extLst"] = node_xml;
      
      
      
      
      
      
      // clean pre and post xml
      xml_post.clear();
      xml_pre.clear();
      
      
      /* --------------------------- sheet Data --------------------------- */
      
      if(has_data){
        
        xml = xml.substr(pos + 11, pos_post - pos - 11);     // get from "sheetData" to the end
        
        // count cells with children
        int ocs = 0;
        string::size_type start = 0;
        while((start = xml.find("<c ", start)) != string::npos){
          ++ocs;
          start += 4;
        }
        
        CharacterVector r(ocs);
        CharacterVector r_nms(ocs);
        CharacterVector v(ocs);
        CharacterVector s(ocs);
        
        std::fill(v.begin(), v.end(), NA_STRING);
        std::fill(s.begin(), s.end(), NA_STRING);
        
        
        int j = 0;
        size_t nextPos = 3;
        pos = xml.find("<c ", 0);
        bool has_v = false;
        bool has_f = false;
        List cells(ocs); // sheetData
        std::string func;
        std::string type;
        
        // PULL OUT CELL AND ATTRIBUTES
        while(j < ocs){
          
          if(pos != std::string::npos){
            
            has_v = false;
            has_f = false;
            type = "n";
            
            nextPos = xml.find("<c ", pos + 9);
            cell = xml.substr(pos, nextPos - pos);
            
            // Pull out ref
            pos = cell.find("r=", 0);  // find r="
            endPos = cell.find(tagEnd, pos + 3);  // find next "
            r[j] = cell.substr(pos + 3, endPos - pos - 3).c_str();
            
            buf = cell.substr(pos + 3, endPos - pos - 3);      
            buf.erase(std::remove_if(buf.begin(), buf.end(), ::isalpha), buf.end());
            r_nms[j] = buf;
            
            // Pull out style
            pos = cell.find(" s=", 0);  // find s="
            if(pos != std::string::npos){
              endPos = cell.find(tagEnd, pos + 4);  // find next "
              s[j] = cell.substr(pos + 4, endPos - pos - 4);
            }
            
            // Pull out type
            pos = cell.find(" t=", 0);  // find t="
            if(pos != std::string::npos){
              endPos = cell.find(tagEnd, pos + 4);  // find next "
              type = cell.substr(pos + 4, endPos - pos - 4);
            }
            
            
            // find <v> tag and </v> end tag
            endPos = cell.find("</v>", 0);
            if(endPos != std::string::npos){
              pos = cell.find("<v", 0);
              pos = cell.find(">", pos);
              v[j] = cell.substr(pos + 1, endPos - pos - 1);
              has_v = true;
            }
            
            // get<f>
            pos = cell.find("<f", 0);
            if(pos != std::string::npos){
              endPos = cell.find("</f>", pos + 3);
              if(endPos == std::string::npos){
                endPos = cell.find("/>", pos + 3);
                func = cell.substr(pos, endPos - pos + 2);
              }else{
                func = cell.substr(pos, endPos - pos + 4);
              }
              has_f = true;
            }
            
            if(has_f & !has_v){
              cells[j] = CharacterVector::create(Named("r") = r[j], Named("t") = NA_STRING, Named("v") = NA_STRING, Named("f") = func); 
            }else if(has_f){
              cells[j] = CharacterVector::create(Named("r") = r[j], Named("t") = type, Named("v") = v[j], Named("f") = func); 
            }else if(has_v){
              cells[j] = CharacterVector::create(Named("r") = r[j], Named("t") = type, Named("v") = v[j], Named("f") = NA_STRING); 
            }else{ //only have s and r
              cells[j] = CharacterVector::create(Named("r") = r[j], Named("t") = NA_STRING, Named("v") = NA_STRING, Named("f") = NA_STRING);
            }
            
            
            
            j++; // INCREMENT OVER OCCURENCES
            pos = nextPos;
            
          }  // end of while loop over occurences
        }  // END OF CELL AND ATTRIBUTION GATHERING
        
        // get names of cells
        
        if(ocs > 0){
          cells.attr("names") = r_nms;
          sheetData[i] = cells;
        }
        
        // count number of rows
        int row_ocs = 0;
        start = 0;
        while((start = xml.find("<row ", start)) != string::npos){
          ++row_ocs;
          start += 4;
        }
        
        CharacterVector rowNumbers(row_ocs);
        CharacterVector heights(row_ocs);
        
        
        // PULL OUT CELL AND ATTRIBUTES
        j = 0;
        pos = xml.find("<row ", 0);
        std::string htTag = " ht=\"";
        std::string attrEnd = "\"";
        
        while(j < row_ocs){
          
          if(pos != std::string::npos){
            
            nextPos = xml.find("<row ", pos + 9);
            cell = xml.substr(pos, nextPos - pos);
            
            
            // Pull out ref
            pos = cell.find("r=", 0);  // find r="
            endPos = cell.find(tagEnd, pos + 3);  // find next "
            rowNumbers[j] = cell.substr(pos + 3, endPos - pos - 3);
            
            
            // find custom height  
            pos = cell.find(htTag, pos);
            if(pos == std::string::npos){
              heights[j] = NA_STRING;
            }else{  
              endPos = cell.find(attrEnd, pos + 5);
              heights[j] = cell.substr(pos + 5, endPos - pos - 5);
            }
            
            
            
            j++; // INCREMENT OVER OCCURENCES
            pos = nextPos;
            
          }  // end of while loop over occurences
        }  // END OF CELL AND ATTRIBUTION GATHERING
        
        
        rowNumbers = rowNumbers[!is_na(heights)];
        if(rowNumbers.size() > 0){
          heights = heights[!is_na(heights)];
          heights.attr("names") = rowNumbers;
          rowHeights[i] = heights;
        }
        
        // styleObjects
        std::string this_sheetname = as<std::string>(sheetNames[i]);
        
        if(any(!is_na(s))){
          
          CharacterVector s_refs = r[!is_na(s)];
          s = s[!is_na(s)];
          
          CharacterVector uStyleInds = sort_unique(s);
          int nsu = uStyleInds.size();
          CharacterVector uStyleInds_j(1);
          
          std::string ref_j;
          CharacterVector styleElementNames = CharacterVector::create("style", "sheet", "rows", "cols");
          
          for(int j = 0; j < nsu; j ++){
            
            List styleElement(4);
            int styleInd = atoi(as<std::string>(uStyleInds[j]).c_str());
            
            if(styleInd != 0){
              
              uStyleInds_j[0] = uStyleInds[j];
              LogicalVector ind = !is_na(match(s, uStyleInds_j));
              CharacterVector s_refs_j = s_refs[ind];
              
              int n_j = s_refs_j.size();
              IntegerVector rows(n_j);
              IntegerVector cols = RcppConvertFromExcelRef(s_refs_j);
              
              for(int k = 0; k < n_j; k++){
                ref_j = s_refs_j[k];
                ref_j.erase(std::remove_if(ref_j.begin(), ref_j.end(), ::isalpha), ref_j.end());
                rows[k] = atoi(ref_j.c_str());  
              }
              
              styleElement[0] = styleObjects[styleInd - 1];
              styleElement[1] = this_sheetname;
              styleElement[2] = rows;
              styleElement[3] = cols;
              
              styleElement.attr("names") = styleElementNames;
              
              wbstyleObjects.push_back(styleElement);
              
            }
            
          }
          
          
        }
        
        dataCount[i] = 1;
      }else{
        dataCount[i] = 0;
      } // end of if(has_data)
      
      
      worksheets[i] = this_worksheet;
      
    } 
  }
  
  // assign back to workbook
  wb.field("worksheets") = worksheets;
  wb.field("sheetData") = sheetData;
  wb.field("freezePane") = freezePane;
  wb.field("rowHeights") = rowHeights;
  wb.field("colWidths") = colWidths;
  wb.field("styleObjects") = wbstyleObjects;
  wb.field("dataCount") = dataCount;
  wb.field("hyperlinks") = hyperLinks;
  
  return wrap(wb);
  
  
}



// [[Rcpp::export]]
SEXP ExcelConvertExpand(const std::vector<int>& cols, const std::vector<std::string>& LETTERS, const std::vector<std::string>& rows){
  
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
List buildMatrixNumeric2(std::vector<std::string> v,
                         const std::vector<int>& rowInd,
                         IntegerVector& colInd,
                         int nRows,
                         int nCols){
  
  int k = v.size();
  NumericMatrix m(nRows, nCols);
  std::fill(m.begin(), m.end(), NA_REAL);
  std::string tmp;
  
  for(int i = 0; i < k; i++){
    tmp = v[i];
    m(rowInd[i], colInd[i]) = atof(tmp.c_str());
  }
  
  List dfList(nCols);
  for(int i=0; i < nCols; ++i)
    dfList[i] = m(_,i);
  
  std::vector<int> rowNames(nRows);
  for(int i = 0;i < nRows; ++i)
    rowNames[i] = i+1;
  
  dfList.attr("row.names") = rowNames;
  dfList.attr("class") = "data.frame";
  
  return Rcpp::wrap(dfList);
  
}


// [[Rcpp::export]]
SEXP readWorkbook2(std::vector<std::string>& v,
                   std::vector<std::string>& r,
                   std::vector<std::string>& string_refs,
                   std::vector<bool>& is_date,
                   int nRows,
                   bool hasColNames,
                   bool skipEmptyRows,
                   Function clean_names
){
  
  
  // Convert r to column number and shift to scale 0:nCols  (from eg. J:AA in the worksheet)
  int nCells = r.size();
  int nDates = is_date.size();
  
  /* do we have any dates */
  bool has_date;
  if(nDates == nCells){
    has_date = true;
  }else{
    has_date = false;
  }
  
  IntegerVector colNumbers = RcppConvertFromExcelRef2(r); 
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
    
    // remove elements in string_refs that appear in the first row (if there are numerics in first row this will remove less elements than pos)
    std::vector<std::string>::iterator first = r.begin();
    std::vector<std::string>::iterator last = r.begin() + pos;
    std::vector<std::string> toRemove(first, last);
    
    for(unsigned int i = 0; i < toRemove.size(); ++i)
      string_refs.erase(std::remove(string_refs.begin(), string_refs.end(), toRemove[i]), string_refs.end());
    
    if(pos > 0)
      r.erase(r.begin(), r.begin() +  pos); 
    
    // tidy up column names
    colNames = clean_names(colNames);
    
    //If nothing left return a data.frame with 0 rows
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
    
    
  }else{ // else colNames is FALSE */
    char name[6];
    for(int i =0; i < nCols; i++){
      sprintf(&(name[0]), "X%d", i+1);
      colNames[i] = name;
    }
  }
  
  // getRow numbers from r 
  std::vector<int> rowNumbers(nCells);  
  if(nCells == nRows*nCols){
    
    int counter = 0;
    for(int i = 0; i < nRows; ++i){
      for(int j = 0; j < nCols; ++j){
        rowNumbers[counter] = i;
        ++counter;
      }
    }
    
  }else{
    
    // remove character from cell ref
    for(int i = 0; i < nCells; i++){
      std::string a = r[i];
      a.erase(std::remove_if(a.begin(), a.end(), ::isalpha), a.end());
      rowNumbers[i] = atoi(a.c_str()) - 1; 
    }
    
    if(skipEmptyRows){
      rowNumbers = matrixRowInds2(rowNumbers);
    }else{
      for(int i = 0; i < nCells; i++)
        rowNumbers[i] = rowNumbers[i] - rowNumbers[0];
    }
    
  }
  
  
  // Possible there are no stringInds to begin with and value of stringInds is 0
  // Possible we have stringInds but they have now all been used up by string_refs
  bool allNumeric = false;
  if((string_refs.size() == 0))
    allNumeric = true;
  
  // check not all string_refs are NA *************************
  
  
  
  
  // if(has_date){
  //    if(is_true(any(is_date)))
  //      allNumeric = false;
  //  }
  
  // If we build colnames we remove the ref & values for the pos number of elements used 
  if(hasColNames & has_date)
    is_date.erase(is_date.begin(), is_date.begin() + pos);
  
  
  
  //Intialise return data.frame
  List m; 
  v.erase(v.begin(), v.begin() + pos);
  
  //   Rcout << "rowNumbers.size(): " << rowNumbers.size() << endl; 
  //   Rcout << "v.size(): " << v.size() << endl; 
  //   Rcout << "colNumbers.size(): " << colNumbers.size() << endl; 
  //   Rcout << "is_date.size(): " << is_date.size() << endl; 
  //   Rcout << "allNumeric: " << allNumeric << endl; 
  
  if(allNumeric){
    
    m = buildMatrixNumeric2(v, rowNumbers, colNumbers, nRows, nCols);
    m.attr("names") = colNames;
    
  }else{
    
    // If it contains any strings it will be a character column
    //     IntegerVector charCols;
    //     if(all(is_na(string_refs))){
    //       charCols = -1;
    //     }else{
    //       charCols = match(string_refs, r);
    //       IntegerVector charColNumbers = colNumbers[charCols-1];
    //       charCols = unique(charColNumbers);
    //     }
    //     
    //Rcout << "colNumbers.size(): " << colNumbers.size() << endl; 
    //Rcout << "is_date.size(): " << is_date.size() << endl; 
    //Rcout << "v.size(): " << v.size() << endl; 
    //Rcout << "vn.size(): " << vn.size() << endl; 
    //Rcout << "rowNumbers.size(): " << rowNumbers.size() << endl; 
    
    //date columns
    //     IntegerVector dateCols(1);
    //     if(has_date){
    //       dateCols = colNumbers[is_date];
    //       dateCols = sort_unique(dateCols);
    //     }else{
    //       dateCols[0] = -1;
    //     }
    
    //    m = buildMatrixMixed(v, rowNumbers, colNumbers, colNames, nRows, nCols, charCols, dateCols, originAdj);
    
  }
  
  return m;
  
}

// [[Rcpp::export]]
SEXP mergeCell2mappingDF(CharacterVector x){
  
  int n = x.size();
  int ref_pos;
  int colon_pos;
  int end_pos;
  
  std::string tag_end = "\"";
  std::string cell;
  std::string left_val;
  std::string right_val;
  
  std::vector<std::string> lv(n);
  std::vector<std::string> rv(n);
  
  IntegerVector start_col(n);
  IntegerVector end_col(n);
  
  IntegerVector start_row(n);
  IntegerVector end_row(n);

  std::vector<std::string> LETTERS = makeLetters();
  
  for(int i = 0; i < n; i++){
    
    cell = as<std::string>(x[i]);
    
    ref_pos = cell.find("ref=", 0);  // find ref=
    colon_pos = cell.find(":", ref_pos + 6);  // find :
    end_pos = cell.find(tag_end, colon_pos + 1);  // find next 
    
    lv[i] = cell.substr(ref_pos + 5, colon_pos - ref_pos - 5).c_str();
    rv[i] = cell.substr(colon_pos + 1, end_pos - colon_pos - 1).c_str();
    
    left_val = lv[i];
    right_val = rv[i];
   
    left_val.erase(std::remove_if(left_val.begin(), left_val.end(), ::isalpha), left_val.end());
    start_row[i] = atoi(left_val.c_str()); 
    
    right_val.erase(std::remove_if(right_val.begin(), right_val.end(), ::isalpha), right_val.end());
    end_row[i] = atoi(right_val.c_str()); 
    
   
  }

  
  start_col = RcppConvertFromExcelRef2(lv);
  end_col = RcppConvertFromExcelRef2(rv);
  
  int n_rows = sum((end_row - start_row + 1) * (end_col - start_col + 1));
  CharacterVector anchor_cell(n_rows);
  CharacterVector ref(n_rows);
  IntegerVector tmpi(1);


  int k = 0;
  for(int i = 0; i < n; ++i){
    
    int n_cols =  (end_col[i] - start_col[i] + 1);
    int n_rows =  (end_row[i] - start_row[i] + 1);
    
    //Rcout << "n_cols: " << n_cols << endl;
    //Rcout << "n_rows: " << n_rows << endl;

    for(int r = 0; r < n_rows; r++){
      for(int c = 0; c < n_cols; c++){
      
        tmpi[0] = start_col[i] + c;
        CharacterVector tmp = convert2ExcelRef(tmpi, LETTERS = LETTERS);
        ref[k] = as<std::string>(tmp[0]) + itos(start_row[i] + r);
        anchor_cell[k] = lv[i];
        k++;
      }
    }
  }
  
  DataFrame mapping = DataFrame::create(Rcpp::Named("anchor_cell") = anchor_cell,
                                    Rcpp::Named("ref") = ref);
  
  return(wrap(mapping));
  
}



