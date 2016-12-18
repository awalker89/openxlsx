
#include "openxlsx.h"







// [[Rcpp::export]]
SEXP write_worksheet_xml(std::string prior
                           , std::string post
                           , Reference sheet_data
                           , std::string R_fileName){
  
  // open file and write header XML
  const char * s = R_fileName.c_str();
  std::ofstream xmlFile;
  xmlFile.open (s);
  xmlFile << "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>";
  xmlFile << prior;

  //NOTES ON WHY THIS WORKS
  // - If no row heights every row has children
  // - If data only added once, sheet_data will be in order
  // - Thus all rows have children (all starting row tags are open) and no need to sort/find
  

  // DEV START
  IntegerVector cell_row = sheet_data.field("rows");
  
  // If no data write childless node and return
  if(cell_row.size() == 0){
    xmlFile << "<sheetData/>";
    xmlFile << post;
    xmlFile.close();
    return Rcpp::wrap(0);
  }
  
  CharacterVector cell_col = int_2_cell_ref(sheet_data.field("cols"));
  CharacterVector cell_types = map_cell_types_to_char(sheet_data.field("t"));
  CharacterVector cell_value = sheet_data.field("v");
  CharacterVector cell_fn = sheet_data.field("f");

  CharacterVector style_id = sheet_data.field("style_id");
  CharacterVector unique_rows(sort_unique(cell_row));

    
  // DEV END

  size_t n = cell_row.size();
  size_t k = unique_rows.size();
  std::string xml;
  std::string cell_xml;

  // write sheet_data
  
  // write xml prior to stringData and opening tag
  xmlFile << "<sheetData>";
  size_t j = 0;
  String currentRow = unique_rows[0];
  
  for(size_t i = 0; i < k; i++){
    
    cell_xml = "";
    
    while(currentRow == unique_rows[i]){
      
      //cell XML strings
      cell_xml += "<c r=\"" + cell_col[j] + itos(cell_row[j]);
      
      if(!CharacterVector::is_na(style_id[j]))
        cell_xml += "\" s=\""+ style_id[j];
      
      //If we have a t value we must have a v value
      if(!CharacterVector::is_na(cell_types[j])){
        
        //If we have a c value we might have an f value (cval[3] is f)
        if(CharacterVector::is_na(cell_fn[j])){ // no function
          
          cell_xml += "\" t=\"" + cell_types[j] + "\"><v>" + cell_value[j] + "</v></c>";
          
        }else{
          if(CharacterVector::is_na(cell_value[j])){ // If v is NA
            cell_xml += "\" t=\"" + cell_types[j] + "\">" + cell_fn[j] + "</c>";
          }else{
            cell_xml += "\" t=\"" + cell_types[j] + "\">" + cell_fn[j] + "<v>" + cell_value[j] + "</v></c>";
          }
        }
        
      }else{
        cell_xml += "\"/>";
        
      }
      
      j += 1;
      
      if(j == n)
        break;
      
      currentRow = cell_row[j];
    }
    
    xmlFile << "<row r=\"" + unique_rows[i] + "\">" + cell_xml + "</row>";
    
  }
  
  // write closing tag, post <sheetData> xml and close
  xmlFile << "</sheetData>";
  xmlFile << post;
  xmlFile.close();
  
  return Rcpp::wrap(0);
  
}












// [[Rcpp::export]]
SEXP buildMatrixNumeric(CharacterVector v, IntegerVector rowInd, IntegerVector colInd,
                        CharacterVector colNames, int nRows, int nCols){
  
  LogicalVector isNA_element = is_na(v);
  if(is_true(any(isNA_element))){
    
    v = v[!isNA_element];
    rowInd = rowInd[!isNA_element];
    colInd = colInd[!isNA_element];
    
  }
  
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
CharacterVector build_table_xml(std::string table, std::string ref, std::vector<std::string> colNames, bool showColNames, std::string tableStyle, bool withFilter){
  
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



