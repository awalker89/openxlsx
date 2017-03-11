

#include "openxlsx.h"


// [[Rcpp::export]]
SEXP write_worksheet_xml_2( std::string prior
                          , std::string post
                          , Reference sheet_data
                          , CharacterVector row_heights
                          , std::string R_fileName){
  
  // open file and write header XML
  const char * s = R_fileName.c_str();
  std::ofstream xmlFile;
  xmlFile.open (s);
  xmlFile << "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>";
  xmlFile << prior;
  
  IntegerVector cell_row = sheet_data.field("rows");
  
  // If no data write childless node and return
  if(cell_row.size() == 0){

    xmlFile << "<sheetData/>";
    xmlFile << post;
    xmlFile.close();
    return Rcpp::wrap(0);
  }
  

  // sheet_data will be in order, jsut need to check for row_heights
  CharacterVector cell_col = int_2_cell_ref(sheet_data.field("cols"));
  CharacterVector cell_types = map_cell_types_to_char(sheet_data.field("t"));
  CharacterVector cell_value = sheet_data.field("v");
  CharacterVector cell_fn = sheet_data.field("f");
  
  CharacterVector style_id = sheet_data.field("style_id");
  CharacterVector unique_rows(sort_unique(cell_row));
  
  
  CharacterVector row_heights_rows = row_heights.attr("names");
  size_t n_row_heights = row_heights.size();
  

  size_t n = cell_row.size();
  size_t k = unique_rows.size();
  std::string xml;
  std::string cell_xml;
  
  size_t j = 0;
  size_t h = 0;
  String current_row = unique_rows[0];
  bool row_has_data = true;
  
  xmlFile << "<sheetData>";
  
  for(size_t i = 0; i < k; i++){
    
    cell_xml = "";
    row_has_data = true;
    
    while(current_row == unique_rows[i]){
      
      row_has_data = true;
      j += 1;
      
      if(CharacterVector::is_na(cell_col[j-1])){ //If r IS NA we have no row data we only have a rowHeight

        row_has_data = false;
        if(j == n)
          break;
        
        current_row = cell_row[j];
        break;
      }

      //cell XML strings
      cell_xml += "<c r=\"" + cell_col[j-1] + itos(cell_row[j-1]);
      
      if(!CharacterVector::is_na(style_id[j-1]))
        cell_xml += "\" s=\"" + style_id[j-1];
      
      
      //If we have a t value we must have a v value
      if(!CharacterVector::is_na(cell_types[j-1])){
        
        //If we have a c value we might have an f value
        if(CharacterVector::is_na(cell_fn[j-1])){ // no function
          
          cell_xml += "\" t=\"" + cell_types[j-1] + "\"><v>" + cell_value[j-1] + "</v></c>";

        }else{
          if(CharacterVector::is_na(cell_value[j-1])){ // If v is NA
            cell_xml += "\" t=\"" + cell_types[j-1] + "\">" + cell_fn[j-1] + "</c>";
          }else{
            cell_xml += "\" t=\"" + cell_types[j-1] + "\">" + cell_fn[j-1] + "<v>" + cell_value[j-1] + "</v></c>";
          }
        }
        
        
      }else if(!CharacterVector::is_na(cell_fn[j-1])){
        cell_xml += "\">" + cell_fn[j-1] + "</c>";
      }else{
        cell_xml += "\"/>";
      }
      
      if(j == n)
        break;
      
      current_row = cell_row[j];
      
    }
    
    
    if(h < n_row_heights){
      
      if((unique_rows[i] == row_heights_rows[h]) & row_has_data){ // this row has a row height and cell_xml data
        
        xmlFile << "<row r=\"" + unique_rows[i] + "\" ht=\"" + row_heights[h] + "\" customHeight=\"1\">" + cell_xml + "</row>";   
        h++;
        
      }else if(row_has_data){
        
        xmlFile << "<row r=\"" + unique_rows[i] + "\">" + cell_xml + "</row>";
        
      }else{
        
        xmlFile << "<row r=\"" + unique_rows[i] + "\" ht=\"" + row_heights[h] + "\" customHeight=\"1\"/>"; 
        h++;
      }
      
    }else{
      
      xmlFile << "<row r=\"" + unique_rows[i] + "\">" + cell_xml + "</row>";
      
    }
    
  }
  
  // write closing tag and XML post data
  xmlFile << "</sheetData>";
  xmlFile << post;
  
  //close file
  xmlFile.close();
  
  return wrap(0);
  
}