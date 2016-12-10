
#include "openxlsx.h"






// [[Rcpp::export]]   
List unique_cell_append(List sheetData, CharacterVector r, List newCells){
  
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
List build_cell_merges(List comps){
  
  size_t nMerges = comps.size(); 
  List res(nMerges);
  
  for(size_t i =0; i < nMerges; i++){
    IntegerVector col = convert_from_excel_ref(comps[i]);  
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