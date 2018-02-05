
I have added calls to R_registerRoutines & R_useDynamicSymbols 
which resolved the NOTE for arch i386 but not x64 when running on local WIndows

This NOTE is not triggered on win-builder devel.


## Test environments
* local Windows 7, R 3.3.3 & R devel
* local ubuntu 16.10, R 3.3.2
* win-builder (devel and release)

## R CMD check results
On Windows there were no ERRORs, WARNINGs.

NOTE on local Windows (did not appear on win-builder devel)

checking compiled code ... NOTE
File 'openxlsx/libs/x64/openxlsx.dll':
  Found no calls to: 'R_registerRoutines', 'R_useDynamicSymbols'

Note on win-builder devel

Possibly mis-spelled words in DESCRIPTION:
  XLSX (3:29)
  xlsx (13:48)
  

On Ubuntu there were no ERRORs or WARNINGs.

A NOTE on Ubuntu:
installed size is 6.5Mb
sub-directoris of 1Mb or more:
  libs  4.7Mb

## Downstream dependencies
I have also run R CMD check on downstream dependencies of openxlsx.
All packages that I could install passed.