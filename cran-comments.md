
## Test environments
* local Windows 7, R 3.3.2
* local ubuntu 16.10, R 3.3.2
* win-builder (devel and release)

## R CMD check results
On Windows there were no ERRORs, WARNINGs or NOTEs
On Ubuntu there were no ERRORs or WARNINGs.

A NOTE on Ubuntu:
installed size is 6.5Mb
sub-directoris of 1Mb or more:
  libs  4.7Mb

## Downstream dependencies
I have also run R CMD check on downstream dependencies of openxlsx.
All packages that I could install passed.