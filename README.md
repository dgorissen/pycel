This is a custom fork of [pycel](https://github.com/dgorissen/pycel) that helps load large excel workbooks where the large sheets within the workbook are static sheets used for VLOOKUPs. Please see [pycel issue](https://github.com/dgorissen/pycel/issues/152)

### Background

When we used pycel with our large spreadsheets 30mb+, it could up to 8 minutes to load. Most of the slowness came from our excel workbooks having massive static tables that were used for vlookups. The main reason the issue seems to occur is that pycel(and the underlying library openpyxl) were loading large static sheet(that do not reference other sheets/cells) the same way they would other cells (that would reference other sheets cells), causing the cell tree/graph structure that gets loaded into memory to be massive. In our case, our massive sheets are static and so we felt there must be a way to improve performance for our use case.

### Solution

- As part of the solution we also had to [fork the openpyxl library ](https://github.com/ObieCRE/openpyxl) so that it does not load large static sheets into memory the same way as other sheets. For a static sheet to be recognized in this way, one simply needs to ensure the name of their sheet as the word "STATIC" in it.

<img width="228" alt="Screen Shot 2022-08-18 at 5 51 14 PM" src="https://user-images.githubusercontent.com/5402488/185508701-13cfaf8d-c5ad-4c88-b0f3-8a317987e730.png">

- We also had to find a way to load and do vlookups in these static sheets _somehow_ ourselves. To do this we load the sheets into memory using openpyxl's readonly feature https://github.com/ObieCRE/pycel/blob/147d287b013634aa48578ae24e2283e44261e500/src/pycel/excelwrapper.py#L244
  - from there we store these large sheets [in a "global"](https://github.com/ObieCRE/pycel/blob/147d287b013634aa48578ae24e2283e44261e500/src/pycel/excelwrapper.py#L30) and then use a [modified vlookup implementation](https://github.com/ObieCRE/pycel/blob/147d287b013634aa48578ae24e2283e44261e500/src/pycel/lib/lookup.py#L473).
  
  
 ### Disclaimer
 
We do not claim to be experts in either the pycel or openpyxl repositories and acknowledge the implementation is not a proper long term solution to this issue, though it does meet our needs. We also are not sure if this solution would work beyond our VLOOKUP use case. Please see [pycel issue](https://github.com/dgorissen/pycel/issues/152) for further discussion/solutions.
