# MergeXLSM

**Usage:** MergeXLSM.exe -i *template.xlsm* -o *result.xlsm* -dir *directory* -first *number* -last *number* -tag *text*

**Arguments:**

  -i *file*     
  Template XLSM file (default: Template.xlsm)

  -o *file*     
  Output XLSM file (default: Result.xlsm)

  -dir *path*   
  Directory containing XLSM files (default: XLSM)

  -first *num*  
  First row to process (default: 6)

  -last *num*   
  Last row to process (default: 1000)

  -tag *text*   
  Tag to search for in the first column (default: Fact)

**Example:**
  MergeXLSM.exe -i Template.xlsm -o Result.xlsm -dir XLSM -first 6 -last 1000 -tag Fact