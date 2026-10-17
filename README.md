     /*---
     c:/utl/jnr-jenner-analytics-part-IV-inline-drop-down-to-poweshell-add-column-to-excell-sheet.sas;
     ---*/
    
    Jenner Analytics part IV inline drop down to poweshell add column to excell sheet
    
    Problem:
       Add column BMI to and existing excel worksheet
    
    Too long to post see
    
    Four of 10 languages
    https://github.com/rogerjdeangelis/jnr-jenner-analytics-part-IV-inline-drop-down-to-poweshell-add-column-to-excell-sheet
    
    Previous
    https://github.com/rogerjdeangelis/jnr-jenner-analytics-part-I-inline-drop-down-to-duckdb-sql
    https://github.com/rogerjdeangelis/jnr-jenner-analytics-part-II-drop-down-to-open-source-spss-run-regression
    https://github.com/rogerjdeangelis/jnr-jenner-analytics-part-III-inline-drop-down-to-open-source-matlab-summarize-data
    
    
    CONTENTS
    
      1  Create input workbook and worksheet have
      2  Powershell add column to sheet have
      3  Jenner print worksheet have powershell created column BMI
         Jenner print macro variable Greetings created y Powershell
      4  Related powershell repositories
      5  Inline drop down macros
    
    /******************************************************************************************************************/
    /* 1  Create input workbook and worksheet have                                                                    */
    /******************************************************************************************************************/
    
    %utlfkil(d:/xls/class.xlsx)
    
    libname xel excel "d:/xls/class.xlsx";
    
    data xel.have;
      input
        name$ hgt wgt;
    cards4;
    Alfred  69.0 112.5
    Alice   56.5  84.0
    Barbara 65.3  98.0
    Carol   62.8 102.5
    Henry   63.5 102.5
    James   57.3  83.0
    ;;;;
    run;quit;
    
    libname xel clear;
    
    INPUT
    =====
    
    EXISTING WORKBOOK
    
    d:\xls\formulasps.xlsx
    
    ----------------------+
    | A1| fx      | NAME  |
    ---------------  ---------+
    [_] |    A    |   B |   C |
    --------------------------|
     1  | NAME    |HGT  |WGT  |
     -- |---------+-----+-----|
     2  |  Alfred |69   |69   |
     -- |---------+-----+-----|
     3  |  Alice  |56.5 |56.5 |
     -- |---------+-----+-----|
     4  |  Barbara|65.3 |65.3 |
     -- |---------+-----+-----|
     5  |  Carol  |62.8 |62.8 |
     -- |---------+-----+-----|
     6  |  Henry  |63.5 |63.5 |
     -- |---------+-----+-----|
     7  |  James  |57.3 |57.3 |
     -------------------------|
    
    [HAVE]
    
    LOG:  16:54:16
    ==============
    
    NOTE: Copyright (c) 2026 Jenner Analytics Ltd., London, England.
    NOTE: Jenner v1.5.101 (build v1.5.101+21ec0ac00f.20261016T072007Z)
          Licensed to Roger DeAngelis, Serial 3BF7B3E6.
    NOTE: DATA _null_
    
    autexec started.
    
    NOTE: DATA elapsed:
      wall  0.00 seconds
      cpu   0.00 seconds
    NOTE: Library WORKX assigned path=d:\wpswrkx.
    NOTE: DATA workx.class
    
    NOTE: Processing inline DATALINES (19 lines)
    
    NOTE: Read 19 rows from DATALINES.
    NOTE: The data set WORKX.CLASS has 19 observations and 5 variables.
    NOTE: DATA elapsed:
      wall  0.00 seconds
      cpu   0.00 seconds
    NOTE: DATA _null_
    
    
    NOTE: DATA _null_ completed. Output written to FILE PRINT
    NOTE: Option SASAUTOS changed to c:/otojnr.
    NOTE: DATA _null_
    
    NOTE: Reading from fileref c:/jnr/runsas_selection.sas (c:/jnr/runsas_selection.sas)
        1 %utlfkil(d:/xls/class.xlsx)
        2
        3 libname xel excel "d:/xls/class.xlsx";
        4
        5 data xel.have;
        6   input
        7     name$ hgt wgt;
        8 cards4;
        9 Alfred  69.0 112.5
       10 Alice   56.5  84.0
       11 Barbara 65.3  98.0
       12 Carol   62.8 102.5
       13 Henry   63.5 102.5
       14 James   57.3  83.0
       15 ;;;;
       16 run;quit;
       17
       18 libname xel clear;
    
    NOTE: Read 18 rows from c:/jnr/runsas_selection.sas.
    NOTE: DATA elapsed:
      wall  0.00 seconds
      cpu   0.00 seconds
    NOTE: DATA _null_
    
    autoexec completed  16:54:16
    
    NOTE: DATA elapsed:
      wall  0.00 seconds
      cpu   0.00 seconds
    NOTE: Library XEL assigned path=d:/xls/class.xlsx.
    NOTE: DATA xel.have
    
    NOTE: Processing inline DATALINES (6 lines)
    
    NOTE: Read 6 rows from DATALINES.
    NOTE: The data set XEL.HAVE has 6 observations and 3 variables.
    NOTE: DATA elapsed:
      wall  0.00 seconds
      cpu   0.02 seconds
    NOTE: Library XEL cleared.
    
    /******************************************************************************************************************/
    /* 2  POWERSHELL ADD COLUMN TO SHEET HAVE                                                                         */
    /******************************************************************************************************************/
    
    %jnr_powershellBegin;
    cards4;
    # Open Excel and the workbook
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $workbook=$excel.Workbooks.Open("d:\xls\class.xlsx")
    $worksheet = $workbook.Sheets.Item("have")
    
    $worksheet.Cells.Item(1, 4) = "BMI"
    
    $lastRow = $worksheet.UsedRange.Rows.Count
    
    for ($row = 2; $row -le $lastRow; $row++) {
        $cell = $worksheet.Cells.Item($row, 4)
        $cell.Formula = "=C$row/(B$row^2)*703"
     }
    
    $workbook.Save()
    $workbook.Close()
    $excel.Quit()
    
    Set-Clipboard -Value "Hello from PowerShell"
    
    # Release COM objects - probablynot needed dor batch cli
    [System.Runtime.Interopservices.Marshal]::`
    ReleaseComObject($worksheet) | Out-Null
    [System.Runtime.Interopservices.Marshal]::`
    ReleaseComObject($workbook) | Out-Null
    [System.Runtime.Interopservices.Marshal]::`
    ReleaseComObject($excel) | Out-Null
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
    ;;;;
    %jnr_powershellEnd(return=greetings);
    
    
    OUTPUT OF POWERSHELL
    ====================
    
    d:\xls\formulasps.xlsx
    
    BMI added                *****
                             Added
    ------------------+      *****
    |A1|fx      | NAME|
    ------------------------------+
    [_]|   A    |  B  |  C  |  D  |
    ------------------------------|
     1 |NAME    | HGT | WGT | BMI |
     --|--------+-----+-----+-----|
     2 | Alfred | 69  | 69  |16.61|
     --|--------+-----+-----+-----|
     3 | Alice  | 56.5| 56.5|18.49|
     --|--------+-----+-----+-----|
     4 | Barbara| 65.3| 65.3|16.15|
     --|--------+-----+-----+-----|
     5 | Carol  | 62.8| 62.8|18.27|
     --|--------+-----+-----+-----|
     6 | Henry  | 63.5| 63.5|17.87|
     --|--------+-----+-----+-----|
     7 | James  | 57.3| 57.3|17.77|
    ------------------------------
    
    
    LOG:  16:56:53
    ==============
    
    NOTE: Copyright (c) 2026 Jenner Analytics Ltd., London, England.
    NOTE: Jenner v1.5.101 (build v1.5.101+21ec0ac00f.20261016T072007Z)
          Licensed to Roger DeAngelis, Serial 3BF7B3E6.
    NOTE: DATA _null_
    
    autexec started.
    
    NOTE: DATA elapsed:
      wall  0.00 seconds
      cpu   0.00 seconds
    NOTE: Library WORKX assigned path=d:\wpswrkx.
    NOTE: DATA workx.class
    
    NOTE: Processing inline DATALINES (19 lines)
    
    NOTE: Read 19 rows from DATALINES.
    NOTE: The data set WORKX.CLASS has 19 observations and 5 variables.
    NOTE: DATA elapsed:
      wall  0.00 seconds
      cpu   0.00 seconds
    NOTE: DATA _null_
    
    
    NOTE: DATA _null_ completed. Output written to FILE PRINT
    NOTE: Option SASAUTOS changed to c:/otojnr.
    NOTE: DATA _null_
    
    NOTE: Reading from fileref c:/jnr/runsas_selection.sas (c:/jnr/runsas_selection.sas)
        1 %jnr_powershellBegin;
        2 cards4;
        3 # Open Excel and the workbook
        4 $excel = New-Object -ComObject Excel.Application
        5 $excel.Visible = $false
        6 $workbook=$excel.Workbooks.Open("d:\xls\class.xlsx")
        7 $worksheet = $workbook.Sheets.Item("have")
        8
        9 $worksheet.Cells.Item(1, 4) = "BMI"
       10
       11 $lastRow = $worksheet.UsedRange.Rows.Count
       12
       13 for ($row = 2; $row -le $lastRow; $row++) {
       14     $cell = $worksheet.Cells.Item($row, 4)
       15     $cell.Formula = "=C$row/(B$row^2)*703"
       16  }
       17
       18 $workbook.Save()
       19 $workbook.Close()
       20 $excel.Quit()
       21
       22 Set-Clipboard -Value "Hello from PowerShell"
       23
       24 # Release COM objects - probablynot needed dor batch cli
       25 [System.Runtime.Interopservices.Marshal]::`
       26 ReleaseComObject($worksheet) | Out-Null
       27 [System.Runtime.Interopservices.Marshal]::`
       28 ReleaseComObject($workbook) | Out-Null
       29 [System.Runtime.Interopservices.Marshal]::`
       30 ReleaseComObject($excel) | Out-Null
       31 [System.GC]::Collect()
       32 [System.GC]::WaitForPendingFinalizers()
       33 ;;;;
       34 %jnr_powershellEnd(return=greetings);
    
    NOTE: Read 34 rows from c:/jnr/runsas_selection.sas.
    NOTE: DATA elapsed:
      wall  0.00 seconds
      cpu   0.02 seconds
    NOTE: DATA _null_
    
    autoexec completed  16:56:53
    
    NOTE: DATA elapsed:
      wall  0.00 seconds
      cpu   0.00 seconds
    NOTE: DATA _null_
    
    NOTE: Writing to fileref D:\wpswrk/powershell_resolve.ps1 (D:\wpswrk/powershell_resolve.ps1)
    NOTE: Processing inline DATALINES (30 lines)
    # Open Excel and the workbook
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $workbook=$excel.Workbooks.Open("d:\xls\class.xlsx")
    $worksheet = $workbook.Sheets.Item("have")
    
    $worksheet.Cells.Item(1, 4) = "BMI"
    
    $lastRow = $worksheet.UsedRange.Rows.Count
    
    for ($row = 2; $row -le $lastRow; $row++) {
        $cell = $worksheet.Cells.Item($row, 4)
        $cell.Formula = "=C$row/(B$row^2)*703"
     }
    
    $workbook.Save()
    $workbook.Close()
    $excel.Quit()
    
    Set-Clipboard -Value "Hello from PowerShell"
    
    # Release COM objects - probablynot needed dor batch cli
    [System.Runtime.Interopservices.Marshal]::`
    ReleaseComObject($worksheet) | Out-Null
    [System.Runtime.Interopservices.Marshal]::`
    ReleaseComObject($workbook) | Out-Null
    [System.Runtime.Interopservices.Marshal]::`
    ReleaseComObject($excel) | Out-Null
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
    
    NOTE: Read 30 rows from DATALINES.
    NOTE: DATA elapsed:
      wall  0.00 seconds
      cpu   0.00 seconds
    WARNING: The NOXWAIT system option does not apply to Jenner and has no effect.
    WARNING 30-12: The NOXSYNC system option is valid only at startup and was ignored.
    NOTE: Fileref RUT assigned to device PIPE (payload: powershell.exe -executionpolicy bypass -file D:\wpswrk/powershell_resolve.ps1).
    NOTE: DATA _null_
    
    NOTE: Reading from fileref rut (D:\wpswrk\jenner_pipe_RUT_48a4e86c65ae475082542b251b9b6a86.tmp)
    
    NOTE: Read 0 rows from rut.
    NOTE: DATA elapsed:
      wall  0.00 seconds
      cpu   0.00 seconds
    NOTE: Fileref RUT cleared.
    NOTE: Fileref CLIP assigned to device CLIPBRD (payload: (none)).
    NOTE: DATA _null_
    
    NOTE: Reading from fileref clip (D:\wpswrk\jenner_clipbrd_CLIP_4fc1021366e94d42bdd3436c267f4eff.tmp)
    *******  Hello from PowerShell   ********
    
    NOTE: Read 1 rows from clip.
    NOTE: DATA elapsed:
      wall  0.00 seconds
      cpu   0.00 seconds
    
    
    /******************************************************************************************************************/
    /* 3  JENNER PRINT WORKSHEET HAVE POWERSHELL CREATED COLUMN BMI                                                   */
    /*    JENNER PRINT MACRO VARIABLE GREETINGS CREATED Y POWERSHELL7  RELATED REOSITORIES                            */
    /******************************************************************************************************************/
    
    libname xel excel "d:/xls/class.xlsx";
    
    %put --- &=greetings ---;
    
    options nocenter;
    proc print data=xel.have;
    title "workbook and worksheet from powershell";
    run;
    
    libname xel clear;
    
    
    JENNER OUTPUT
    
    WORKBOOK AND WORKSHEET FROM POWERSHELL
    
      Obs  name      hgt    wgt            BMI
    
        1  Alfred     69  112.5  16.6115311909
        2  Alice    56.5     84  18.4985511786
        3  Barbara  65.3     98  16.1567884355
        4  Carol    62.8  102.5  18.2708984137
        5  Henry    63.5  102.5  17.8702957406
        6  James    57.3     83  17.7715035833
    
    
    LOG:  16:58:22
    ==============
    
    NOTE: Copyright (c) 2026 Jenner Analytics Ltd., London, England.
    NOTE: Jenner v1.5.101 (build v1.5.101+21ec0ac00f.20261016T072007Z)
          Licensed to Roger DeAngelis, Serial 3BF7B3E6.
    NOTE: DATA _null_
    
    autexec started.
    
    NOTE: DATA elapsed:
      wall  0.00 seconds
      cpu   0.00 seconds
    NOTE: Library WORKX assigned path=d:\wpswrkx.
    NOTE: DATA workx.class
    
    NOTE: Processing inline DATALINES (19 lines)
    
    NOTE: Read 19 rows from DATALINES.
    NOTE: The data set WORKX.CLASS has 19 observations and 5 variables.
    NOTE: DATA elapsed:
      wall  0.00 seconds
      cpu   0.00 seconds
    NOTE: DATA _null_
    
    
    NOTE: DATA _null_ completed. Output written to FILE PRINT
    NOTE: Option SASAUTOS changed to c:/otojnr.
    NOTE: DATA _null_
    
    NOTE: Reading from fileref c:/jnr/runsas_selection.sas (c:/jnr/runsas_selection.sas)
        1 libname xel excel "d:/xls/class.xlsx";
        2
        3 %put --- &=greetings ---;
        4
        5 options nocenter;
        6 proc print data=xel.have;
        7 title "workbook and worksheet from powershell";
        8 run;
        9
       10 libname xel clear;
    
    NOTE: Read 10 rows from c:/jnr/runsas_selection.sas.
    NOTE: DATA elapsed:
      wall  0.00 seconds
      cpu   0.00 seconds
    NOTE: DATA _null_
    
    autoexec completed  16:58:22
    
    NOTE: DATA elapsed:
      wall  0.00 seconds
      cpu   0.00 seconds
    NOTE: Library XEL assigned path=d:/xls/class.xlsx.
    WARNING: Apparent symbolic reference GREETINGS not resolved.
    --- GREETINGS=&greetings ---
    NOTE: PROC PRINT data=xel.have
    
    NOTE: PROC PRINT completed: 6 observations printed, 4 variables
    NOTE: Library XEL cleared.
    
    
    /******************************************************************************************************************/
    /* 4  RELATED POWERSHELL REPOSITORIES                                                                             */
    /******************************************************************************************************************/
    
    repo
    ------------------------------------------------------------------------------------------------------------------------------------
    https://github.com/rogerjdeangelis/utl-altair-slc-drop-down-to-powershell-and-programatically-add-and-execute-an-excel-macro
    https://github.com/rogerjdeangelis/utl-get-text-file-attributes-and-window-file-properties-powershell-python-sas
    https://github.com/rogerjdeangelis/utl_dropping-down-to-powershell-and-converting-doc-and-rtf-files-to-pdfs
    https://github.com/rogerjdeangelis/utl-powershell-unzip-one-meber-of-a-winzip-archive
    https://github.com/rogerjdeangelis/utl-drop-down-to-powershell-and-programatically-create-an-odbc-data-source-for-excel-wps-r-rodbc
    https://github.com/rogerjdeangelis/utl-dropping-down-to-powershell-from-personal-altair-slc
    https://github.com/rogerjdeangelis/utl-export-excel-sheet-to-sas-table-using-one-line-macro-invocation-without-sas-access-powershell
    https://github.com/rogerjdeangelis/utl-scraping-AI-results-without-restriction-or-API-with-powershell-and-perplexity
    https://github.com/rogerjdeangelis/utl-altair-slc-language8-powershell-read-write-text-files-summarize-and-looping-with-sqlite
    https://github.com/rogerjdeangelis/utl-create-ms-access-mdb-table-from-sas-table-without-sas-access-using-powershell
    https://github.com/rogerjdeangelis/utl-read-print-file-backwards-in-perl-powershell-sas-r-and-python
    https://github.com/rogerjdeangelis/utl-add-formula-to-an-existing-excel-worksheet-using-powershell
    https://github.com/rogerjdeangelis/utl-altair-slc-add-column-bmi-to-an-existing-excel-worksheet-using-powershell
    https://github.com/rogerjdeangelis/utl-altair-slc-powershell-calculating-and-validating-file-hash-using-different-algorithms
    https://github.com/rogerjdeangelis/utl-drop-down-using-dosubl-from-sas-datastep-to-wps-r-perl-powershell-python-msr-vb
    https://github.com/rogerjdeangelis/utl-examples-of-drop-downs-from-sas-to-wps-r-microsoftR-python-perl-powershell
    https://github.com/rogerjdeangelis/utl-importing-a-sas-created-rtf-file-into-ms-word-using-powershell
    https://github.com/rogerjdeangelis/utl-export-ms-access-table-from-mdb-database-to-sas-table-without-sas-access-using-powershell
    https://github.com/rogerjdeangelis/utl-altair-slc-caculate-the-sha256-or-sha512-checksum-of-a-file-powershell
    https://github.com/rogerjdeangelis/utl-how-to-remove-a-record-from-a-csv-file-in-place-with-powershell-fast
    https://github.com/rogerjdeangelis/utl-drop-down-to-powershell-and-return-the-number-of-lines-in-a-file-to-wps
    https://github.com/rogerjdeangelis/zip-and-unzip-using-ms-powershell
    
    
    /******************************************************************************************************************/
    /* 5  Inline drop down macros                                                                                     */
    /******************************************************************************************************************/
    
    filename ft15f001 "c:/otojnr/jnr_powershellBegin.sas";
    parmcards4;
    %macro jnr_powershellBegin(
        resolve=Y
        )/des="No additional quote(backtic) because powershell needs it";
    
      %local pth;
      %let pth=%sysfunc(pathname(work));
    
      %utlfkil(&pth/powershell_resolve.ps1);
    
      data _null_;
        length _infile_ $255;
        file "&pth/powershell_resolve.ps1";
        input;
        if upcase("&resolve")=:"Y" then _infile_=resolve(_infile_);
        put _infile_;
        putlog _infile_;
    
    %mend jnr_powershellBegin;
    ;;;;
    run;
    
    filename ft15f001 "c:/otojnr/jnr_powershellEnd.sas";
    parmcards4;
    %macro jnr_powershellEnd(
        return=
       )/des="execute powershell script";
    run;
      %local pth;
      %let pth=%sysfunc(pathname(work));
    
      options noxwait noxsync;
      filename rut pipe "powershell.exe -executionpolicy bypass -file &pth/powershell_resolve.ps1";
    run;quit;
      run;quit;
      data _null_;
        infile rut;
        input;
        putlog _infile_;
        file print;
        put _infile_;
      run;
    
      filename rut clear;
    
      * use the clipboard to create macro variable;
      %if "&return" ^= "" %then %do;
        filename clip clipbrd;
        data _null_;
         infile clip;  /*--- created from the SPSS script ---*/
         input;
         putlog "*******  " _infile_ "  ********";
         call symputx("&return",_infile_,"G");
        run;quit;
      %end;
    
    %mend jnr_powershellEnd;
    ;;;;
    run;
    
    
    LOG:  16:52:50
    ==============
    
    NOTE: Copyright (c) 2026 Jenner Analytics Ltd., London, England.
    NOTE: Jenner v1.5.101 (build v1.5.101+21ec0ac00f.20261016T072007Z)
          Licensed to Roger DeAngelis, Serial 3BF7B3E6.
    NOTE: DATA _null_
    
    autexec started.
    
    NOTE: DATA elapsed:
      wall  0.00 seconds
      cpu   0.00 seconds
    NOTE: Library WORKX assigned path=d:\wpswrkx.
    NOTE: DATA workx.class
    
    NOTE: Processing inline DATALINES (19 lines)
    
    NOTE: Read 19 rows from DATALINES.
    NOTE: The data set WORKX.CLASS has 19 observations and 5 variables.
    NOTE: DATA elapsed:
      wall  0.00 seconds
      cpu   0.00 seconds
    NOTE: DATA _null_
    
    
    NOTE: DATA _null_ completed. Output written to FILE PRINT
    NOTE: Option SASAUTOS changed to c:/otojnr.
    NOTE: DATA _null_
    
    NOTE: Reading from fileref c:/jnr/runsas_selection.sas (c:/jnr/runsas_selection.sas)
        1 filename ft15f001 "c:/otojnr/jnr_powershellBegin.sas";
        2 parmcards4;
        3 %macro jnr_powershellBegin(
        4     resolve=Y
        5     )/des="No additional quote(backtic) because powershell needs it";
        6
        7   %local pth;
        8   %let pth=%sysfunc(pathname(work));
        9
       10   %utlfkil(&pth/powershell_resolve.ps1);
       11
       12   data _null_;
       13     length _infile_ $255;
       14     file "&pth/powershell_resolve.ps1";
       15     input;
       16     if upcase("&resolve")=:"Y" then _infile_=resolve(_infile_);
       17     put _infile_;
       18     putlog _infile_;
       19
       20 %mend jnr_powershellBegin;
       21 ;;;;
       22 run;
       23
       24 filename ft15f001 "c:/otojnr/jnr_powershellEnd.sas";
       25 parmcards4;
       26 %macro jnr_powershellEnd(
       27     return=
       28    )/des="execute powershell script";
       29 run;
       30   %local pth;
       31   %let pth=%sysfunc(pathname(work));
       32
       33   options noxwait noxsync;
       34   filename rut pipe "powershell.exe -executionpolicy bypass -file &pth/powershell_resolve.ps1";
       35 run;quit;
       36   run;quit;
       37   data _null_;
       38     infile rut;
       39     input;
       40     putlog _infile_;
       41     file print;
       42     put _infile_;
       43   run;
       44
       45   filename rut clear;
       46
       47   * use the clipboard to create macro variable;
       48   %if "&return" ^= "" %then %do;
       49     filename clip clipbrd;
       50     data _null_;
       51      infile clip;  /*--- created from the SPSS script ---*/
       52      input;
       53      putlog "*******  " _infile_ "  ********";
       54      call symputx("&return",_infile_,"G");
       55     run;quit;
       56   %end;
       57
       58 %mend jnr_powershellEnd;
       59 ;;;;
       60 run;
    
    NOTE: Read 60 rows from c:/jnr/runsas_selection.sas.
    NOTE: DATA elapsed:
      wall  0.00 seconds
      cpu   0.00 seconds
    NOTE: DATA _null_
    
    autoexec completed  16:52:50
    
    NOTE: DATA elapsed:
      wall  0.00 seconds
      cpu   0.00 seconds
    NOTE: Fileref FT15F001 assigned to c:/otojnr/jnr_powershellBegin.sas.
    NOTE: PARMCARDS4 wrote 18 record(s) to fileref FT15F001.
    NOTE: Fileref FT15F001 assigned to c:/otojnr/jnr_powershellEnd.sas.
    NOTE: PARMCARDS4 wrote 33 record(s) to fileref FT15F001.
    
    
    /*--- end ---*/# jnr-jenner-analytics-part-IV-inline-drop-down-to-poweshell-add-column-to-excell-sheet
