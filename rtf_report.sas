/*******************************************************************************
|
| Program Name:   qu_rtf_report.sas
| Adapted From:   rtf_report.sas
| Program Version: 1.0
|
| Program Purpose: To create a Macro which produces Simple Report Tool
|
| SAS Version:  9.3
|
| Created By:   Thomas Drury
| Date:         16-11-16   
|
|--------------------------------------------------------------------------------
| Output: Printed output.
|--------------------------------------------------------------------------------
| Global macro variables created: NONE
|
| Local macros parameters:
|
|
|--------------------------------------------------------------------------------
| Change Log
|
| Modified By:               
| Date of Modification:      
| New version number:        
| Modification ID:           
| Reason For Modification:    
| 
********************************************************************************/

********************************************************************************;
***                          CREATE REPORT MACRO                             ***;
********************************************************************************; 

%macro rtf_report(indata            = 
                 ,outfile           =
                 ,outtemplate       =
                 ,topleftvars       = 
                 ,topleftlabels     = 
                 ,columnvars        = 
                 ,columnwidths      = 
                 ,columnlabels      = 
                 ,grouplabels       = 
                 ,ordervars         = 
                 ,noprintvars       = 
                 ,skipvars          = 
                 ,pagevars          =
                 ,alignheaders      = 
                 ,aligncolumns      = 
                 ,header1           =
                 ,header2           = 
                 ,header3           =
                 ,header4           = 
                 ,footer1           =
                 ,footer1           =
                 ,footer1           =
                 ,footer1           =
                 ,fonttype          =%str(Courier New)
                 ,fontsize          =10
                 )/minoperator;

  %*******************************************************************************;
  %*** MACRO SECTION 1: CREATE MACRO PARAMETERS AND BASIC ERROR CHECKS         ***;
  %*******************************************************************************;


  %if %length(&columnvars.) = 0 %then %do;
   %put ER%upcase(ror): (RTF_REPORT MACRO): No variables specified in the COLUMNS parameter.;
   %abort cancel;
  %end;



  %************************************;
  %***  1.1: COUNT LIST PARAMETERS  ***;
  %************************************;


  %let topleftvars_n   = %sysfunc(countw(&topleftvars.,%str(,))); 
  %let topleftlabels_n = %sysfunc(countw(&topleftlabels.,%str(,))); 
  %let columnvars_n    = %sysfunc(countw(&columnvars.,%str(,))); 
  %let columnwidths_n  = %sysfunc(countw(&columnwidths.,%str(,))); 
  %let columnlabels_n  = %sysfunc(countw(&columnlabels.,%str(,))); 
  %let grouplabels_n   = %sysfunc(countw(&grouplabels.,%str(,))); 
  %let ordervars_n     = %sysfunc(countw(&ordervars.,%str(,)));
  %let noprintvars_n   = %sysfunc(countw(&noprintvars.,%str(,))); 
  %let skipvars_n      = %sysfunc(countw(&skipvars.,%str(,)));
  %let pagevars_n      = %sysfunc(countw(&pagevars.,%str(,)));
  %let alignheaders_n  = %sysfunc(countw(&alignheaders.,%str(,))); 
  %let aligncolumns_n  = %sysfunc(countw(&aligncolumns.,%str(,)));

  %**********************************************;
  %***  1.2: CREATE ALL PARAMETERS AND LISTS  ***;
  %**********************************************;


  %*** CREATE INDIVDUAL TOPLEFTVAR NAMES AND LISTS ***;
  %let topleftvarlist=;
  %let topleftnumvarslist=;
  %let dsopen = %sysfunc(open(&indata.,i));

  %do ii = 1 %to &topleftvars_n;
   %let topleftvars&ii. = %scan(&topleftvars.,&ii.,%str(,));
   %let topleftvarlist = &topleftvarlist. &&topleftvars&ii.;
   %let varnum = %sysfunc(varnum(&dsopen.,&&topleftvars&ii.));
   %if (%sysfunc(vartype(&dsopen.,&varnum.)) = N) %then %let topleftnumvarslist = &topleftnumvarslist. &&topleftvars&ii.;
  %end;

   %put &topleftnumvarslist.;
   %put &topleftvarlist.;

  %let rc = %sysfunc(close(&dsopen.));


  %*** CREATE INDIVDUAL TOPLEFTVAR LABELS ***;
  %do ii = 1 %to &topleftlabels_n;
   %let topleftlabelterm&ii.  = %scan(&topleftlabels.,&ii.,%str(,));
   %let topleftlabelvar&ii.   = %scan(&&topleftlabelterm&ii.,1,%str(=));
   %let topleftlabelvalue&ii. = %scan(&&topleftlabelterm&ii.,2,%str(=));
  %end;


  %*** CREATE INDIVIDUAL COLUMN NAMES AND LIST ***;
  %let columnlist=;
  %do ii = 1 %to &columnvars_n;
   %let columnvars&ii. = %scan(&columnvars.,&ii.,%str(,));
   %let columnlist = &columnlist. &&columnvars&ii.;
  %end;


  %*** CREATE INDIVIDUAL COLUMN LABELS AND LIST ***;
  %let labelvarlist =;
  %let labellist    =;
  %do ii = 1 %to &columnlabels_n;
    %let labelterm&ii.  = %sysfunc(scan(&columnlabels.,&ii.,%str(,)));
    %let labelvarlist   = &labelvarlist. %scan(&&labelterm&ii.,1,%str(=));
    %let labelvalue&ii. = %scan(&&labelterm&ii.,2,%str(=));
  %end;

   
  %*** CREATE VARS FOR GROUP LABELS AND GROUP VARNAMES AND VARLIST ***;
  %do ii = 1 %to &grouplabels_n;

   %let groupstatement&ii. = %scan(&grouplabels.,&ii.,%str(,));
   %let grouplabel&ii.     = %scan(&&groupstatement&ii.,2,%str(""''));
   %let grouplabelvars&ii. = %scan(&&groupstatement&ii.,-1,%str(""''()));

   %let grouplabelvars&ii._n = %sysfunc(countw(&&grouplabelvars&ii.,%str( )));

   %let grouplabellist&ii. =;
   %do jj = 1 %to &&grouplabelvars&ii._n;
    %let grouplabelvar&ii.&jj. = %scan(&&grouplabelvars&ii.,&jj., %str( ));
    %let grouplabellist&ii.    =  &&grouplabellist&ii. &&grouplabelvar&ii.&jj.;
   %end;

  %end;


  %*** CREATE INDIVDUAL ORDERVAR NAMES AND LIST ***;
  %let ordervarlist=;
  %do ii = 1 %to &ordervars_n;
   %let ordervarlist = &ordervarlist. %scan(&ordervars.,&ii.,%str(,));
  %end;


  %*** CREATE MACRO VARS FOR NOPRINTING AND LIST ***;
  %let noprintlist=;
  %do ii = 1 %to &noprintvars_n;
   %let noprintlist = &noprintlist. %scan(&noprintvars.,&ii.,%str(,));
  %end;


  %*** CREATE MACRO VARS FOR PAGING     ***;
  %do ii = 1 %to &pagevars_n;
   %let pagevar&ii. = %scan(&pagevars.,&ii.,%str(,));
  %end;


  %*** CREATE MACRO VARS FOR SKIPPING   ***;
  %do ii = 1 %to &skipvars_n;
   %let skipvar&ii. = %scan(&skipvars.,&ii.,%str(,));
  %end;


  %*** CREATE MACRO VARS FOR HEADER ALIGNMENT ***;
  %let alignheadervarlist =;
  %let alignheaderlist    =;
  %do ii = 1 %to &alignheaders_n;
    %let alignheaderterm&ii.  = %sysfunc(scan(&alignheaders.,&ii.,%str(,)));
    %let alignheadervarlist   = &alignheadervarlist. %scan(&&alignheaderterm&ii.,1,%str(=));
    %let alignheaderlist      = &alignheaderlist. %scan(&&alignheaderterm&ii.,2,%str(=));
  %end;


  %*** CREATE MACRO VARS FOR COLUMN ALIGNMENT ***;
  %let aligncolumnvarlist =;
  %let aligncolumnlist    =;
  %do ii = 1 %to &aligncolumns_n;
    %let aligncolumnterm&ii.  = %sysfunc(scan(&aligncolumns.,&ii.,%str(,)));
    %let aligncolumnvarlist   = &aligncolumnvarlist. %scan(&&aligncolumnterm&ii.,1,%str(=));
    %let aligncolumnlist      = &aligncolumnlist. %scan(&&aligncolumnterm&ii.,2,%str(=));
  %end;

  %*** CHECK IF ALL THE VARIABLES IN THE LIST ARE IN COLUMNS STATEMENT ***;



  %**************************************************;
  %***  1.3: BUILD PROC REPORT COLUMNS STATEMENT  ***;
  %**************************************************;
  
  %*** IF GROUP LABELS SPECIFIED ***;
  %if %length(&grouplabels.) ne 0 %then %do;
  
    %*** CHECK VARIABLES EXIST IN COLUMN STATEMENT ***;
    %let mismatches=; 
    %do ii = 1 %to &grouplabels_n;
    %do jj = 1 %to &&grouplabelvars&ii._n;
     %if not (&&grouplabelvar&ii.&jj. in &columnlist.) %then %do;
      %let mismatches= &mismatches. &&grouplabelvar&ii.&jj.;
     %end;
    %end;
    %end;
  
    %if %length(&mismatches.) ne 0 %then %do;
       %put ER%upcase(ror): (RTF_REPORT MACRO):;
       %put ER%upcase(ror): (RTF_REPORT MACRO): Parameters in GROUPLABEL (&mismatches.) not specified in COLUMNVARS.;
       %put ER%upcase(ror): (RTF_REPORT MACRO):;
       %abort cancel;
    %end;

    %*** CHECK FOR EACH GROUPING THAT ALL THE GROUPED VARS ARE NEXT TO EACH OTHER IN THE COLUMN STATEMENT ***;
    %let ordererr=;
    %do ii = 1 %to &grouplabels_n;
     %do jj = 1 %to &&grouplabelvars&ii._n;
      %do kk = 1 %to &columnvars_n;
       %if &&grouplabelvar&ii.&jj. = &&columnvars&kk. %then %do;
          %let v&ii.&jj. = &kk.;
       %end;
      %end;
     %end;
     %do jj = 1 %to %eval(&&grouplabelvars&ii._n - 1);
      %let kk = %eval(&jj.+1);
      %if &&v&ii.&kk. ne &&v&ii.&jj. + 1 %then %do;
         %let ordererr = &ordererr. &&grouplabelvar&ii.&jj. (columnvar position = &&v&ii.&jj.) &&grouplabelvar&ii.&kk. (columnvar position = &&v&ii.&kk.);
      %end;
     %end;
    %end;

    %if %length(&ordererr.) ne 0 %then %do;
       %put ER%upcase(ror): (RTF_REPORT MACRO):;
       %put ER%upcase(ror): (RTF_REPORT MACRO): Parameters in GROUPLABEL not in the same order as specified in COLUMNVARS.;
       %put ER%upcase(ror): (RTF_REPORT MACRO): Check order in both GROUPLABEL and COLUMNVARS.;
       %put ER%upcase(ror): (RTF_REPORT MACRO): Variables affected: &ordererr.;
       %put ER%upcase(ror): (RTF_REPORT MACRO):;
       %abort cancel;
    %end;

  %end;

  %*** BUILD PROC REPORT COLUMNS STATEMENT ***;
  %let columnstatement =;
  %do ii = 1 %to &columnvars_n;
   %let newterm = &&columnvars&ii.;
   %do jj = 1 %to &grouplabels_n;
    %let kk = &&grouplabelvars&jj._n;
    %if  &&columnvars&ii. = &&grouplabelvar&jj.1 %then %do;
      %let newterm = ("^S={borderbottomcolor=black borderbottomwidth=1} &&grouplabel&jj." &&columnvars&ii.  ;
    %end;
    %else %if &&columnvars&ii. = &&grouplabelvar&jj.&kk. %then %do;
      %let newterm = &&columnvars&ii. );
    %end;
   %end;
   %let columnstatement = &columnstatement. &newterm.;
  %end;


  %**************************************************;
  %***  1.4: BUILD PROC REPORT DEFINE STATEMENT   ***;
  %**************************************************;

  %do ii = 1 %to &columnvars_n;

    %*** CREATE NOPRINT TEXT ***;
    %let noprint&ii=;
    %if &&columnvars&ii. in &noprintlist. %then %let noprint&ii. = NOPRINT;

    %*** CREATE ORDER TYPE TEXT (DISPLAY OR ORDER=DATA) ***;
    %let order&ii = DISPLAY;
    %if &&columnvars&ii. in &ordervarlist. %then %let order&ii. = %str(ORDER ORDER=DATA);

    %*** CREATE LABEL TEXT ***;
    %let label&ii. =;
    %do jj = 1 %to &columnlabels_n.;
      %if &&columnvars&ii. = %scan(&labelvarlist.,&jj.) %then %let label&ii. = &&labelvalue&jj.;
    %end;

    %*** CREATE HEADER ALIGNMENTS ***;
    %let alignheader&ii. = style(header)=[just=L];
    %do jj = 1 %to &alignheaders_n.;
      %if &&columnvars&ii. = %scan(&alignheadervarlist.,&jj.) %then %let alignheader&ii. = style(header)=[just=%scan(&alignheaderlist.,&jj.)];
    %end;

    %*** CREATE COLUMN ALIGNMENTS ***;
    %let aligncolumn&ii. = style(column)=[just=D asis=on];
    %do jj = 1 %to &aligncolumns_n.;
      %if &&columnvars&ii. = %scan(&aligncolumnvarlist.,&jj.) %then %let aligncolumn&ii. = style(column)=[just=%scan(&aligncolumnlist.,&jj.) asis=on];
    %end;

    %*** CREATE FULL DEFINE STATEMENT FOR EACH COLUMN VARIABLE ***;
    %let definestatement&ii. = &&columnvars&ii. / &&order&ii. &&noprint&ii. &&alignheader&ii. &&aligncolumn&ii. &&label&ii.;

  %end;


  %**************************************************;
  %***  1.5: BUILD PROC REPORT COMPUTE BLOCKS     ***;
  %**************************************************;

  %do ii = 1 %to &topleftvars_n.;
  
    %let valuetxt&ii. = put(&&topleftvars&ii.,$50.);
    %if %length(&topleftnumvarslist.) ne 0 %then %do;
     %if (&&topleftvars&ii. in &topleftnumvarslist.) %then %let valuetxt&ii. = put(&&topleftvars&ii.,8.2);
    %end;

    %do jj = 1 %to &topleftlabels_n.;
      %if &&topleftvars&ii. = &&topleftlabelvar&jj. %then %let valuetxt&ii. = &&topleftlabelvalue&jj.||&&valuetxt&ii.;
    %end;

    %let topleftcomputeblock&ii. = %str(valuetxt&ii. = &&valuetxt&ii.; line @1 valuetxt&ii. $50.;);

  %end;

  %*******************************************************************************;
  %*** MACRO SECTION 2: SET UP HEADERS AND FOOTERS                              ***;
  %*******************************************************************************;

  %if %length(&header1.) ne 0 %then %do; title4 "&header1." %end;
  %if %length(&header2.) ne 0 %then %do; title5 "&header2." %end;
  %if %length(&header3.) ne 0 %then %do; title6 "&header3." %end;
  %if %length(&header4.) ne 0 %then %do; title7 "&header4." %end;

  %if %length(&footer1.) ne 0 %then %do; footnote1 "&footer1." %end;
  %if %length(&footer2.) ne 0 %then %do; footnote2 "&footer2." %end;
  %if %length(&footer3.) ne 0 %then %do; footnote3 "&footer3." %end;
  %if %length(&footer4.) ne 0 %then %do; footnote4 "&footer4." %end;


  *******************************************************************************;
  *** CODE SECTION 1: CREATE RTF STYLE SIMILAR TO PHARMA REPORTING STANDARDS  ***;
  *******************************************************************************;

  ods escapechar="^";
  options orientation = landscape;

  proc template ;
    define style template_1; 
      parent = styles.rtf ; 
      replace fonts / 
        'TitleFont'           = (&fonttype.,&fontsize.pt) /* TITLE statement */
        'TitleFont2'          = (&fonttype.,&fontsize.pt) /* PROC titles */
        'headingFont'         = (&fonttype.,&fontsize.pt) /* Table column/row headings */
        'docFont'             = (&fonttype.,&fontsize.pt) /* data in table cells */
        'footFont'            = (&fonttype.,&fontsize.pt) /* FOOTNOTE statements */
        'StrongFont'          = (&fonttype.,&fontsize.pt)
        'EmphasisFont'        = (&fonttype.,&fontsize.pt)
        'headingEmphasisFont' = (&fonttype.,&fontsize.pt)
        'FixedFont'           = (&fonttype.,&fontsize.pt)
        'FixedEmphasisFont'   = (&fonttype.,&fontsize.pt)
        'FixedStrongFont'     = (&fonttype.,&fontsize.pt)
        'FixedHeadingFont'    = (&fonttype.,&fontsize.pt)
        'BatchFixedFont'      = (&fonttype.,&fontsize.pt);
      style table from table /
        background  = _UNDEF_ /* REMOVES TABLE BACKGROUND COLOR */
        rules       = groups  /* INTERNAL BORDERS: SET TO BOTTOM BORDER ON ROW HEADERS */
        frame       = above   /* EXTERNAL BORDERS: SET TO TOP LINE OF TABLE ONLY */
        cellspacing = 0       /* SPACE BETWEEN TABLE CELLS */
        cellpadding = 0       /* REMOVES PARAGRAPH SPACING BEFORE/AFTER CELL CONTENTS */
        borderwidth = 1pt;    /* SET WIDTH OF BORDER IN FRAME= */ 
      replace HeadersAndFooters from Cell /
        background = _undef_
        font       = Fonts('TitleFont') ;
      replace SystemFooter from TitlesAndFooters /
        font = Fonts('footFont')
        just = LEFT; 
      replace Body from Document /
        bottommargin = .75in
        topmargin    = 1in
        rightmargin  = 1in
        leftmargin   = 1in; 
    end ; 
  run ; 


  *******************************************************************************;
  *** CODE SECTION 2: READ IN DATA                                            ***;
  *******************************************************************************;

  *** READ IN REPORT DATASET ***;
  proc sort data = &indata.
            out  = r_data;
   by &topleftvarlist. &ordervarlist.;
  run;


  *******************************************************************************;
  *** SECTION 3: CREATE PROC REPORT                                           ***;
  *******************************************************************************;

  *** SET UP OPTIONS AND ODS ***;
  options mprint 
          nobyline 
          orientation = landscape 
          missing = '';

  ods rtf file="&outfile." style=template_1 startpage=yes;
  ods results off;
  ods listing close;

  proc report data    = r_data nowd 
              spacing = 1 
              split   = "#" 
              headskip
              missing
              ;

    %if %length(&topleftvars.) ne 0 %then %do; 
    by &topleftvarlist.; 

    compute before _page_; 
    %do ii = 1 %to &topleftvars_n.;
    &&topleftcomputeblock&ii.;
    %end;
    line @1 " ";
    endcomp;

    %end;

    column &topleftvarlist. &columnstatement.;

    %do ii = 1 %to &topleftvars_n.;
    define &&topleftvars&ii. / order order=data noprint;
    %end;

    %do ii = 1 %to &columnvars_n;
    define &&definestatement&ii.;
    %end;

    %do ii = 1 %to &skipvars_n.;
    break after &&skipvar&ii. / skip;
    compute after &&skipvar&ii. ;
     line ' ' ;
    endcomp;
    %end;

    %do ii = 1 %to &pagevars_n.;
    break after &&pagevar&ii. / page;
    %end;

  run;

  ods results;
  ods rtf close;
  ods listing;
  options nomprint;

  
  %*******************************************************************************;
  %***                          CLEAN UP DATASETS                              ***;
  %*******************************************************************************;

  *** DELETE TEMPORARY WORK DATASETS ***;
  proc datasets lib = work nolist;
    delete r_:;
  quit;
  run;

%mend;

























