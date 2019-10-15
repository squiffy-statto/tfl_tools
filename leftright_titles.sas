/* Read text string from DATALINES.  To right justify the string,     */
/* position it in a character string so that the last byte is at      */
/* the last position of the linesize setting from the LINESIZE=       */
/* system option.  The character variable is then copied into a macro */
/* variable R_NOTE.  Character strings align to the left by default,  */
/* so create the macro variable L_NOTE without adding any leading     */
/* spaces.                                                            */
/*                                                                    */
/* Note:  The length of the NOTE variable should be at least the      */
/* length of the LINESIZE system variable.  The informat used to read */
/* the NOTE variable in the INPUT statement must be long enough       */
/* to include the entire text of the title or footnote.               */

options nonumber nodate nosymbolgen center ps=30 ls=100;

data _null_;
  infile datalines truncover;
  length jnote note $100;
  input note $50. ;

  /* Determine the number of spaces needed to pad NOTE to right justify */
  /* the value of NOTE                                                  */
  spaces=(input(getoption('ls'),8.) - length(note))-1;

  /* Pad NOTE and assign the result into JNOTE */
  jnote=repeat(' ',spaces)||note;

  /* Create a macro variable to resolve in a left justified TITLE  */
  /* and a macro variable to resolve in a right justified FOOTNOTE */
  call symput('r_note',jnote);
  call symput('l_note',note);
datalines;
This is the text string I want to justify
;

title "&l_note";
footnote "&r_note";

proc print data=sashelp.class;
run;
