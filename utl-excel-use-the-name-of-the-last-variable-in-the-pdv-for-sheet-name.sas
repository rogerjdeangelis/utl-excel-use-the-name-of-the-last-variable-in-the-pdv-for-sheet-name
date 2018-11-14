Excel use the name of the last variable in the pdv for sheet name

github
https://tinyurl.com/y9s3y5dw
https://github.com/rogerjdeangelis/utl-excel-use-the-name-of-the-last-variable-in-the-pdv-for-sheet-name

Other excel repos
https://tinyurl.com/ybnm6azh
https://github.com/rogerjdeangelis?utf8=%E2%9C%93&tab=repositories&q=excel+in%3Aname&type=&language=

StackOverflow
https://tinyurl.com/y7tnbxg5
https://stackoverflow.com/questions/53253745/sas-export-subset-of-column-to-worksheet-with-a-column-name

SQl dictionaries are often to slow on an EG server so I suggest you
use  proc contents.


INPUT
=====

 SASHELP.CLASS total obs=19

  NAME       SEX    AGE    HEIGHT    WEIGHT

  Alfred      M      14     69.0      112.5
  Alice       F      13     56.5       84.0
  Barbara     F      13     65.3       98.0
  Carol       F      14     62.8      102.5
  Henry       M      14     63.5      102.5
  James       M      12     57.3       83.0
  Jane        F      12     59.8       84.5
 ...

EXAMPLE OUTPUT (sheet name WEIGHT)
----------------------------------

  +----------------------------------------------------------------+
  |     A      |    B       |     C      |    D       |    E       |
  +----------------------------------------------------------------+
1 | NAME       |   SEX      |    AGE     |  HEIGHT    |  WEIGHT    |
  +------------+------------+------------+------------+------------+
2 | ALFRED     |    M       |    14      |    69      |  112.5     |
  +------------+------------+------------+------------+------------+
   ...
  +------------+------------+------------+------------+------------+
N | WILLIAM    |    M       |    15      |   66.5     |  112       |
  +------------+------------+------------+------------+------------+


[WEIGHT]  ***** NOTE SHEETNAME IS LAST VARIABLE NAME


PROCESS
=======

%utlfkil(d:/xls/want.xlsx);
%symdel havMax / nowarn;
data log;

  * get meta data at compile time;
  if _n_=0 then do; %let rc=%sysfunc(dosubl('
      ods output variables=havCon;
      proc contents data=sashelp.class;
      run;quit;
      proc sql;
        select variable into :havMax from havCon having num=max(num)
      ;quit;
      '));
   end;

   *export with weight as sheet name;
   rc=dosubl('
     libname xel "d:/xls/want.xlsx";
     data xel.&havMax;
       set sashelp.class;
     run;quit;
     %let cc=&syserr;
     libname xel clear;
  ');

  if symgetn('cc') = 0 then status="sheet &havMax created";
  else status="sheet &havMax failed";

run;quit;


*                _               _       _
 _ __ ___   __ _| | _____     __| | __ _| |_ __ _
| '_ ` _ \ / _` | |/ / _ \   / _` |/ _` | __/ _` |
| | | | | | (_| |   <  __/  | (_| | (_| | || (_| |
|_| |_| |_|\__,_|_|\_\___|   \__,_|\__,_|\__\__,_|

;

JUST USE SASHELP.CLASS


