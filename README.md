# utl-update-a-master-sheet-with-transaction-sheet-using-excel-and-r-openxls-package-and-sqldf
Update a master sheet with transaction sheet using excel and r openxls package and sqldf
    %let pgm=utl-update-a-master-sheet-with-transaction-sheet-using-excel-and-r-openxls-package-and-sqldf;

    %stop_submission;

    Update a master sheet with transaction sheet using excel and r openxls package and sqldf

      TWO SOLUTIONS
          1 r sql  (classic left join)
          2 r base

    SOAPBOX ON
      This adds excel to the list of languages that support sqlite (using sqldf).
      So we are close to ansi standard sql in sas, r, python and excel...
      Just edit the sqldf code for data wrangling.
    SOAPBOX OFF

    github
    https://tinyurl.com/mpknktkm
    https://github.com/rogerjdeangelis/utl-update-a-master-sheet-with-transaction-sheet-using-excel-and-r-openxls-package-and-sqldf

    excel output
    https://tinyurl.com/2s38va78
    https://github.com/rogerjdeangelis/utl-update-a-master-sheet-with-transaction-sheet-using-excel-and-r-openxls-package-and-sqldf/blob/main/want.xlsx

    stackoverflow
    https://tinyurl.com/5n7uz6k3
    https://stackoverflow.com/questions/79344767/how-to-replace-multiple-column-values-based-on-another-tables-column-where-each

    /*               _     _
     _ __  _ __ ___ | |__ | | ___ _ __ ___
    | `_ \| `__/ _ \| `_ \| |/ _ \ `_ ` _ \
    | |_) | | | (_) | |_) | |  __/ | | | | |
    | .__/|_|  \___/|_.__/|_|\___|_| |_| |_|
    |_|
    */

    /**************************************************************************************************************************/
    /*                                               |                                      |                                 */
    /*      INPUT EXCEL WORKBOOK                     |        PROCESS                       |           OUTPUT                */
    /*        ( TWO SHEETS )                         |(SELF EXPLANATORY ASCII STANDARD SQL) |                                 */
    /*                                               |                                      |                                 */
    /*  d:/xls/want.xlsx sheet=MASTER                |WORKS EVERYWHERE SAS, R AND PYTHON... |                                 */
    /* +---------------                              |                                      |                                 */
    /* |  A1  |  RIDE |                              | select                               |                                 */
    /* |--------------------------------+            |   l.ride                             | d:/xls/want.xlsx sheet=TRANSACT */
    /* |   A  |   B   |       C         |            |  ,l.street                           ||---------------                 */
    /* |--------------------------------+            |  ,r.coord                            || A1  | RUDE   |                 */
    /* |RIDE  |STREET |COORD            |            | from                                 ||-------------------------------+*/
    /* |------+-------+-----------------|            |   sd1.master as l                    ||RIDE |STREET  |COORD           |*/
    /* |001   |Ash    |-100.123,98.123  | /replace   | left join                            ||-------------------------------+*/
    /* |------+-------+-----------------+| transact  |   sd1.transact as r                  ||  A  |   B    |     C          |*/
    /* |002   |Ash    |-100.100,98.123  | \street    | on                                   ||-------------------------------+*/
    /* |------+-------+-----------------+            |   l.street = r.street                ||RIDE |STREET  |COORD           |*/
    /* |003   |Brooke |90.456,91.456    | /no match  |                                      ||-----+--------+----------------|*/
    /* |------+-------+-----------------+| no change |Reasonably fast in R                  ||001  |Ash     |-100.123,98.123 |*/
    /* |004   |Brooke |90.400,91.987    | \          |                                      ||-----+--------+----------------+*/
    /* |------+-------+-----------------+            |1 R SQL                               ||002  |Ash     |-100.123,98.123 |*/
    /* |005   |9th    |20.567,-100.654  | /          |=======                               ||-----+--------+----------------+*/
    /* |------+-------+-----------------+| replace   |                                      ||003  |Brooke  |90.456,91.456   |*/
    /* |006   |9th    |21.123,-100.654  || transact  |%utl_rbeginx;                         ||-----+--------+----------------+*/
    /* |------+-------+-----------------+| street    |parmcards4;                           ||004  |Brooke  |90.456,91.456   |*/
    /* |007   |9th    |20.567,-101.100  | \          |library(openxlsx)                     ||-----+--------+----------------+*/
    /* ----------------------------------            |library(sqldf)                        ||005  |9th     |20.567,-100.654 |*/
    /* [MASTER}                                      |wb<-loadWorkbook("d:/xls/want.xlsx")  ||-----+--------+----------------+*/
    /*                                               | master<-read.xlsx(wb,"master")       ||006  |9th     |20.567,-100.654 |*/
    /*                                               | transact<-read.xlsx(wb, "transact")  ||-----+--------+----------------+*/
    /*  d:/xls/want.xlsx sheet=TRANSACT              | addWorksheet(wb, "want")             ||007  |9th     |20.567,-100.654 |*/
    /* +----------------                             | want<-sqldf('                        |---------------------------------*/
    /* |  A1  |STREET  |                             |  select                              |                                 */
    /* --------------------------------              |   l.ride                             |                                 */
    /* |      A       |      B        |              |  ,l.street                           |                                 */
    /* |------------------------------+              |  ,r.coord                            |                                 */
    /* |STREET        |COORD          |              | from                                 |                                 */
    /* |--------------+---------------|              |   master as l left join transact as r|                                 */
    /* |Ash           |-100.123,98.123|              |  on                                  |                                 */
    /* |--------------+---------------+              |   l.street = r.street                |                                 */
    /* |9th           |20.567,-100.654|              |                                      |                                 */
    /* |--------------+---------------+              |  ')                                  |                                 */
    /* |Brooke        |90.456,91.456  |              | writeData(wb,sheet="want",x=want)    |                                 */
    /* --------------------------------              | saveWorkbook(                        |                                 */
    /*                                               |     wb                               |                                 */
    /* CREATE INPUT NO NEED TO RUN IF                |    ,"d:/xls/want.xlsx"               |                                 */
    /* YOU HAVE THE WORKBOOK                         |    ,overwrite=TRUE)                  |                                 */
    /*                                               | ;;;;                                 |                                 */
    /* SAS DATASETS                                  | %utl_rendx;                          |                                 */
    /*                                               |                                      |                                 */
    /* options validvarname=upcase;                  |--------------------------------------|                                 */
    /* libname sd1 "d:/sd1";                         |                                      |
    /* data sd1.master;                              |2 R BASE                              |                                 */
    /*  informat ride  street coord $24.;            |========                              |                                 */
    /*  input ride  street coord;                    |                                      |                                 */
    /* cards4;                                       |utl_rbeginx;                          |                                 */
    /* 001 Ash -100.123,98.123                       |parmcards4;                           |                                 */
    /* 002 Ash -100.100,98.123                       | library(openxlsx)                    |                                 */
    /* 003 Brooke 90.456,91.456                      | wb<-loadWorkbook("d:/xls/want.xlsx") |                                 */
    /* 004 Brooke 90.400,91.987                      | master<-read.xlsx(wb,"master")       |                                 */
    /* 005 9th 20.567,-100.654                       | transact<-read.xlsx(wb, "transact")  |                                 */
    /* 006 9th 21.123,-100.654                       | wb<-loadWorkbook("d:/xls/want.xlsx") |                                 */
    /* 007 9th 20.567,-101.100                       | addWorksheet(wb, "want")             |                                 */
    /* ;;;;                                          | master$COORD <- transact$COORD[      |                                 */
    /* run;quit;                                     | match(master$STREET, transact$STREET)|                                 */
    /*                                               | writeData(wb,sheet="want",x=master)  |                                 */
    /* data sd1.transact;                            | saveWorkbook(                        |                                 */
    /*  informat street coord $24.;                  |     wb                               |                                 */
    /*  input street coord ;                         |    ,"d:/xls/want.xlsx"               |                                 */
    /* cards4;                                       |    ,overwrite=TRUE)                  |                                 */
    /* Ash -100.123,98.123                           | ;;;;                                 |                                 */
    /* Brooke 90.456,91.456                          | %utl_rendx;                          |                                 */
    /* 9th 20.567,-100.654                           |                                      |                                 */
    /* ;;;;                                          |                                      |                                 */                                          |                                 */
    /* run;quit;                                     |                                      |                                 */
    /*                                               |                                      |                                 */
    /* EXCEL SHEETS                                  |                                      |                                 */
    /*                                               |                                      |                                 */
    /* %utlfkil(d:/xls/want.xlsx);                   |                                      |                                 */
    /*                                               |                                      |                                 */
    /* %utl_rbeginx;                                 |                                      |                                 */
    /* parmcards4;                                   |                                      |                                 */
    /* library(openxlsx)                             |                                      |                                 */
    /* library(sqldf)                                |                                      |                                 */
    /* library(haven)                                |                                      |                                 */
    /* master<-read_sas("d:/sd1/master.sas7bdat")    |                                      |                                 */
    /* transact<-read_sas("d:/sd1/transact.sas7bdat")|                                      |                                 */
    /* master                                        |                                      |                                 */
    /* transact                                      |                                      |                                 */
    /* wb <- createWorkbook()                        |                                      |                                 */
    /* addWorksheet(wb, "master")                    |                                      |                                 */
    /* addWorksheet(wb, "transact")                  |                                      |                                 */
    /* writeData(wb,sheet="master",x=master)         |                                      |                                 */
    /* writeData(wb,sheet="transact",x=transact)     |                                      |                                 */
    /* saveWorkbook(                                 |                                      |                                 */
    /*     wb                                        |                                      |                                 */
    /*    ,"d:/xls/want.xlsx"                        |                                      |                                 */
    /*    ,overwrite=TRUE)                           |                                      |                                 */
    /* ;;;;                                          |                                      |                                 */
    /* %utl_rendx;                                   |                                      |                                 */
    /*                                               |                                      |                                 */
    /**************************************************************************************************************************/

    /*                   _
    (_)_ __  _ __  _   _| |_
    | | `_ \| `_ \| | | | __|
    | | | | | |_) | |_| | |_
    |_|_| |_| .__/ \__,_|\__|
            |_|
    */

    options validvarname=upcase;
    libname sd1 "d:/sd1";
    data sd1.master;
     informat ride  street coord $24.;
     input ride  street coord;
    cards4;
    001 Ash -100.123,98.123
    002 Ash -100.100,98.123
    003 Brooke 90.456,91.456
    004 Brooke 90.400,91.987
    005 9th 20.567,-100.654
    006 9th 21.123,-100.654
    007 9th 20.567,-101.100
    ;;;;
    run;quit;

    data sd1.transact;
     informat street coord $24.;
     input street coord ;
    cards4;
    Ash -100.123,98.123
    Brooke 90.456,91.456
    9th 20.567,-100.654
    ;;;;
    run;quit;

    /**************************************************************************************************************************/
    /*                                                                                                                        */
    /*  MASTER.SAS7BDAT                       TRASACT.SAS7BDAT                                                                */
    /*                                                                                                                        */
    /*                                                                                                                        */
    /*  RIDE    STREET         COORD          STREET         COORD                                                            */
    /*                                                                                                                        */
    /*  001     Ash       -100.123,98.123     Ash       -100.123,98.123                                                       */
    /*  002     Ash       -100.100,98.123     Brooke    90.456,91.456                                                         */
    /*  003     Brooke    90.456,91.456       9th       20.567,-100.654                                                       */
    /*  004     Brooke    90.400,91.987                                                                                       */
    /*  005     9th       20.567,-100.654                                                                                     */
    /*  006     9th       21.123,-100.654                                                                                     */
    /*  007     9th       20.567,-101.100                                                                                     */
    /*                                                                                                                        */
    /**************************************************************************************************************************/

    /*                 _   _                   _
      _____  _____ ___| | (_)_ __  _ __  _   _| |_
     / _ \ \/ / __/ _ \ | | | `_ \| `_ \| | | | __|
    |  __/>  < (_|  __/ | | | | | | |_) | |_| | |_
     \___/_/\_\___\___|_| |_|_| |_| .__/ \__,_|\__|
                                  |_|
    */

    %utlfkil(d:/xls/want.xlsx);

    %utl_rbeginx;
    parmcards4;
    library(openxlsx)
    library(sqldf)
    library(haven)
    master<-read_sas("d:/sd1/master.sas7bdat")
    transact<-read_sas("d:/sd1/transact.sas7bdat")
    master
    transact
    wb <- createWorkbook()
    addWorksheet(wb, "master")
    addWorksheet(wb, "transact")
    writeData(wb, sheet = "master", x = master)
    writeData(wb, sheet = "transact", x = transact)
    saveWorkbook(
        wb
       ,"d:/xls/want.xlsx"
       ,overwrite=TRUE)
    ;;;;
    %utl_rendx;

    /*                    _
    / |  _ __   ___  __ _| |
    | | | `__| / __|/ _` | |
    | | | |    \__ \ (_| | |
    |_| |_|    |___/\__, |_|
                       |_|
    EXCEL INPUT NO SAS DATASETS INPUT
    */

    %utl_rbeginx;
    parmcards4;
    library(openxlsx)
    library(sqldf)
     wb<-loadWorkbook("d:/xls/want.xlsx")
     master<-read.xlsx(wb,"master")
     transact<-read.xlsx(wb, "transact")
     addWorksheet(wb, "want")
     want<-sqldf('
      select
       l.ride
      ,l.street
      ,r.coord
     from
       master as l left join transact as r
      on
       l.street = r.street

      ')
     writeData(wb,sheet="want",x=want)
     saveWorkbook(
         wb
        ,"d:/xls/want.xlsx"
        ,overwrite=TRUE)
     ;;;;
     %utl_rendx;

    /**************************************************************************************************************************/
    /*                                                                                                                        */
    /*    d:/xls/want.xlsx sheet=WANT                                                                                         */
    /*   |---------------                                                                                                     */
    /*   | A1  | RUDE   |                                                                                                     */
    /*   |-------------------------------+                                                                                    */
    /*   |RIDE |STREET  |COORD           |                                                                                    */
    /*   |-------------------------------+                                                                                    */
    /*   |  A  |   B    |     C          |                                                                                    */
    /*   |-------------------------------+                                                                                    */
    /*   |RIDE |STREET  |COORD           |                                                                                    */
    /*   |-----+--------+----------------|                                                                                    */
    /*   |001  |Ash     |-100.123,98.123 |                                                                                    */
    /*   |-----+--------+----------------+                                                                                    */
    /*   |002  |Ash     |-100.123,98.123 |                                                                                    */
    /*   |-----+--------+----------------+                                                                                    */
    /*   |003  |Brooke  |90.456,91.456   |                                                                                    */
    /*   |-----+--------+----------------+                                                                                    */
    /*   |004  |Brooke  |90.456,91.456   |                                                                                    */
    /*   |-----+--------+----------------+                                                                                    */
    /*   |005  |9th     |20.567,-100.654 |                                                                                    */
    /*   |-----+--------+----------------+                                                                                    */
    /*   |006  |9th     |20.567,-100.654 |                                                                                    */
    /*   |-----+--------+----------------+                                                                                    */
    /*   |007  |9th     |20.567,-100.654 |                                                                                    */
    /*   ---------------------------------                                                                                    */
    /*                                                                                                                        */
    /**************************************************************************************************************************/

    /*___           _
    |___ \   _ __  | |__   __ _ ___  ___
      __) | | `__| | `_ \ / _` / __|/ _ \
     / __/  | |    | |_) | (_| \__ \  __/
    |_____| |_|    |_.__/ \__,_|___/\___|

    */

    %utl_rbeginx;
    parmcards4;
     library(openxlsx)
     wb<-loadWorkbook("d:/xls/want.xlsx")
     master<-read.xlsx(wb,"master")
     transact<-read.xlsx(wb, "transact")
     wb<-loadWorkbook("d:/xls/want.xlsx")
     addWorksheet(wb, "want")
     master$COORD <- transact$COORD[
      match(master$STREET, transact$STREET)]
     writeData(wb,sheet="want",x=master)
     saveWorkbook(
         wb
        ,"d:/xls/want.xlsx"
        ,overwrite=TRUE)
    ;;;;
    %utl_rendx;

    /**************************************************************************************************************************/
    /*                                                                                                                        */
    /*    d:/xls/want.xlsx sheet=WANT                                                                                         */
    /*   |---------------                                                                                                     */
    /*   | A1  | RUDE   |                                                                                                     */
    /*   |-------------------------------+                                                                                    */
    /*   |RIDE |STREET  |COORD           |                                                                                    */
    /*   |-------------------------------+                                                                                    */
    /*   |  A  |   B    |     C          |                                                                                    */
    /*   |-------------------------------+                                                                                    */
    /*   |RIDE |STREET  |COORD           |                                                                                    */
    /*   |-----+--------+----------------|                                                                                    */
    /*   |001  |Ash     |-100.123,98.123 |                                                                                    */
    /*   |-----+--------+----------------+                                                                                    */
    /*   |002  |Ash     |-100.123,98.123 |                                                                                    */
    /*   |-----+--------+----------------+                                                                                    */
    /*   |003  |Brooke  |90.456,91.456   |                                                                                    */
    /*   |-----+--------+----------------+                                                                                    */
    /*   |004  |Brooke  |90.456,91.456   |                                                                                    */
    /*   |-----+--------+----------------+                                                                                    */
    /*   |005  |9th     |20.567,-100.654 |                                                                                    */
    /*   |-----+--------+----------------+                                                                                    */
    /*   |006  |9th     |20.567,-100.654 |                                                                                    */
    /*   |-----+--------+----------------+                                                                                    */
    /*   |007  |9th     |20.567,-100.654 |                                                                                    */
    /*   ---------------------------------                                                                                    */
    /*                                                                                                                        */
    /**************************************************************************************************************************/

    /*              _
      ___ _ __   __| |
     / _ \ `_ \ / _` |
    |  __/ | | | (_| |
     \___|_| |_|\__,_|

    */
