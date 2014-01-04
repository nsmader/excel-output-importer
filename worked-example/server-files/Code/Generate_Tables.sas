/*--------------------------------------------------------*/
/*                                                        */
/* EXAMPLE CODE FOR AUTOMATING TABLE GENERATION IN EXCEL  */
/* Generate Descriptive Summary Tables and Export to csv  */
/*                                                        */
/*  Author: Nicholas Mader (nmader@chapinhall.org)        */
/*                                                        */
/*--------------------------------------------------------*/

/*----------------------------------------------*/
/* * * SET ENVIRONMENTAL AND RUN PARAMETERS * * */
/*----------------------------------------------*/

	%LET MyPath = /tmp/Excel_Seminar;
	%LET PlayerList = dunna001 konep001 youkk001 ; /* */
	%LET OutcomesList = AbOut_Bb AbOut_K AbOut_Hit AbOut_1B AbOut_2B AbOut_3B AbOut_Hr;

/*------------------------------------------------------------------------*/
/* * * READ DATA, AND SUBSET TO ONLY THE PLAYERS WE ARE INTERESTED IN * * */
/*------------------------------------------------------------------------*/
	
	PROC IMPORT DATAFILE = "&MyPath./Data/Baseball Data for Excel Seminar.csv"
				OUT = WORK.BaseballData
				DBMS = CSV REPLACE;
		GETNAMES = YES;
		DATAROW = 2;
	RUN;
	
	%LET QuotedPlayers = %SYSFUNC(TRANWRD(%SYSFUNC(TRIM("&PlayerList.")), %STR( ), ", "));
	
	DATA BBDataSubset;
		SET BaseballData (WHERE = ( (BatterId IN (&QuotedPlayers.)) & (Runner1BInd = 1 | Runner2BInd = 1 | Runner3BInd = 1) )); /*  */
	RUN;

/*------------------------------------------------*/
/* * * RUN THE DESCRIPTIVE SUMMARY AND OUTPUT * * */
/*------------------------------------------------*/
	
	%MACRO RunAvg;
		%LET i = 1;
		%DO %UNTIL (%SCAN(&PlayerList., &i.) = );
			%LET MyPlyr = %SCAN(&PlayerList, &i.);
			%PUT Working on processing player: &MyPlyr.;
			
			PROC MEANS DATA = BBDataSubset (WHERE = (BatterId = "&MyPlyr.")) NOPRINT;
				VAR &OutcomesList.;
				BY Year;
				OUTPUT OUT = HitStats_ByYear_&MyPlyr. MEAN = &OutcomesList.;
			RUN;
			PROC EXPORT DATA = HitStats_ByYear_&MyPlyr. OUTFILE = "&MyPath/Output/HitStats_&MyPlyr._ByYear_SAS.csv" DBMS = CSV REPLACE;
			RUN;
			%LET i = %EVAL(&i. + 1);
		%END;
	%MEND;
	%RunAvg;
	