/*--------------------------------------------------------*/
/*                                                        */
/* EXAMPLE CODE FOR AUTOMATING TABLE GENERATION IN EXCEL  */
/* Generate Descriptive Summary dTables and Export to csv */
/*                                                        */
/*  Author: Nicholas Mader (nmader@chapinhall.org)        */
/*                                                        */
/*--------------------------------------------------------*/

/* * * Set Environmental Parameters * * */

	clear all
	local MyPath  /tmp/Excel_Seminar
	set more off, perm
	set trace off
	set tracedepth 1
	pause on

/* * * Set Run Parameters * * */
	
	local PlayerList dunna001 konep001
	local OutcomesList AbOut_Bb AbOut_K AbOut_Hit AbOut_1B AbOut_2B AbOut_3B AbOut_Hr

/* * * Import and Subset Data * * */
	
	insheet  using "`MyPath'/Data/Baseball Data for Excel Seminar.csv", comma clear case
	save "`MyPath'/Data/Baseball Data for Excel Seminar.dta", replace

	keep if (strpos("`PlayerList'", BatterId)>0)
	tab BatterId

/* * * Run data summaries and export descriptive tables * * */

foreach MyPlyr in `PlayerList' {
	forvalues y = 2009(1)2011 {
		estpost tabstat `OutcomesList' if BatterId == "`MyPlyr'" & Year == `y', statistics(mean sd count) columns(statistics)
		esttab using "`MyPath'/Output/HitStats_`MyPlyr'_`y'_STATA.csv", cells("mean sd count") csv replace nonumbers
	}
}



