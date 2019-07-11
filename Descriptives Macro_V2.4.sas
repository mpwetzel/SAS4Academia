
/*---------------------------------------------------------


Title: Descriptives Macro
Date: 2018/07/10
Author: Martha Wetzel
Purpose: Run basic statistics by group

Usage Notes: The output code for this macro was based in part on the
	Winship Cancer Institute's macro located here:
	https://bbisr.winship.emory.edu/

	Currently, categorical variables must be character type. 

	Macro parameters:
		DATASET=,  Input data set 
	DRIVER=, File path and name for Excel driver file - no quotes
	OVERALL=Y, Request overall summary statistics 
	BYCLASS=Y, Request statistic by class variable 
	OUTLIB= ,  Name of library to save out SAS data sets to 
	OUTPATH=, File pathway for RTF output, no quotes 
	FNAME=,  File name for RTF output 
	CLASSVAR=, CAN ONLY TAKE ONE CLASS VARIABLE 
	ODSGRAPHS=N,  Create or suppress ODS graphics (Y/N)
	FEXACT=Y,	Run Fisher's exact tests. By default, Fisher's exact tests are run on all categorical variables and the 
		p-values from the exact tests are automatically used if the Chi squared small cell size warning appears.
		However, since the exact tests can take a while to run, they can be globally suppressed using this option.
		If any Fisher's Exact tests are requested in the driver, this option will be automatically set to Yes.
	ROUNDTO=0.01 Decimal place to round results to  
	DISPLAY_PVAL=Y,  Display p-values in RTF report (Y/N)
	DISPLAY_METRIC=Y,  Display metric names in RTF report (Y/N)
	DISPLAY_N =Y  Display variable N's in RTF report (Y/N)
	CLEARTEMP = Y, Delete all the temp files created by the macro. 


	Input:
		Excel driver: The macro is operated using a Excel driver that includes a list of variable names,
		labels, types, and requested test. 
		In data: The macro requires an input SAS data set with categorical variables in character format.

	Output:  
		The macro outputs two SAS data sets and one RTF table. The SAS data sets 
		contain a variety of statistics of interest to the analyst, including 
		N's and results of alternative statistical tests and (for continuous variables) means, 
		medians, percentiles, min, max, mode, normality test statistics, and counts of missing values. The RTF output contains 
		a limited number of columns that can easily be used in a demographic table for a paper.

	Caveat: This macro begins by clearing files that start with an underscore.

Updates:

	- v2 - 2018/07/30: Added option to suppress Fisher's Exact test 
						Fixed issue that would arise with long variable names
						Added capabilities to include classification variables with more than three levels
						Added Shapiro-Wilk test to out data sets
						Formats p-values in output
	- v2.2 - 2019/02/08: adding capabilities to optionally run overall summary stats in addition to results by 
						class. 
						Added total N's to headers in output
						Added check to ensure categorical variables are character type
	- v2.2 - 2019/03/05: Changed length of variable labels for output report to allow labels of up to 200 characters
						Fixed an error occurring when exclusively continuous variables were requested
	- v2.3 - 2019/03/18: Add an N column to output and parameters to control output display
		   - 2019/07/11: Moved all intermediate data created by macro to a temp folder and created option
						 to clear or keep temp files.


-----------------------------------------------------------*/


%MACRO DESCRIPTIVE(
	DATASET=, /* Input data set */
	DRIVER=, /* File path and name for Excel driver file - no quotes*/ 
	OVERALL=Y, /* Request overall summary statistics */
	BYCLASS=Y, /* Request statistic by class variable */
	OUTLIB= WORK, /* Name of library to save out SAS data sets to */
	OUTPATH=, /* File pathway for RTF output, no quotes */
	FNAME=,  /* File name for RTF output */
	CLASSVAR=, /* CAN ONLY TAKE ONE CLASS VARIABLE */
	ODSGRAPHS=N, /* Create or suppress ODS graphics (Y/N) */
	FEXACT=Y,	/* Run Fisher's exact tests - See note at top of program */
	ROUNDTO=0.01, /* Decimal place to round results to */ 
	DISPLAY_PVAL=Y, /* Display p-values in RTF report */
	DISPLAY_METRIC=Y, /* Display metric names in RTF report */
	DISPLAY_N =Y, /* Display variable N's in RTF report */
	CLEARTEMP = Y /* Delete temp files created by macro */
);  

	option label;
	ods html close;
	ods html close;
	ods listing;

	/* Establish a temp directory for macro files */
	%let work_path=%sysfunc(pathname(work));
	options dlcreatedir;
	libname Table1 "&work_path.\Table1";

	/* Get rid of any existing files Table1 directory */
	proc datasets lib=Table1 nolist kill;
	run;
	quit;


	/* Set ODS graphics option */
	%if %sysfunc(upcase(&ODSGRAPHS.)) = Y %then %do;
		ods graphics;
		%put NOTE: ODS graphics are on;
	%end;
	
	%else %do;
		ods graphics off;
		%put NOTE: ODS graphics will be suppressed;
	%end;

	/* Set up Fisher's switch */
	%if &Fexact. = Y %then %do;
		%let  Fish = Fishers;
		%let Exact = Exact;
	%end;
	%else %do;
		%let Fish = ;
		%let Exact = ;
	%end;

 	proc import datafile = "&Driver." out = Table1.driver
		dbms = xlsx replace;
		sheet = "Variables";
	run;

	/* Check counts of variables */
	data _null_;
		set Table1.driver;
		retain CatCount ContCount AnyFish;
		if _N_ = 1 then do;
			CatCount = 0;
			ContCount = 0;
		end;
		if upcase(Type) = "CATEGORICAL" then CatCount +1;
		if upcase(Type) = "CONTINUOUS" then ContCount +1;
		call symputx("CatCount", CatCount);
		call symputx("ContCount", ContCount);
		if upcase(strip(statistical_test)) = "FISHERS" then AnyFish = "Y";
		if missing(AnyFish) then AnyFish = "N";
		call symputx("AnyFish", AnyFish);
	run;

	%put Note: There are &ContCount. continuous variables and  &CatCount. categorical variables;

	/* Check to make sure driver options don't contradict parameter options */
	%if &Fexact. = N and &AnyFish. = Y %then %do;
		%put WARNING: Discrepancy found between driver settings and FEXACT parameter. Resetting Fishers exact parameter to Y.;
		%let Fexact = Y;
	%end;

	/* Set up variable lists */
	%if %eval(&CatCount. >= 1) %then %do;
		proc sql noprint;
			select Variable_Name, 
					upcase(strip(Variable_name))
			into :CATLIST separated by " ",
				:CatListQ separated by '" "'

			from Table1.driver
			where upcase(Type) = "CATEGORICAL";
		quit;
	%end;

	%else %do;
		%let catlist = ;
	%end;

	%if %eval(&ContCount. >= 1) %then %do; 
		/* List of continuous variables */
		proc sql noprint;
			select Variable_Name
			into :NLIST separated by " "
			from Table1.driver
			where upcase(Type) = "CONTINUOUS";

		quit;

	%end;

	%else %do;
		%let NLIST = ;
	%end;

	%put The following categorical variables will be included: &catlist.;
	%put The following numeric variables will be included: &nlist.;

	/* Set up for determining which statistical tests to apply */
	data Table1.testfmt;
		set Table1.driver (keep = Variable_Name Statistical_test );
		where not missing(Variable_name);
		Start = upcase(strip(Variable_name));
		Type = "C";
		Fmtname = "Test";
		Label = upcase(Statistical_test);
	run;

	proc format cntlin = Table1.testfmt;
	run;

	/* Set up format for labeling and ordering variables */
	data Table1.labelfmt (drop = order);
		set Table1.driver (keep = Variable_name Label );
		where not missing(Variable_name);
		Start = upcase(strip(Variable_name));
		/* Label */
		Type = "C";
		Fmtname = "Varlab";
		output;
		/* Order */
		Order +1;
		Label = input(Order, $50.);
		FmtName = "Order";
		output;

	run;

	proc sort data = Table1.labelfmt;
		by Fmtname;
	run;

	proc format cntlin = Table1.labelfmt;
	run;

	%if &ByClass = Y %then %do;
		/* Set up standardized naming scheme for class levels - needed for output following transpose */
		proc sort data = &dataset. out = Table1.insort ;
			by &classvar.;
		run;

		data Table1.indata;
			set Table1.insort;
			by &classvar.;

			if first.&classvar then ClassNum +1;

			length StandClass $10;

			StandClass = cats("Class_",ClassNum);
		run;

		/* Generate label statement */
		proc sql noprint;
			select distinct cats(StandClass, "='",&classvar.,"'"), /* For regular data set labeling */
				count(unique(StandClass)) /* count number of class levels */
		
			into :ClassLabel separated by " ",
				:ClassCount
			from Table1.indata;
		quit;

		proc format;
			value $ ClassFmt 
			&classlabel.;
		run;
			
		%put &Classlabel.;
		%put Number Classes = &ClassCount.;
	%end;

	%else %do; /* Don't prep class parts if it's only an overall analysis */
		data Table1.indata;
			set &dataset.;
		run;
	%end;


	/* Check that the categorical variables are character format - abort if wrong */
	%if &CatCount. >= 1 %then %do;
		proc sql noprint;
			select 
				 sum(case when type = "num" then 1 else 0 end),
					cats(case when type = "num" then name end)
			into :BadVarCount,
				:BadVars separated by " "
			from dictionary.columns
			where upcase(libname) = "TABLE1" and UPCASE(memname) = "INDATA" and upcase(name) in ("&CatListQ.");
		quit;

		%if &BadVarCount. > 0 %then %do;
			%put ERROR: Categorical analysis requested for the following numeric variables: &BadVars.. 
				Convert these variables to character type for categorical analysis.
				Aborting analysis;
			%goto ExitNow;	
		%end;
	%end;


	/* ------------------------------------------------------*/
	/*					Numeric Variables					*/
	/* ------------------------------------------------------*/


    %IF &NLIST NE  %THEN %DO; 

   		/* Run overall stats, even if not requested, to pick up the N's */
			%put NOTE: Running overall statistics for continuous data;

	   		/* Calculate basic stats for all variables */
		   	proc univariate data = Table1.indata normal outtable = Table1.Overall_numsummary (keep =_VAR_  _NOBS_ _NMISS_ _SUM_ _MEAN_ _STD_ _MIN_ 
				_Q1_ _MEDIAN_ _Q3_ _MAX_ _RANGE_ _qRANGE_ _MODE_ _NORMAL_ 
				) noprint;
				var &nlist.;
			run;

			data &outlib..AllStats_Numeric_overall;
				set Table1.overall_numsummary (rename = (_VAR_ = Variable));

				/* Round */
				array unround (5) _Q1_ _MEDIAN_ _Q3_ _MEAN_ _STD_;
				array rounded (5) r_Q1_ r_MEDIAN_ r_Q3_ r_MEAN_ r_STD_;

				do i = 1 to 5;
					rounded(i) = round(unround(i),&ROUNDTO.);
				end;

				StatTest =  put(upcase(strip(Variable)), $test.);
				format Metrics $20.;
				if StatTest = "NON-PARAMETRIC" then do;
					Overall_Result = cat(r_MEDIAN_," (", r_Q1_, ", ",r_Q3_,")");
					Metrics = "Median (Q1, Q3)";
				end;
				else do;
					Overall_Result = cat(r_MEAN_, " (", r_STD_, ")");
					Metrics = "Mean (SD)";
				end;
	
			run;

			proc sort data = &outlib..AllStats_Numeric_overall ;
				by Variable;
			run;

	
		%if &byclass. = Y %then %do;
			/* Calculate basic stats for all variables */
		   	proc univariate data = Table1.indata normal outtable = Table1.numsummary (keep =_VAR_ StandClass _NOBS_ _NMISS_ _SUM_ _MEAN_ _STD_ _MIN_ 
				_Q1_ _MEDIAN_ _Q3_ _MAX_ _RANGE_ _qRANGE_ _MODE_ _NORMAL_ _PROBN_
				) noprint;
				var &nlist.;
				class StandClass;
			run;


			/* All the tests are run for analyst review - only requested test is included in output table */
			/* Run non-parametric stats */
			proc npar1way data=Table1.indata wilcoxon;
				var &nlist.;
				class StandClass;
				output out = Table1.nonpar_tests /*(keep = _VAR_ _WIL_ PL_WIL PR_WIL P2_WIL _KW_ P_KW 	)*/;
			run;

			data Table1.nonpar_tests2;
				set Table1.nonpar_tests;
				%if &ClassCount = 2 %then %do;
					rename P2_WIL = pvalue_np;
				%end;
				%else %do;
					rename P_KW = pvalue_np;
				%end;
			run;

			/* Run parametric stats */
			/* T-tests: only if class variable has 2 levels */
			%if &CLassCount = 2 %then %do;
				ods output Ttests = Table1.numttests equality = Table1.numeqvar;
				proc ttest data = Table1.indata ;
					var &nlist.;
					class StandClass;
				run;

				/* Choose correct ttest method based on variance */
				proc sort data = Table1.numttests;
					by Variable;
				run;

				proc sort data = Table1.numeqvar;
					by Variable;
				run;

				data Table1.Parm_tests (drop = ProbF DF  rename = (Method = TTestMethod Probt = Parm_pvalue));
					merge Table1.numttests Table1.numeqvar (keep = Variable ProbF) ;
					by Variable;
					if (ProbF < 0.05 and variances = "Unequal") or (ProbF > 0.05 and Variances = "Equal");
				run;
			%end;

			%else %if &ClassCount >= 3 %then %do;
			
				/*----	This section iterates through each  variable ----*/
				/* Create ANOVA driver data set */
				data Table1.driver_con;
					set Table1.driver;
					where upcase(Type) = "CONTINUOUS" ;
					Counter +1;
					call symputx('totparm', Counter);
				run;

				%let parmsets = ; /* this will be a list of all the final data sets that need to be combined */

				%let a = 1;
				%do %while (&a. <= &totparm.);

					data _null_;
						set Table1.driver_con;
						where counter = &a.;

						/* Set up a short name so as to avoid issues with data set names being too long */
						if length(strip(Variable_name)) > 25 then ShortName = cats(substr(Variable_name, 1,15),&a.);
							else ShortName = Variable_name;

						call symputx("varcon", Variable_name);
						call symputx("varcon_s", ShortName);
					run;

					/* Run the ANOVA */
					ods table ModelANOVA   = Table1.Anova_&varcon_s.;
					proc glm data = Table1.indata;
						class &classvar.;
						model &varcon. = &classvar.;
					run;
					quit;

					/* Add to list of data sets to stack at end */
					%let parmsets = &parmsets. Table1.Anova_&varcon_s.;

					/* Iterate to next variable */
					%let a = %eval(&a. +1);

				%end; /* End variable loop for ANOVA  */

				/* Combine ANOVA output */
				data Table1.Parm_tests (rename = (ProbF = Parm_pvalue Dependent = Variable));
					length Dependent $32.;
					set &parmsets.;
					if HypothesisType = 1; /* Doesn't matter which one b/c it's a one-way ANOVA */
				run;

			%end; /* End ANOVA analysis */


			/* Merge stat test results to summary stats */
				proc sort data = Table1.numsummary;
					by _VAR_;
				run;

				proc sort data = Table1.nonpar_tests2;
					by _VAR_;
				run;
		
				proc sort data = Table1.Parm_tests;
					by Variable;
				run;

			/* Save detailed summary table */
			data &outlib..AllStats_Numeric_class (drop = i r_:);

				format StandClass2 $100. Variable $32.  Result $50. Metrics $20. pvalue pvalue6.3 StandClass $100.;

				merge Table1.numsummary (rename = (_VAR_ = Variable)) 
					Table1.nonpar_tests2 (rename = (_VAR_ = Variable))
					Table1.Parm_tests;
				by Variable;

				/* Round for output */
				array unround (5) _Q1_ _MEDIAN_ _Q3_ _MEAN_ _STD_;
				array rounded (5) r_Q1_ r_MEDIAN_ r_Q3_ r_MEAN_ r_STD_;

				do i = 1 to 5;
					rounded(i) = round(unround(i),&ROUNDTO.);
				end;

				/* Summary stats depend on parametric or nonparametric */
				StatTest =  put(upcase(strip(Variable)), $test.);
				if StatTest = "NON-PARAMETRIC" then do;
					Result = cat(r_MEDIAN_," (", r_Q1_, ", ",r_Q3_,")");
					Metrics = "Median (Q1, Q3)";
					pvalue = pvalue_np;
				end;
				else do;
					Result = cat(r_MEAN_, " (", r_STD_, ")");
					Metrics = "Mean (SD)";
					pvalue = Parm_pvalue;
				end;

				label Parm_pvalue = "Parametric Test p-value";
				StandClass2 = put(StandClass, $Classfmt.);
				label StandClass2 = "&classvar.";

			run;

			/* Transpose so class variables go across */
			proc transpose data = &outlib..AllStats_Numeric_class out = Table1.numeric_t;
				by Variable pvalue Metrics;
				var  Result;
				id StandClass;
			run;

			%end; /* End "by class" analysis */

	
			/* Combine overall and by class analysis */
			data Table1.numeric_t2;
				merge 
					&outlib..AllStats_Numeric_overall  (keep = Variable _NOBS_
						%if &Overall = Y %then %do;
						Overall_Result Metrics
						%end; )
				%if &BYCLASS = Y %then %do;
					Table1.numeric_t
				%end;
					;
				by Variable;
			run;
			


	%end; /* End numeric variable analysis */

	/* ------------------------------------------------------*/
	/*					Categorical Variables				*/
	/* ------------------------------------------------------*/

	%if &catlist. ne %then %do;

		/*----	This section iterates through each categorical variable ----*/
		/* Create categorical driver data set that renumbers the categorical variables */
		data Table1.driver_cat;
			set Table1.driver;
			where upcase(Type) = "CATEGORICAL";
			Counter +1;
		run;

		/* Find number of variables */
		proc sql noprint;
			select max(counter) into :totcats
			from Table1.driver_cat;
		quit;

		%let list = ; /* this will be a list of all the final data sets that need to be combined */
		%let list_overall = ;

		%let c = 1;
		%do %while (&c. <= &totcats.);

			data _null_;
				set Table1.driver_cat;
				where counter = &c.;

				/* Set up a short name so as to avoid issues with data set names being too long */
				if length(strip(Variable_name)) > 25 then ShortName = cats(substr(Variable_name, 1,15),&c.);
					else ShortName = Variable_name;

				call symputx("var", Variable_name);
				call symputx("varcat_s", ShortName);
			run;

			%put Note: Calculating &var. statistics;

			/* Run overall regardless of options to get overall N's by variable */
				proc freq data = Table1.indata (where = (not missing(&var.))) ;
					table &var. / out = Table1.ovfreq_&varcat_s. ;
				run;

				/* Manipulate frequency table */
				data Table1.ovfreq2_&varcat_s. (rename = (&var. = Category));
					retain Variable;
					set Table1.ovfreq_&varcat_s.;
					/* combine count and percent into single variable */
					Overall_Result = cat(count, " (",round(PERCENT,&ROUNDTO.), "%)");
					Variable = "&var.";
				run;

				%let list_overall = &list_overall. Table1.ovfreq2_&varcat_s.;

			/* Run by class statistics */
			%if &byclass. = Y %then %do;
				/* Calculate stats */
				proc freq data = Table1.indata (where = (not missing(&var.))) ;
					table &var.*StandClass / out = Table1.freq_&varcat_s. outpct chisq &exact. warn = output;
					output out = Table1.chisq_&varcat_s. chisq N &fish. ;
				run;

				%if %sysfunc(exist(Table1.chisq_&varcat_s.)) = 0 %then %do;
					data Table1.chisq_&varcat_s.;
					run;
				%end;


				/* Manipulate frequency table */
				data Table1.freq2_&varcat_s. (rename = (&var. = Category));
					retain Variable;
					set Table1.freq_&varcat_s.;
					/* combine count and percent into single variable */
					NPercent = cat(count, " (",round(PCT_COL,&ROUNDTO.), "%)");
					Variable = "&var.";
				run;

				proc transpose data = Table1.freq2_&varcat_s. out = Table1.freqt_&varcat_s. (drop = _NAME_);
					by Variable Category;
					var NPercent;
					id StandClass;
				run;

				/* Merge with stats */
				data Table1.chisq_&varcat_s.2;
					set Table1.chisq_&varcat_s.;
					Variable = "&var.";
				run;

				data Table1.freq3_&varcat_s.;
					merge Table1.freqt_&varcat_s. Table1.chisq_&varcat_s.2;
					format pvalue pvalue8.3;

					by Variable;

					/* Initialize XP2_Fish to avoid unnecessary notes in log */
					%if &fexact. = N %then %do;
						XP2_Fish = .; /* Fisher's results set to missing since they weren't requested */
					%end;

					/* Select test and p-value */
					/* If requested, or if Chi inappropriate, use Fisher's */
					if upcase(put(upcase("&var."), $test.)) = "FISHERS" or (WARN_PCHI = 1 and "&fexact." = "Y") then do;
						pvalue = XP2_FISH;
						Test_Used = "Fishers";
					end;

					/* Else, use CHISQ but warn if inappropriate */
					else do;
						pvalue = p_pchi;
						Test_Used = "Chisq    ";
						if upcase(put(upcase("&var."), $test.)) = "CHI" and WARN_PCHI = 1 and
							"&fexact." = "N" then put "WARNING: Chi Squared may be inappropriate for &var. but the exact test was suppressed. Change the macrovariable parameter FEXACT to Y to calculate the exact test";
					end;


				run;

				/* Add to list of data sets to stack at end */
				%let list = &list. Table1.freq3_&varcat_s.;

			%end; /* Closes "by class" segment */

			/* Iterate to next variable */
			%let c = %eval(&c. +1);

		%end; /* End categorical loop */


		/* Stack Overall data sets */
		data &outlib..allstats_cat_overall ;
			length Variable $32 Category $100;
			format Variable $32. Category $100.;
			set &list_overall.;

			if missing(Overall_Result) then Overall_Result = "0 (0.00%)";
			Metrics = "N (%)";
		run;

		/* Counts of non-missing values */
		proc means data = &outlib..allstats_cat_overall noprint nway;
			output out = Table1.char_Ns
			sum(COUNT)=;
			class Variable;
		run;

		proc sort data = &outlib..allstats_cat_overall ;
			by Variable Category;
		run;


		/* Stack by class data sets */
		%if &byclass. = Y %then %do;
			/* Create single data set with all categorical information */
			data &outlib..allstats_cat_class ;
				length Variable $32 Category $100;
				format Variable $32. Category $100.;
				set &list.;

				Metrics = "N (%)";

				/* Fill in zero cells */
				array zerofill (*) class:;
				do i = 1 to hbound(zerofill);
					if missing(zerofill(i)) then zerofill(i) = "0 (0.00%)";
				end;

				label &classlabel.;

			run;

			proc sort data =  &outlib..allstats_cat_class;
				by Variable Category;
			run;

		%end;

		data Table1.Allcat_a;
			merge
			%if &overall. = Y %then %do;
				&outlib..allstats_cat_overall (keep = Variable Category Overall_Result Metrics)
			%end;
			%if &byclass. = Y %then %do;
				&outlib..allstats_cat_class (keep = Variable Category Metrics Class_: pvalue)
			%end;
			;
			by Variable Category;
		run;

		/* Add N's */
		data Table1.Allcat;
			merge Table1.allcat_a 	Table1.char_Ns (keep = variable count rename = (Count = _NOBS_));
			by variable;
		run;


   %end; /* End categorical analysis */


   /*---------------------------------------------------*/
   /*----		Print for Analyst Review			----*/
   /*---------------------------------------------------*/

   ods html;
	%IF &NLIST NE  %THEN %DO; 

		%if &byclass. = Y  %then %do;
		   proc print data = &outlib..AllStats_Numeric_class (keep = StandClass2 variable _NOBS_ _NMISS_ _MEAN_ _STD_ _MIN_
				_Q1_ _MEDIAN_ _Q3_ 	_MAX_ _MODE_ _PROBN_);
			run;
		%end;

		%if &overall. = Y %then %do;
		   proc print data = &outlib..AllStats_Numeric_overall (keep = variable _NOBS_ _NMISS_ _MEAN_ _STD_ _MIN_
				_Q1_ _MEDIAN_ _Q3_ 	_MAX_ _MODE_ );
			run;
		%end;

	%end;


	ods html close;

   /*---------------------------------------------------*/
   /*----			Prep for Output					----*/
   /*---------------------------------------------------*/

   	/* Generate label for proc report */
	%if &byclass. = Y %then %do;
		proc sql noprint;
			select 
				distinct cats(StandClass, "='",&classvar.,"~ N=", count(StandClass), "'")  /* For proc report column headers */
			into :ReportLabel separated by " "
			from Table1.indata
			group by &classvar.;
		quit;
	%end;

	/* Get the overall observation count */
	data _null_;
		set Table1.indata;
		call symputx('totalobs',_N_);
	run;

   /* Combine character and numeric results */
   data Table1.combo (drop = Variable rename = (Variable2=Variable));
   		format Variable2 $200. %if &CatCount. > 0 %then %do; Category %end; Class_: $100. Metrics $20.;
		%if &Byclass = Y %then %do;
			format pvalue pvalue8.3;
		%end;

   		set 
			%if &CatCount. > 0 %then %do;
				Table1.Allcat 
			%end;
			%if &ContCount. > 0 %then %do;
				Table1.numeric_t2 ;
			%end;
			;

		/* Add observation counts to header */
		label 
			/* Apply class labels */
			%if &byclass. = Y %then %do;
				&ReportLabel
			%end;
			
			%if &overall. = Y %then %do;
				Overall_Result = "Overall~N=&totalobs."
			%end;
			;
		Order = input(put(strip(upcase(Variable)), $order.), 8.);
		Variable2 = put(strip(upcase(Variable)),$varlab.); 

	run;

	proc sort data = Table1.combo;
		by Order;
	run;

	/*---------------------------------------------------*/
   /*----			Create RTF					----*/
   /*---------------------------------------------------*/

ODS PATH WORK.TEMPLAT(UPDATE)
   SASUSR.TEMPLAT(UPDATE) SASHELP.TMPLMST(READ);

   PROC TEMPLATE;
	   DEFINE STYLE STYLES.TABLES;
	   NOTES "MY TABLE STYLE"; 
	   PARENT=STYLES.MINIMAL;

	     STYLE SYSTEMTITLE /FONT_SIZE = 12pt     FONT_FACE = "TIMES NEW ROMAN";

	     STYLE HEADER /
	           FONT_FACE = "TIMES NEW ROMAN"
	            CELLPADDING=8
	            JUST=C
	            VJUST=C
	            FONT_SIZE = 10pt
	           FONT_WEIGHT = BOLD; 

	     STYLE TABLE /
	            FRAME=HSIDES            /* outside borders: void, box, above/below, vsides/hsides, lhs/rhs */
	            RULES=GROUP              /* internal borders: none, all, cols, rows, groups */
	            CELLPADDING=6            /* the space between table cell contents and the cell border */
	            CELLSPACING=6           /* the space between table cells, allows background to show */
	            JUST=C
	            FONT_SIZE = 10pt
	            BORDERWIDTH = 0.5pt;  /* the width of the borders and rules */

	     STYLE DATAEMPHASIS /
	           FONT_FACE = "TIMES NEW ROMAN"
	           FONT_SIZE = 10pt
	           FONT_WEIGHT = BOLD;

	     STYLE DATA /
	           FONT_FACE = "TIMES NEW ROMAN" 
	           FONT_SIZE = 10pt;

	     STYLE SYSTEMFOOTER /FONT_SIZE = 9pt FONT_FACE = "TIMES NEW ROMAN" JUST=C;
	   END;

   RUN; 

   *------- build the table -----;

   OPTIONS ORIENTATION=PORTRAIT MISSING = "-" NODATE;

     ODS RTF STYLE=tables FILE= "&OUTPATH.\&FNAME &SYSDATE..DOC"; 


	PROC REPORT DATA=Table1.combo HEADLINE HEADSKIP CENTER STYLE(REPORT)={JUST=CENTER} SPLIT='~' nowd 
	          SPANROWS LS=256;
	      COLUMNS order variable 

 			%if &CatCount. > 0 %then %do; category %end;
			%if &DISPLAY_N. = Y %then %do; _NOBS_ %end;
			%if &overall. = Y %then %do;
				Overall_Result
			%end;
			%if &byclass. = Y %then %do;
				Class: 
				%if &DISPLAY_PVAL = Y %then %do; pvalue %end;
			%end;
			%if &DISPLAY_METRIC = Y %then %do; metrics %end;

			;	

	      DEFINE order/order order=internal noprint;
	      DEFINE variable/ Order order=data  "Variable"  STYLE(COLUMN) = {JUST = L CellWidth=15%};

			%if &CatCount. > 0 %then %do;
				DEFINE category/ DISPLAY   "Level"   STYLE(COLUMN) = {JUST = L CellWidth=15%};
			%end;

			%if &DISPLAY_N. = Y %then %do; DEFINE _NOBS_ / "N" order style(Column) = {JUST = C }; %end;

			%if &overall. = Y %then %do;
				DEFINE Overall_Result / DISPLAY STYLE(COLUMN) = {JUST = C } ;
			%end;

			%if &byclass. = Y %then %do;
			  %let z = 1;
			  
			  %do %while (&z. <= &ClassCount.);
			  	DEFINE Class_&z. / DISPLAY STYLE(COLUMN) = {JUST = C } ;
		     	%let z = %eval(&z. +1);
			  %end;

			   %if &DISPLAY_PVAL. = Y %then %do;
			      /* Bold p-values under 0.05 */
			       DEFINE pvalue/ORDER MISSING "P-Value" STYLE(COLUMN)={JUST = C CellWidth=10%} ;

		         COMPUTE pvalue; 
		              IF . < pvalue <0.05 THEN 
		               CALL DEFINE("pvalue", "STYLE", "STYLE=[FONT_WEIGHT=BOLD]");
		         ENDCOMP; 
				%end; /* End p-value column stuff */

			%end; /* End Class Variable Columns */
	       
		%if &DISPLAY_METRIC. = Y %then %do; DEFINE Metrics/ DISPLAY    STYLE(COLUMN) = {JUST = C }; %end;

	     compute after variable; line ''; endcomp; /* Inserts a blank line after each variable */
	       
	   RUN; 

 
      ODS RTF CLOSE; 

	/* Get rid of temp files created by macro in Table1 directory */
	%if %upcase(&CLEARTEMP.) = Y %then %do;
		proc datasets lib=Table1 nolist  ;
			save combo;
		run;
		quit;
	%end;

	  %ExitNow:

%mend; 




/*--------------------------------------------------------------*/
/*							Sample Call							*/
/*--------------------------------------------------------------*/

/*libname out "C:\Users\mwetze2\Box Sync\Standard Code\Sandbox\Test Output";*/
/**/
/*%DESCRIPTIVE(*/
/*	DATASET=analytic2, /* Input data set */
/*	DRIVER=C:\Users\mwetze2\Box Sync\Standard Code\Sandbox\Test Output\Test Var List Code Driver.xlsx, /* File path and name for Excel driver file - no quotes*/
/*	OUTLIB=out , /* Name of library to save out SAS data sets to */
/*	OUTPATH=C:\Users\mwetze2\Box Sync\Standard Code\Sandbox\Test Output, /* File pathway for RTF output, no quotes */
/*	FNAME=Testing,  /* File name for RTF output */
/*	CLASSVAR=ERAS, /* CAN ONLY TAKE ONE CLASS VARIABLE */
/*	ODSGRAPHS=N, /* Create or suppress ODS graphics (Y/N) */
/*	FEXACT=Y,	/* Run Fisher's exact tests - See note at top of program */
/*	ROUNDTO=0.01); /* Decimal place to round results to */
