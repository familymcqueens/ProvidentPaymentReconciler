##
## Script to read AutoManager and QuickBooks Bank Register Files
##

my $YELLOW = "yellow";
my $BLUE   = "blue";
my $GREEN  = "green";
my $RED = "red";

my $TableIterator = 1;

$numArgs = $#ARGV + 1;

## Make sure we have the right command line arguments
if ( $numArgs != 2 )
{
	print "ProvidentPaymentReconciler.pl <AutoManager .csv file> <Quickbooks .csv file>\n";
	exit 1;
}

## Open file and sure the file AutoManager file exists
if (open(AM_INPUT_FILE,$ARGV[0]) == 0) {
   print "Error opening: ",$ARGV[0];
   exit -1;  
}

## Open file and sure the file Quickbooks file exists
if (open(QB_INPUT_FILE,$ARGV[1]) == 0) {
   print "Error opening: ",$ARGV[1];
   exit -1;  
}

## Open AutoManager output file 
if (open(AM_OUTPUT_FILE,'>AutoManager_output.csv') == 0) {
   print "Error opening: AutoManager_output.csv";
   exit -1;  
}

## Open Quickbooks output file 
if (open(QB_OUTPUT_FILE,'>QuickBooks_output.csv') == 0) {
   print "Error opening: QuickBooks_output.csv";
   exit -1;  
}

## Read the AutoManager input file 
while (<AM_INPUT_FILE>) {
 chomp;
 ($date,$total,$expense,$name,$payment_type,$notes,$invoice,$check,$part,$balance,$interest,$principal,$lot,$vehicle,$status,$receipt) = split(",");
 
 	 $amountabs = sprintf('%.2f',abs($amount));
	 
	 # Check to see if entry is a payment; if so, write out payment to output file	 
	 if ($payment_type eq "PAYMENT" )
	 {  
		print AM_OUTPUT_FILE "$date",",","$name",",Interest Income,",sprintf('%.2f',abs($interest)),"\n";
		print AM_OUTPUT_FILE "$date",",","$name",",TOTAL NOTES RECEIVABLE,",sprintf('%.2f',abs($principal)),"\n";
	 }
	 
	 # Check to see if entry is a late file; if so, write to output file
	 elsif ($payment_type eq "LATE FEE" )
	 {
		print AM_OUTPUT_FILE "$date",",","$name",",LATE FEE,\n";
	 } 	 
 }
 
 #print "\n\n*** END OF AUTO MANAGER FILE PARSING ***\n\n";
 #close(AM_INPUT_FILE);
 
 # QUICKBOOKS DEPOSIT FORMAT  
 # 4/7/2014		Deposit	WASHINGTON MUTUAL	300
 #				
 #	Martinez-Lopez, Salvador	Deposit	Late Fees	-38.14
 #	Martinez-Lopez, Salvador	Deposit	Interest Income	-261.86
 # TOTAL				-300
 
 ## Start reading of Quickbooks file
 while (<QB_INPUT_FILE>) {
 chomp;
 ($type, $num, $date,$name,$memo,$payment_type,$amount) = split(",");

  $date=~ s/"//;	
  $date=~ s/"//;
  $name=~ s/"//;	
  $name=~ s/"//;
  $payment_type=~ s/"//;	
  $payment_type=~ s/"//;


 	## If the name element is non-empty, then we have found an actual deposit row.  
	if ($name ne "")
	{
		($type,$num,$date,$lastname, $firstname,$memo,$payment_type,$amount) = split(",");
		$ucfirstname = uc($firstname);
		$uclastname = uc($lastname);
		$ucfirstname =~ s/"//;	
		$uclastname =~ s/"//;
		$ucfirstname =~ s/ //;  
		$memo=~ s/"//;	
		$memo=~ s/"//;
		$payment_type=~ s/"//;	
		$payment_type=~ s/"//;
              		
		$name = sprintf("%s %s",$uclastname,$ucfirstname);
		}
	## 
 
	## Since the date is listed ONCE per deposit, check to see if the date value 
	## is of non-zero length and store off in $deposit_date variable
        if (length($date) && ($date) != "TOTAL")
	{
		$deposit_date = $date;
	}
	
	## Use absolute value
	$amountabs = sprintf('%.2f',abs($amount));
	
	## Set name to uppercase for text comparision
	$ucname = uc($name);
		
	## Check for Interest portion of payment	
	 if ($payment_type eq "Interest Income" )
	 {
		# print $deposit_date,",",$ucname,",Interest Income,",$amountabs,",",$memo,"\n";
		print QB_OUTPUT_FILE $deposit_date,",",$ucname,",Interest Income,",$amountabs,",",$memo,"\n";		
	 }
	 
	 ## Check for principal portion of payment
	 elsif ($payment_type eq "TOTAL NOTES RECEIVABLE" )
	 {
		# print $deposit_date,",",$ucname,",TOTAL NOTES RECEIVABLE,",$amountabs,",",$memo,"\n";
		print QB_OUTPUT_FILE $deposit_date,",",$ucname,",TOTAL NOTES RECEIVABLE,",$amountabs,",",$memo,"\n";
	 }
	 
	 ## Check for late fee portion of payment
	 elsif ($payment_type eq "Late Fees" )
	 {
		print QB_OUTPUT_FILE $deposit_date,",",$ucname,",LATE FEE,",$amountabs,",",$memo,"\n";
	 }

	 ## Check to see if payment was incorrectly entered unter Interest Expense account
	 elsif ($payment_type eq "Interest Expense" )
	 {
		print QB_OUTPUT_FILE $deposit_date,",",$ucname,",","Interest Expense",",",$amountabs,"\n";
	 }
	 else
	 {
		if ($amountabs != 0)
		{
			#print "----> ", $date," ",$ucname," ",$payment_type, " ",$amountabs," ",$memo,"\n";
	 }	}
 }
 
close(QB_INPUT_FILE);
# print "\n\n*** END OF QUICKBOOKS FILE PARSING ***\n\n";

print "\nParsing output sent to file 'AutoManager_output.csv'";
print "\nParsing output sent to file 'QuickBooks_output.csv'\n";

close(AM_OUTPUT_FILE);
close(QB_OUTPUT_FILE);


##
## This script will read the output files from AutoManager and Quickbooks
## and find the possible unmatched entries.
##
open (AM_OUTPUT_FILE, 'AutoManager_output.csv');
open (QB_OUTPUT_FILE, 'QuickBooks_output.csv');

## Open Payment Reconciler HTML output file 
if (open(PR_HTML_OUTPUT_FILE,'>ProvidentPaymentReconciler.html') == 0) {
   print "Error opening: ProvidentPaymentReconciler.html";
   exit -1;  
}

my $htmlFileHandle = \*PR_HTML_OUTPUT_FILE;

print $htmlFileHandle "<html>\n";
print $htmlFileHandle "<head><title>Provident Financial Payment Reconciler Utility </title></head>\n";
print $htmlFileHandle "<body>\n";
print $htmlFileHandle "<h1>Provident Financial Payment Reconciler Utility</h1><br>\n";

HtmlTableTopSection($htmlFileHandle,$BLUE);
print   $htmlFileHandle "<tr><th>Error Key</th><th>Description</th></tr>\n";
print   $htmlFileHandle "<tr>\n";
print   $htmlFileHandle "<td align=\"left\">","PAYMENT TYPE MISMATCH</td>\n";
print   $htmlFileHandle "<td align=\"left\">","Payments with the 'Total Notes Receivable' and 'Interest Income' swapped","</td>\n";
print   $htmlFileHandle "</tr>\n";
print   $htmlFileHandle "<tr>\n";
print   $htmlFileHandle "<td align=\"left\">","PAYMENT PROBABLE MATCH</td>\n";
print   $htmlFileHandle "<td align=\"left\">","First letter of first and last name, payment type, and payment amount match, but exact match","</td>\n";
print   $htmlFileHandle "</tr>\n";
print   $htmlFileHandle "<tr>\n";
print   $htmlFileHandle "<td align=\"left\">","PAYMENT WEAK MATCH</td>\n";
print   $htmlFileHandle "<td align=\"left\">","First letter of last name and payment amount match, but not exact match","</td>\n";
print   $htmlFileHandle "</tr>\n";
print   $htmlFileHandle "<tr>\n";
print   $htmlFileHandle "<td align=\"left\">","PAYMENT EXACT MATCH</td>\n";
print   $htmlFileHandle "<td align=\"left\">","First name, last name, payment type, and payment amount match","</td>\n";
print   $htmlFileHandle "</tr>\n";

HtmlTableTailSection($htmlFileHandle);
print   $htmlFileHandle "<br><br><br>\n";

$num_exact_matches = 0;
$num_weak_matches = 0;
$num_missing_payments = 0;
$num_probable_matches = 0;
$num_payment_type_mismatches = 0;
$num_missing_payment_unknown_reason = 0;
$num_late_fee_payments = 0;
$num_total_entries = 0;

print "\nProcessing Quickbooks deposits...\n\n";

# Walk over Quickbook output file
 while (<QB_OUTPUT_FILE>) {
 chomp;
 ($qb_date,$qb_name, $qb_payment_type, $qb_amount,$qb_memo) = split(",");

 	$num_total_entries++;
  	$found_qb_entry = 0;
	$found_name_amount_match = 0;
	$found_shortname_entry = 0;
	$found_weak_match_entry = 0;
	
	print {STDERR} ".";
	@tokens = split(/ /, $qb_name);
	$loop = 0;
	foreach my $token (@tokens) 
	{
		if ($loop eq 0)
		{
			$qb_last_name = $token;
		}
		
		if ($loop eq 1)
		{
			$qb_first_name = $token;
		}			
		$loop++;
	}
	
	
	# Walk over AutoManager output file
	while (<AM_OUTPUT_FILE>) {
		chomp;
		($am_date,$am_name, $am_payment_type, $am_amount) = split(",");
	  
		$am_name = uc($am_name);
	  
	    @tokens = split(/ /, $am_name);
		$loop = 0;
		foreach my $token (@tokens) 
		{
		    if ($loop eq 0)
			{
				$am_last_name = $token;
			}
			
			if ($loop eq 1)
			{
				$am_first_name = $token;
			}			
			$loop++;
	    }		
		
				
		# Look for EXACT match with NAME, PAYMENT TYPE, and AMOUNT
	    if (($am_name eq $qb_name) && ($am_payment_type eq  $qb_payment_type) && ($am_amount eq $qb_amount) && ($qb_memo eq "Deposit"))
		{
			$num_exact_matches++;

			$exact_match_log = sprintf("%s:%s:%s:%s:%s",uc($qb_name), uc($qb_payment_type), $qb_amount, $qb_date, $am_date);
			push (@exact_match_log, $exact_match_log);	
			$found_qb_entry = 1;
		}
		
		## Look for match with only NAME and AMOUNT
	    if (($am_name eq $qb_name) && ($am_amount eq $qb_amount))
		{
			$found_name_amount_match = 1;
			$am_date_match = $am_date;
			$am_name_match = $am_name;
			$am_payment_type_match = $am_payment_type;
			$qb_payment_type_match = $qb_payment_type;
			$am_amount_match = $am_amount;
			$qb_amount_match = $qb_amount;
			$qb_memo_match = $qb_memo;
		}
		
		$am_first_shortname = sprintf("%.1s", $am_first_name);
		$qb_first_shortname = sprintf("%.1s", $qb_first_name);
		
		$am_last_shortname = sprintf("%.1s", $am_last_name);
		$qb_last_shortname = sprintf("%.1s", $qb_last_name);
			
		## Look for match with "short" last name and amount only
		if (($am_last_shortname eq $qb_last_shortname) && ($am_amount eq $qb_amount))
		{
			## Look for further match of "short" first name and payment type
			if (($am_first_shortname eq $qb_first_shortname) && 
				($am_payment_type_match) eq ($qb_payment_type_match))
			{
				$found_shortname_entry = 1;
				$am_shortname_date_match = $am_date;
				$qb_shortname_date_match = $qb_date;
				$am_shortname_match = $am_name;
				$qb_shortname_match = $qb_name;
				$am_shortname_payment_type_match = $am_payment_type;
				$am_shortname_amount_match = $am_amount;
			}
			elsif (uc($qb_memo) eq "DEPOSIT")
			{
				$found_weak_match_entry = 1;
				$am_weak_match_date = $am_date;
				$qb_weak_match_date = $qb_date;
				$am_weak_match_name = $am_name;
				$qb_weak_match_name = $qb_name;
				$qb_weak_match_memo = $qb_memo;
				$am_weak_match_payment_type = $am_payment_type;
				$qb_weak_match_payment_type = $qb_payment_type;
				$am_weak_match_amount = $am_amount;
			}
		}       		
    }
	
	if ($found_qb_entry eq 0)
	{
	    if ($found_name_amount_match)
		{
			$num_payment_type_mismatches++;		
			
			HtmlTableTopSection($htmlFileHandle,$RED);
			print   $htmlFileHandle "<tr><th>Error Description</th><th>Date</th><th>Name</th><th>Payment</th><th>QuickBooks Type</th><th>AutoManager Type</th><th>Quickbooks Memo</th>\n";
			print   $htmlFileHandle "<tr>\n";
			print   $htmlFileHandle "<td align=\"left\">","PAYMENT TYPE MISMATCH or QB MEMO != DEPOSIT</td>\n";
			print   $htmlFileHandle "<td align=\"center\">",$qb_date,"</td>\n";
			print   $htmlFileHandle "<td align=\"left\">",uc($am_name_match),"</td>\n";
			print   $htmlFileHandle "<td align=\"center\">",$am_amount_match,"</td>\n";
			print   $htmlFileHandle "<td align=\"center\">",uc($qb_payment_type_match),"</td>\n";
			print   $htmlFileHandle "<td align=\"center\">",uc($am_payment_type_match),"</td>\n";
			print   $htmlFileHandle "<td align=\"center\">",uc($qb_memo_match),"</td>\n";
			print   $htmlFileHandle "</tr>\n";			
			HtmlTableTailSection($htmlFileHandle);
			
			print "\n_______________________________________________________\n";
			print "\nERROR: Payment Type Mismatch\n";
			print "   --> Date: ", $am_date_match," Name: ",$am_name_match," Amount: ",$am_amount_match,"\n";
			print "   --> Quickbooks Entry: [", $qb_payment_type_match, "]  Memo: [",uc($qb_memo_match),"]\n";
			print "   --> AutoManager Entry: [", $am_payment_type_match, "]\n";
		}
		
		elsif ($found_shortname_entry && $am_shortname_amount_match)
		{
		    $num_probable_matches++;

			HtmlTableTopSection($htmlFileHandle,$RED);
			print   $htmlFileHandle "<br><tr><th>Error Description</th><th>Payment</th><th>Payment Type</th><th>QuickBooks Date-Name</th><th>AutoManager Date-Name</th>\n";
			print   $htmlFileHandle "<tr>\n";
			print   $htmlFileHandle "<td align=\"left\">","PAYMENT PROBABLE MATCH</td>\n";
			print   $htmlFileHandle "<td align=\"left\">",$am_shortname_amount_match,"</td>\n";
			print   $htmlFileHandle "<td align=\"center\">",uc($am_shortname_payment_type_match),"</td>\n";
			print   $htmlFileHandle "<td align=\"center\">",$qb_shortname_date_match, " ",uc($qb_shortname_match),"</td>\n";
			print   $htmlFileHandle "<td align=\"center\">",$am_shortname_date_match, " ",uc($am_shortname_match),"</td>\n";
			print   $htmlFileHandle "</tr><br>\n";			
			HtmlTableTailSection($htmlFileHandle);
			
			print "\n_______________________________________________________\n";
			print "\nProbable Match\n";
			print "   --> Amount: [",$am_shortname_amount_match,"] Payment Type: [",$am_shortname_payment_type_match,"]\n";
			print "   --> Quickbooks Entry:  [",$qb_shortname_date_match," ",$qb_shortname_match,"]\n";
			print "   --> AutoManager Entry: [",$am_shortname_date_match," ",$am_shortname_match,"]\n";
		}
		elsif ($found_weak_match_entry && ($am_weak_match_amount ) && (uc($qb_payment_type) ne "LATE FEE"))
		{
			$num_weak_matches++;
		    
			HtmlTableTopSection($htmlFileHandle,$RED);
			print   $htmlFileHandle "<br><tr><th>Error Description</th><th>Payment</th><th>Quickbooks Memo</th><th>QuickBooks Date-Name-Type</th><th>AutoManager Date-Name-Type</th>\n";
			print   $htmlFileHandle "<tr>\n";
			print   $htmlFileHandle "<td align=\"left\">","PAYMENT WEAK MATCH</td>\n";
			print   $htmlFileHandle "<td align=\"left\">",$am_weak_match_amount,"</td>\n";
			print   $htmlFileHandle "<td align=\"center\">",uc($qb_weak_match_memo),"</td>\n";
			print   $htmlFileHandle "<td align=\"center\">",$qb_weak_match_date, " ",uc($qb_weak_match_name)," ", uc($qb_weak_match_payment_type),"</td>\n";
			print   $htmlFileHandle "<td align=\"center\">",$am_weak_match_date, " ",uc($am_weak_match_name)," ", uc($am_weak_match_payment_type),"</td>\n";
			print   $htmlFileHandle "</tr><br>\n";			
			HtmlTableTailSection($htmlFileHandle);
			
			print "\n_______________________________________________________\n";
			print "\nWeak Match\n";
			print "   --> Amount: [",$am_weak_match_amount,"]\n";
			print "   --> Quickbooks Memo: [",$qb_weak_match_memo,"]\n";
			print "   --> Quickbooks Entry:  [",$qb_weak_match_date," ",$qb_weak_match_name," ", $qb_weak_match_payment_type,"]\n";
			print "   --> AutoManager Entry: [",$am_weak_match_date," ",$am_weak_match_name," ", $am_weak_match_payment_type,"]\n";
		}
		else
		{
			if ((uc($qb_memo) eq "DEPOSIT"))
			{
				## In AutoManager, late fee amounts are not listed, so, we have to make sure we make a case for this exception.
				if (uc($qb_payment_type) eq "LATE FEE")
				{
					$num_late_fee_payments++;
				}
				else
				{
					$num_missing_payments++;
					
					HtmlTableTopSection($htmlFileHandle,$RED);
					print   $htmlFileHandle "<br><tr><th>Error Description</th><th>Date</th><th>Name</th><th>Payment Type</th><th>Amount</th>\n";
					print   $htmlFileHandle "<tr>\n";
					print   $htmlFileHandle "<td align=\"left\">","QBOOKS DEPOSIT DUPLICATE OR MISSING IN AUTOMANAGER, NSF??</td>\n";
					print   $htmlFileHandle "<td align=\"left\">",$qb_date,"</td>\n";
					print   $htmlFileHandle "<td align=\"center\">",uc($qb_name),"</td>\n";
					print   $htmlFileHandle "<td align=\"center\">",uc($qb_payment_type),"</td>\n";
					print   $htmlFileHandle "<td align=\"center\">",$qb_amount,"</td>\n";
					print   $htmlFileHandle "</tr><br>\n";			
					
					
					# START : Find possible match and put in a drop down select
					# TBD: Change this to read a memory array
					seek(AM_INPUT_FILE,0,0);
					my $myNumFoundEntries = 0;
					while (<AM_INPUT_FILE>) {
					chomp;
					($date,$total,$expense,$name,$payment_type,$notes,$invoice,$check,$part,$balance,$interest,$principal,$lot,$vehicle,$status,$receipt) = split(",");
					 
						 if ((abs($interest) eq $qb_amount) || (abs($principal) eq $qb_amount))
						 {
							print $htmlFileHandle "<tr>\n";
							print $htmlFileHandle "<td align=\"left\">","AUTOMANAGER POSSIBLE MATCH</td>\n";
							print $htmlFileHandle "<td align=\"left\">",$date,"</td>\n";
							print $htmlFileHandle "<td align=\"center\">",uc($name),"</td>\n";
							print $htmlFileHandle "<td align=\"left\">","Principal: ",$principal," Interest: ",$interest,"</td>\n";		
							$myNumFoundEntries++;								
						 }						  	 
					}					
					if ($myNumFoundEntries gt 0)
					{
						print $htmlFileHandle "</tr><br>\n";
					}
					# END : Find possible match and put in a drop down select
					
					
					HtmlTableTailSection($htmlFileHandle);
					
				}
			}
			else
			{
				if (uc($qb_payment_type) ne "LATE FEE")
				{
					$num_missing_payment_unknown_reason++;				
					$unknown_missing_log = sprintf("%s:%s:%s:%s:%s",
						uc($qb_memo), $qb_date, uc($qb_name), uc($qb_payment_type), $qb_amount);
						
					push (@unknown_missing_log, $unknown_missing_log);
				}
			}
		}
	}
	
	seek AM_OUTPUT_FILE, 0, 0; 
 } 
 
print  PR_HTML_OUTPUT_FILE "</body>\n";
print  PR_HTML_OUTPUT_FILE "</html>\n";
 
close(AM_OUTPUT_FILE);
close(QB_OUTPUT_FILE);

print "\n\n -> Number of Late Payments (Skipped) = ", $num_late_fee_payments, "\n\n";

print " Exact Match:\n";
print " First name, last name, payment type, and payment amount match.\n";
print " -> Number of Exact Matches = ", $num_exact_matches, "\n\n";

print " Probable Match:\n";
print " First letter of first and last name, payment type, and payment amount match.\n";
print " -> Number of Probable Matches = ", $num_probable_matches, "\n\n";

print " Weak Match:\n";
print " First letter of last name and payment amount match.\n";
print " -> Number of Weak Matches = ", $num_weak_matches, "\n\n";

print " Unknown Missing Payment:\n";
print " These payments are usually of type (NSF,REG,REPO,etc..)\n";
print " -> Number of Unknown Missing Payments = ", $num_missing_payment_unknown_reason, "\n\n";

print " Missing Payment:\n";
print " Fails all other Auto Manager payment match criteria.\n";
print " -> Number of Missing Entries = ", $num_missing_payments, "\n\n";

print " Payment Type Mismatch:\n";
print " Payments with the 'Total Notes Receivable' and 'Interest Income' swapped\n";
print " -> Number of Payment Type Mismatches = ", $num_payment_type_mismatches, "\n\n";

print "\nTotal Payment Entries Read:    ", $num_total_entries;
print "\nTotal Payment Entries Matched: ", $num_exact_matches + $num_late_fee_payments + $num_probable_matches + $num_weak_matches + $num_missing_payment_unknown_reason + $num_missing_payments + $num_payment_type_mismatches,"\n";

##print "\n\nDisplay unknown (missing) entries? (NSF/REG/REPO,etc...) --> (y/n) ";
##$promptForEachText = <STDIN>;
##chomp $promptForEachText;

##if ($promptForEachText eq "y")
{
	##print(@unknown_missing_log); 
	
	print $htmlFileHandle "<br><br><br>";
	HtmlTableTopSection($htmlFileHandle,$GREEN);
	print $htmlFileHandle "<tr><th colspan=5>The following Quickbook entries are MISSING from AutoManager</th>\n";	
	print $htmlFileHandle "<tr><th>Memo</th><th>Date</th><th>Name</th><th>Payment Type</th><th>Amount</th>\n";
	
	for my $entry (@unknown_missing_log) 
	{
		my @values = split(':', $entry);
		print   $htmlFileHandle "<tr>\n";
		
		foreach my $val (@values) 
		{
			print $htmlFileHandle "<td align=\"left\">",$val,"</td>\n";		
		}		
		print $htmlFileHandle "</tr>\n";		
	}	
	HtmlTableTailSection($htmlFileHandle);
	
	
	
	
	print $htmlFileHandle "<br><br><br>";	
	HtmlTableTopSection($htmlFileHandle,$GREEN);	
	print $htmlFileHandle "<tr><th colspan=5>The following entries are EXACT matches in QuickBooks and AutoManager</th>\n";
	print $htmlFileHandle "<tr><th>Name</th><th>Payment Type</th><th>Amount</th><th>Date (Quickbooks)</th><th>Date (AutoManager)</th>\n";
	
	for my $entry (@exact_match_log) 
	{
		my @values = split(':', $entry);
		print   $htmlFileHandle "<tr>\n";
		
		foreach my $val (@values) 
		{
			print $htmlFileHandle "<td align=\"left\">",$val,"</td>\n";		
		}		
		print $htmlFileHandle "</tr>\n";		
	}
	
	HtmlTableTailSection($htmlFileHandle);
	
	
}

close(PR_HTML_OUTPUT_FILE);


sub HtmlTableTopSection 
{
    my $fileHandle = $_[0];
	my $headerColor = $_[1];	
	
	print $fileHandle "<head><style>\n";
	print $fileHandle "table  { width:80%;}\n";
	print $fileHandle "th, td { padding:10px;}\n";
	print $fileHandle "table#table",$TableIterator," tr:nth-child(even) { background-color: #eee; }\n";
	print $fileHandle "table#table",$TableIterator," tr:nth-child(odd)  { background-color: #fff; }\n";
	print $fileHandle "table#table",$TableIterator," th { background-color: ",$headerColor,"; color: white; }\n";
	print $fileHandle "</style></head>\n";
	print $fileHandle "<table border=5 id=\"table",$TableIterator,"\" >\n";
	$TableIterator++;
}

sub HtmlTableTailSection
{
    my $fileHandle = $_[0];		
	print  $fileHandle "</table><br>\n";
}


