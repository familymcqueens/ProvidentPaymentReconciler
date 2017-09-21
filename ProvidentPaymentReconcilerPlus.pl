##
## Script to read AutoManager and QuickBooks Bank Register Files
##

my $YELLOW = "yellow";
my $BLUE   = "blue";
my $GREEN  = "green";
my $RED = "red";

my $TableIterator = 1;
my $numArgs = $#ARGV + 1;

my @AMA;  # Auto Manager Array
my @QBA;  # QuickBooks Array

my $amIndex = 0;
my $qbIndex = 0;

## Make sure we have the right command line arguments
if ( $numArgs != 2 )
{
	print "ProvidentPaymentReconciler.pl <AutoManager.csv file> <Quickbooks.csv file>\n";
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

use constant {
	INTEREST_INCOME_TYPE   => 1,
	PRINCIPAL_INCOME_TYPE => 2,
	LATE_FEE_INCOME_TYPE => 3,
	INTEREST_EXPENSE_INCOME_TYPE  => 4,
};
 
## Read the AutoManager input file 
use constant {
	AM_DATE_INDEX => 0,
	AM_NAME_INDEX => 1,
	AM_PAYMENT_TYPE_INDEX => 2,
	AM_PAYMENT_AMOUNT_INDEX => 3,
	AM_NOTES_INDEX => 4,
	AM_QB_EXACT_MATCH_INDEX => 5
};


while (<AM_INPUT_FILE>) {
 chomp;
 my ($date,$total,$expense,$name,$payment_type,$notes,$invoice,$check,$part,$balance,$interest,$principal,$lot,$vehicle,$status,$receipt) = split(",");
 
 	if ($payment_type eq "PAYMENT")
	{
		$AMA[$amIndex][AM_DATE_INDEX] = $date;
		$AMA[$amIndex][AM_NAME_INDEX] = uc($name);
		$AMA[$amIndex][AM_NOTES_INDEX] = uc($notes);		
		$AMA[$amIndex][AM_PAYMENT_TYPE_INDEX] = INTEREST_INCOME_TYPE;
		$AMA[$amIndex][AM_PAYMENT_AMOUNT_INDEX] = sprintf('%.2f',abs($interest));
		$AMA[$amIndex][AM_QB_EXACT_MATCH_INDEX] = -1;
		
		$amIndex++;
		
		$AMA[$amIndex][AM_DATE_INDEX] = $date;
		$AMA[$amIndex][AM_NAME_INDEX] = uc($name);
		$AMA[$amIndex][AM_NOTES_INDEX] = uc($notes);		
		$AMA[$amIndex][AM_PAYMENT_TYPE_INDEX] = PRINCIPAL_INCOME_TYPE;
		$AMA[$amIndex][AM_PAYMENT_AMOUNT_INDEX] = sprintf('%.2f',abs($principal));	
		$AMA[$amIndex][AM_QB_EXACT_MATCH_INDEX] = -1;		
	}
	elsif ($payment_type eq "LATE FEE")
	{
		$AMA[$amIndex][AM_DATE_INDEX] = $date;
		$AMA[$amIndex][AM_NAME_INDEX] = uc($name);
		$AMA[$amIndex][AM_NOTES_INDEX] = uc($notes);		
		$AMA[$amIndex][AM_PAYMENT_TYPE_INDEX] = LATE_FEE_INCOME_TYPE;
		$AMA[$amIndex][AM_PAYMENT_AMOUNT_INDEX] = sprintf('%.2f',abs($total));
		$AMA[$amIndex][AM_QB_EXACT_MATCH_INDEX] = -1;
	}
	else 
	{
		# Make sure to set down payments, warranties, etc.. to index -1.
		#print "Payment Type: ", $payment_type, " Index: ",$amIndex," Name: ",$name, " Amt:",$total,"\n";
		$AMA[$amIndex][AM_QB_EXACT_MATCH_INDEX] = -1;
	}
	
	$amIndex++;		
 }

 
 # QUICKBOOKS DEPOSIT FORMAT  
 # 4/7/2014		Deposit	WASHINGTON MUTUAL	300
 #				
 #	Martinez-Lopez, Salvador	Deposit	Late Fees	-38.14
 #	Martinez-Lopez, Salvador	Deposit	Interest Income	-261.86
 # TOTAL				-300

use constant {
	QB_DATE_INDEX   => 0,
	QB_NAME_INDEX => 1,
	QB_AMOUNT_INDEX => 2,
	QB_PAYMENT_TYPE_INDEX  => 3,
	QB_MEMO_INDEX  => 4,
	QB_AM_EXACT_MATCH_INDEX => 5,
	QB_AM_NAME_AMOUNT_MATCH_INDEX => 6,
	QB_AM_SHORTNAME_AMOUNT_MATCH_INDEX => 7,
	QB_AM_WEAKNAME_MATCH_INDEX => 8,
	QB_AM_NAME_TYPE_MATCH_INDEX => 9
};

## Start reading of Quickbooks file

my $deposit_date;

 while (<QB_INPUT_FILE>) {
 chomp;
 ($type, $num, $date,$name,$memo,$payment_type,$amount) = split(",");

	## Since the date is listed ONCE per deposit, check to see if the date value 
	## is of non-zero length and store off in $deposit_date variable
    if (length($date) && ($date ne "TOTAL"))
	{
		$deposit_date = $date;
	}
	
	if ($name eq "" || ($amount =~ /^[a-zA-Z]+$/))
	{
		next;
	}
	
	$date=~ s/"//;	
	$date=~ s/"//;
	$name=~ s/"//;	
	$name=~ s/"//;
	$payment_type=~ s/"//;	
	$payment_type=~ s/"//;
	
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
 		
	## Use absolute value
	$amountabs = sprintf('%.2f',abs($amount));
	
	my $income_type;
	
	## Check for Interest portion of payment	
	if ($payment_type eq "Interest Income" )
	{
		$income_type = INTEREST_INCOME_TYPE;		
	}

	## Check for principal portion of payment
	elsif ($payment_type eq "TOTAL NOTES RECEIVABLE" )
	{
		$income_type = PRINCIPAL_INCOME_TYPE;
	}

	## Check for late fee portion of payment
	elsif ($payment_type eq "Late Fees" )
	{
		$income_type = LATE_FEE_INCOME_TYPE;
	}

	## Check to see if payment was incorrectly entered unter Interest Expense account
	elsif ($payment_type eq "Interest Expense" )
	{
		$income_type = INTEREST_EXPENSE_INCOME_TYPE;
	}
		
	$QBA[$qbIndex][QB_DATE_INDEX] = $deposit_date;
	$QBA[$qbIndex][QB_NAME_INDEX] = uc($name);
	$QBA[$qbIndex][QB_AMOUNT_INDEX] = $amountabs;
	$QBA[$qbIndex][QB_MEMO_INDEX] = uc($memo);
	$QBA[$qbIndex][QB_PAYMENT_TYPE_INDEX] = $income_type;
	$QBA[$qbIndex][QB_AM_EXACT_MATCH_INDEX] = -1;
	
	$qbIndex++;	
 }
  
# print "\n\n*** END OF QUICKBOOKS FILE PARSING ***\n\n";

close (QB_INPUT_FILE);
close (AM_INPUT_FILE);
 
## Loop over all QB array entries, for each entry, loop in the AM array entries
## If an EXACT match with NAME, PAYMENT TYPE, and AMOUNT
## If found NAME and AMOUNT match, mark AM array indexes that match 
## If found NAME match, mark AM array indexes that match
## If found AMOUNT match, mark AM array indexes that match
##
## If found NAME and AMOUNT match, then, it is a type mismatch
## Else If found only NAME match(es), the AMOUNT could be entered wrong, list all possible matches
## Else If found only AMOUNT match(es), the NAME could be entered wrong, list all possible matches
## Else If found only SHORT NAME match(es) and AMOUNT, the NAME could be entered wrong, list all possible matches


##
## This script will read the output files from AutoManager and Quickbooks
## and find the possible matched entries.
##

## Open Payment Reconciler HTML output file 
if (open(PR_HTML_OUTPUT_FILE,'>ProvidentPaymentReconciler.html') == 0) {
   print "Error opening: ProvidentPaymentReconciler.html";
   exit -1;  
}

my $htmlFileHandle = \*PR_HTML_OUTPUT_FILE;

print $htmlFileHandle "<html>\n";
print $htmlFileHandle "<head><title>Provident Financial Payment Reconciler Utility </title></head>\n";
print $htmlFileHandle "<body>\n";
print $htmlFileHandle "<h1>Provident Financial Payment Reconciler Utility</h1>\n";

my $num_exact_matches = 0;
my $num_weak_matches = 0;
my $num_missing_payments = 0;
my $num_probable_matches = 0;
my $num_payment_type_mismatches = 0;
my $num_missing_payment_unknown_reason = 0;
my $num_late_fee_payments = 0;
my $num_total_entries = 0;

print "\nProcessing Quickbooks deposits...\n\n";

my $nameAmountMatchIndex = 0;
my $weakNameAmountMatchIndex = 0;
my $shortNameAmountMatchIndex = 0;
my $found_exact_match = 0;

print "Size of QBooks Array: ",$#QBA,"\n";
print "Size of AutoManager Array: ",$#AMA,"\n";

my $debug = 1;

# Walk over Quickbook entries
for my $i (0 .. ($#QBA)) 
{
	my $qb_date         = $QBA[$i][QB_DATE_INDEX];
	my $qb_name         = $QBA[$i][QB_NAME_INDEX];
	my $qb_amount       = $QBA[$i][QB_AMOUNT_INDEX];
	my $qb_memo         = $QBA[$i][QB_MEMO_INDEX];
	my $qb_payment_type = ConvertPaymentTypeToString($QBA[$i][QB_PAYMENT_TYPE_INDEX]);
	
	$num_total_entries++;
	
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

	$nameAmountMatchIndex = 0;
	$weakNameAmountMatchIndex = 0;
	$shortNameAmountMatchIndex = 0;
	$namePaymentTypeMatchIndex = 0;
	$found_exact_match = 0;
	
	# Look for matching Auto Manager entries
	for my $j (0 .. ($#AMA - 1)) 
	{
		my $am_date         = $AMA[$j][AM_DATE_INDEX];
		my $am_name         = $AMA[$j][AM_NAME_INDEX];
		my $am_notes        = $AMA[$j][AM_NOTES_INDEX];
		my $am_payment_type = ConvertPaymentTypeToString($AMA[$j][AM_PAYMENT_TYPE_INDEX]);
		my $am_amount       = $AMA[$j][AM_PAYMENT_AMOUNT_INDEX];
		  
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
		
		$am_first_shortname = sprintf("%.1s", $am_first_name);
		$am_last_shortname  = sprintf("%.1s", $am_last_name);
		$qb_first_shortname = sprintf("%.1s", $qb_first_name);
		$qb_last_shortname  = sprintf("%.1s", $qb_last_name);
		
		# Look for EXACT match with NAME, PAYMENT TYPE, and AMOUNT
	    if (($am_name eq $qb_name) && ($am_payment_type eq  $qb_payment_type) && 
		    ($am_amount eq $qb_amount) && (($qb_memo eq "DEPOSIT") || ($qb_memo eq $am_notes)))
		{
			$num_exact_matches++;			
			$QBA[$i][QB_AM_EXACT_MATCH_INDEX] = $j;	
			$ABA[$i][QB_AM_EXACT_MATCH_INDEX] = $j;
			
			##print "EXACT MATCH: AmIndex: ",$j," QBIndex:",$i,":",$qb_name,":",$qb_amount,":",$qb_memo,":",$qb_payment_type,"\n";
			
			$exact_match_log = sprintf("%s:%s:%s:%s:%s:%s",$j,$qb_name, $qb_payment_type, $qb_amount, $qb_date, $am_date);
			push (@exact_match_log, $exact_match_log);
			
			$found_exact_match = 1;
		}
		else
		{
			$found_exact_match = 0;
		}
		
		if ($found_exact_match)
		{
			last;
		}
		
		## Look for match with only NAME and AMOUNT
	    elsif (($am_name eq $qb_name) && ($am_amount eq $qb_amount))
		{
			$QBA[$i][QB_AM_NAME_AMOUNT_MATCH_INDEX][$nameAmountMatchIndex++] = $j;
		}
		## Look for match with only LAST NAME and AMOUNT
		elsif (($am_last_name eq $qb_last_name) && ($am_amount eq $qb_amount))
		{
			$QBA[$i][QB_AM_WEAKNAME_MATCH_INDEX][$weakNameAmountMatchIndex++] = $j;
		} 
		
		## Look for match with "short" last name and amount only
		elsif (($am_last_shortname eq $qb_last_shortname) && ($am_amount eq $qb_amount))
		{
			## Look for further match of "short" first name and payment type
			if (($am_first_shortname eq $qb_first_shortname) && ($am_payment_type) eq ($qb_payment_type))
			{
				$QBA[$i][QB_AM_SHORTNAME_AMOUNT_MATCH_INDEX][$shortNameAmountMatchIndex++] = $j;
			}
			elsif ( ($qb_memo eq "DEPOSIT") && (($am_payment_type) eq ($qb_payment_type)) )
			{
				$QBA[$i][QB_AM_WEAKNAME_MATCH_INDEX][$weakNameAmountMatchIndex++] = $j;
			}
		}
		elsif (($am_name eq $qb_name) && ($am_payment_type) eq ($qb_payment_type))
		{
			$QBA[$i][QB_AM_NAME_TYPE_MATCH_INDEX][$namePaymentTypeMatchIndex++] = $j;
		}					
    }
 } 

print "Num Exact Matches: ",$num_exact_matches,"\n"; 
$num_exact_matches=0;

print "--------------------------------------------------\n";
 
for my $m (1 .. ($#QBA))  
{
	my $qb_date   = $QBA[$m][QB_DATE_INDEX];
	my $qb_name   = $QBA[$m][QB_NAME_INDEX];
	my $qb_amount = $QBA[$m][QB_AMOUNT_INDEX];
	my $qb_memo   = $QBA[$m][QB_MEMO_INDEX];
	my $qb_payment_type = ConvertPaymentTypeToString($QBA[$m][QB_PAYMENT_TYPE_INDEX]);

	if (($QBA[$m][QB_MEMO_INDEX]) ne "DEPOSIT" && ($QBA[$m][QB_MEMO_INDEX] ne ""))
	{
		$num_missing_payment_unknown_reason++;				
		$unknown_missing_log = sprintf("%s:%s:%s:%s:%s",
			$qb_memo,$qb_date,$qb_name,$qb_payment_type,$qb_amount);
		push (@unknown_missing_log, $unknown_missing_log);
		next;
	}
    
	if ( ($QBA[$m][QB_AM_EXACT_MATCH_INDEX] < 0) )
	{
		print "-------------------------------------------------------\n";
		print "\n-> MISSING AutoManager Match <-\n";
		print " Name: ",$qb_name,"\n";
		print " Date: ",$qb_date,"\n";
		print " Amount: ",$qb_amount,"\n";
		print " Memo:",$qb_memo,"\n";
		print " Payment Type: ",$qb_payment_type,"\n\n";
		
		HtmlTableTopSection($htmlFileHandle,$RED);
		print   $htmlFileHandle "<br><br><tr>\n";
		print   $htmlFileHandle "<tr><th>Quickbooks Description</th><th>Date</th><th>Name</th><th>Payment</th><th colspan=2>Payment Type</th>\n";
		print   $htmlFileHandle "<tr>\n";
		print   $htmlFileHandle "<td align=\"left\">",$qb_memo,"</td>\n";
		print   $htmlFileHandle "<td align=\"center\">",$qb_date,"</td>\n";
		print   $htmlFileHandle "<td align=\"left\">",$qb_name,"</td>\n";
		print   $htmlFileHandle "<td align=\"center\">",$qb_amount,"</td>\n";
		print   $htmlFileHandle "<td align=\"center\">",$qb_payment_type,"</td>\n";
		print   $htmlFileHandle "</tr>\n";
		print   $htmlFileHandle "<tr><th>AutoManager Description</th><th>Date</th><th>Name</th><th>Payment</th><th>Payment Type</th><th>Match Type</th>\n";			
		
		my $numPaymentTypeMatches = scalar @{ $QBA[$m][QB_AM_NAME_AMOUNT_MATCH_INDEX] };
		
		if ( $numPaymentTypeMatches > 0 )
		{
			print "-> PAYMENT TYPE MISMATCH COUNT:",$numPaymentTypeMatches,"\n";
			
			for my $k (0 .. ($numPaymentTypeMatches-1))
			{
				my $index = $QBA[$m][QB_AM_NAME_AMOUNT_MATCH_INDEX][$k];
				my $am_date         = $AMA[$index][AM_DATE_INDEX];
				my $am_name         = $AMA[$index][AM_NAME_INDEX];
				my $am_notes        = $AMA[$index][AM_NOTES_INDEX];
				my $am_payment_type = ConvertPaymentTypeToString($AMA[$index][AM_PAYMENT_TYPE_INDEX]);
				my $am_amount       = $AMA[$index][AM_PAYMENT_AMOUNT_INDEX];
				
				if ( $AMA[$m][AM_QB_EXACT_MATCH_INDEX] < 0 )
				{
					print "\n   Payment Type Mismatch\n";
					print "     --> Date: ", $am_date," Name: ",$am_name," Amount: ",$am_amount,"\n";
					print "     --> Quickbooks Entry: [", $qb_payment_type, "]  Memo: [",$qb_memo,"]\n";
					print "     --> AutoManager Entry: [", $am_payment_type, "]\n\n";
									
					print   $htmlFileHandle "<tr>\n";
					print   $htmlFileHandle "<td align=\"left\">",$am_notes,"</td>\n";
					print   $htmlFileHandle "<td align=\"center\">",$am_date,"</td>\n";
					print   $htmlFileHandle "<td align=\"left\">",$am_name,"</td>\n";
					print   $htmlFileHandle "<td align=\"center\">",$am_amount,"</td>\n";
					print   $htmlFileHandle "<td align=\"center\">",$am_payment_type,"</td>\n";
					print   $htmlFileHandle "<td align=\"center\">","Payment Type Mismatch","</td>\n";
					print   $htmlFileHandle "</tr>\n";	
				}
			}
			
		}
		
		my $numShortNameAmountMatches = scalar @{ $QBA[$m][QB_AM_SHORTNAME_AMOUNT_MATCH_INDEX] };
		
		if ( $numShortNameAmountMatches > 0 )
		{
			print "--> PAYMENT SHORTNAME AMT TYPE MATCH COUNT:",$numShortNameAmountMatches,"\n";
			
			for my $k (0 .. ($numShortNameAmountMatches-1))
			{
				my $index = $QBA[$m][QB_AM_SHORTNAME_AMOUNT_MATCH_INDEX][$k];
				my $am_date         = $AMA[$index][AM_DATE_INDEX];
				my $am_name         = $AMA[$index][AM_NAME_INDEX];
				my $am_notes        = $AMA[$index][AM_NOTES_INDEX];
				my $am_payment_type = ConvertPaymentTypeToString($AMA[$index][AM_PAYMENT_TYPE_INDEX]);
				my $am_amount       = $AMA[$index][AM_PAYMENT_AMOUNT_INDEX];
				
				if ( $AMA[$m][AM_QB_EXACT_MATCH_INDEX] < 0 )
				{
					print   $htmlFileHandle "<tr>\n";
					print   $htmlFileHandle "<td align=\"left\">",$am_notes,"</td>\n";
					print   $htmlFileHandle "<td align=\"center\">",$am_date,"</td>\n";
					print   $htmlFileHandle "<td align=\"left\">",$am_name,"</td>\n";
					print   $htmlFileHandle "<td align=\"center\">",$am_amount,"</td>\n";
					print   $htmlFileHandle "<td align=\"center\">",$am_payment_type,"</td>\n";
					print   $htmlFileHandle "<td align=\"center\">","Shortname Amount Type Match","</td>\n";
					print   $htmlFileHandle "</tr>\n";	
			
					print "\nShortname Amount Type Match\n";
					print "   --> Amount: [",$am_name,"] Payment Type: [",$am_payment_type,"]\n";
					print "   --> Quickbooks Entry:  [",$qb_date," ",$qb_name,"]\n";
					print "   --> AutoManager Entry: [",$am_date," ",$am_name,"]\n";			
				}
			}
		}
		
		my $numWeakMatches = scalar @{ $QBA[$m][QB_AM_WEAKNAME_MATCH_INDEX] };
		
		if ( $numWeakMatches > 0 )
		{
			print "--> PAYMENT WEAKNAME MATCH COUNT:",$numWeakMatches,"\n";	
			
			for my $k (0 .. ($numWeakMatches-1))
			{
				my $index = $QBA[$m][QB_AM_WEAKNAME_MATCH_INDEX][$k];
				my $am_date         = $AMA[$index][AM_DATE_INDEX];
				my $am_name         = $AMA[$index][AM_NAME_INDEX];
				my $am_notes        = $AMA[$index][AM_NOTES_INDEX];
				my $am_payment_type = ConvertPaymentTypeToString($AMA[$index][AM_PAYMENT_TYPE_INDEX]);
				my $am_amount       = $AMA[$index][AM_PAYMENT_AMOUNT_INDEX];
				
				if ( $AMA[$m][AM_QB_EXACT_MATCH_INDEX] < 0 )
				{
					print   $htmlFileHandle "<tr>\n";
					print   $htmlFileHandle "<td align=\"left\">",$am_notes,"</td>\n";
					print   $htmlFileHandle "<td align=\"center\">",$am_date,"</td>\n";
					print   $htmlFileHandle "<td align=\"left\">",$am_name,"</td>\n";
					print   $htmlFileHandle "<td align=\"center\">",$am_amount,"</td>\n";
					print   $htmlFileHandle "<td align=\"center\">",$am_payment_type,"</td>\n";
					print   $htmlFileHandle "<td align=\"center\">","Weak Match","</td>\n";
					print   $htmlFileHandle "</tr>\n";	
					
					print "\nWeak Match\n";
					print "   --> Amount: [",$am_amount,"]\n";
					print "   --> Quickbooks Memo: [",$qb_memo,"]\n";
					print "   --> Quickbooks Entry:  [",$qb_date," ",$qb_name," ", $qb_payment_type,"]\n";
					print "   --> AutoManager Entry: [",$am_date," ",$am_name," ", $am_payment_type,"]\n";
				}
			}
		}
		
				
		my $numNameTypeMatches = scalar @{ $QBA[$m][QB_AM_NAME_TYPE_MATCH_INDEX] };
		
		if ( $numNameTypeMatches > 0 )
		{
			print "--> PAYMENT NAME TYPE MATCH COUNT:",$numNameTypeMatches,"\n";	
			
			for my $k (0 .. ($numNameTypeMatches-1))
			{
				my $index = $QBA[$m][QB_AM_NAME_TYPE_MATCH_INDEX][$k];
				my $am_date         = $AMA[$index][AM_DATE_INDEX];
				my $am_name         = $AMA[$index][AM_NAME_INDEX];
				my $am_notes        = $AMA[$index][AM_NOTES_INDEX];
				my $am_payment_type = ConvertPaymentTypeToString($AMA[$index][AM_PAYMENT_TYPE_INDEX]);
				my $am_amount       = $AMA[$index][AM_PAYMENT_AMOUNT_INDEX];
				
				if ( $AMA[$m][AM_QB_EXACT_MATCH_INDEX] < 0 )
				{
					print   $htmlFileHandle "<tr>\n";
					print   $htmlFileHandle "<td align=\"left\">",$am_notes,"</td>\n";
					print   $htmlFileHandle "<td align=\"center\">",$am_date,"</td>\n";
					print   $htmlFileHandle "<td align=\"left\">",$am_name,"</td>\n";
					print   $htmlFileHandle "<td align=\"center\">",$am_amount,"</td>\n";
					print   $htmlFileHandle "<td align=\"center\">",$am_payment_type,"</td>\n";
					print   $htmlFileHandle "<td align=\"center\">","Name Payment Type Match","</td>\n";
					print   $htmlFileHandle "</tr>\n";			
					
					print "\n  Name Match\n";
					print "   --> Amount: [",$am_amount,"]\n";
					print "   --> Quickbooks Memo: [",$qb_memo,"]\n";
					print "   --> Quickbooks Entry:  [",$qb_date," ",$qb_name," ", $qb_payment_type,"]\n";
					print "   --> AutoManager Entry: [",$am_date," ",$am_name," ", $am_payment_type,"]\n";
				}
			}
			
			HtmlTableTailSection($htmlFileHandle);
			print "\n-------------------------------------------------------\n";			
		}		
	}
	else
	{
		$num_exact_matches++;
		my $qb_date   = $QBA[$m][QB_DATE_INDEX];
		my $qb_name   = $QBA[$m][QB_NAME_INDEX];
		my $qb_amount = $QBA[$m][QB_AMOUNT_INDEX];
		my $qb_memo   = $QBA[$m][QB_MEMO_INDEX];
		my $qb_payment_type = ConvertPaymentTypeToString($QBA[$m][QB_PAYMENT_TYPE_INDEX]);
	}
}

print  PR_HTML_OUTPUT_FILE "</body>\n";
print  PR_HTML_OUTPUT_FILE "</html>\n";

{
	print $htmlFileHandle "<br><br>\n";
	
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
	
	print $htmlFileHandle "<br><br>\n";
	
	HtmlTableTopSection($htmlFileHandle,$GREEN);	
	print $htmlFileHandle "<tr><th colspan=6>The following entries are EXACT matches in QuickBooks and AutoManager</th>\n";
	print $htmlFileHandle "<tr><th>Index</th><th>Name</th><th>Payment Type</th><th>Amount</th><th>Date (Quickbooks)</th><th>Date (AutoManager)</th>\n";
	
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
	
	##TBD - print out AMA table
	##for my $j (0 .. ($#AMA - 1))
	##$AMA[$j][AM_DATE_INDEX] = $date;
	##$AMA[$j][AM_NAME_INDEX] = uc($name);
	##$AMA[$j][AM_NOTES_INDEX] = uc($notes);
	##$AMA[$j][AM_PAYMENT_TYPE_INDEX] = LATE_FEE_INCOME_TYPE;	
	##$AMA[$j][AM_PAYMENT_AMOUNT_INDEX] = sprintf('%.2f',abs($total));
	##$AMA[$j][AM_QB_EXACT_MATCH_INDEX] = -1;
		
}

close(PR_HTML_OUTPUT_FILE);
## end of main script logic




sub ConvertPaymentTypeToString
{
	my $paymentType = $_[0];
	
	if ( $paymentType eq INTEREST_INCOME_TYPE)
	{
		return "INTEREST INCOME";
	}
	elsif ($paymentType eq PRINCIPAL_INCOME_TYPE)
	{
		return "TOTAL NOTES RECEIVABLE";
	}
	elsif ($paymentType eq LATE_FEE_INCOME_TYPE)
	{
		return "LATE_FEE";
	}
	elsif ($paymentType eq INTEREST_EXPENSE_INCOME_TYPE)
	{
		return "INTEREST_EXPENSE";
	}	
}

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
	print  $fileHandle "</table>\n";
}


