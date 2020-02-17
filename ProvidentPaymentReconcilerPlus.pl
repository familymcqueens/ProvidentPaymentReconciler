##
## Script to read AutoManager and QuickBooks Bank Register Files
##

use Time::Piece;

my @AMA;  # Auto Manager Array
my @QBA;  # QuickBooks Array

my $amIndex = 0;
my $qbIndex = 0;
my $qbFirstDate;
my $numArgs = $#ARGV + 1;

## Toogle debug for console verbosity
my $debug = 0;


## Make sure we have the right command line arguments
if ( $numArgs != 2 )
{
	print "PaymentReconciler.pl <AutoManager.csv file> <Quickbooks.csv file>\n";
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
	AM_EXACT_MATCH_INDEX => 5,
};

while (<AM_INPUT_FILE>) {
 chomp;
 my ($date,$total,$expense,$name,$payment_type,$notes,$invoice,$check,$part,$balance,$interest,$principal,$lot,$vehicle,$status,$receipt) = split(",");
 
	# Remove white space before and after
	$date =~ s/^\s+|\s+$//g;

	if ($payment_type eq "PAYMENT")
	{
		$AMA[$amIndex][AM_EXACT_MATCH_INDEX] = -1;
		$AMA[$amIndex][AM_DATE_INDEX] = $date;
		$AMA[$amIndex][AM_NAME_INDEX] = uc($name);
		$AMA[$amIndex][AM_NOTES_INDEX] = uc($notes);		
		$AMA[$amIndex][AM_PAYMENT_TYPE_INDEX] = INTEREST_INCOME_TYPE;
		$AMA[$amIndex][AM_PAYMENT_AMOUNT_INDEX] = sprintf('%.2f',abs($interest));
		
		if ($debug)
		{
			print "AM[",$amIndex,"]"," Date:",$AMA[$amIndex][AM_DATE_INDEX]," Name:",$AMA[$amIndex][AM_NAME_INDEX]," Amt:",$AMA[$amIndex][AM_PAYMENT_AMOUNT_INDEX]," Type:",$AMA[$amIndex][AM_PAYMENT_TYPE_INDEX],"\n";
		}
		
		$amIndex++;
		
		$AMA[$amIndex][AM_EXACT_MATCH_INDEX] = -1;
		$AMA[$amIndex][AM_DATE_INDEX] = $date;
		$AMA[$amIndex][AM_NAME_INDEX] = uc($name);
		$AMA[$amIndex][AM_NOTES_INDEX] = uc($notes);		
		$AMA[$amIndex][AM_PAYMENT_TYPE_INDEX] = PRINCIPAL_INCOME_TYPE;
		$AMA[$amIndex][AM_PAYMENT_AMOUNT_INDEX] = sprintf('%.2f',abs($principal));	
		
		if ($debug)
		{
			print "AM[",$amIndex,"]"," Date:",$AMA[$amIndex][AM_DATE_INDEX]," Name:",$AMA[$amIndex][AM_NAME_INDEX]," Amt:",$AMA[$amIndex][AM_PAYMENT_AMOUNT_INDEX]," Type:",$AMA[$amIndex][AM_PAYMENT_TYPE_INDEX],"\n";
		}
	}
	elsif ($payment_type eq "LATE FEE")
	{
		$AMA[$amIndex][AM_EXACT_MATCH_INDEX] = -1;
		$AMA[$amIndex][AM_DATE_INDEX] = $date;
		$AMA[$amIndex][AM_NAME_INDEX] = uc($name);
		$AMA[$amIndex][AM_NOTES_INDEX] = uc($notes);		
		$AMA[$amIndex][AM_PAYMENT_TYPE_INDEX] = LATE_FEE_INCOME_TYPE;
		$AMA[$amIndex][AM_PAYMENT_AMOUNT_INDEX] = sprintf('%.2f',abs($total));
		
		if ($debug)
		{
			print "AM[",$amIndex,"]"," Date:",$AMA[$amIndex][AM_DATE_INDEX]," Name:",$AMA[$amIndex][AM_NAME_INDEX]," Amt:",$AMA[$amIndex][AM_PAYMENT_AMOUNT_INDEX]," Type:",$AMA[$amIndex][AM_PAYMENT_TYPE_INDEX],"\n";
		}
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
};

use constant {
	NAME_AMOUNT_MATCH => 1,
	SHORTNAME_AMOUNT_MATCH => 2,
	WEAKNAME_MATCH => 3,
	NAME_TYPE_MATCH => 4,
	NAME_DATE_MATCH => 5,
	QB_MEMO_ERROR => 6,
	PAYMENT_TYPE_MISMATCH => 7,
};



## Start reading of Quickbooks file

my $deposit_date;

 while (<QB_INPUT_FILE>) {
 chomp;
 ($type,$num,$date,$name,$memo,$payment_type,$amount) = split(",");    
	
	# Remove quotes from date
	$date =~ s/"/ /i;
	$date =~ s/"/ /i;
	
	# Remove white space before and after
	$date =~ s/^\s+|\s+$//g;
	
	## Since the date is listed ONCE per deposit, check to see if the date value 
	## is of non-zero length and store off in $deposit_date variable
    if (length($date) && ($date ne "TOTAL"))
	{
		$deposit_date = $date;
	}
	
	if (($name eq "") || ($amount =~ /^[a-zA-Z]+$/))
	{
		next;
	}
	
	($type,$num,$date,$lastname, $firstname,$memo,$payment_type,$amount) = split(",");
	
	$count = ($lastname =~ tr/"//);
	
	if ($count eq 2)
	{
		$memo_old = $memo;
		$memo = $firstname;
		
		$payment_type_old = $payment_type;
		$payment_type = $memo_old;
		
		$amount_old = $amount;
		$amount = $payment_type_old;		
		
		$firstname = "";
		$lastname =~ s/"//;	
	}
	
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
	elsif (uc($payment_type) eq "TOTAL NOTES RECEIVABLE" )
	{
		$income_type = PRINCIPAL_INCOME_TYPE;
	}

	## Check for late fee portion of payment
	elsif (uc($payment_type) eq "LATE FEES" )
	{
		$income_type = LATE_FEE_INCOME_TYPE;
	}

	## Check to see if payment was incorrectly entered unter Interest Expense account
	elsif (uc($payment_type) eq "INTEREST EXPENSE" )
	{
		$income_type = INTEREST_EXPENSE_INCOME_TYPE;
	}
	
	## This will catch the last totals line in the Qbook input file and not record it in array
	else
	{
		next;
	}
	
	if ($qbIndex == 0)
	{
		$qbFirstDate = $deposit_date;
		
		if ($debug)
		{
			print "QB report first date found: [", $qbFirstDate,"]\n";		
		}
	}
		
	$QBA[$qbIndex][QB_DATE_INDEX] = $deposit_date;
	$QBA[$qbIndex][QB_NAME_INDEX] = uc($name);
	$QBA[$qbIndex][QB_AMOUNT_INDEX] = $amountabs;
	$QBA[$qbIndex][QB_MEMO_INDEX] = uc($memo);
	$QBA[$qbIndex][QB_PAYMENT_TYPE_INDEX] = $income_type;
	$QBA[$qbIndex][AM_EXACT_MATCH_INDEX] = -1;
	
	if ($memo eq "")
	{
		print "QB MEMO ERROR: [",$qbIndex,"]"," Date:",$QBA[$qbIndex][QB_DATE_INDEX]," Name:",$QBA[$qbIndex][QB_NAME_INDEX]," Amt:",$QBA[$qbIndex][QB_AMOUNT_INDEX]," Memo:",$QBA[$qbIndex][QB_MEMO_INDEX]," Type:",$QBA[$qbIndex][QB_PAYMENT_TYPE_INDEX],"\n";
		exit -1;
	}
	
	if ($debug)
	{
		print "QB[",$qbIndex,"]"," Date:",$QBA[$qbIndex][QB_DATE_INDEX]," Name:",$QBA[$qbIndex][QB_NAME_INDEX]," Amt:",$QBA[$qbIndex][QB_AMOUNT_INDEX]," Memo:",$QBA[$qbIndex][QB_MEMO_INDEX]," Type:",$QBA[$qbIndex][QB_PAYMENT_TYPE_INDEX],"\n";
	}
	
	$qbIndex++;	
 }
  
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

print "\nProcessing Quickbooks deposits...\n\n";
print "Size of QBooks Array: ",scalar(@QBA),"\n";
print "Size of AutoManager Array: ",scalar(@AMA),"\n";

#
# Walk over Quickbook entries
#
for my $i (0 .. scalar(@QBA)-1) 
{
	my $qb_date         = $QBA[$i][QB_DATE_INDEX];
	my $qb_name         = $QBA[$i][QB_NAME_INDEX];
	my $qb_amount       = $QBA[$i][QB_AMOUNT_INDEX];
	my $qb_memo         = $QBA[$i][QB_MEMO_INDEX];
	my $qb_payment_type = ConvertPaymentTypeToString($QBA[$i][QB_PAYMENT_TYPE_INDEX]);
	
	if (!$debug)
	{
		print {STDERR} ".";
	}	
	
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

	$found_exact_match = 0;
	$found_partial_match = 0;
	
	# Look for matching Auto Manager entries
	for my $j (0 .. scalar(@AMA)-1) 
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
	    if (($am_name eq $qb_name) && ($am_payment_type eq  $qb_payment_type) && ($am_amount eq $qb_amount) && 
			(($qb_memo eq "DEPOSIT") || ($qb_memo eq $am_notes) || ($qb_memo eq "SIG") || ($qb_memo eq "PNM") || ($qb_memo eq "PRO FIN")) )
		{
			$QBA[$i][AM_EXACT_MATCH_INDEX] = $j;
			$AMA[$j][AM_EXACT_MATCH_INDEX] = $i;
			
			if ($debug)
			{
				print "QBIndex:",$i," AmIndex:",$j, " EXACT MATCH: NAME:",$qb_name," TYPE:",$qb_payment_type," AMOUNT:",$qb_amount," MEMO:",$qb_memo,"\n";			
			}
			$exact_match_log = sprintf("%d:%d:%s:%s:%s:%s:%s:%s",$i,$j,$qb_date,$am_date,$qb_name,$qb_payment_type,$qb_amount,$am_notes);
			push (@exact_match_log, $exact_match_log);
			
			$found_exact_match = 1;
			last;
		}
		else
		{
			$found_exact_match = 0;
		}
		
		if (($am_name eq $qb_name) && ($am_payment_type ne $qb_payment_type) && ($am_amount eq $qb_amount))
		{
			if ($debug)
			{	
				print "QBIndex:",$i," PAYMENT_TYPE_MISMATCH: NAME:",$am_name," TYPE:",$am_payment_type," AMOUNT:",$am_amount,"\n";
			}
			
			$partial_match_log = sprintf("%d:%s:%s:%s:%s:%s:%s",
			PAYMENT_TYPE_MISMATCH,$i,$am_date, $am_name, $am_notes, $am_payment_type, $am_amount);
			push (@partial_match_log, $partial_match_log);
			
			$found_partial_match=1;
		}		
		## Look for match with only NAME and AMOUNT
	    elsif (($am_name eq $qb_name) && ($am_amount eq $qb_amount))
		{
			if ($debug)
			{	
				print "QBIndex:",$i," NAME_AMOUNT_MATCH:",$am_name," TYPE:",$am_payment_type," AMOUNT:",$am_amount,"\n";
			}
			
			$partial_match_log = sprintf("%d:%s:%s:%s:%s:%s:%s",
			NAME_AMOUNT_MATCH,$i,$am_date, $am_name, $am_notes, $am_payment_type, $am_amount);
			push (@partial_match_log, $partial_match_log);
			
			$found_partial_match=1;
		}
		## Look for match with only LAST NAME and AMOUNT
		elsif (($am_last_name eq $qb_last_name) && ($am_amount eq $qb_amount))
		{
			if ($debug)
			{	
				print "QBIndex:",$i," LASTNAME_WEAKNAME_MATCH: ",$am_name," TYPE:",$am_payment_type," AMOUNT:",$am_amount,"\n";
			}
			
			$partial_match_log = sprintf("%d:%s:%s:%s:%s:%s:%s",
			WEAKNAME_MATCH,$i,$am_date, $am_name, $am_notes, $am_payment_type, $am_amount);
			push (@partial_match_log, $partial_match_log);
			
			$found_partial_match=1;
			
		} 
		## Look for match with only FIRST NAME and AMOUNT
		elsif (($am_first_name eq $qb_first_name) && ($am_amount eq $qb_amount))
		{
			if ($debug)
			{	
				print "QBIndex:",$i," FIRSTNAME_WEAKNAME_MATCH: ",$am_name," TYPE:",$am_payment_type," AMOUNT:",$am_amount,"\n";
			}
			
			$partial_match_log = sprintf("%d:%s:%s:%s:%s:%s:%s",
			WEAKNAME_MATCH,$i,$am_date, $am_name, $am_notes, $am_payment_type, $am_amount);
			push (@partial_match_log, $partial_match_log);
			
			$found_partial_match=1;
		}

		## Look for match with only FIRST NAME and AMOUNT
		elsif (($am_first_name eq $qb_first_name) && ($am_last_name eq $qb_last_name) && ($am_date eq $qb_date))
		{
			if ($debug)
			{	
				print "QBIndex:",$i," NAME_DATE_MATCH: ",$am_first_name," ",$am_last_name," TYPE:",$am_payment_type," AMOUNT:",$am_amount,"\n";
			}
			
			$partial_match_log = sprintf("%d:%s:%s:%s:%s:%s:%s",
			NAME_DATE_MATCH,$i,$am_date, $am_name, $am_notes, $am_payment_type, $am_amount);
			push (@partial_match_log, $partial_match_log);
			
			$found_partial_match=1;
		}	
		
		## Look for match with "short" last name and amount only
		elsif (($am_last_shortname eq $qb_last_shortname) && ($am_amount eq $qb_amount))
		{
			if ($debug)
			{	
				print "QBIndex:",$i," SHORTNAME_AMOUNT_MATCH: ",$am_name," TYPE:",$am_payment_type," AMOUNT:",$am_amount,"\n";
			}
			
			## Look for further match of "short" first name and payment type
			if (($am_first_shortname eq $qb_first_shortname) && ($am_payment_type) eq ($qb_payment_type))
			{
				$partial_match_log = sprintf("%d:%s:%s:%s:%s:%s:%s",
				SHORTNAME_AMOUNT_MATCH,$i,$am_date, $am_name, $am_notes, $am_payment_type, $am_amount);
				push (@partial_match_log, $partial_match_log);
				
			}
			elsif ( ($qb_memo eq "DEPOSIT") && (($am_payment_type) eq ($qb_payment_type)) )
			{
				$partial_match_log = sprintf("%d:%s:%s:%s:%s:%s:%s",
				SHORTNAME_AMOUNT_MATCH,$i,$am_date, $am_name, $am_notes, $am_payment_type, $am_amount);
				push (@partial_match_log, $partial_match_log);
			}
			
			$found_partial_match=1;
		}
		
		## Look for name and payment type match
		elsif (($am_name eq $qb_name) && ($am_payment_type eq $qb_payment_type))
		{
			if ($debug)
			{	
				print "QBIndex:",$i," NAME_TYPE_MATCH: NAME:",$am_name," TYPE:",$am_payment_type," AMOUNT:",$am_amount,"\n";				
				$partial_match_log = sprintf("%d:%s:%s:%s:%s:%s:%s",
				NAME_TYPE_MATCH,$i,$am_date, $am_name, $am_notes, $am_payment_type, $am_amount);
				push (@partial_match_log, $partial_match_log);
			}
			
			$found_partial_match=1;
		}	
    }
	
	# Unable to find an exact match or partial match in AM for this QB entry
	if (!$found_partial_match && !$found_exact_match)
	{
		if ($debug)
		{	
			print "QBIndex:",$i," NO AM MATCH: QB_NAME:",$qb_name," QB_TYPE:",$qb_payment_type," QB_AMOUNT:",$qb_amount,"\n";
		}	
		
		$qb_no_match_log = sprintf("%s:%s:%s:%s:%s:%s",$qb_memo,$qb_date,$qb_name,$qb_payment_type,$qb_amount);
		push (@qb_no_match_log, $qb_no_match_log);
		
	}
	
	if (($qb_memo ne "DEPOSIT") || ($qb_memo eq ""))
	{
		if ($debug)
		{	
			print "QBIndex:",$i," QB MEMO UNKNOWN: QB_NAME:",$qb_name," QB_TYPE:",$qb_payment_type," QB_AMOUNT:",$qb_amount,"\n";
		}		
		
		$qb_no_match_log = sprintf("%s:%s:%s:%s:%s:%s",$qb_memo,$qb_date,$qb_name,$qb_payment_type,$qb_amount);
		push (@qb_no_match_log, $qb_no_match_log);
	}
} 




##
## Open Payment Reconciler HTML output file 
##
if (open(PR_HTML_OUTPUT_FILE,'>ProvidentPaymentReconciler.html') == 0) {
   print "Error opening: ProvidentPaymentReconciler.html";
   exit -1;  
}

my $htmlFileHandle = \*PR_HTML_OUTPUT_FILE;

print $htmlFileHandle "<html>\n";
print $htmlFileHandle "<head><meta http-equiv=\"refresh\" content=\"500\"><title>Provident Financial Payment Reconciler Utility </title></head>\n";
print $htmlFileHandle "<body>\n";
print $htmlFileHandle "<h1><i>Provident Financial Payment Reconciler Utility</i></h1>\n";
PrintCssStyle($htmlFileHandle);


for my $m (0 .. scalar(@QBA)-1)  
{
	my $qb_date   = $QBA[$m][QB_DATE_INDEX];
	my $qb_name   = $QBA[$m][QB_NAME_INDEX];
	my $qb_amount = $QBA[$m][QB_AMOUNT_INDEX];
	my $qb_memo   = $QBA[$m][QB_MEMO_INDEX];
	my $qb_payment_type = ConvertPaymentTypeToString($QBA[$m][QB_PAYMENT_TYPE_INDEX]);	
	
	# Skip over non-deposit QB entries
	if (($qb_memo ne "DEPOSIT") || ($qb_memo eq ""))
	{
		next;
	}
	
	# Look for QB entries with no matches
	if ( ($QBA[$m][AM_EXACT_MATCH_INDEX] < 0) )
	{
		if ($debug)
		{
			print "\n-> No AutoManager exact match for the following QBook entry <-\n";
			print "QB Idx: [",$m,"]\n";
			print "QB Date:[",$qb_date,"]\n"; 
			print "QB Name [",$qb_name,"]\n";
			print "QB Amt: [",$qb_amount,"]\n";
			print "QB Memo:[",$qb_memo,"]\n";
			print "QB Type:[",$qb_payment_type,"]\n";
		}
		
		print   $htmlFileHandle "<div class=\"datagrid\"><table class=\"table1\">\n";
        print   $htmlFileHandle "<br><thead><tr><th>Quickbooks Description</th><th>Date</th><th>Name</th><th>Payment</th><th colspan=2>Payment Type</th></tr></thead>\n";
		print   $htmlFileHandle "<tbody>\n";
		print   $htmlFileHandle "<tr>\n";
		print   $htmlFileHandle "<td align=\"left\">",$qb_memo,"</td>\n";
		print   $htmlFileHandle "<td align=\"left\">",$qb_date,"</td>\n";
		print   $htmlFileHandle "<td align=\"left\">",$qb_name,"</td>\n";
		print   $htmlFileHandle "<td align=\"left\">",$qb_amount,"</td>\n";
		print   $htmlFileHandle "<td align=\"left\">",$qb_payment_type,"</td>\n";
		print   $htmlFileHandle "</tr>\n";
		print   $htmlFileHandle "</tbody>\n";
        print   $htmlFileHandle "</table></div>\n";
		
		print   $htmlFileHandle "<div class=\"datagrid\"><table class=\"table2\">\n";
		print   $htmlFileHandle "<thead><tr><th>AutoManager Matches</th><th>Date</th><th>Name</th><th>Payment</th><th>Payment Type</th><th>Match Type</th></tr></thead>\n";
		print   $htmlFileHandle "<tbody>\n";

		# Look for partial matches
		for my $entry (@partial_match_log) 
		{
			my @values = split(':', $entry);
			
			my $tmp_match   = $values[0];
			my $tmp_qbIndex = $values[1];
			my $tmp_date    = $values[2];
			my $tmp_name    = $values[3];
			my $tmp_notes   = $values[4];
			my $tmp_type    = $values[5];
			my $tmp_amount  = $values[6];
			
			if ( $m eq $tmp_qbIndex )
			{
				if ($debug) 
				{
					print "\n\tMatch: ",ConvertPartialMatchToString($tmp_match),"\n";
					print "\tQB Idx: ",$tmp_qbIndex,"\n";
					print "\tAM Date: ",$tmp_date,"\n"; 
					print "\tAM Name: ",$tmp_name,"\n";
					print "\tAM Amt: ",$tmp_amount,"\n";
					print "\tAM Memo: ",$tmp_notes,"\n";
					print "\tAM Type: ",$tmp_type,"\n";
				}
				
				print $htmlFileHandle "<tr>\n";
				print $htmlFileHandle "<td align=\"left\">",$tmp_notes,"</td>\n";
				print $htmlFileHandle "<td align=\"left\">",$tmp_date,"</td>\n";
				print $htmlFileHandle "<td align=\"left\">",$tmp_name,"</td>\n";
				print $htmlFileHandle "<td align=\"left\">",$tmp_amount,"</td>\n";
				print $htmlFileHandle "<td align=\"left\">",$tmp_type,"</td>\n";
				print $htmlFileHandle "<td align=\"left\">",ConvertPartialMatchToString($tmp_match),"</td>\n";
				print $htmlFileHandle "</tr>\n";
			}				
		}

		print $htmlFileHandle "</tbody>\n";
		print $htmlFileHandle "</table></div>\n";					
	}	
}

#
# Print out the Automanager entries missing a QBooks match
#
print $htmlFileHandle "<div class=\"datagrid\"><table class=\"table1\">\n";
print $htmlFileHandle "<br><thead><tr><th colspan=5>AutoManager entries missing a QuickBooks match</th>\n";
print $htmlFileHandle "<tr><th>Notes</th><th>Date</th><th>Name</th><th>Payment Type</th><th>Amount</th>\n";
print $htmlFileHandle "<tbody>\n";

for my $m (0 .. scalar(@AMA)-1)  
{
	my $am_date         = $AMA[$m][AM_DATE_INDEX];
	my $am_name         = $AMA[$m][AM_NAME_INDEX];
	my $am_notes        = $AMA[$m][AM_NOTES_INDEX];
	my $am_payment_type = ConvertPaymentTypeToString($AMA[$m][AM_PAYMENT_TYPE_INDEX]);
	my $am_amount       = $AMA[$m][AM_PAYMENT_AMOUNT_INDEX];
	my $am_qbmatch      = $AMA[$m][AM_EXACT_MATCH_INDEX];
	
	# Put the QB first payment date and AM current payment into the right format
	my $qbFirstDate = Time::Piece->strptime($qbFirstDate,'%m/%d/%YY');
	my $amCurrDate  = Time::Piece->strptime($am_date,'%m/%d/%y');
	
	# Compare these dates and only monitor the AM entries that are at least 
	# greater than or equal to the first QB payment date in report.
	if ($amCurrDate < $qbFirstDate )
	{
		if ($debug)
		{
			print "Skipping AutoManager payment with Date:[", $am_date,"] Name:[",$am_name,"[ Amount:[",$am_amount,"]\n";
		}
		next;
	}	
	
	if ( ($am_payment_type eq "TOTAL NOTES RECEIVABLE") || ($am_payment_type eq "INTEREST INCOME") )
	{
		if ( $am_amount ne "0.00" && $am_qbmatch eq -1)
		{
			if ($debug)
			{
				print "\n-> No QBooks Match for the following AutoManager entry <-\n";
				print " Name: ",$am_name,"\n";
				print " Date: ",$am_date,"\n";
				print " Amount: ",$am_amount,"\n";
				print " Payment Type: ",$am_payment_type,"\n";
			}	
				
			print $htmlFileHandle "<tr>\n";
			print $htmlFileHandle "<td align=\"left\">",$am_notes,"</td>\n";
			print $htmlFileHandle "<td align=\"left\">",$am_date,"</td>\n";
			print $htmlFileHandle "<td align=\"left\">",$am_name,"</td>\n";
			print $htmlFileHandle "<td align=\"left\">",$am_payment_type,"</td>\n";
			print $htmlFileHandle "<td align=\"left\">",$am_amount,"</td>\n";
			print $htmlFileHandle "</tr>\n";			
		}
	}
}

print $htmlFileHandle "</tbody>\n";
print $htmlFileHandle "</table></div>\n";

print $htmlFileHandle "</body>\n";
print $htmlFileHandle "</html>\n";



# Print out the missing QB entries from AutoManager
print $htmlFileHandle "<div class=\"datagrid\"><table class=\"table2\">\n";
print $htmlFileHandle "<br><thead><tr><th colspan=5>Quickbook entries with no matches from AutoManager</th>\n";
print $htmlFileHandle "<tr><th>Memo</th><th>Date</th><th>Name</th><th>Payment Type</th><th>Amount</th>\n";
print $htmlFileHandle "<tbody>\n";

my $iteration = 0;

for my $entry (@qb_no_match_log) 
{
	my @values = split(':', $entry);
	
	if ($iteration++ % 2)
	{
		print $htmlFileHandle "<tr class=\"alt\">\n";
	}
	else
	{
		print $htmlFileHandle "<tr>\n";
	}
	
	foreach my $val (@values) 
	{
		print $htmlFileHandle "<td align=\"left\">",$val,"</td>\n";		
	}		
	print $htmlFileHandle "</tr>\n";		
}	

print $htmlFileHandle "</tbody>\n";
print $htmlFileHandle "</table></div>\n";
print $htmlFileHandle "<br>\n";




# Print out the exact match log
print $htmlFileHandle "<div class=\"datagrid\"><table class=\"table3\">\n";
print $htmlFileHandle "<br><thead><tr><th colspan=8>The following entries are EXACT matches in QuickBooks and AutoManager</th>\n";
print $htmlFileHandle "<tr><th>Index(QB)</th><th>Index(AM)</th><th>Date(QB)</th><th>Date(AM)</th><th>Name</th><th>Payment Type</th><th>Amount</th><th>Note(AM)</th>\n";
print $htmlFileHandle "<tbody>\n";

my $iteration = 0;

for my $entry (@exact_match_log) 
{
	my @values = split(':', $entry);
	
	if ($iteration++ % 2)
	{
		print $htmlFileHandle "<tr class=\"alt\">\n";
	}
	else
	{
		print $htmlFileHandle "<tr>\n";
	}
	
	foreach my $val (@values) 
	{
		print $htmlFileHandle "<td align=\"left\">",$val,"</td>\n";		
	}		
	print $htmlFileHandle "</tr>\n";		
}

print   $htmlFileHandle "</tbody>\n";
print   $htmlFileHandle "</table></div>\n";
print   $htmlFileHandle "<br><br><br>\n";		



close(PR_HTML_OUTPUT_FILE);
## end of main script logic






sub ConvertPartialMatchToString
{
	my $partialMatch = $_[0];
	
	if ( $partialMatch eq NAME_AMOUNT_MATCH)
	{
		return "NAME+AMOUNT MATCH";
	}
	elsif ($partialMatch eq SHORTNAME_AMOUNT_MATCH)
	{
		return "SHORTNAME+AMOUNT MATCH";
	}
	elsif ($partialMatch eq WEAKNAME_MATCH)
	{
		return "WEAKNAME_MATCH";
	}
	elsif ($partialMatch eq NAME_TYPE_MATCH)
	{
		return "NAME+TYPE_MATCH";
	}
	elsif ($partialMatch eq NAME_DATE_MATCH)
	{
		return "NAME+DATE_MATCH";
	}
	elsif ($partialMatch eq PAYMENT_TYPE_MISMATCH)
	{
		return "PAYMENT+TYPE_MISMATCH";
	}
}



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



sub PrintCssStyle
{
	my $fileHandle = $_[0];		
	print  $fileHandle "<style>\n";
	print  $fileHandle "body {background-color: lightblue;}\n";
	print  $fileHandle "h1 {color: navy;margin-left: 20px;}\n";
		
	print  $fileHandle ".datagrid table.table1 { border-collapse: collapse; text-align: left; width: 100%; }\n"; 
	print  $fileHandle ".datagrid table.table1 {font: normal 12px/150% Arial, Helvetica, sans-serif; background: #fff; overflow: hidden; border: 4px solid #991821; -webkit-border-radius: 3px; -moz-border-radius: 3px; border-radius: 5px; }\n";
	print  $fileHandle ".datagrid table.table1 td,\n"; 
	print  $fileHandle ".datagrid table.table1 th { padding: 3px 10px; }\n";
	print  $fileHandle ".datagrid table.table1 thead th {background:-webkit-gradient( linear, left top, left bottom, color-stop(0.05, #991821), color-stop(1, #80141C) );background:-moz-linear-gradient( center top, #991821 5%, #80141C 100% );filter:progid:DXImageTransform.Microsoft.gradient(startColorstr='#991821', endColorstr='#80141C');background-color:#991821; color:#FFFFFF; font-size: 15px; font-weight: bold; border-left: 1px solid #B01C26; }\n"; 
	print  $fileHandle ".datagrid table.table1 thead th:first-child { border: none; }\n";
	print  $fileHandle ".datagrid table.table1 tbody td { color: #80141C; border-left: 1px solid #F7CDCD;font-size: 12px;font-weight: normal; }\n";
	print  $fileHandle ".datagrid table.table1 tbody .alt td { background: #F7CDCD; color: #80141C; }\n";
	print  $fileHandle ".datagrid table.table1 tbody td:first-child { border-left: none; }\n";
	print  $fileHandle ".datagrid table.table1 tbody tr:last-child td { border-bottom: none; }\n";

	print  $fileHandle ".datagrid table.table2 { border-collapse: collapse; text-align: left; width: 100%; }\n"; 
	print  $fileHandle ".datagrid table.table2 {font: normal 12px/150% Arial, Helvetica, sans-serif; background: #fff; overflow: hidden; border: 4px solid #006699; -webkit-border-radius: 3px; -moz-border-radius: 3px; border-radius: 5px;}\n";
	print  $fileHandle ".datagrid table.table2 td,\n"; 
	print  $fileHandle ".datagrid table.table2 th { padding: 3px 10px; }\n";
	print  $fileHandle ".datagrid table.table2 thead th {background:-webkit-gradient( linear, left top, left bottom, color-stop(0.05, #006699), color-stop(1, #00557F) );background:-moz-linear-gradient( center top, #006699 5%, #00557F 100% );filter:progid:DXImageTransform.Microsoft.gradient(startColorstr='#006699', endColorstr='#00557F');background-color:#006699; color:#FFFFFF; font-size: 15px; font-weight: bold; border-left: 1px solid #0070A8; }\n"; 
	print  $fileHandle ".datagrid table.table2 thead th:first-child { border: none; }\n";
	print  $fileHandle ".datagrid table.table2 tbody td { color: #80141C; border-left: 1px solid #CCE5FF;font-size: 12px;font-weight: normal; }\n";
	print  $fileHandle ".datagrid table.table2 tbody .alt td { background: #CCE5FF; color: #80141C; }\n";
	print  $fileHandle ".datagrid table.table2 tbody td:first-child { border-left: none; }\n";
	print  $fileHandle ".datagrid table.table2 tbody tr:last-child td { border-bottom: none; }\n";

	print  $fileHandle ".datagrid table.table3 { border-collapse: collapse; text-align: left; width: 100%; }\n"; 
	print  $fileHandle ".datagrid table.table3 {font: normal 12px/150% Arial, Helvetica, sans-serif; background: #fff; overflow: hidden; border: 4px solid #36752D; -webkit-border-radius: 3px; -moz-border-radius: 3px; border-radius: 5px; }\n";
	print  $fileHandle ".datagrid table.table3 td,\n"; 
	print  $fileHandle ".datagrid table.table3 th { padding: 3px 10px; }\n";
	print  $fileHandle ".datagrid table.table3 thead th {background:-webkit-gradient( linear, left top, left bottom, color-stop(0.05, #36752D), color-stop(1, #275420) );background:-moz-linear-gradient( center top, #36752D 5%, #275420 100% );filter:progid:DXImageTransform.Microsoft.gradient(startColorstr='#36752D', endColorstr='#275420');background-color:#36752D; color:#FFFFFF; font-size: 15px; font-weight: bold; border-left: 1px solid #36752D; }\n"; 
	print  $fileHandle ".datagrid table.table3 thead th:first-child { border: none; }\n";
	print  $fileHandle ".datagrid table.table3 tbody td { color: #275420; border-left: 1px solid #C6FFC2;font-size: 12px;font-weight: normal; }\n";
	print  $fileHandle ".datagrid table.table3 tbody .alt td { background: #DFFFDE; color: #275420; }\n";
	print  $fileHandle ".datagrid table.table3 tbody td:first-child { border-left: none; }\n";
	print  $fileHandle ".datagrid table.table3 tbody tr:last-child td { border-bottom: none; }\n";
		
	print  $fileHandle "</style>\n";
}


