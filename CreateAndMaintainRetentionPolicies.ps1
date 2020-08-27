# This script divides users by department, and then alphabetically into different
# Retention Policies.  Each policy has a cap of 1000 mailboxes.
# As written the script assumes that only active mailboxes are part of the retention policies.
# It is assumed that inactive mailboxes can be removed from the policies, and that
# other procedures in the organization will see to their needs.
# 
# Since replication of retention policies can be slow the script will save out the 
# list of users in each policy.  If that list does not need to change from one day
# to the next then that policy and the associated file is not updeated.
# 
# For each policy that is created or updated a CSV file is written to the folder from
# which the script runs.  The files are overwritten on each execution where the policy 
# is created or changed.
# 
# It is assumed each division has a different policy.  The Divisions should be saved to a CSV
# file folliwng this format (field headings and example):
# DivisionName, RetentionDuration, DateOption, Action
# accounting,2557,CreationAgeInDays,KeepAndDelete
# Development,2557,ModificationAgeInDays,Keep
#
# This script utilizes the admin's existing connection to Security and Compliance Center PowerShell.
# It makes no attempt to connect to Exchange PowerShell.
# Prior to running the script the list of mailboxes must be pulled into a CSV file.  That CSV file must conatin
# the email address to be used, and the property that specifies which division the mailbox belongs to.
# At time of writing this is assumed to be the fields PrimarySMTPAddress and CustomAttribute10.
# This will need to be customized by each admin.
# Recommend the file be pulled with invoke-command or the new V2 get-mailbox that runs a lot faster. 


# ******************************************************************************
function update_policies {
param ([string]$PolicyName, [array]$UserList, [int]$RetentionDuration, [string]$DateOption, [string]$Action )

# Function loads the file listing people who were in the policy previously.
# if no file found assume this is initial creation.
# if file is found compare to current list. If different overwrite the current list on the policy.  If same proceed to next policy if there are mailboxes left.


  $MbxRemaining = $UserList.count
  $low_bound = 0
  $loop_count = 0

  while ($MbxRemaining -gt 0 ) {
    #   This loop creates one search per loop and decrements $MbxRemaining by the
    # number of mailboxes included in the search.

    $loop_count += 1
    $this_policy = $PolicyName + $loop_count
    write-host "Checking Policy: " $this_policy -foregroundcolor cyan

	
	
    if($MbxRemaining -gt $mailboxesPerObject) {
      #  There are at least $mailboxesPerObject mailboxes left. 
      # This IF branch fills $search_list with $mailboxesPerObject mailboxes.

      # Clear $search_list
      [array]$search_list = $NULL
      # Place mailboxes from the current value of low_bound to low_bound plus $mailboxesPerObject
      #minus one in $search_list
      $UserList[$low_bound..($mailboxesPerObject * $loop_count -1)] | foreach {
         [array]$search_list=$search_list + [string]$_.PrimarySMTPAddress
      }
     } else {
        # This branch fills $search_list with the mailboxes that are left when
        # there are less than $mailboxesPerObject remaining.

        # Clear $search_list
        [array]$search_list = $NULL
        #Place mailboxes from the current value of low_bound to the end of the array in $search_list
        $UserList[$low_bound..($UserList.count -1)] | foreach {
          [array]$search_list=$search_list + [string]$_.PrimarySMTPAddress
      }
    }
    $filename= $this_policy+".csv"
    [array]$OldList = import-csv $filename
	#trap error
	if($OldList.count -eq 0 ) { # create new policy, because we didn't have one yesterday
      # The two new-ret* lines that follow may need to be customized for individual environments.
	  new-retentioncompliancepolicy -name $this_policy -enabled $true -Exchangelocation $search_list
      $rule_name = $this_policy+"rule"
	  new-retentioncompliancerule -name $rule_name -policy $this_policy -retentionduration $RetentionDuration -ExpirationDateOption $DateOption -RetentionComplianceAction $Action
	  $search_list | export-csv $filename -notypeinformation
	  write-host "Created policy $this_policy."
	} else {
	
	  [array]$differences = compare-object -ReferenceObject $OldList -DifferenceObject $UserList

      if($differences.count -ne 0 ) {  #If there is a difference we need to change the list of users in the policy 

        set-retentioncompliancepolicy -name $this_policy -Exchangelocation $search_list
	    $search_list | export-csv $filename -notypeinformation
		write-host "Updated policy $this_policy."
	  } # Endif.   Do nothing if the count is zero.  Just move on to the next batch of mailboxes.
	  
    $low_bound = $mailboxesPerObject * $loop_count
    $MbxRemaining -= $mailboxesPerObject
  } # end while
} # end function


# ******************************************************************************

# MAIN Body
#
# Import list of divisions with their retention settings.
# Loop through them
#  In each loop break out each division by a letter of the alphabet, and create a policy for that letter.
#  If certain letters are uncommon and you wish to combine them the loop and criteria can be customized as needed.


[array]$DivisionsList = get-content "c:\o365\GroupDivisions.txt"  
$DivisionsCount = $DivisionsList.count
$DivisionsLoop = 0

#   Set the number of items that will be in each object.  Acceptable range is 1 to
# 1000.  3 is selected for illustration purposes in a lab.  
$mailboxesPerObject = 3                       

$base_object_name = "RetentionPolicy"

$mbxs = import-csv "mailbox_list.csv"
# Prior to running the script the list of mailboxes must be pulled into a CSV file.  That CSV file must conatin
# the email address to be used, and the property that specifies which division the mailbox belongs to.
# At time of writing this is assumed to be the fields PrimarySMTPAddress and CustomAttribute10.
# This will need to be customized by each admin.
 
while ($DivisionsLoop -lt $DivisionsList.Count ) {  # Loop through the Divisions 

  #Get list of users for current division
  [array]$CurrentDivision = $mbxs | ?{$_.CustomAttribute10 -eq $DivisionsList[$DivisionsLoop].DivisionName} 
  
  for($i=65 ; $i -lt 96; $i++) {  # Need to loop through the letters of the alphabet.
                                  # If there are other starting characters will need a separate loop for them.
    $CurrentLetter = [char]$i+"*" # set the wild card for each letter
    [array]$CurrentLetterList = $CurrentDivision | ?{$_.primarysmtpaddress -like $CurrentLetter}  # Find all items in the current division that start with the current letter.
    $PolicyName = $base_object_name+$DivisionsList[$DivisionsLoop]+"_"+[char]$i
	update_policies $PolicyName, $CurrentLetterList, $DivisionsList[$DivisionsLoop].RetentionDuration, $DivisionsList[$DivisionsLoop].DateOption, $DivisionsList[$DivisionsLoop].Action 

  } # Next $i 

   $DivisionsLoop++
} # End While ($DivisionsLoop -lt $DivisionsList.Count )
