# SCCRetentionPolicyCapof1000Mailboxes
Purpose

This script is meant to deal with the situation where an Office 365 admin has thousands of users that need to be individually added to Security and Compliance Center policies.  At this time those policies are limited to 1000 mailboxes per policy.
This script divides users by department, and then alphabetically into different Retention Policies.  
If there are 10 departments then the script will create at least 260 retention policies.  If some departments employ more than 1000 people, whose email address starts with the same letter the script will create multiple policies.

Implementation Details

Since replication of retention policies can be slow the script will save out the list of users in each policy.  If that list does not need to change from one day to the next, then that policy and the associated file is not updated.

For each policy that is created or updated a CSV file is written to the folder from which the script runs.  The files are overwritten on each execution where the policy is created or changed.

It is assumed each division has a different policy.  The Divisions should be saved to a CSV file following this format (field headings and 2 example lines):
DivisionName, RetentionDuration, DateOption, Action
accounting,2557,CreationAgeInDays,KeepAndDelete
Development,2557,ModificationAgeInDays,Keep

This script utilizes the admin's existing connection to Security and Compliance Center PowerShell.  It makes no attempt to connect to Exchange PowerShell.   Prior to running the script, the list of mailboxes must be pulled into a CSV file.  That CSV file must contain the email address to be used, and the property that specifies which division the mailbox belongs to.
At time of writing this is assumed to be the fields PrimarySMTPAddress and CustomAttribute10.  This will need to be customized by each admin.
Recommend the list of mailboxes be pulled with invoke-command or the new V2 get-mailbox that runs a lot faster than a regular get-mailbox. 

