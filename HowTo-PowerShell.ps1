# This is something I made to help a Jr. Admin get started with Powershell.



﻿<# <--- This starts a comment block.
Single line comments start with # e.g. #THIS IS A COMMENT. Nothing else on this line will work past the #.
Every { needs a } at the end.
Every ( needs a ) at the end.
Neary all output can be piped into another command so Get-Date can be piped into Write-Host e.g. Get-Date | Write-Host
This ends the comment block --> #>

# How to use: Just follow the directions and don't run this as one big script.

# We'll need the addin to make stuff happen. So add this.
Add-PSSnapin Quest.ActiveRoles.ADManagement -WarningAction SilentlyContinue -ErrorAction SilentlyContinue

# Let's start by looking at ad groups.
Get-QADGroup

# HOLY CRAP THAT'S A LOT OF DATA! Let's make it more managable.
# We'll start by putting all of that into a variable.
# Variables start with $ e.g. $ThisIsAVariable. ThisIsNotAVariable

$Groups = Get-QADGroup

# Now that we have our groups in $Groups. Let's do something with it.
# How about we look at the managers of each group?
# To start with let's use the foreach command.

ForEach ($Group in $Groups) {$Group.ManagedBy}

# Now that's interesting, but what can we do with it? Well how about we format that a little better?
# Well start with a single group

$Group = Get-QADGroup -Identity "DL - Blood Drive"
$Group.ManagedBy

# Now we know who manages that group... let's get some better info by piping our data.
# Nearly everything in power shell can be piped.

$Group.ManagedBy | Get-QADUser

# How how about making a a little cleaner since we only need the name?
# That's right, you can pipe the pipes to each other. #YODAWG.

$Group.ManagedBy | Get-QADUser | Format-List -Property "Name"

# Now let's look at our list again this till with formating.

ForEach ($Group in $Groups) {$Group.ManagedBy | Get-QADUser | Format-List -Property "Name"}

# Wow, that was crap! What the hell went wrong?
# Well, we didn't get the group names only the person who managed it. Also we got A LOT of errors because some groups don't have managers.
# Let's see if we can clean this up a bit.

$Groups | foreach {
	$Group = $_  # $_ is a special variable that means what ever object in the pipe. We're giving that Object a name here.
	$Group | Format-Table -Property Name, ManagedBy
	}

# Good, but not good enough.

$Groups | Format-Table -Property Name, ManagedBy

#Better....

$Group | Format-Table -Property Name, @{Name='ManagedBy';expression={$_.ManagedBy | Get-QADUser}}

#Best

$Groups | Format-Table -Property Name, @{Name='ManagedBy';expression={
		IF ($_.ManagedBy -eq $null) {"No Manager"}
		ELSE {$User = $_.ManagedBy | Get-QADUser
			$User.Name}
		}
	}

# There you have it. A few ways of getting this done and making lists. For more information head to your friendly neighborhood search engine.
# If you want to know more about some commands check out the command listing on http://ss64.com/
