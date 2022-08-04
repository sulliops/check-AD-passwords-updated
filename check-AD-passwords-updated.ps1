# Dependencies
Import-Module ActiveDirectory
Add-Type -AssemblyName System.Windows.Forms

# Function for prompting for file input (account names)
$fileBrowser = New-Object System.Windows.Forms.OpenFileDialog -Property @{

    Multiselect = $false
    Filter = 'Text file (*.txt)|*.txt'

}

# Call file browser window to select accounts file
Write-Host "Using system file browser to prompt for file with usernames..."
$fileBrowserValid = $fileBrowser.ShowDialog()

# Check to make sure user has not canceled the system file browser
if ($fileBrowserValid -eq [System.Windows.Forms.DialogResult]::OK) { # If a file was selected, continue

    # Store file browser selection as accountsFile
    $accountsFile = $fileBrowser.FileName
    Write-Host "Selected input file: $accountsFile"

    # Write new line to terminal (for formatting)
    Write-Host "`n"

    # Get parent directory of input file
    $workingPath = Split-Path -Parent $accountsFile

    # Prompt for starting time (date when temporary password was set)
    $startingDate = Read-Host -Prompt 'Input date when temporary passwords were set (format MM/DD/YYYY)'

    # Write new line to terminal (for formatting)
    Write-Host "`n"
    
    # Convert starting time to epoch time for comparison
    $startingEpochTime = Get-Date -Date $startingDate -UFormat %s
    
    # Offset the starting epoch time by 6 hours to allow for delays in ECAD account creation
    $offsetStartingTime = $startingEpochTime + 21600

    # Get the contents of accountsFile as usernames by line
    $names = Get-Content -Path $accountsFile

    # Counter variables for password change successes/failures, non-existent accounts, and duplicate accounts
    $numSuccesses = 0
    $numFailures = 0
    $numNonexistent = 0
    $numDuplicates = 0

    # Arrays to store names of failures, non-existent accounts, and duplicates
    $namesFailures = @()
    $namesNonexistent = @()
    $namesDuplicates = @()

    # For each line with username in accountsFile, check if the user exists by name, find their account username if they exist, and check if they set their password since the start time
    ForEach ($name in $names) {

        # Trim whitespace from username being read
        $name = $name.Trim()
        
        # Check if the user exists
        $userExists = Get-AdUser -Filter {Name -like $name}

        if (!$userExists) { # If user does not exist
            
            # Increment the number of non-existent accounts
            $numNonexistent = $numNonexistent + 1

            # Add name to list of non-existent names
            $namesNonexistent = $namesNonexistent + $name

            # Write to terminal that user doesn't exist
            Write-Host -ForegroundColor Red "ERROR: User with name $name does not exist."

        } else { # If user exists
           
           # Get the user's account username
           $user = Get-AdUser -Filter {Name -like $name} | select -ExpandProperty SamAccountName

           if ($user.Count -gt 1) { # If name matches more than one user

                # Increment the number of duplicate accounts
                $numDuplicates = $numDuplicates + 1

                # Add name to list of duplicate names
                $namesDuplicates = $namesDuplicates + $name

                # Write to terminal that name has duplicates
                Write-Host -ForegroundColor Red "ERROR: More than one user with name $name found in AD. Manual check required."

           } else { # If name matches only one user

                # Get the date of user's last password change
                $passwordLastSet = Get-AdUser $user -properties PwdLastSet | select-object @{Name='PwdLastSet';Expression={[DateTime]::FromFileTime($_.PwdLastSet)}} | select -ExpandProperty PwdLastSet

                # Format date of user's last password change as epoch time
                $passwordLastSet = Get-Date -Date $passwordLastSet -Uformat %s

                if (!($passwordLastSet -gt $offsetStartingTime)) { # If user's last password change is not more recent than starting time

                    # Increment number of failures
                    $numFailures = $numFailures + 1

                    # Add name to list of failures
                    $namesFailures = $namesFailures + $name
                    
                    # Write to terminal that user has not changed their password
                    Write-Host -ForegroundColor Yellow "FAILURE: $name ($user) has not reset their password yet."

                } else { # If user's last password change is more recent than starting time

                    # Increment number of successes
                    $numSuccesses = $numSuccesses + 1

                }

           }

        }

    }

    # Write new line to terminal (for formatting)
    Write-Host "`n"

    # Store and output the total number of entries
    $numTotalEntries = $numSuccesses + $numFailures + $numNonexistent + $numDuplicates
    Write-Host "A total of $numTotalEntries entries were processed."

    # Write new line to terminal (for formatting)
    Write-Host "`n"

    # Store and output the number of valid and invalid entries
    $numValidEntries = $numSuccesses + $numFailures
    $numInvalidEntries = $numNonexistent + $numDuplicates
    Write-Host "Of these, " -NoNewline
    Write-Host -ForegroundColor Green "$numValidEntries" -NoNewLine
    Write-Host " entries were valid and " -NoNewLine
    Write-Host -ForegroundColor Red "$numInvalidEntries" -NoNewLine
    Write-Host " were invalid ($numNonexistent non-existent and $numDuplicates duplicates)."

    if ($numNonexistent -gt 0) { # If there are any non-existent entries

        # Print non-existent names
        Write-Host "Non-existent names: "
        ForEach ($nameNonexistent in $namesNonexistent) {

            Write-Host -ForegroundColor Red "$nameNonexistent"

        }

    }

    if ($numDuplicates -gt 0) { # If there are any duplicates

        # Print duplicate names
        Write-Host "Names with duplicates: "
        ForEach ($nameDuplicates in $namesDuplicates) {
            
            Write-Host -ForegroundColor Red "$nameDuplicates"
        
        }

    }

    # Write new line to terminal (for formatting)
    Write-Host "`n"

    # Output the number of successes and failures
    Write-Host "Of the valid entries, " -NoNewLine
    Write-Host -ForegroundColor Green "$numSuccesses" -NoNewLine
    Write-Host "/$numValidEntries changed their password and " -NoNewLine
    Write-Host -ForegroundColor Yellow "$numFailures" -NoNewLine
    Write-Host "/$numValidEntries have not changed their password."

    if ($numFailures -gt 0) { # If there are any failures

        # Print names of failures
        Write-Host "Names that did not change password: "
        ForEach ($nameFailures in $namesFailures) {

            Write-Host -ForegroundColor Yellow "$nameFailures"

        }

    }

} else { # If user did not select an input file

    Write-Host -ForegroundColor Red "ERROR: System file browser exited without selection. Aborting script."

}
