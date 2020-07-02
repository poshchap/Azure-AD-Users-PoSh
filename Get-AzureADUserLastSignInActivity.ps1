function Get-AzureADUserLastSignInActivity {

    ############################################################################

    <#
    .SYNOPSIS

        Gets Azure Active Directory user last interactive sign-in activity details.


    .DESCRIPTION

        Gets Azure Active Directory user last interactive sign-in activity details
        using the signInActivity.lastSignInDateTime attribute.

            Use -All to get details for all users in the target tenant.

            Use -UserObjectId to target a single user or groups of users.

            Use -StaleThreshold to see details of users whose sign-in activity is before
            a certain datetime threshold.

            Use -GuestInfo to include additional information specific to guest accounts

        Can also produce a date and time stamped CSV file as output.


        PRE-REQUISITE - the function uses the MSAL.ps module from the PS Gallery:
        
                        https://www.powershellgallery.com/packages/MSAL.ps


    .EXAMPLE

        Get-AzureADUserLastSignInActivity -TenantId b446a536-cb76-4360-a8bb-6593cf4d9c7f -All

        Gets the last interactive sign-in activity for all users on the tenant.


    .EXAMPLE

        Get-AzureADUserLastSignInActivity -TenantId b446a536-cb76-4360-a8bb-6593cf4d9c7f 
        -UserObjectId 69447235-0974-4af6-bfa3-d0e922a92048 -CsvOutput

        Gets the last interactive sign-in activity for the user, targeted by their object ID.

        Writes the output to a date and time stamped CSV file in the execution directory.


    .EXAMPLE

        Get-AzureADUserLastSignInActivity -TenantId b446a536-cb76-4360-a8bb-6593cf4d9c7f 
        -StaleThreshold 60 -GuestInfo -CsvOutput

        Gets all users whose last interactive sign-in activity is before the stale threshold of 60 days. 

        Writes the output to a date and time stamped CSV file in the execution directory.#

        Includes additional attributes for guest user insight.


    .EXAMPLE

        Get-AzureADUserLastSignInActivity -TenantId b446a536-cb76-4360-a8bb-6593cf4d9c7f
        -StaleThreshold 30

        Gets all users whose last interactive sign-in activity is before the stale threshold of 30 days. 


    .EXAMPLE

        Get-AzureADUserLastSignInActivity -TenantId b446a536-cb76-4360-a8bb-6593cf4d9c7f
        -StaleThreshold 30 -GuestInfo

        Gets all users whose last interactive sign-in activity is before the stale threshold of 30 days. 

        Includes additional attributes for guest user insight.


    #>

    ############################################################################

    [CmdletBinding()]
    param(

        #The tenant ID in GUID form
        [Parameter(Mandatory,Position=0)]
        [guid]$TenantId,

        #Get sign-in activity for all users in the tenant
        [Parameter(Mandatory,Position=1,ParameterSetName="All")]
        [switch]$All,

        #Get the sign-in activity for a single user by object ID
        [Parameter(Mandatory,Position=2,ParameterSetName="UserObjectId")]
        [string]$UserObjectId,

        #The number of days before which accounts are considered stale
        [Parameter(Mandatory,Position=3,ParameterSetName="Threshold")]
        [ValidateSet(30,60,90)] 
        [int32]$StaleThreshold,

        #Include additio al information for guest accounts
        [Parameter(Position=4)]
        [switch]$GuestInfo,

        #Use this switch to create a date and time stamped CSV file
        [Parameter(Position=5)]
        [switch]$CsvOutput

    )


    ############################################################################

    ##################
    ##################
    #region FUNCTIONS

    function Get-AzureADApiToken {

        ############################################################################

        <#
        .SYNOPSIS

            Get an access token for use with the API cmdlets.


        .DESCRIPTION

            Uses MSAL.ps to ontain an access token. Has an option to refresh an existing token.

        .EXAMPLE

            Get-AzureADApiToken -TenantId b446a536-cb76-4360-a8bb-6593cf4d9c7f

            Gets or refreshes an access token for making API calls for the tenant ID
            b446a536-cb76-4360-a8bb-6593cf4d9c7f.


        .EXAMPLE

            Get-AzureADApiToken -TenantId b446a536-cb76-4360-a8bb-6593cf4d9c7f -ForceRefresh

            Gets or refreshes an access token for making API calls for the tenant ID
            b446a536-cb76-4360-a8bb-6593cf4d9c7f.

        #>

        ############################################################################

        [CmdletBinding()]
        param(

            #The tenant ID
            [Parameter(Mandatory,Position=0)]
            [guid]$TenantId,

            #The tenant ID
            [Parameter(Position=1)]
            [switch]$ForceRefresh

        )


        ############################################################################


        #Get an access token using the PowerShell client ID
        $ClientId = "1b730954-1685-4b74-9bfd-dac224a7b894"
        $RedirectUri = "urn:ietf:wg:oauth:2.0:oob"
        $Authority = "https://login.microsoftonline.com/$TenantId"
    
        if ($ForceRefresh) {

            Write-Verbose -Message "$(Get-Date -f T) - Attempting to refresh an existing access token"

            #Attempt to refresh access token
            try {

                $Response = Get-MsalToken -ClientId $ClientId -RedirectUri $RedirectUri -Authority $Authority -ForceRefresh
            }
            catch {}

            #Error handling for token acquisition
            if ($Response) {

                Write-Verbose -Message "$(Get-Date -f T) - API Access Token refreshed - new expiry: $(($Response).ExpiresOn.UtcDateTime)"

                return $Response

            }
            else {
            
                Write-Warning -Message "$(Get-Date -f T) - Failed to refresh Access Token - try re-running the cmdlet again"

            }

        }
        else {

            Write-Verbose -Message "$(Get-Date -f T) - Checking token cache"

            #Run this to obtain an access token - should prompt on first run to select the account used for future operations
            try {

                $Response = Get-MsalToken -ClientId $ClientId -RedirectUri $RedirectUri -Authority $Authority -Prompt SelectAccount
            }
            catch {}

            #Error handling for token acquisition
            if ($Response) {

                Write-Verbose -Message "$(Get-Date -f T) - API Access Token obtained"

                return $Response

            }
            else {

                Write-Warning -Message "$(Get-Date -f T) - Failed to obtain an Access Token - try re-running the cmdlet again"
                Write-Warning -Message "$(Get-Date -f T) - If the problem persists, start a new PowerShell session"

            }

        }


    }   #end function


    function Get-AzureADHeader {
    
        [CmdletBinding()]
        param(

            #The tenant ID
            [Parameter(Mandatory,Position=0)]
            [string]$Token,

            #Switch to include ConsistencyLevel = Eventual for $count operations
            [Parameter(Position=1)]
            [switch]$ConsistencyLevelEventual

            )

        if ($ConsistencyLevelEventual) {

            return @{

                "Authorization" = ("Bearer {0}" -f $Token);
                "Content-Type" = "application/json";
                "ConsistencyLevel" = "eventual";

            }

        }
        else {

            return @{

                "Authorization" = ("Bearer {0}" -f $Token); 
                "Content-Type" = "application/json";

            }

        }

    }   #end function


    #endregion


    ############################################################################

    ##################
    ##################
    #region MAIN

    #Try and get MSAL.ps module 
    $MSAL = Get-Module -ListAvailable MSAL.ps -Verbose:$false -ErrorAction SilentlyContinue

    if ($MSAL) {

        Write-Verbose -Message "$(Get-Date -f T) - MSAL.ps installed"

        #Deal with different search criterea
        if ($All) {

            #API endpoint
            $Filter = "?`$select=displayName,userPrincipalName,Id,signInActivity,userType,externalUserState,creationType,createdDateTime"

            Write-Verbose -Message "$(Get-Date -f T) - All user mode selected"

        }
        elseif ($UserObjectId) {

            #API endpoint
            $Filter = "?`$filter=ID eq '$UserObjectId'&`$select=displayName,userPrincipalName,Id,signInActivity,userType,externalUserState,creationType,createdDateTime"

            Write-Verbose -Message "$(Get-Date -f T) - Single user mode selected"

        }
        elseif ($StaleThreshold) {

            Write-Verbose -Message "$(Get-Date -f T) - Stale mode selected"

            #Obtain a datetime object before which accounts are considered stale
            $DaysAgo = (Get-Date (Get-Date).AddDays(-$StaleThreshold) -Format s) + "Z"

            Write-Verbose -Message "$(Get-Date -f T) - Stale threshold set to $DaysAgo"

            #API endpoint
            $Select = "&`$select=displayName,userPrincipalName,Id,signInActivity,userType,externalUserState,creationType,createdDateTime"
            $Filter = "?`$filter=signInActivity/lastSignInDateTime le $DaysAgo$Select"

        }


        ############################################################################
    
        $Url = "https://graph.microsoft.com/beta/users$Filter"


        ############################################################################

        #Get / refresh an access token
        $global:Token = (Get-AzureADApiToken -TenantId $TenantId).AccessToken

        if ($Token) {

            if ($All) {

                #Construct header with access token and ConsistencyLevel = Eventual
                $Headers = Get-AzureADHeader -Token $Token -ConsistencyLevelEventual

                $CountUrl = "https://graph.microsoft.com/beta/users/`$count"

                Write-Verbose -Message "$(Get-Date -f T) - Invoking web request for $CountUrl"

                #Now make a call to get the number of users
                try {
                 
                    $UserCount = (Invoke-WebRequest -Headers $Headers -Uri $CountUrl -Verbose:$false)

                }
                catch {}

                if ($UserCount) {

                    #Estimate execution time
                    if ($CsvOutput) {

                        $ExTime = (0.03 * $UserCount.Content)

                    }
                    else {

                        $ExTime = (0.035 * $UserCount.Content)

                    }


                    $ExTimeSpan = [timespan]::FromSeconds($ExTime)

                    Write-Verbose -Message "$(Get-Date -f T) - Estimated function execution time is $($ExTimeSpan.Hours) hours, $($ExTimeSpan.Minutes) minutes, $($ExTimeSpan.Seconds) seconds"
                    

                    #Light up the progress bar in the later loop
                    $ShowProgress = $true


                }
                else {

                    Write-Warning -Message "$(Get-Date -f T) - User count unobtainable - unable to estimate function execution time"
                }

            }
            else {

                #Construct header with access token
                $Headers = Get-AzureADHeader -Token $Token

            }

            #Tracking variables
            $Count = 0
            $RetryCount = 0
            $OneSuccessfulFetch = $false
            $TotalReport = $null
            $i = 1


            #Do until the fetch URL is null
            do {

                Write-Verbose -Message "$(Get-Date -f T) - Invoking web request for $Url"

                ##################################
                #Do our stuff with error handling
                try {

                    #Invoke the web request
                    $MyReport = (Invoke-WebRequest -UseBasicParsing -Headers $Headers -Uri $Url -Verbose:$false)

                }
                catch [System.Net.WebException] {
        
                    $StatusCode = [int]$_.Exception.Response.StatusCode
                    Write-Warning -Message "$(Get-Date -f T) - $($_.Exception.Message)"

                    #Check what's gone wrong
                    if (($StatusCode -eq 401) -and ($OneSuccessfulFetch)) {

                        #Token might have expired; renew token and try again
                        $Token = (Get-AzureADApiToken -TenantId $TenantId).AccessToken
                        $Headers = Get-AzureADHeader -Token $Token
                        $OneSuccessfulFetch = $False

                    }
                    elseif (($StatusCode -eq 429) -or ($StatusCode -eq 504) -or ($StatusCode -eq 503)) {

                        #Throttled request or a temporary issue, wait for a few seconds and retry
                        Start-Sleep -Seconds 5

                    }
                    elseif (($StatusCode -eq 403) -or ($StatusCode -eq 401)) {

                        Write-Warning -Message "$(Get-Date -f T) - Please check the permissions of the user"
                        break

                    }
                    elseif ($StatusCode -eq 400) {

                        Write-Warning -Message "$(Get-Date -f T) - Please check the query used"
                        break

                        }
                    else {
            
                        #Retry up to 5 times
                        if ($RetryCount -lt 5) {
                
                            write-output "Retrying..."
                            $RetryCount++

                        }
                        else {
                
                            #Write to host and exit loop
                            Write-Warning -Message "$(Get-Date -f T) - Download request failed. Please try again in the future"
                            break

                        }

                    }

                }
                catch {

                    #Write error details to host
                    Write-Warning -Message "$(Get-Date -f T) - $($_.Exception)"


                    #Retry up to 5 times    
                    if ($RetryCount -lt 5) {

                        write-output "Retrying..."
                        $RetryCount++

                    }
                    else {

                        #Write to host and exit loop
                        Write-Warning -Message "$(Get-Date -f T) - Download request failed - please try again in the future"
                        break

                    }

                } # end try / catch


                ###############################
                #Convert the content from JSON
                $ConvertedReport = ($MyReport.Content | ConvertFrom-Json).value

                $TotalObjects = @()

                foreach ($User in $ConvertedReport) {

                    if ($GuestInfo) {

                        #Construct a custom object
                        $Properties = [PSCustomObject]@{

                            displayName = $User.displayName
                            userPrincipalName = $User.userPrincipalName
                            objectId = $User.Id
                            lastSignInDateTime = $User.signInActivity.lastSignInDateTime
                            lastSignInRequestId = $User.signInActivity.lastSignInRequestId
                            userType = $User.userType
                            createdDateTime = $User.createdDateTime
                            externalUserState = $User.externalUserState
                            creationType = $User.creationType

                        }
            
                    }
                    else {

                        #Construct a custom object
                        $Properties = [PSCustomObject]@{

                            displayName = $User.displayName
                            userPrincipalName = $User.userPrincipalName
                            objectId = $User.Id
                            lastSignInDateTime = $User.signInActivity.lastSignInDateTime
                            lastSignInRequestId = $User.signInActivity.lastSignInRequestId

                        }

                    }

                    $TotalObjects += $Properties

                    #Progress bar when targeting all users
                    if ($ShowProgress) {

                        Write-Progress -Activity "Processing..." `
                                    -Status ("Checked {0}/{1} user accounts" -f $i++, $UserCount.Content) `
                                    -PercentComplete ((($i -1) / $UserCount.Content) * 100)

                    }

                }


                #Add to concatenated findings
                [array]$TotalReport += $TotalObjects

                #Update the fetch url to include the paging element
                $Url = ($myReport.Content | ConvertFrom-Json).'@odata.nextLink'

                #Update the access tokenon the second iteration
                if ($OneSuccessfulFetch) {
                
                    $Token = (Get-AzureADApiToken -TenantId $TenantId).AccessToken
                    $Headers = Get-AzureADHeader -Token $Token

                }

                #Update count and show for this cycle
                $Count = $Count + $ConvertedReport.Count
                Write-Verbose -Message "$(Get-Date -f T) - Total records fetched: $count"

                #Update tracking variables
                $OneSuccessfulFetch = $true
                $RetryCount = 0


            } while ($Url) #end do / while


        }

        #See if we need to write to CSV
        if ($CsvOutput) {

            #Output file
            $now = "{0:yyyyMMdd_hhmmss}" -f (Get-Date)
            $CsvName = "UserLastSignInDetails_$now.csv"

            Write-Verbose -Message "$(Get-Date -f T) - Generating a CSV for last user Sign-In details"

            $TotalReport | Export-Csv -Path $CsvName -NoTypeInformation

            Write-Verbose -Message "$(Get-Date -f T) - Last user sign-in details written to $(Get-Location)\$CsvName"

        }
        else {

            #Return stuff
            $TotalReport

        }

    }
    else {

        Write-Warning -Message "$(Get-Date -f T) - Please install the MSAL.ps PowerShell module (Find-Module MSAL.ps)"    

    }

    #endregion


}   #end function