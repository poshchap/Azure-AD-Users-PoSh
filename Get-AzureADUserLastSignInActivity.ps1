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

        Can also produce a date and time stamped CSV file as output.

        IMPORTANT: 

            * The -Verbose switch will help you understand what the function is doing


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
        -StaleThreshold 60 -CsvOutput

        Gets all users whose last interactive sign-in activity is before the stale threshold of 60 days. 

        Writes the output to a date and time stamped CSV file in the execution directory.


    .EXAMPLE

        Get-AzureADUserLastSignInActivity -TenantId b446a536-cb76-4360-a8bb-6593cf4d9c7f
        -StaleThreshold 30

        Gets all users whose last interactive sign-in activity is before the stale threshold of 30 days. 


    .NOTES

    THIS CODE-SAMPLE IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESSED 
    OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR 
    FITNESS FOR A PARTICULAR PURPOSE.

    This sample is not supported under any Microsoft standard support program or service. 
    The script is provided AS IS without warranty of any kind. Microsoft further disclaims all
    implied warranties including, without limitation, any implied warranties of merchantability
    or of fitness for a particular purpose. The entire risk arising out of the use or performance
    of the sample and documentation remains with you. In no event shall Microsoft, its authors,
    or anyone else involved in the creation, production, or delivery of the script be liable for 
    any damages whatsoever (including, without limitation, damages for loss of business profits, 
    business interruption, loss of business information, or other pecuniary loss) arising out of 
    the use of or inability to use the sample or documentation, even if Microsoft has been advised 
    of the possibility of such damages, rising out of the use of or inability to use the sample script, 
    even if Microsoft has been advised of the possibility of such damages. 

    #>

    ############################################################################

    [CmdletBinding()]
    param(

        #The tenant ID
        [Parameter(Mandatory,Position=0)]
        [string]$TenantId,

        #The user or users initiating the action by ID
        [Parameter(Mandatory,Position=1,ParameterSetName="All")]
        [switch]$All,

        #The service principal or principals initiating the action by Display Name
        [Parameter(Mandatory,Position=2,ParameterSetName="UserObjectId")]
        [string]$UserObjectId,

        #The number of days before which accounts are considered stale
        [Parameter(Mandatory,Position=3,ParameterSetName="Threshold")]
        [ValidateSet(30,60,90)] 
        [int32]$StaleThreshold,

        #Use this switch to create a date and time stamped CSV file
        [Parameter(Position=4)]
        [switch]$CsvOutput

    )


    ############################################################################

    #Function to construct a header for the web request (with token)

    function Get-Headers {
    
        param($Token)

        return @{

            "Authorization" = ("Bearer {0}" -f $Token);
            "Content-Type" = "application/json";

        }

    }   #end function


    ############################################################################

    #Function to get a token for MS Graph with PowerShell client ID

    function Get-AzureADApiToken {

        ############################################################################

        <#
        .SYNOPSIS

            Get an access token for use with the API cmdlets.


        .DESCRIPTION

           Check the global $TokenObtained variable. 
       
           If true, i.e. we've previously obtained a token, will attempt a refresh. 

           If false, i.e. we haven't previously obtained a token, will attempt an 
           interactive authentication. 


        .EXAMPLE

            Get-AzureADApiToken -TenantId b446a536-cb76-4360-a8bb-6593cf4d9c7f

            Gets or refreshes an access token for making API calls for the tenant ID
            b446a536-cb76-4360-a8bb-6593cf4d9c7f.


        #>

        ############################################################################

        [CmdletBinding()]
        param(

            #The tenant ID
            [Parameter(Mandatory,Position=0)]
            [string]$TenantId

        )


        ############################################################################


        #Get an access token using the PowerShell client ID
        $ClientId = "1b730954-1685-4b74-9bfd-dac224a7b894"
        $RedirectUri = "urn:ietf:wg:oauth:2.0:oob"
        $Authority = "https://login.microsoftonline.com/$TenantId"
    
        if ($TokenObtained) {

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

            Write-Verbose -Message "$(Get-Date -f T) - Please input a credential or select an existing account"

            #Run this to interactvely obtain an access token
            try {

                $Response = Get-MsalToken -ClientId $ClientId -RedirectUri $RedirectUri -Authority $Authority -Interactive
            }
            catch {}

            #Error handling for token acquisition
            if ($Response) {

                Write-Verbose -Message "$(Get-Date -f T) - API Access Token obtained"

                #Global variable to show we've already obtained a token
                $TokenObtained = $true

                return $Response

            }
            else {

                Write-Warning -Message "$(Get-Date -f T) - Failed to obtain an Access Token - try re-running the cmdlet again"

            }

        }


    }   #end function


    ############################################################################

    #Try and get MSAL.ps module 
    $MSAL = Get-Module -ListAvailable MSAL.ps -Verbose:$false -ErrorAction SilentlyContinue

    if ($MSAL) {

        #Deal with different search criterea
        if ($All) {

            #API endpoint
            $Filter = "?`$select=displayName,userPrincipalName,Id,signInActivity"

            Write-Verbose -Message "$(Get-Date -f T) - All user mode selected"

        }
        elseif ($UserObjectId) {

            #API endpoint
            $Filter = "?`$filter=ID eq '$UserObjectId'&`$select=displayName,userPrincipalName,Id,signInActivity"

            Write-Verbose -Message "$(Get-Date -f T) - Single user mode selected"

        }
        elseif ($StaleThreshold) {

            Write-Verbose -Message "$(Get-Date -f T) - Stale mode selected"

            #Obtain a datetime object before which accounts are considered stale
            $DaysAgo = (Get-Date (Get-Date).AddDays(-$StaleThreshold) -Format s) + "Z"

            Write-Verbose -Message "$(Get-Date -f T) - Stale threshold set to $DaysAgo"

            #API endpoint
            $Filter = "?`$filter=signInActivity/lastSignInDateTime le $DaysAgo&`$select=displayName,userPrincipalName,Id,signInActivity"

        }


        ############################################################################
    
        $Url = "https://graph.microsoft.com/beta/users$Filter"


        ############################################################################

        #Get / refresh an access token
        $Token = (Get-AzureADApiToken -TenantId $TenantId).AccessToken

        if ($Token) {

            #Construct header with access token
            $Headers = Get-Headers($Token)

            #Tracking variables
            $Count = 0
            $RetryCount = 0
            $OneSuccessfulFetch = $false
            $TotalReport = $null


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
                        $Headers = Get-Headers($Token)
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
                
                            Write-Host "Retrying..."
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

                        Write-Host "Retrying..."
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

                    #Construct a custom object
                    $Properties = [PSCustomObject]@{

                        displayName = $User.displayName
                        userPrincipalName = $User.userPrincipalName
                        objectId = $User.Id
                        lastSignInDateTime = $User.signInActivity.lastSignInDateTime
                        lastSignInRequestId = $User.signInActivity.lastSignInRequestId

                    } 
            
                    $TotalObjects += $Properties

                }

                #Add to concatenated findings
                [array]$TotalReport += $TotalObjects

                #Update the fetch url to include the paging element
                $Url = ($myReport.Content | ConvertFrom-Json).'@odata.nextLink'

                #Update count and show for this cycle
                $Count = $Count + $ConvertedReport.Count
                Write-Verbose -Message "$(Get-Date -f T) - Total records fetched: $count"

                #Update tracking variables
                $OneSuccessfulFetch = $true
                $RetryCount = 0


            } while ($Url -ne $null) #end do / while


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

}   #end function