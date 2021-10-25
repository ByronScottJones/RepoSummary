#requires -version 6
<#
.SYNOPSIS
  Queries Github to get the Pull Requests in the last two weeks for a project, and email a summary.
.DESCRIPTION
  Queries Github to get the Pull Requests in the last two weeks for a project, and email a summary.
.PARAMETER Owner
    The Owner Name for the Repository
.PARAMETER Repository
    The Repository Name
.PARAMETER EmailAddresses
    The email addresses to send the summary report to
.Parameter Days
    OPTIONAL The number of days to go back. Defaults to 14
.Parameter ReturnResults
    OPTIONAL FLAG Whether to return the Query Results to the Pipeline as a PSCustomObject
.INPUTS
  None
.OUTPUTS
  Sends an email to the Recipients
.NOTES
  Version:        1.0
  Author:         Byron Jones
  Creation Date:  2021-10-23
  Purpose/Change: Initial script development
  
.EXAMPLE
  send-reposummary -Owner "Dotnet" -Repository "SDK" -EmailAddresses @{"byronjones@outlook.com"} -Days 14 -ReturnResults
#>

param (
    [string]$Owner  = $(throw "-Repository is required."), 
    [string]$Repository  = $(throw "-Repository is required."),    
    [string[]]$EmailAddresses,
    [int]$Days = 14,
    [switch]$ReturnResults = $false
)

#Check for the Powershell Secrets Modules
if (!((Get-Module "Microsoft.Powershell.SecretManagement") -and (Get-Module "Microsoft.Powershell.SecretStore"))) {
    write-host "Prerequisites:"
    write-host " Install-Module Microsoft.PowerShell.SecretManagement"
    write-host " Install-Module Microsoft.PowerShell.SecretStore"
    write-host " Import-Module Microsoft.Powershell.SecretManagement"
    write-host " Import-Module Microsoft.Powershell.SecretStore"
    write-host " Register-SecretVault -Name SecretStore -ModuleName Microsoft.PowerShell.SecretStore -DefaultVault"
    write-host " Set-Secret -Name ""GitHubAPI"" -Secret ""<Insert GitHub API Token Here>"""
    
}

#Import the Powershell Secrets Modules
Import-Module Microsoft.PowerShell.SecretManagement
Import-Module Microsoft.PowerShell.SecretStore

#Try to get the GitHub API Token
try{
    $GitHubAPIToken = get-secret -Vault SecretStore -name "GitHubAPI" -AsPlainText
    $GitHub_Oauth_Header = @{"Authorization" = "Bearer $GitHubAPIToken"}
}
catch {
    Throw "Unable to retrieve the Github API Token from the Secrets Store. Make sure it is stored as 'GitHubAPI'"
}

#Set Variables
$url = "https://api.github.com/graphql"
$EmailFrom = "get-reposummary@mycompany.com"
$SMTPServer = "smtp.company.com"
$StartDate = (get-date).AddDays(-$Days).ToString("yyyy-MM-dd")
$LastCount = 100

#Create the Github GraphQL Query
$body = @{"query" = @"
{
    search(query: "repo:$Owner/$Repository is:pr created:>$StartDate", type: ISSUE, last:$LastCount) {
    edges {
      node {
        ... on PullRequest {
          url
          title
          createdAt
          state
          isDraft
          author{
            login
          }
          bodyText
          commits(last:10) {
            edges {
              node {
                url
              }
            }
          }
          files(last:10) {
            edges {
              node {
                path
              }
            }
          }
          comments(last:2) {
            edges {
              node {
                author {
                  login
                }
                bodyText
              }
            }
          }
        }
      }
    }
  }
}
"@
} | convertto-json


$result = Invoke-RestMethod -Method 'Post' -Uri $url -Headers $GitHub_Oauth_Header -ContentType "application/json; charset=utf-8" -Body $body


#Prepare the Email Body
[System.Collections.ArrayList]$Results = @()

foreach ($node in $result.data.search.edges.node) {
    
    $commits = ""
    foreach($url in $node.commits.edges.node){
        $urltext = $url.url
        $commits += "$urltext`r`n"
    }

    $tempobj = [PSCustomObject]@{
        url = $node.url
        title = $node.title
        author = $node.author.login
        state = $node.state
        draft = $node.isDraft
        created = $node.createdAt
        description = $node.bodyText
        commits = $commits

    }

    # Redirection to null is necessary as ArrayList outputs an index number for each added item.
    $Results.Add($tempobj) > $null
}

$ResultsText = $Results | format-list | out-string

#Send the Emails
if($null -ne $EmailAddresses){
    $MailMessage = @{
        To = $EmailAddresses
        From = $EmailFrom
        Subject = “GitHub Repository Summary: $Owner/$Repository for $((get-date).ToString("yyyy-MM-dd"))”
        Body = $ResultsText
        Smtpserver = $SMTPServer
        ErrorAction = “SilentlyContinue”
    }
    Send-MailMessage @MailMessage
}


#Return the Results Object if requested
if($ReturnResults){
    return $results
}