## Domains are set in global variables

$Domains = $ExchangeOnlineSharedMailboxDomain.split(';')

foreach ($Domain in $Domains) {

    $returnObject = @{
        name               = $domain
        domain            = $domain
    }

    Write-Output $returnObject
}

