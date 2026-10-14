# Structured to scale org wide 

Connect-ExchangeOnline -UserPrincipalName admin@ICAN.org.uk

# Hash table avoids - prefix and easier to scale
$fontSettings = @{
    DefaultFontName = 'Century Gothic'
    DefaultFontSize = 11
}

Set-MailboxMessageConfiguration -Identity richard.hepburn@speechandlanguage.org.uk @fontSettings

# debug configuration
Get-MailboxMessageConfiguration -Identity richard.hepburn@speechandlanguage.org.uk |
    Select-Object DefaultFontName, DefaultFontSize