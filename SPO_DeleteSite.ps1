<# 
THIS SCRIPT IS CREATED BY HATIM DHILA 
Contact me on hatimadhila@gmail.com for any queries.
#>

Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

$Form                            = New-Object system.Windows.Forms.Form
$Form.ClientSize                 = New-Object System.Drawing.Point(444,429)
$Form.text                       = "SharePoint Online - Delete Site"
$Form.TopMost                    = $true

$TextBox1                        = New-Object system.Windows.Forms.TextBox
$TextBox1.multiline              = $false
$TextBox1.width                  = 358
$TextBox1.height                 = 20
$TextBox1.location               = New-Object System.Drawing.Point(59,7)
$TextBox1.Font                   = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$TextBox2                        = New-Object system.Windows.Forms.TextBox
$TextBox2.multiline              = $false
$TextBox2.width                  = 337
$TextBox2.height                 = 20
$TextBox2.location               = New-Object System.Drawing.Point(80,33)
$TextBox2.Font                   = New-Object System.Drawing.Font('Microsoft Sans Serif',10)
<#
$TextBox3                        = New-Object system.Windows.Forms.TextBox
$TextBox3.multiline              = $false
$TextBox3.width                  = 339
$TextBox3.height                 = 20
$TextBox3.location               = New-Object System.Drawing.Point(80,58)
$TextBox3.Font                   = New-Object System.Drawing.Font('Microsoft Sans Serif',10)
#>

$Label1                          = New-Object system.Windows.Forms.Label
$Label1.text                     = "URL"
$Label1.AutoSize                 = $true
$Label1.width                    = 25
$Label1.height                   = 10
$Label1.location                 = New-Object System.Drawing.Point(11,8)
$Label1.Font                     = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$Label2                          = New-Object system.Windows.Forms.Label
$Label2.text                     = "Username"
$Label2.AutoSize                 = $true
$Label2.width                    = 25
$Label2.height                   = 10
$Label2.location                 = New-Object System.Drawing.Point(8,33)
$Label2.Font                     = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

<#
$Label3                          = New-Object system.Windows.Forms.Label
$Label3.text                     = "Password"
$Label3.AutoSize                 = $true
$Label3.width                    = 25
$Label3.height                   = 10
$Label3.location                 = New-Object System.Drawing.Point(11,60)
$Label3.Font                     = New-Object System.Drawing.Font('Microsoft Sans Serif',10)
#>

$Button1                         = New-Object system.Windows.Forms.Button
$Button1.text                    = "Get SPO Sites"
$Button1.width                   = 407
$Button1.height                  = 30
$Button1.location                = New-Object System.Drawing.Point(17,85)
$Button1.Font                    = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$ListBox1                        = New-Object system.Windows.Forms.ListBox
$ListBox1.text                   = "listBox"
$ListBox1.width                  = 211
$ListBox1.height                 = 274
$ListBox1.location               = New-Object System.Drawing.Point(10,129)

$Button2                         = New-Object system.Windows.Forms.Button
$Button2.text                    = "Delete SPO Site"
$Button2.width                   = 124
$Button2.height                  = 41
$Button2.location                = New-Object System.Drawing.Point(265,249)
$Button2.Font                    = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$Form.controls.AddRange(@($TextBox1,$TextBox2,$TextBox3,$Label1,$Label2,$Label3,$Button1,$ListBox1,$Button2))

$path = $PSScriptRoot
Import-Module "$path\SPO_module\Microsoft.Online.SharePoint.PowerShell.psd1" 

$Button1.Add_Click({ Get_SPO_Sites })
$Button2.Add_Click({ Delete_SPO_Site })

$Global:SPOSiteArray = $null


function Get_SPO_Sites { 
    $URL = $TextBox1.Text
    $Username = $TextBox2.Text

    Connect-SPOService -Url $URL -Credential $Username
    ListSites
}

function ListSites{
    
    $Global:SPOSiteArray = Get-SPOSite -Detailed -WarningAction SilentlyContinue | Sort-Object Title
    
    $ListBox1.Items.Clear()
    if($Global:SPOSiteArray -eq $null){
         $ListBox1.Items.Add("Please Click on Get SPO Sites first")
    }
    else{
        foreach($SPOSite in $Global:SPOSiteArray){
            $ListBox1.Items.Add("$($SPOSite.Title)")
        }
    }
}


function Delete_SPO_Site {
    $index = $ListBox1.SelectedIndex
    $SiteURL = $Global:SPOSiteArray[$index].url
    #$Global:MIGSelected = "$($Global:MIGArray[$index].Name)"
    Remove-SPOSite -Identity $SiteURL -Confirm:$false
    ListSites
 }


[void]$Form.ShowDialog()