Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

$Form                            = New-Object system.Windows.Forms.Form
$Form.ClientSize                 = New-Object System.Drawing.Point(574,400)
$Form.text                       = "Form"
$Form.TopMost                    = $false

$ListBox1                        = New-Object system.Windows.Forms.ListBox
$ListBox1.text                   = "listBox"
$ListBox1.width                  = 195
$ListBox1.height                 = 207
$ListBox1.location               = New-Object System.Drawing.Point(12,98)
$ListBox1.ScrollAlwaysVisible    = $true
$listBox.SelectionMode = 'MultiExtended'

$Form.controls.AddRange(@($ListBox1))


#region Logic 


#Write your logic code here

#endregion
$testvalue = 1..20
foreach($value in $testvalue){
    $ListBox1.Items.Add($Value)
}
[void]$Form.ShowDialog()