########################################################################
# Date: 5/19/2010
# Author: Rich Prescott
# Blog: blog.richprescott.com
# Twitter: @Rich_Prescott
########################################################################

function GenerateForm {

[reflection.assembly]::loadwithpartialname("System.Windows.Forms") | Out-Null
[reflection.assembly]::loadwithpartialname("System.Drawing") | Out-Null

$formMain = New-Object System.Windows.Forms.Form
$btnQuery = New-Object System.Windows.Forms.Button
$rtbMain = New-Object System.Windows.Forms.RichTextBox
$lvMain = New-Object System.Windows.Forms.ListView
$InitialFormWindowState = New-Object System.Windows.Forms.FormWindowState

function Query
{
$rtbMain.Text = ""
$btnQuery.Text = "Processing..."
$SelectedClasses = '$lvMain.SelectedItems | %{$_.Text}'
$WMIClassQuery = iex $SelectedClasses
ForEach ($Class in $WMIClassQuery)
    {
    $WMIQuery = Get-WmiObject -class $Class
    $rtbMain.Text += "### $Class `r`n`r`n"
    ForEach ($WMIitem in $WMIQuery)
        {
        ForEach ($Property in $WMIitem.Properties)
            {
            $rtbMain.Text += "{0,-60}`t{1} `r`n" -f $Property.Name, $Property.Value
            }
        $rtbMain.Text += "`r`n"
        }
    }
$btnQuery.Text = "Query"
} #End Function Query

function OnLoad_ListWMI
{
$WMIList = gwmi -list | sort name
ForEach ($Class in $WMIList)
    {
    $lvItem = New-Object System.Windows.Forms.ListViewItem
    $lvItem.Tag = $Class
    $lvItem.Text = $Class.Name
    $lvMain.Items.Add($lvItem)
    }
} #End Function OnLoad_ListWMI

$OnLoadForm_StateCorrection=
{#Correct the initial state of the form to prevent the .Net maximized form issue
    $formMain.WindowState = $InitialFormWindowState
}

#----------------------------------------------
#region Generated Form Code
$formMain.Text = "Arposh WMI Retriever"
$formMain.Name = "formMain"
$formMain.AutoScaleMode = 0
$formMain.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Width = 1000
$System_Drawing_Size.Height = 593
$formMain.ClientSize = $System_Drawing_Size

$btnQuery.TabIndex = 2
$btnQuery.Name = "btnQuery"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Width = 99
$System_Drawing_Size.Height = 23
$btnQuery.Size = $System_Drawing_Size
$btnQuery.UseVisualStyleBackColor = $True
$btnQuery.Text = "Query WMI"
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 54
$System_Drawing_Point.Y = 551
$btnQuery.Location = $System_Drawing_Point
$btnQuery.DataBindings.DefaultDataSourceUpdateMode = 0
$btnQuery.add_Click({Query})
$formMain.Controls.Add($btnQuery)

$rtbMain.Name = "rtbMain"
$rtbMain.Text = ""
$rtbMain.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 235
$System_Drawing_Point.Y = 12
$rtbMain.Location = $System_Drawing_Point
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Width = 756
$System_Drawing_Size.Height = 569
$rtbMain.Size = $System_Drawing_Size
$rtbMain.TabIndex = 1
$formMain.Controls.Add($rtbMain)

$lvMain.UseCompatibleStateImageBehavior = $False
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Width = 215
$System_Drawing_Size.Height = 533
$lvMain.Size = $System_Drawing_Size
$lvMain.DataBindings.DefaultDataSourceUpdateMode = 0
$lvMain.Name = "lvMain"
$lvMain.View = 1
$lvMain.TabIndex = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 14
$System_Drawing_Point.Y = 12
$lvMain.Location = $System_Drawing_Point
$lvMain.FullRowSelect = $True
$lvMain.Columns.Add("Classes",193) | out-null
$formMain.Controls.Add($lvMain)

$InitialFormWindowState = $formMain.WindowState
$formMain.add_Load($OnLoadForm_StateCorrection)
$formMain.add_Load({OnLoad_ListWMI})
$formMain.ShowDialog()| Out-Null

} #End Function

GenerateForm