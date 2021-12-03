################################################################################
#                              File System Reports                             #
#                           Written By: MSgt Brechtel                          #
#                                                                              #
################################################################################
#####Global Variables###########################################################
################################################################################
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
$scriptpath = $MyInvocation.MyCommand.Path
$dir = Split-Path $scriptpath
Set-Location $dir
################################################################################
clear-host
$version="1.1"
$script:prompt_return = "Null";
$script:excel_report = "Null" 
$loading = New-Object System.Windows.Forms.Label
$loading.Font = New-Object System.Drawing.Font("Copperplate Gothic Bold",10,[System.Drawing.FontStyle]::Regular)
################################################################################
function duplicate_file_finder
{


    
    $form = New-Object System.Windows.Forms.Form
    $form.FormBorderStyle = 'Fixed3D'
    $form.BackColor = "#434343"
    #$Form.Opacity = 0.9
    $form.MaximizeBox = $false
    $form.Icon = $icon
    $Form.SizeGripStyle = "Hide"
    $form.Size='400,230'
    $form.Text = "Duplicate File Finder"
    $form.TopMost = $True
    $form.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen
    ################################################################################
    
    
    $title1 = New-Object System.Windows.Forms.Label
    $title2 = New-Object System.Windows.Forms.Label
    ######################################
    $title1.Font = New-Object System.Drawing.Font("Copperplate Gothic Bold",17,[System.Drawing.FontStyle]::Regular)
    $title1.Text="Duplicate File Finder  "
    $title1.TextAlign = 'MiddleCenter'
    $title1.Width=$form.Width
    $title1.Top = 6
    $title1.ForeColor = "white"
    #$title1.BackColor = "darkGray"
    $title1.Left = (($form.width / 2) - ($form.width / 2))
    $form.Controls.Add($title1)

    ###########Title Written By
    $title2.Font = New-Object System.Drawing.Font("Copperplate Gothic",7.5,[System.Drawing.FontStyle]::Regular)
    $title2.Text="Written by: Anthony Brechtel`nVer $version"
    $title2.TextAlign = 'MiddleCenter'
    $title2.ForeColor = "white"
    #$title1.BackColor = "darkGray"
    $title2.Width=$form.Width
    $title2.Height=40
    $title2.Top = 25
    $title2.Left = (($form.width / 2) - ($form.width / 2))
    $form.Controls.Add($title2)



    ##########################
    $report_button = New-Object System.Windows.Forms.Button
    $report_button.Width=150
    $report_button.top = 165
    $report_button.forecolor = "white"
    $report_button.backcolor = "#606060"
    $report_button.Left = ($form.width / 2) - ($report_button.width / 2);   
    $report_button.Text='View Report'
    $report_button.Add_Click({
        Invoke-Item "$script:excel_report"
    })
    



    ##########################
    $scan_target_button = New-Object System.Windows.Forms.Button
    $scan_target_button.Width=150
    $scan_target_button.top = 140 
    $scan_target_button.forecolor = "white"
    $scan_target_button.backcolor = "#606060"
    $scan_target_button.Left = ($form.width / 2) - ($scan_target_button.width / 2);   
    $scan_target_button.Text='Scan Target'
    $scan_target_button.Add_Click({
        $form.Controls.Remove($report_button)
        $scan_target_button.Enabled = $false
        scan_target $script:prompt_return
        $scan_target_button.Enabled = $true
        $form.Controls.Add($report_button)
    })
    ################################################################################
    $target_box = New-Object System.Windows.Forms.TextBox
    $target_box.Location = New-Object System.Drawing.Point(18,107)
    $target_box.Size = New-Object System.Drawing.Size(350,20)
    $target_box.Text = "Browse or Enter a file path"
    $target_box.Add_Click({
        if($target_box.Text -eq "Browse or Enter a file path")
        {
            $target_box.Text = ""
        }
    })
    $target_box.Add_TextChanged({
    
        [string]$script:prompt_return = $target_box.text
        if(($script:prompt_return -ne $null) -and ($script:prompt_return -ne ""))
        {
            if(Test-Path $target_box.text)
            {
                $form.Controls.Add($scan_target_button)
            }
            else
            {
                $form.Controls.Remove($scan_target_button)
            }
        }
        else
        {
            $form.Controls.Remove($scan_target_button)
        }
    })
    $form.Controls.Add($target_box)
    ##################################################################################

    
    $file1_dialog_button = New-Object System.Windows.Forms.Button
    $file1_dialog_button.Location='15,80'
    $file1_dialog_button.Width=200
    $file1_dialog_button.Text='Browse for Target Directory'
    $form.Controls.Add($file1_dialog_button)
    #$file1_dialog_button.BackColor ="darkGray"
    $file1_dialog_button.ForeColor = "White"
    $file1_dialog_button.Backcolor = "#606060"
    $file1_dialog_button.Add_Click(
    {    
		    $script:prompt_return = prompt_for_folder
            
            if(($prompt_return -ne $Null) -and ($prompt_return -ne "") -and ((Test-Path $prompt_return) -eq $True))
            {
                write-host $prompt_return
                $target_box.Text="$prompt_return"
            }
    }
    )

    $form.TopMost = $false;
    $form.ShowDialog()
    
}
################################################################################
function prompt_for_folder()
{  
    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms")|Out-Null

    $foldername = New-Object System.Windows.Forms.FolderBrowserDialog
    $foldername.Description = "Select a folder"
    $foldername.rootfolder = "MyComputer"

    if($foldername.ShowDialog() -eq "OK")
    {
        $folder += $foldername.SelectedPath
    }
    return $folder
}
################################################################################
function scan_target ($target)
{
    $status = "Preparing Scan..."
    loading_start $status
    $output = build_output_file_name $target
    $form.Refresh();

    $files_folders_count = ( Get-ChildItem -literalPath "$target" -Recurse -File -ErrorAction SilentlyContinue | Measure-Object ).Count;
    $status = "Scanning... $files_folders_count Files"
    loading_start $status
    $form.Refresh();
    clear-host

    $files = Get-ChildItem -LiteralPath "$target" -Recurse -File -ErrorAction SilentlyContinue | 
    Group-Object Length | 
    Where-Object { $_.Count -gt 1 } | 
    select -ExpandProperty group | 
    foreach  {get-filehash -literalpath $_.fullname} | 
    group -property hash | 
    where { $_.count -gt 1 }
 

    add-content -literalpath "$dir\Results\$output" "Hash,Path,File,Directory,File,Type,Size GB,Size MB,Size KB,Last Modified Date,Last Accessed Date,Creation Date,Full Path"
    $writer = new-object system.IO.StreamWriter("$dir\Results\$output",$true)
    foreach ($group in $files) 
    {
        foreach ($item in $group.group) 
        {
            #Write-Host ($group.group | Format-Table | Out-String)
            #write-host $item.hash
            #write-host $item.path
            $hash = $item.hash        
            $directory = (Get-Item -literalpath $item.path).Directory
            $name  = (Get-Item -literalpath $item.path).name
            $extention = (Get-Item -literalpath $item.path).Extension
            $sizeGB      = (Get-Item -literalpath $item.path).Length /1Gb
            $sizeGB = [math]::Round($sizeGB,2)
            $sizeMB      = (Get-Item -literalpath $item.path).Length /1Mb
            $sizeMB = [math]::Round($sizeMB,2)
            $sizeKB      = (Get-Item -literalpath $item.path).Length /1Kb
            $sizeKB = [math]::Round($sizeKB,2)
            $modified  = (Get-Item -literalpath $item.path).LastWriteTime
            $accessed  = (Get-Item -literalpath $item.path).LastAccessTime
            $created   = (Get-Item -literalpath $item.path).CreationTime
            $full_path = (Get-Item -literalpath $item.path).FullName


            $dir_link = "=HYPERLINK(`"`"$directory`"`",`"`"Path`"`")";
            $file_link = "=HYPERLINK(`"`"$directory\$name`"`",`"`"File`"`")";

            if($name -match "^-")
            {
                $write_line = "$hash,`"$dir_link`",`"$file_link`",`"$directory`",=`"$name`",$extention,$sizeGB,$sizeMB,$sizeKB,$modified,$accessed,$created,`"$full_path`"";
            }
            else
            {
                $write_line = "$hash,`"$dir_link`",`"$file_link`",`"$directory`",`"$name`",$extention,$sizeGB,$sizeMB,$sizeKB,$modified,$accessed,$created,`"$full_path`"";
            }
        
            $writer.write("$write_line`r`n");
        }
    }
    $writer.Close()


    Write-Progress -Activity "Complete" -status "Complete" -Completed
    $status = "Building Report"
    loading_start $status
    $form.Refresh();

    csv_to_xlsx $output 
    write-host "Done";
    loading_stop
}
################################################################################
################################################################################
function loading_start ($status)
{
    $loading.Text= $status
    $loading.TextAlign = 'MiddleCenter'
    $loading.Width=$form.Width
    $loading.top = $form.height - 60
    $loading.Height = 20;
    #$title1.ForeColor = "white"
    $loading.BackColor = "Red"
    $loading.Left = (($form.width / 2) - ($form.width / 2))
    $form.Controls.Add($loading)
    
}
################################################################################
function loading_stop($loading)
{
    #$script:loading.Text=""
    $form.Controls.remove($script:loading)
    $form.refresh();
}
################################################################################
################Build Output File###############################################
function build_output_file_name($target)
{
    if(!(test-path "$dir\Results"))
    {
        New-Item -ItemType Directory "$dir\Results"
    }
    #$target_name = [System.IO.Path]::GetFileNameWithoutExtension($target)
    $target_name = $target
    $target_name = $target_name.replace(':\',")");
    $target_name = $target_name.replace('/',")");
    $target_name = $target_name.replace('\',")");
    $target_name = $target_name.replace('--',"");
    
    $date = Get-Date -Format G
    [regex]$pattern = " "
    $date = $pattern.replace($date, " @ ", 1);
    $date = $date.replace('/',"-");
    $date = $date.replace(':',".");

    $output = "$target_name      ($date)" + ".csv";
    return $output
}
################################################################################
function csv_to_xlsx($output)
{
    ### Set input and output path
    $inputCSV = "$dir\Results\$output"
    $output2 = [io.path]::GetFileNameWithoutExtension($inputCSV)
    $outputXLSX = "$dir\Results\$output2.xlsx"

    $objExcel = New-Object -ComObject Excel.Application
    $workbook = $objExcel.Workbooks.Open("$inputCSV")
    $worksheet = $workbook.worksheets.item(1) 
    $objExcel.Visible=$false
    $objExcel.DisplayAlerts = $False


    ### Make it pretty
    $worksheet.UsedRange.Columns.Autofit();
    
    $worksheet.Columns.item("B").NumberFormat = "@"
    $worksheet.Columns.item("C").NumberFormat = "@"
    $worksheet.Columns.item("G").NumberFormat = "0"
    $worksheet.Columns.item("H").NumberFormat = "0"
    $worksheet.Columns.item("I").NumberFormat = "0"
    $headerRange = $worksheet.Range("a1","m1")
    $headerRange.AutoFilter() | Out-Null
    $headerRange.Interior.ColorIndex =48
    $headerRange.Font.Bold=$True
    $row_count = $worksheet.UsedRange.Rows.Count
    #$objRange = $worksheet.Range("C2:C$row_count")  
    #[void] $objRange.Sort($objRange) 

    $empty_Var = [System.Type]::Missing
    $sort_col = $worksheet.Range("A1:A$row_count")
    $worksheet.UsedRange.Sort($sort_col,1,$empty_Var,$empty_Var,$empty_Var,$empty_Var,$empty_Var,1)

    $borderrange = $worksheet.Range(“A1","M$row_count")
    $borderrange.Borders.Color = 0
    $borderrange.Borders.Weight = 2




    $workbook.SaveAs($outputXLSX,51)
    $objExcel.Quit()
    $script:excel_report = $outputXLSX;
    Remove-Item "$dir\Results\$output"


}
################################################################################
function csv_to_xlsx1($output)
{
    ### Set input and output path
    $inputCSV = "$dir\Results\$output"
    $output2 = [io.path]::GetFileNameWithoutExtension($inputCSV)
    $outputXLSX = "$dir\Results\$output2.xlsx"

    ### Create a new Excel Workbook with one empty sheet
    $excel = New-Object -ComObject excel.application 
    $workbook = $excel.Workbooks.Add(1)
    $worksheet = $workbook.worksheets.Item(1)

    ### Build the QueryTables.Add command
    ### QueryTables does the same as when clicking "Data » From Text" in Excel
    $TxtConnector = ("TEXT;" + $inputCSV)
    $Connector = $worksheet.QueryTables.add($TxtConnector,$worksheet.Range("A1"))
    $query = $worksheet.QueryTables.item($Connector.name)

    ### Set the delimiter (, or ;) according to your regional settings
    $query.TextFileOtherDelimiter = $Excel.Application.International(5)

    ### Set the format to delimited and text for every column
    ### A trick to create an array of 2s is used with the preceding comma
    $query.TextFileParseType  = 1
    #$query.TextFileColumnDataTypes = ,2 * $worksheet.Cells.Columns.Count
    $query.AdjustColumnWidth = 1
    
    ### Execute & delete the import query
    $query.Refresh()
    $query.Delete()

    ### Add Links
    #$row_count = ($worksheet.UsedRange.Rows.Count - 1)
    #$row = 1;
    #while($row -le $row_count)
    #{
    #    $row++;
    #    $column = 1
    #    $worksheet.Hyperlinks.Add($worksheet.Cells.Item($row,$column),$worksheet.Cells.Item($row,$column).Text,"",$worksheet.Cells.Item($row,$column).Text,$worksheet.Cells.Item($row,$column).Text) | Out-Null
    #    $column = 2
    #   
    #    ###Fix Formula Dashes
    #    if($worksheet.Cells.Item($row,$column).text -match ":") 
    #    {
    #        $text = $worksheet.Cells.Item($row,$column).Text
    #        $text = $text -replace ":","'"
    #        $worksheet.Cells.Item($row,$column) = $text
    #    }
    #
    #
    #    $link = $worksheet.Cells.Item($row,1).Text + "\" + $worksheet.Cells.Item($row,$column).Text
    #   
    #    $worksheet.Hyperlinks.Add($worksheet.Cells.Item($row,$column),$link,"",$worksheet.Cells.Item($row,$column).Text,$worksheet.Cells.Item($row,$column).Text) | Out-Null
    #    $column = 10
    #    $worksheet.Hyperlinks.Add($worksheet.Cells.Item($row,$column),$worksheet.Cells.Item($row,$column).Text,"",$worksheet.Cells.Item($row,$column).Text,$worksheet.Cells.Item($row,$column).Text) | Out-Null
    #}


    
    ### Make it pretty
    $headerRange = $worksheet.Range("a1","M1")
    $headerRange.AutoFilter() | Out-Null
    $headerRange.Interior.ColorIndex =48
    $headerRange.Font.Bold=$True



    ### Save & close the Workbook as XLSX. Change the output extension for Excel 2003
    $workbook.SaveAs($outputXLSX)
    $excel.Quit()
    $script:excel_report = $outputXLSX;
    #Remove-Item "$dir\Results\$output"
}
################################################################################
function Show-Console
{
    param ([Switch]$Show,[Switch]$Hide)
    if (-not ("Console.Window" -as [type])) { 

        Add-Type -Name Window -Namespace Console -MemberDefinition '
        [DllImport("Kernel32.dll")]
        public static extern IntPtr GetConsoleWindow();

        [DllImport("user32.dll")]
        public static extern bool ShowWindow(IntPtr hWnd, Int32 nCmdShow);
        '
    }

    if ($Show)
    {
        $consolePtr = [Console.Window]::GetConsoleWindow()

        # Hide = 0,
        # ShowNormal = 1,
        # ShowMinimized = 2,
        # ShowMaximized = 3,
        # Maximize = 3,
        # ShowNormalNoActivate = 4,
        # Show = 5,
        # Minimize = 6,
        # ShowMinNoActivate = 7,
        # ShowNoActivate = 8,
        # Restore = 9,
        # ShowDefault = 10,
        # ForceMinimized = 11

        $null = [Console.Window]::ShowWindow($consolePtr, 5)
    }

    if ($Hide)
    {
        $consolePtr = [Console.Window]::GetConsoleWindow()
        #0 hide
        $null = [Console.Window]::ShowWindow($consolePtr, 0)
    }
}
################################################################################
Show-Console -Hide
duplicate_file_finder
