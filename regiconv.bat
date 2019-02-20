@powershell -STA -NoProfile -ExecutionPolicy Unrestricted "$s=[scriptblock]::create((gc \"%~f0\"|?{$_.readcount -gt 1})-join\"`n\");&$s "%~dp0 %*&goto:eof

Write-Host "処理開始"

#環境を設定------------------
set USB_PLACE "C:\Users\yuki\Desktop\regiconv\" -option constant
set EXCEL_PLACE "C:\Users\yuki\Desktop\regiconv\temp.xlsx" -option constant
set SAVE_PLACE "C:\Users\yuki\Desktop\"  -option constant
#----------------------------

Add-Type -AssemblyName PresentationFramework
$xaml = @'
<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="売店売上USBデータ変換"
    Height="189.879" Width="369.13">

    <Grid>
        <Button x:Name="input" Content="CSVファイルを選択" HorizontalAlignment="Left" Height="36" Margin="31,23,0,0" VerticalAlignment="Top" Width="101"/>
        <Button x:Name="output" Content="変換" HorizontalAlignment="Left" Height="35" Margin="222,24,0,0" VerticalAlignment="Top" Width="101"/>
        <Expander Header="ファイル" HorizontalAlignment="Left" Height="70" Margin="33,64,0,0" VerticalAlignment="Top" Width="289">

            <Grid Background="#FFE5E5E5">
                <Grid.RowDefinitions>
                    <RowDefinition Height="43*"/>
                    <RowDefinition Height="36*"/>
                </Grid.RowDefinitions>
                <Label x:Name="filename" Content="未選択" HorizontalAlignment="Left" Margin="2,1,0,0" VerticalAlignment="Top"/>
            </Grid>
        </Expander>
        <Label Content="&gt;&gt;" HorizontalAlignment="Left" Height="36" Margin="160,21,0,0" VerticalAlignment="Top" Width="42" FontSize="20" FontFamily="Microsoft YaHei Light"/>

    </Grid>

</Window>
'@


$frm = [System.Windows.Markup.XamlReader]::Parse($xaml)
$btn1 = $frm.FindName("input")
$btn2 = $frm.FindName("output")
$lbl1 = $frm.FindName("filename")

Add-Type -assemblyName System.Windows.Forms
$dialog = New-Object System.Windows.Forms.OpenFileDialog
$dialog.Filter = "CSV ファイル(*.CSV)|*.CSV"
$dialog.InitialDirectory = $USB_PLACE
$dialog.Title = "ファイルを選択してください"

$headerstr = "jan,unused1,name,sales100,unused2,revenue"
$header = $headerstr -split ","

$btn1.Add_Click({
    if ($dialog.ShowDialog() -eq "OK") {
        $lbl1.Content = "選択したファイル " + $dialog.FileName + '"'
    } else {
        $lbl1.Content = "選択なし。"
    }
})

$btn2.Add_Click({
    $csvName = [System.IO.Path]::GetFileNameWithoutExtension($dialog.FileName)
    $csv = Get-Content $dialog.FileName | ConvertFrom-Csv -Header $header

    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $true
    $excel.DisplayAlerts = $true

    $book = $excel.Workbooks.Open($EXCEL_PLACE)
    #ソート順のひな型となるシートを開く
    $sheet = $excel.Worksheets.Item("マスタ")

    #商品行数を取得
    [int]$itemcnt = $sheet.Cells.Item(2,6).Text
    Write-Host $itemcnt

    $eArr = @()
    for ($i=2; $i -lt $itemcnt+2; $i++){
        $eArr += ,[array]@($i,$sheet.Cells.Item($i,2).text,$sheet.Cells.Item($i,3).text,0)
    }

    $eArr | foreach{
         $eJan = $_[2]
         $eSales = $csv | Where-Object {$_.jan -eq $eJan} | ForEach-Object {$_.sales100}
         Write-Host $eSales
         $_[3] = $eSales/100
    }
    Write-Host $eArr

    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel)
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($sheet)
    
    $path = $SAVE_PLACE + $csvName + "変換後売上.xlsx"

    Write-Host "変換データを出力中"
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $true

    $book = $excel.Workbooks.Add()
    $sheet = $excel.Worksheets.Item(1)
    $i=1
    $eArr | foreach{
        $sheet.Cells.Item($i,1) = $_[1]
        $sheet.Cells.Item($i,1).Interior.ColorIndex = 38
        $sheet.Cells.Item($i,2) = $_[3]
        $i++
    }
    $null = $sheet.Columns.AutoFit()
    $book.SaveAs($path)

    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel)
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($sheet)

    Write-Host "変換データを"  $path  "に出力しました"
    Write-Host "終了"
})

$frm.ShowDialog()
