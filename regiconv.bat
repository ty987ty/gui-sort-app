@powershell -STA -NoProfile -ExecutionPolicy Unrestricted "$s=[scriptblock]::create((gc \"%~f0\"|?{$_.readcount -gt 1})-join\"`n\");&$s "%~dp0 %*&goto:eof

Write-Host "�����J�n"

#����ݒ�------------------
set USB_PLACE "C:\Users\yuki\Desktop\regiconv\" -option constant
set EXCEL_PLACE "C:\Users\yuki\Desktop\regiconv\temp.xlsx" -option constant
set SAVE_PLACE "C:\Users\yuki\Desktop\"  -option constant
#----------------------------

Add-Type -AssemblyName PresentationFramework
$xaml = @'
<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="���X����USB�f�[�^�ϊ�"
    Height="189.879" Width="369.13">

    <Grid>
        <Button x:Name="input" Content="CSV�t�@�C����I��" HorizontalAlignment="Left" Height="36" Margin="31,23,0,0" VerticalAlignment="Top" Width="101"/>
        <Button x:Name="output" Content="�ϊ�" HorizontalAlignment="Left" Height="35" Margin="222,24,0,0" VerticalAlignment="Top" Width="101"/>
        <Expander Header="�t�@�C��" HorizontalAlignment="Left" Height="70" Margin="33,64,0,0" VerticalAlignment="Top" Width="289">

            <Grid Background="#FFE5E5E5">
                <Grid.RowDefinitions>
                    <RowDefinition Height="43*"/>
                    <RowDefinition Height="36*"/>
                </Grid.RowDefinitions>
                <Label x:Name="filename" Content="���I��" HorizontalAlignment="Left" Margin="2,1,0,0" VerticalAlignment="Top"/>
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
$dialog.Filter = "CSV �t�@�C��(*.CSV)|*.CSV"
$dialog.InitialDirectory = $USB_PLACE
$dialog.Title = "�t�@�C����I�����Ă�������"

$headerstr = "jan,unused1,name,sales100,unused2,revenue"
$header = $headerstr -split ","

$btn1.Add_Click({
    if ($dialog.ShowDialog() -eq "OK") {
        $lbl1.Content = "�I�������t�@�C�� " + $dialog.FileName + '"'
    } else {
        $lbl1.Content = "�I���Ȃ��B"
    }
})

$btn2.Add_Click({
    $csvName = [System.IO.Path]::GetFileNameWithoutExtension($dialog.FileName)
    $csv = Get-Content $dialog.FileName | ConvertFrom-Csv -Header $header

    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $true
    $excel.DisplayAlerts = $true

    $book = $excel.Workbooks.Open($EXCEL_PLACE)
    #�\�[�g���̂ЂȌ^�ƂȂ�V�[�g���J��
    $sheet = $excel.Worksheets.Item("�}�X�^")

    #���i�s�����擾
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
    
    $path = $SAVE_PLACE + $csvName + "�ϊ��㔄��.xlsx"

    Write-Host "�ϊ��f�[�^���o�͒�"
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

    Write-Host "�ϊ��f�[�^��"  $path  "�ɏo�͂��܂���"
    Write-Host "�I��"
})

$frm.ShowDialog()
