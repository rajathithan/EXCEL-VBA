$filepath = split-path $SCRIPT:MyInvocation.MyCommand.Path -parent
$filepathN = $filepath + '\uploadFile.xlsm'
$app = New-Object -comobject Excel.Application
$wb = $app.Workbooks.Open($filepathN)
$app.Run("uploadFileMacro")
$app.Quit()
