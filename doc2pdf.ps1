## PowerShell script to convert word files to pdf, in a folder
$wordApp = New-Object -ComObject Word.Application
$docFiles = Get-ChildItem -Path "C:\word" -Filter "*.doc"
$pdfFolder = "C:\pdf"

ForEach ($docFile in $docFiles) {
    $doc = $wordApp.Documents.Open($docFile.FullName)
    $pdfFilePath = [System.IO.Path]::Combine($pdfFolder, [System.IO.Path]::ChangeExtension($docFile.Name, ".pdf"))
    $doc.SaveAs([ref] $pdfFilePath, [ref] 17)
    $doc.Close()
}

$wordApp.Quit()
