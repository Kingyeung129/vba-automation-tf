@"
This powershell script attempts to import all *.bas modules in the same working directory into VBA project in Personal Macro Workbook.
If module of the same name exists, it will be overwritten by first removing then importing the new VBA code.

Pre-requisites:
1. Excel must be installed of course.
1. Ensure that 'Trust access to the VBA project object model' is enabled in Excel options.
"@

# Get all .bas files in the script's directory
$scriptDirectory = Split-Path -Parent $MyInvocation.MyCommand.Definition
$basFiles = Get-ChildItem -Path $scriptDirectory -Filter *.bas

try {
    # Create a new Excel application instance and Open the Personal Macro Workbook
    $excel = New-Object -ComObject Excel.Application
    $personalWorkbook = $excel.Workbooks.Open($excel.Application.StartupPath + "\Personal.xlsb")

    # Get the VBProject
    $vbProject = $personalWorkbook.VBProject
    if (-not $vbProject) {
        throw "Failed to access the VBProject. Ensure that 'Trust access to the VBA project object model' is enabled in Excel options."
    }

    # Import all .bas files in script's directory into the Personal Macro Workbook
    foreach ($file in $basFiles) {
        $vbaModulePath = $file.FullName
        $moduleName = [System.IO.Path]::GetFileNameWithoutExtension($file.Name)
        $vbComponents = $vbProject.VBComponents

        # Check if the module already exists and remove it if found
        foreach ($vbComponent in $vbComponents) {
            if ($vbComponent.Name -eq $moduleName) {
                $vbComponents.Remove($vbComponent)
                break
            }
        }

        # Import the new module
        $vbComponents.Import($vbaModulePath)
    }

    # Save the Personal Macro Workbook and quit Excel
    $personalWorkbook.Save()
    $personalWorkbook.Close($true)
    $excel.Quit()
} catch {
    Write-Error "An error occurred: $_"
} finally {
    # Release the COM objects to free up resources and Garbage Collection
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($vbComponents) | Out-Null
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($vbProject) | Out-Null
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($personalWorkbook) | Out-Null
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}