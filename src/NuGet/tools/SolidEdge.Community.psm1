$regasm32 = 'C:\Windows\Microsoft.NET\Framework\v4.0.30319\RegAsm.exe'
$regasm64 = 'C:\Windows\Microsoft.NET\Framework64\v4.0.30319\RegAsm.exe'

function GetAssemblyPath
{
    $project = (Get-Project)
    $localPath = $project.Properties.Item("LocalPath").Value
    $outputPath = $project.ConfigurationManager.ActiveConfiguration.Properties.Item("OutputPath").Value
    $outputFileName = $project.Properties.Item("OutputFileName").Value
    return Join-Path ( Join-Path ($localPath) $outputPath) $outputFileName
}

function RegisterAddIn
{
    param([string]$regasm, [string]$assembly, [bool]$register)
    
    $al = ''

    if ($register -eq $true)
    {
        $al = "/codebase " + '"' + "$($assembly)" + '"'
    }
    else
    {
        $al = "/u " + '"' + "$($assembly)" + '"'
    }

    #http://blog.coretech.dk/rja/capture-output-from-command-line-tools-with-powershell/
    Write-Host "Executing: $($regasm) $($al)"
    Start-Process -FilePath $regasm -verb runas -ArgumentList $al | Wait-Process
}

function Start-SolidEdge
{
    $application = New-Object -Com SolidEdge.Application
    $application.Visible = $true
    $application.Activate()
}

function Stop-SolidEdge
{
    try
    {
        $application = [System.Runtime.InteropServices.Marshal]::GetActiveObject('SolidEdge.Application')
        $application.Quit()
    }
    catch [System.Exception]
    {
    }
}

function Register-SolidEdgeAddIn
{
    $assembly = GetAssemblyPath
    
    RegisterAddIn $regasm32 $assembly $true
    RegisterAddIn $regasm64 $assembly $true
}

function Unregister-SolidEdgeAddIn
{
    $assembly = GetAssemblyPath
        
    RegisterAddIn $regasm32 $assembly $false
    RegisterAddIn $regasm64 $assembly $false
}

function Install-SolidEdgeAddInRibbonSchema
{
    $project = (Get-Project)
    $projectDir = $project.Properties.Item("LocalPath").Value
    
    $xsdSource = (Join-Path $PSScriptRoot 'Ribbon.xsd')
    $xsdDestination = (Join-Path $projectDir 'Ribbon.xsd')
    
    Copy-Item $xsdSource $xsdDestination -Force
    
    $project.ProjectItems.AddFromFile($xsdDestination) | Out-Null
    $project.Save()
                
    #Write-Host "projectDir: $($projectDir)"
    #Write-Host "PSScriptRoot: $($PSScriptRoot)"
    #Write-Host "xsdSource: $($xsdSource)"
    Write-Host "Added $($xsdDestination) to project."
}

function Set-DebugSolidEdge
{
	$edgePath = ''
	
    try
    {
		$installData = New-Object -Com SolidEdge.InstallData
		$installPath = $installData.GetInstalledPath()
		$edgePath = Join-Path $installPath 'Edge.exe'
		Write-Host "Detected $edgePath"
    }
    catch [System.Exception]
    {
		Write-Host 'Exception while querying Solid Edge install information.'
    }

	if (Test-Path -Path $edgePath)
	{
		$project = (Get-Project)
		$projectName = $project.Name
		$activeConfiguration = $project.ConfigurationManager.ActiveConfiguration	
		#[VSLangProj.prjStartAction]::prjStartActionProgram
		$activeConfiguration.Properties.Item("StartAction").Value = 1
		$activeConfiguration.Properties.Item("StartProgram").Value = $edgePath
		Write-Host "$projectName modified to debug $edgePath."
	}
}

Export-ModuleMember Register-SolidEdgeAddIn
Export-ModuleMember Unregister-SolidEdgeAddIn
Export-ModuleMember Start-SolidEdge
Export-ModuleMember Stop-SolidEdge
Export-ModuleMember Install-SolidEdgeAddInRibbonSchema
Export-ModuleMember Set-DebugSolidEdge