 <#

.Synopsis
    Office suite installation during OOBE

.DESCRIPTION
     Extract the ZIP to specified TEMP folder and run the Office setup.exe 
     with the configuration file as a parameter during the initial setup 
     of Windows 10 device. 

     Original work https://docs.microsoft.com/cs-cz/archive/blogs/beanexpert/how-to-install-office-using-a-provisioning-package
     which has tobe modified, as it´s using DeviceContext provisioning commands. 

.AUTHOR
    Jan Čížek
    j.cizek@hotmail.com

#>
[CmdletBinding()]
[Alias()]
[OutputType([int])]

Param
(
 [Parameter(Mandatory=$false,
 ValueFromPipelineByPropertyName=$true,
 Position=0)]
 $Log = "$env:windir\debug\Start-ProvisioningCommands.log"
)

Begin{

    # Start logging
    Start-Transcript -Path $Log -Force -ErrorAction SilentlyContinue

    #Specify TEMP folder and create
    $path = "c:\TEMP\O2019"
    new-item $path -ItemType Directory 

    #For logging purposes in Transcript - uncomment
    write-host "ARCHIVES directory: " $PSScriptRoot
    write-host "TARGET directory: " $path 

    #Load all archives...
    $Archives = Get-ChildItem -Path $PSScriptRoot -Filter *.zip | Select-Object -Property FullName
    #For logging purposes in Transcript - uncomment
    write-host "ARCHIVES list: " $Archives.Fullname

    #and UNZIP
    ForEach-Object -InputObject $Archives -Process { Expand-Archive -Path $_.FullName -DestinationPath $path -Force}

    }

Process {
    #specify the TEMP folder you used in Begin region
    $WorkingDirectory = "c:\TEMP\O2019"

    #Update Configuration.XML with the TEMP folder as source
    $Configuration = Get-ChildItem -Path $WorkingDirectory -Filter *.xml | Select-Object -Property FullName
    [XML]$XML = Get-Content -Path $Configuration.FullName

    write-host "writing sourcepath attribute to XML... " 
    $XML.Configuration.Add.SourcePath = $WorkingDirectory
    $XML.Save($Configuration.FullName)


    $setupexe = Get-ChildItem -Path $WorkingDirectory -Filter *.exe

    #For logging purposes in Transcript - uncomment
    write-host "WORKING directory: " $WorkingDirectory
    write-host "XML fullName: " $XML.fullname
    write-host "SETUP fullname: " 
     
    # Run Office 2016 setup.exe
    write-host "Starting Office 365 installation..." -NoNewline

    try {

        Set-Location $WorkingDirectory

        #Start-Process -FilePath .\setup.exe `
        #-ArgumentList ('/Configure "{0}"' -f $Configuration.FullName)  `
        #-WorkingDirectory $WorkingDirectory  `
        #-Wait -WindowStyle Hidden 

        cmd /c "setup.exe  /Configure Configuration.xml"

        write-host " DONE" 
        }

    catch{ write-host "Office installation failed"}

    # Cleanup Installation
    write-host "Cleaning installation data..."
    Remove-Item -Path $WorkingDirectory -Force -Confirm:$false
     write-host "TEMP cleared..."
    }

end {
    # Stop logging
    write-host "Stopping transcript..."
    Stop-Transcript -ErrorAction SilentlyContinue
     }








