<#
    Script  : ImportPSModule
    Author  : Riwut Libinuko
    Blog    : http://blog.libinuko.com 
    Copyright :© IdeasFree - cakriwut@gmail.com, 2011. All rights reserved.
#>

<# 
   Installation and Configuration
   1. Edit profile.ps1 to load ImportPSModule on start.
   2. You can change $globalPkgSource to local repository if necessary.
   3. Publish your PowerShell Module as NuGet package, tag with at least PSModule.
#>
$globalPkgSource = "https://go.microsoft.com/fwlink/?LinkID=230477"

#Auto configuration
$basePath = (split-path -parent $MyInvocation.MyCommand.Definition)
$globalPkgInstallationDir = Join-Path $basePath "Modules"
$globalPkgDistDir = Join-Path $env:temp "Distribution"  

$userModulePath = $env:PSModulePath -split ";"
$webClient = New-Object System.Net.WebClient 
$webClient.UseDefaultCredentials = $true
$webClient.Proxy.Credentials = $webClient.Credentials
#End autoconfiguration

function GetNugetPath
{
  $nugetPath = (gci $basePath -recurse -filter "nuget.exe").FullName
  
  if(!$nugetPath)
  {
     $installDir = $null
     if($matches) {
       $matches.Clear()
     }
     if(($isModule = $userModulePath |? { $basePath.StartsWith($_) } ))     
     {     
        if(!($nugetPath = (gci $userModulePath[0] -recurse -filter "nuget.exe" | sort LastWriteTime -Descending | select -First 1).FullName))
        {
           $installDir = $userModulePath[0] # Nuget.exe can not run from within GlobalModule. A bug or design limitation??
        } else {        
           return $nugetPath
        }
     } else {
         $installDir = $basePath        
     }
     
     if(($nugetPkg = GetPackage -PackageName "Nuget.CommandLine" -Source $globalPkgSource -distFolder $installDir))
     {
         $zipName = [IO.Path]::ChangeExtension( $nugetPkg, ".zip" )
         gi $nugetPkg | move-item -destination $zipName -force              
         Unzip $zipName (join-path $installDir (gi $zipName).BaseName)
         remove-item $zipName -force |out-null
         $nugetPath = (gci $installDir -recurse -filter "nuget.exe" | sort LastWriteTime -Descending | select -First 1).FullName                  
         if(!$nugetPath)
         {
           Write-Host "Can not find mandatory Nuget.exe"           
         } else 
         {
            return $nugetPath
         }          
     }
  }  
  return $nugetPath
}

function Unzip($zipFile, $dest)
{    
    new-item $dest -ItemType Directory -force | out-null
    $shellApp = New-Object -Com Shell.Application      
    $ZipFileRef = $shellApp.namespace([String]$zipFile)             
    $destination = $shellApp.namespace($dest)             
    $destination.Copyhere($ZipFileRef.items())
 }

function Import-PSModule
{ 
<#
.SYNOPSIS
Import-PSModule from central NuGet repository into memory. 
		
.DESCRIPTION
The function import and load PowerShell module from NuGet repository. It can register the new module to current/global user profile, 
so that the module will be available on any future session.
    
.INPUTS
None. You can not pipe objects to Import-PSModule.
	
.OUTPUTS
None. Operation status,new module name and exported commands are displayed in the screen.
		
.PARAMETER PackageName
MANDATORY parameter to specify the PackageName. For example RemoteStsAdm. 
	
.PARAMETER Source
OPTIONAL parameter to specify NuGet repository URL. Default: https://go.microsoft.com/fwlink/?LinkID=206669. 
You can override this value by specifying global variable $globalPkgSource of Import-PSModule. 
        
.PARAMETER Version
OPTIONAL parameter to load PS Module from specific Package version. This value will be ignored if Latest = TRUE.

.PARAMETER Latest
OPTIONAL parameter to load PS Module from latest Package version.

.PARAMETER Tags
OPTIONAL parameter to identify special NuGet package tags. Default : PSModule.

.PARAMETER Install
OPTIONAL parameter to specify install/un-install operation. Default : TRUE.
			

.EXAMPLE
	PS>  Import-PSModule RemoteStsAdm 
			

	Description
	-----------
	Download, extract and register PS Module from RemoteStsAdm package stored in default NuGet repository.
		
.EXAMPLE
	PS> Import-PSModule -PackageName RemoteStsAdm -Source  "http://code.contoso.com/nuget"


	Description
	-----------
	Download, extract and register PS Module from RemoteStsAdm package stored in http://http://code.contoso.com/nuget.
            
.LINK
    Author blog  : IdeasForFree  (http://blog.libinuko.com)
.LINK
    Author email : cakriwut@gmail.com
#> 
   param(
      [Parameter(Mandatory=$true,Position=0)]
      [string] $packageName,      
      [Parameter(Mandatory=$false,Position=1)]
      [string] $source= $globalPkgSource,
      [string] $version =$null,
      [bool] $latest= $true,
      [string] $tags='PSModule',
      [bool] $install = $true
   )

   if((test-path $source))
   {
      gci $source | remove-item -recurse -force
   }
   #Get feed url
   $serviceBase = GetPackageUrl $source
   $feedUrl = $serviceBase + "Packages"
     
   
   if((GetPackage -PackageName $packageName -Source $source -distFolder $globalPkgDistDir -Latest $latest -Version $version -ExtraFilter " substringof(`'$tags`',Tags)"))
   {   
     #call InstallPkg , from local distribution to avoid proxy auth problem     
     if(($pkgLocation = InstallPkg -packageName $packageName -nodeps $false -Source $globalPkgDistDir))
     {        
       #in the installation folder, find PSModuleInstall.ps1 / PSModuleUninstall.ps1
       if($install)
       {         
         if(($ps1 = gci $pkgLocation -recurse |? { $_.name -match "PSModuleInstall\.ps1" } | sort name -Descending | select -First 1))
         {
            Invoke-Expression "$($ps1.FullName)"            
         } 
         
         write-host "Trying to automatically load all module files."
         $pkgLocation |% {
            RegisterModule -ModuleDirectory $_
         }                  
         Write-host ""
         Write-host "$packageName install operation has been completed."
         
       } else {
         if(($ps1 = gci $pkgLocation -recurse |? { $_.name -match "PSModuleUnInstall\.ps1" } | sort name -Descending | select -First 1))
         {
            Invoke-Espression "$($ps1.FullName)"            
         } 
         
         UnregisterModule -ModuleDirectory $pkgLocation -removeAll $true         
         split-path $pkgLocation | gci |? { $_ -match "$packageName(\.\d)*$" } |% {
             Write-Host "..Removing installation history $($_.FullName)"
             remove-item $_.FullName -recurse -force #| out-null 
        }
         Write-Host "$packageName has been removed."
         Write-Host "All dependencies item are not removed. You may remove it manually if no other component are using it."
       }
       
     } else{
        write-host "Can not find $packageName in the $source. Ensures both package and dependencies are exists." -foregroundcolor red
     }
   }
      
}


function RegisterModule
{
<#
.SYNOPSIS
Locate and register all PS Module from specific path. 
		
.DESCRIPTION
Registering all PS Module from given path.
    
.INPUTS
None. You can not pipe objects to Import-PSModule.
	
.OUTPUTS
None. Operation status,new module name and exported commands are displayed in the screen.
		
.PARAMETER ModuleDirectory
MANDATORY parameter in which all PS Module in ModuleDirectory will be loaded (and/or registered). 			

.EXAMPLE
	PS>  RegisterModule "D:\LocalModule"
			

	Description
	-----------
	Register all PS Module in D:\LocalModule and its sub folder.
		            
.LINK
    Author blog  : IdeasForFree  (http://blog.libinuko.com)
.LINK
    Author email : cakriwut@gmail.com
#> 

   param (
      [string] $ModuleDirectory      
   )
                             
   if(!($moduleLists = GetAllModuleFiles $ModuleDirectory))
   {
      Write-Warning "Can not find any modules. Check if PowerShell module exists in $ModuleDirectory"
      return 
   }

   $package = GetPackageInfo $ModuleDirectory
   
   $single = new-object System.Management.Automation.Host.ChoiceDescription "Once","Load module for this session only"
   $personal = new-object System.Management.Automation.Host.ChoiceDescription "Personal","Load and register module in PowerShell current user profile"
  # $global = new-object System.Management.Automation.Host.ChoiceDescription "Global","Load and register module in PowerShell ALL user profile"
   $options = [System.Management.Automation.Host.ChoiceDescription[]]($single,$personal) #,$global)
   
   $selectedStatus = $host.ui.PromptForChoice("New-PSModule Registration","Do you want to register the module?",$options,1)
     
   $moduleLists |% {     
      if(get-module -name $_.ModuleName)
      {
          remove-module -name $_.ModuleName
       }
       Write-Host "..Loading module $($_.ModuleName)" -NoNewLine
       Import-Module -name $_.ModulePath -global
       Write-Host "..Sucessful." -foregroundcolor DarkGreen
       Get-Module $_.ModuleName
   } 
   
   if($selectedStatus -eq 0)
   {  
     # No update required      
   } elseif ($selectedStatus -eq 1)
   {
      Write-Host "Register personal $($package.Name) modules"
      UpdateProfile -profilePath $profile.CurrentUserAllHosts -PSModuleName $package.Name -PSModuleVer $package.Version -Content $moduleLists
      UpdateProfile -profilePath $profile.CurrentUserCurrentHost -PSModuleName $package.Name -PSModuleVer $package.Version -Content $moduleLists  
                 
   } else {
      Write-Host "Register global $($package.Name) modules"
      #UpdateProfile -profilePath $profile.AllUsersAllHosts -PSModuleName $package.Name -PSModuleVer $package.Version -Content $moduleLists
      #UpdateProfile -profilePath $profile.AllUsersCurrentHost -PSModuleName $package.Name -PSModuleVer $package.Version -Content $moduleLists  
   }
}

function GetAllModuleFiles
{
    param (
      [string] $ModuleDirectory
    )
        
    return ( gci $ModuleDirectory -recurse -filter "*.psm" |% {
               new-object PSObject -prop @{ 
                       ModuleName=$_.BaseName; 
                       ModulePath=(join-path $_.DirectoryName $_.BaseName) }
              })
}

function GetPackageInfo
{
   param (
     [string] $ModuleDirectory
   )
   
   $pkgLocation = gi $ModuleDirectory
   return  ( $pkgLocation.Name -match "(?<PkgName>.+[^(\.\d)])\.(?<PkgVer>(\d\.){1,}(\d))" |% { 
                        new-object PSObject -prop @{
                           Name =($matches["PkgName"]).Trim();
                           Version=($matches["PkgVer"]).Trim()} 
                        })
}

function UnregisterModule
{
   param (
      [string] $ModuleDirectory,
      [bool] $removeAll=$false     
   )
                            
   if(!($moduleLists = GetAllModuleFiles $ModuleDirectory))
   {
      Write-host "Can not find any modules. Check if PowerShell module exists in $ModuleDirectory"
      return 
   }
      
   $package = GetPackageInfo $ModuleDirectory 
   $moduleLists |% {     
      if(get-module -name $_.ModuleName)
      {
          remove-module -name $_.ModuleName
          Write-Host "..Removing module $($_.ModuleName)"
      }
   } 
   
   if(!$removeAll)
   {
     $personal = new-object System.Management.Automation.Host.ChoiceDescription "Personal","Remove module in PowerShell current user profile"
     $global = new-object System.Management.Automation.Host.ChoiceDescription "Global","Remove module in PowerShell ALL user profile"
     $options = [System.Management.Automation.Host.ChoiceDescription[]]($personal,$global)
   
     $selectedStatus = $host.ui.PromptForChoice("New-PSModule UnRegistration","Do you want to unregister the module from profile?",$options,0)
   }
        
   if (($selectedStatus -eq 0) -or $removeAll)
   {
      Write-Host "Removing from current user profile. $($package.Name) modules"
      RemoveProfile -profilePath $profile.CurrentUserAllHosts -PSModuleName $package.Name 
      RemoveProfile -profilePath $profile.CurrentUserCurrentHost -PSModuleName $package.Name  
                 
   } 
   if (($selectedStatus -eq 1) -or $removeAll)
   {
      Write-Host "Removing from all user profile. $($package.Name) modules"
      RemoveProfile -profilePath $profile.AllUsersAllHosts -PSModuleName $package.Name
      RemoveProfile -profilePath $profile.AllUsersCurrentHost -PSModuleName $package.Name   
   }
   
   Write-Host "Unregistering $($package.Name) has been completed."
}

function PreparePath
{
    param (
        [string] $path
    )
    
    if(!(test-path $path))
    {
       new-item -type directory -path $path -force     
    }
    return $path
}

function GetPackage
{
    param(
        [string] $packageName,        
        [string] $source= $globalPkgSource,
        [string] $distFolder = $globalPkgDistDir,
        [bool] $latest=$true,
        [string] $version,
        [string] $extrafilter=$null        
    )
    
    # set up feed URL    
    $serviceBase = GetPackageUrl $source
    $feedUrl = $serviceBase + "Packages"

    if($latest) {
        $feedUrl = $feedUrl +"?`$filter=(IsLatestVersion eq true) and (Id eq `'$packageName`')"
    } else {
        $feedUrl = $feedUrl + "?`$filter=(Version eq $version) and (Id eq `'$packageName`')"
    }
    
    if($extrafilter)
    {
       $feedUrl = $feedUrl + " and $extrafilter"
    }

    PreparePath $distFolder | Out-Null
    Write-Host "Download package $packageName" -NoNewLine
    DownloadEntries $feedUrl $distFolder | Out-Null
    if(($pkgPath = gci $distFolder |? { $_.Name -match "$packageName\." } | sort LastWriteTime -Descending | select -First 1))
    {
       write-host "..Sucessful" -foreground DarkGreen
       return $pkgPath.FullName
    } else {
       write-host "..Unsucessful." -foreground red
       write-host "..Error: Can not download package $packageName" -foreground red
       return $null
    }
}


function GetPackageUrl 
{  
   param ([string]$source)  
   $resp = [xml]$webClient.DownloadString($source)  
   return $resp.service.GetAttribute("xml:base") 
}

function RemoveProfile
{
    param (
      [string] $profilePath,
      [string] $PSModuleName
    )
    
    if(!(test-path $profilePath))
    {
       Write-Host "Profile does not exists!"
       return
    }
    
    $profileContent = gc $profilePath
   if($profilecontent |? { $_ -match "\#PSMODULE\:$PSModuleName(\.\d)*$" } )
   {
     Write-Host "Removing profile entries..." -NoNewLine
     $profileContent |% {
        if($_ -match "\#PSMODULE\:$PSModuleName(\.\d)*$")
        {
            $contentStart = $true
            # Remove package registration
        } elseif(($_ -match "\#PSMODULE") -and $contentStart)
        {
            $contentStart = $false
            # Remove package registration
        } elseif($contentStart)
        {
           #Remove existing entries in the block
        } else {
            "$_"
        }
      } | Set-Content $profilePath
      Write-Host "Completed"
   } else {
     Write-Host "Package $PSModuleName has not been registered to profile!"
     return
   }
}

function UpdateProfile
{
   param (
      [string] $profilePath,
      [string] $PSModuleName,
      [string] $PSModuleVer,
      $content
   )

   if(!(test-path $profilePath))
   {
      new-item -type file -path $profilePath -force | Out-Null
   }
   if($PSModuleVer)
   {
      $PSModuleVer = ".$PSModuleVer"
   }
   
   $profileContent = gc $profilePath
   if($profilecontent |? { $_ -match "#PSMODULE\:$PSModuleName(\.\d)*$" } )
   {
      Write-Host "Modifying existing profile entries"
      $profileContent |% {
        if($_ -match "\#PSMODULE\:$PSModuleName(\.\d)*$")
        {
            $contentStart = $true
            "`n#PSMODULE:$PSModuleName$PSModuleVer"
            $content |% {
               "Import-Module -name `"$($_.ModulePath)`""
            }            
        } elseif(($_ -match "\#PSMODULE") -and $contentStart)
        {
            $contentStart = $false
            "$_`n"
        } elseif($contentStart)
        {
           #Remove existing entries in the block
        } else {
            if($_) { 
               "$_"
            }
        }
      } | Set-Content $profilePath
      
   } else {
      
      $newScript = "`n#PSMODULE:$PSModuleName$PSModuleVer" 
      $newScript += $content |% {
             "`nImport-Module -name `"$($_.ModulePath)`""
          } 
      $newScript += "`n#PSMODULE" 
      
      $newScript | Add-Content $profilePath
   }
   
}

function InstallPkg
{
<#
.SYNOPSIS
Install package from specific location. 
		
.DESCRIPTION
Instal package from specific location.
    
.INPUTS
None. You can not pipe objects to InstallPkg.
	
.OUTPUTS
None. 
		
.PARAMETER PackageName
MANDATORY parameter . 			

.PARAMETER Source
MANDATORY parameter .

.PARAMETER InstallDir
MANDATORY parameter .


.EXAMPLE
	PS>  InstallPkg RemoteStsAdm "D:\LocalModule"
			

	Description
	-----------
	Install RemoteStsAdm from D:\LocalModule.
		            
.LINK
    Author blog  : IdeasForFree  (http://blog.libinuko.com)
.LINK
    Author email : cakriwut@gmail.com
#> 
    param (
       [Parameter(Mandatory=$true,Position=0)]
       [string] $packageName,
       [Parameter(Mandatory=$false,Position=1)]
       [string] $source = $globalPkgSource,
       [string] $installDir = $globalPkgInstallationDir,
       [bool] $nodeps = $false
    )
    write-host "Trying to extract package $packageName" -NoNewLine
    
    #Nuget.exe compensation. Can not install/read Global PSModule path
    if($installDir.StartsWith($userModulePath[1]))
    {
       $installDir = $installDir.Replace($userModulePath[1].TrimEnd("\"),$userModulePath[0].TrimEnd("\"))
    }
    
    $nugetExe = GetNugetPath   
    $nugetArgs = "install $packageName -source `"$source`" -OutputDirectory `"$installDir`""
    $result = invoke-expression "& `"$nugetExe`" $nugetArgs 2>&1"     
    $pattern = "Unable to resolve dependency \'(?<Pkg>.+)\s?\(\≥(?<Ver>.+)\)\'"
    $unresolved = $result |? {
                      $_ -match $pattern } |% { 
                           new-object PSObject -prop @{ Package=($matches['Pkg']).Trim(); Version=($matches["Ver"]).Trim() }}                               
    $localSource = ( $source |? { $_ -match "(http|https)\:\/\/.+?" }) -eq $null    
    
    if($unresolved -eq $null)
    {
      if(($result |? { $_ -match "Unable to find package \'(.+?)\'" }))
      {
         write-host "..Unsucessful." -foregroundcolor red
         write-host "..Error: $result" -foregroundcolor red         
      } else {
         write-host "..Successful." -foregroundcolor DarkGreen
         $pkgLocation = gci $installDir |? { $_.Name -match "$packageName\.(\d\.){1,}(\d)" } | sort name -Descending | select -First 1
         return (join-path $installDir $pkgLocation)
      }
    } elseif ($nodeps)
    {
      write-host "..Unsucessful."  -foregroundcolor red 
      write-host "..Error: $result" -foregroundcolor red      
    } else {       
      $unresolved | foreach { 
        write-host ""
        write-host "..Downloading dependencies $($_.Package) ($($_.Version) or later)"
        $depPkg=$null
        # Call to download pkg
        if($localSource)
        {
           $depPkg = GetPackage -Packagename $_.Package -distFolder $source
        } else {
           # This rarely happened because Nuget will handle itself.
           $depPkg = GetPackage -Packagename $_.Package
        }
        # Call to install pkg
        if($depPkg)
        {
           InstallPkg -PackageName $_.Package -Source $source   
           $nodeps = $false         
        } else {
           $nodeps = $true
        }
      }
      #Final to install current package. $nodeps
      InstallPkg -PackageName $packageName -Nodeps $nodeps -Source $source
    }
}

function DownloadEntries 
{  
  param (
      [string]$feedUrl,
      $destinationDirectory
  )    
  $feed = [xml]$webClient.DownloadString($feedUrl) 
  $entries = $feed | select -ExpandProperty feed | select -ExpandProperty entry -ErrorAction SilentlyContinue  
  if(!$entries)
  {      
     return "Can not find any matched packages."
  }
  $progress = 0                
  foreach ($entry in $entries) 
  {     
    $url = $entry.content.src   
    if($entry.properties.id)
    {  
      $fileName = $entry.properties.id + "." + $entry.properties.version + ".nupkg"    
    } else {
      $fileName = $entry.title."#text" + "." + $entry.properties.version + ".nupkg" 
    }
    
    $saveFileName = join-path $destinationDirectory $fileName       
    $pagepercent = ((++$progress)/@($entries).Length*100)     
    if ((-not $overwrite) -and (Test-Path -path $saveFileName))     
    {         
        write-progress -activity "$fileName already downloaded, using cached file" -status "$pagepercent% of current page complete" -percentcomplete $pagepercent        
        continue     
    }     
    write-progress -activity "Downloading $fileName" -status "$pagepercent% of current page complete" -percentcomplete $pagepercent    
    $webClient.DownloadFile($url, $saveFileName)   
  }   
  
  $link = $feed.feed.link | where { $_.rel.startsWith("next") } | select href   
  if ($link -ne $null) 
  {         
    $feedUrl = $link.href     
    DownloadEntries $feedUrl $destinationDirectory
  } 
} 

GetNugetPath | out-null