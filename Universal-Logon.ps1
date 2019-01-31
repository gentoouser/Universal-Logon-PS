#region Global Variables
############################ Start Global Variables ##########################################
$uiHash = [hashtable]::Synchronized(@{})
$GlobalHash = [hashtable]::Synchronized(@{})
$runspaceHash = [hashtable]::Synchronized(@{})
$jobs = [system.collections.arraylist]::Synchronized((New-Object System.Collections.Arraylist))
$GlobalHash.ScriptVersion = 1.0.0
$GlobalHash.PrintersScript=@()
$GlobalHash.LocalPrinter=@()
$GlobalHash.UserID = $env:username
$GlobalHash.RealWorkstation = $env:computername
$GlobalHash.AppData = [Environment]::GetFolderPath("ApplicationData")
$GlobalHash.ProgramFiles = [Environment]::GetFolderPath("ProgramFiles")
If ([environment]::Is64BitOperatingSystem) {
    $GlobalHash.ProgramFilesX86 = [Environment]::GetFolderPath("ProgramFilesX86")
    $GlobalHash.AddressWidth = 64
} else {
    $GlobalHash.AddressWidth = 32
}
$GlobalHash.WindowsDir = [Environment]::GetFolderPath("Windows")
$GlobalHash.CommonStartMenu = [Environment]::GetFolderPath("CommonStartMenu")
$GlobalHash.StartMenu = [Environment]::GetFolderPath("StartMenu")
$GlobalHash.UserProfile = [Environment]::GetFolderPath("UserProfile")
$GlobalHash.DateYMD = Get-Date -format yyyyMMdd
$GlobalHash.OSVersionMajor = [environment]::OSVersion.Version.Major
$GlobalHash.OSVersionMinor = [environment]::OSVersion.Version.Minor
# Check and set of clientname is used
If ($env:clientname) {
    $GlobalHash.Workstation = $env:clientname
}else {
    $GlobalHash.Workstation = $GlobalHash.RealWorkstation
}
# Set status of Printer Spooler is running. 
If ((Get-Service -Name Spooler).Status -eq "Running") {$GlobalHash.PrintSpooler = $True}
############################ End Global Variables ##########################################
#endregion
#region Global User Settings
############################ Start Global User Settings ##########################################
$GlobalHash.CompanyName = "Example Inc." 
$GlobalHash.FileServer = "fs.example.com"
$GlobalHash.FileServerHome = "fshome.example.com"
$GlobalHash.PrintServer = "ps.example.com"
$GlobalHash.PrintGroupsOU = "OU=Printers,OU=Groups,DC=example,DC=com"
$GlobalHash.Icon = "http://www.example.com/logos/example.ico"
$GlobalHash.Logo = "http://www.example.com/logos/example.png"
$GlobalHash.LogoHeight = 120
$GlobalHash.LogoWidth = 350
$GlobalHash.ExemptWorkstation = 0
$GlobalHash.UserHomeShare =  ("\\" + $GlobalHash.FileServerHome + "\users$\" + $GlobalHash.UserID)
$GlobalHash.PrintSpooler = $False
$GlobalHash.ForcePrinter= "PDFCREATOR"
$GlobalHash.StrPaperVisionWebAssistant = "C:\Program Files (x86)\Digitech Systems\PaperVision\PVWA\DSI.PVWA.Host.exe"
$GlobalHash.ChangeDefault = $True
$GlobalHash.LogUNC="\\wwt-fshome.example.com\logs$\sessions_csv"
$GlobalHash.CompanyDefaultWallpaper = ($env:SystemRoot + "\system32\oobe\info\backgrounds\background1920x1200.jpg")
$GlobalHash.CompanyDefaultWallpaperStyle="2"
$GlobalHash.StrContact = "Please contact Support Desk at: support@example.com"
If ([environment]::Is64BitOperatingSystem) {
    $GlobalHash.ShLibExe = ([Environment]::GetFolderPath("ProgramFilesX86") + "\SysinternalsSuite\ShLib.exe")
}else{
    $GlobalHash.ShLibExe = ([Environment]::GetFolderPath("ProgramFiles") + "\SysinternalsSuite\ShLib.exe")
}
# Local printers to ignore
$GlobalHash.PrinterLocalExclude = @{
     "SHRFAX"="FAX";
	 "WEBEX DOCUMENT LOADER PORT" = "WEBEX DOCUMENT LOADER";
	 "PDFCREATOR" = "PDFCMON" ;
	 "MICROSOFT PRINT TO PDF" ="PORTPROMPT";
	 "MICROSOFT XPS DOCUMENT WRITER" ="XPSPORT";
	 "FAX" = "FAX";
	 "WEBEX DOCUMENT LOADER" = "WEBEX DOCUMENT LOADER PORT";
	 "SEND TO ONENOTE 2010" = "NUL";
	 "SEND TO ONENOTE 2013" = "NUL";
	 "SEND TO ONENOTE 16" = "NUL";
	 "SEND TO ONENOTE 2016" = "NUL";
	 "CANON GENERIC FAX DRIVER (FAX)" = "CANON GENERIC FAX DRIVER (FAX)";
	 "HP EPRINT" = "HP EPRINT";
     "Adobe PDF" = "DOCUMENTS\*.PDF";
     "CutePDF Writer" = "CPW3";
     "PDF-XCHANGE5" = "PDF-XCHANGE PRINTER 2012";
}
############################ End Global User Settings ##########################################
#endregion
#region Create GUI
$uiHash.jobFlag = $True
$newRunspace =[runspacefactory]::CreateRunspace()
$newRunspace.ApartmentState = "STA"
$newRunspace.ThreadOptions = "ReuseThread"          
$newRunspace.Open()
$newRunspace.SessionStateProxy.SetVariable("GlobalHash",$GlobalHash)  
$newRunspace.SessionStateProxy.SetVariable("uiHash",$uiHash)         
$newRunspace.SessionStateProxy.SetVariable("runspaceHash",$runspaceHash)     
$newRunspace.SessionStateProxy.SetVariable("jobs",$jobs)     
$psCmd = [PowerShell]::Create().AddScript({  
    #Build the GUI
    Add-Type –assemblyName PresentationFramework
    Add-Type –assemblyName PresentationCore
    Add-Type –assemblyName WindowsBase  
    [void][System.Reflection.Assembly]::LoadWithPartialName('presentationframework')
    [xml]$xaml = @"
    <Window 
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        x:Name="Window" Title="Initial Window" WindowStartupLocation = "CenterScreen" ResizeMode="NoResize"
        Width = "500" Height = "625" ShowInTaskbar = "True" Background = "lightgray"> 
         <Window.TaskbarItemInfo>
                <TaskbarItemInfo/>
        </Window.TaskbarItemInfo>
        <StackPanel >  
            <Image Name = "Logo"/>
            <ScrollViewer x:Name = "scrollviewer" VerticalScrollBarVisibility="Visible"  Height="365">    
                <TextBlock x:Name = 'textblock' TextWrapping = "Wrap" />
            </ScrollViewer >            
        </StackPanel>
    </Window>
"@
    $reader=(New-Object System.Xml.XmlNodeReader $xaml)
    $uiHash.Window=[Windows.Markup.XamlReader]::Load( $reader )
    $uiHash.Window.Title = ($GlobalHash.CompanyName + " - Loading ...")


    $uiHash.Window.Icon = $GlobalHash.Icon
    $uiHash.Window.TaskbarItemInfo.Overlay = $GlobalHash.Icon
    $uiHash.Window.TaskbarItemInfo.Description = $uiHash.Window.Title
    $uiHash.Window.FindName('Logo').Source=$GlobalHash.Logo
    $uiHash.Window.FindName('Logo').Width=$GlobalHash.LogoWidth
    $uiHash.Window.FindName('Logo').Height=$GlobalHash.LogoHeight
    $uiHash.Window.FindName('Logo').HorizontalContentAlignment = 'Center'
    $uiHash.Window.FindName('scrollviewer').Height=(625 - $GlobalHash.LogoHeight)
    
    #Connect to Controls
    $uiHash.textblock = $uiHash.Window.FindName('textblock')
    $uiHash.button = $uiHash.Window.FindName('button')
    $uiHash.scrollviewer = $uiHash.Window.FindName('scrollviewer')
#endregion
  #region Main thread Start
    $MainRunspace =[runspacefactory]::CreateRunspace()      
    $MainRunspace.Open()
    $MainRunspace.SessionStateProxy.SetVariable("GlobalHash",$GlobalHash)  
    $MainRunspace.SessionStateProxy.SetVariable("uiHash",$uiHash)         
    $MainRunspace.SessionStateProxy.SetVariable("runspaceHash",$runspaceHash)     
    $MainRunspace.SessionStateProxy.SetVariable("jobs",$jobs)     
    $MainpsCmd = "" | Select-Object PowerShell,Handle
    $MainpsCmd.PowerShell = [PowerShell]::Create().AddScript({  
  #endregion
        #region Main Functions
        ############################ Start Main Functions ##########################################
        function LogtoGUI () {
             param(
                [Parameter(Mandatory=$true)][string]$Message,
                [Parameter(Mandatory=$false)][string]$Color,
                [Parameter(Mandatory=$false)][string]$Postion,
                [Parameter(Mandatory=$false)][string]$FontStyle,
                [Parameter(Mandatory=$false)][string]$FontWeight,
                [Parameter(Mandatory=$false)][string]$FontSize,
                [Parameter(Mandatory=$false)][switch]$NoNewLine
            )
            $uiHash.textblock.Dispatcher.Invoke("Normal",[action]{
                $Run = New-Object System.Windows.Documents.Run
                if ($Color) {
                    $Run.Foreground = $Color
                }else{
                    $Run.Foreground = "Black"
                }
                if ($FontWeight) {
                    $Run.FontWeight = $FontWeight
                }else{
                    $Run.FontWeight = "Normal"
                }
                if ($FontStyle) {
                    $Run.FontStyle = $FontStyle
                }else{
                    $Run.FontStyle = "Normal"
                }
                if ($FontSize) {
                    $Run.FontSize = $FontSize
                }

                $Run.Text = ("{0}" -f  $Message)
                $uiHash.TextBlock.Inlines.Add($Run)
                If (-Not $NoNewLine) {
                    $uiHash.TextBlock.Inlines.Add((New-Object System.Windows.Documents.LineBreak))   
                }
                $Run = $null    
            })

            $uiHash.scrollviewer.Dispatcher.Invoke("Normal",[action]{
                $uiHash.scrollviewer.ScrollToEnd()
            })
        }
        function LogtoGUILine () {
                $uiHash.textblock.Dispatcher.Invoke("Normal",[action]{
                $Run = New-Object System.Windows.Documents.Run
                $Run.Background = "Black"
                #$Run.FontStyle = $FontStyle
                $Run.FontSize = "1"
                $Run.Text = ("-------------------------------------------------------------------------" + `
                "--------------------------------------------------------------------------------------" + `
                "--------------------------------------------------------------------------------------" + `
                "--------------------------------------------------------------------------------------" + `
                "--------------------------------------------------------------------------------------" + `
                "--------------------------------------------------------------------------------------" + `
                "--------------------------------------------------------------------------------------" + `
                "--------------------------------------------------------------------------------------" + `
                "--------------------------------------------------------------------------------------" + `
                "--------------------------------------------------------------------------------------" + `
                "--------------------------------------------------------------------------------------" + `
                "--------------------------------------------------------------------------------------" + `
                "--------------------------------------------------------------------------------------" + `
                "--------------------------------------------------------------------------------------")

                $uiHash.TextBlock.Inlines.Add($Run)
                $uiHash.TextBlock.Inlines.Add((New-Object System.Windows.Documents.LineBreak))   
                $Run = $null    
            })

            $uiHash.scrollviewer.Dispatcher.Invoke("Normal",[action]{
                $uiHash.scrollviewer.ScrollToEnd()
            })
        }
        function CloseGUI () {
            $uiHash.textblock.Dispatcher.Invoke("Normal",[action]{ $uiHash.Window.Close()})
        }
        function Test-is64Bit {
            param($FilePath)
            #Source: https://superuser.com/questions/358434/how-to-check-if-a-binary-is-32-or-64-bit-on-windows
            [int32]$MACHINE_OFFSET = 4
            [int32]$PE_POINTER_OFFSET = 60

            [byte[]]$data = New-Object -TypeName System.Byte[] -ArgumentList 4096
            $stream = New-Object -TypeName System.IO.FileStream -ArgumentList ($FilePath, 'Open', 'Read')
            $stream.Read($data, 0, 4096) | Out-Null

            [int32]$PE_HEADER_ADDR = [System.BitConverter]::ToInt32($data, $PE_POINTER_OFFSET)
            [int32]$machineUint = [System.BitConverter]::ToUInt16($data, $PE_HEADER_ADDR + $MACHINE_OFFSET)

            $result = "" | Select-Object FilePath, FileType, Is64Bit
            $result.FilePath = $FilePath
            $result.Is64Bit = $false

            switch ($machineUint) 
            {
                0      { $result.FileType = 'Native' }
                0x014c { $result.FileType = 'x86' }
                0x0200 { $result.FileType = 'Itanium' }
                0x8664 { $result.FileType = 'x64'; $result.is64Bit = $true; }
            }

            $result
        }
        function GetADObject {
            param([string] $objectName, [string] $objectType)
    
            switch ($objectType)
                {
                "User" {$Filter = "(&(objectCategory=$objectType)(samAccountName=$objectName))"}
                "Group" {$Filter = "(&(objectCategory=$objectType)(samAccountName=$objectName))"}
                "Computer" {$Filter = "(&(objectCategory=$objectType)(Name=$objectName))"}
                default {$Filter = "(&(objectCategory=$objectType)(samAccountName=$objectName))"}
                }
            $objSearcher = New-Object System.DirectoryServices.DirectorySearcher
            $objSearcher.Filter = $Filter
            ($objSearcher.FindOne()).GetDirectoryEntry()
        }
        function GetBrowsers {
               $Temp= @{}
               #Chrome x64
               $TestProg = ($GlobalHash.ProgramFiles + "\Google\Chrome\Application\chrome.exe")
               if (Test-Path $TestProg) {
                    $TempVer = ((Get-Item $TestProg).VersionInfo).ProductVersion
                    $TempBit = (Test-is64Bit -FilePath $TestProg).FileType
                    $Temp.Add(("Chrome-" + $TempBit),@{Version = $TempVer; IntSize = $TempBit;Path = $TestProg; Browser = "Chrome"})
               }
               #Chrome x32
               $TestProg = ($GlobalHash.ProgramFilesX86 + "\Google\Chrome\Application\chrome.exe")
               if (Test-Path $TestProg) {
                    $TempVer = ((Get-Item $TestProg).VersionInfo).ProductVersion
                    $TempBit = (Test-is64Bit -FilePath $TestProg).FileType
                    $Temp.Add(("Chrome-" + $TempBit),@{Version = $TempVer; IntSize = $TempBit;Path = $TestProg; Browser = "Chrome"})
               }
               #Firefox x64
               $TestProg = ($GlobalHash.ProgramFiles + "\Mozilla Firefox\firefox.exe")
               if (Test-Path $TestProg) {
                    $TempVer = ((Get-Item $TestProg).VersionInfo).ProductVersion
                   $TempBit = (Test-is64Bit -FilePath $TestProg).FileType
                    $Temp.Add(("Firefox-" + $TempBit),@{Version = $TempVer; IntSize = $TempBit;Path = $TestProg; Browser = "Firefox"})
               }
               #Firefox x32
               $TestProg = ($GlobalHash.ProgramFilesX86 + "\Mozilla Firefox\firefox.exe")
               if (Test-Path $TestProg) {
                    $TempVer = ((Get-Item $TestProg).VersionInfo).ProductVersion
                    $TempBit = (Test-is64Bit -FilePath $TestProg).FileType
                    $Temp.Add(("Firefox-" + $TempBit),@{Version = $TempVer; IntSize = $TempBit;Path = $TestProg;Browser = "Firefox"})
               }
               #Internet Explorerr x64
               $TestProg = ($GlobalHash.ProgramFiles + "\Internet Explorer\iexplore.exe")
               if (Test-Path $TestProg) {
                    $TempVer = ((Get-Item $TestProg).VersionInfo).ProductVersion
                    $TempBit = (Test-is64Bit -FilePath $TestProg).FileType
                    $Temp.Add(("IE-" + $TempBit),@{Version = $TempVer; IntSize = $TempBit;Path = $TestProg;Browser = "Internet Explorer"})
               }
               #Internet Explorer x32
               $TestProg = ($GlobalHash.ProgramFilesX86 + "\Internet Explorer\iexplore.exe")
               if (Test-Path $TestProg) {
                    $TempVer = ((Get-Item $TestProg).VersionInfo).ProductVersion
                    $TempBit = (Test-is64Bit -FilePath $TestProg).FileType
                    $Temp.Add(("IE-" + $TempBit),@{Version = $TempVer; IntSize = $TempBit;Path = $TestProg;Browser = "Internet Explorer"})
               }
               #Microsoft Edge
               $TestProg = ($GlobalHash.WindowsDir + "\SystemApps\Microsoft.MicrosoftEdge_*\MicrosoftEdge.exe")
               if (Test-Path $TestProg) {
                    $TempVer = ((Get-Item $TestProg).VersionInfo).ProductVersion
                    $TempBit = (Test-is64Bit -FilePath ((Get-Item $TestProg).FullName)).FileType
                    $Temp.Add(("Edge-" + $TempBit),@{Version = $TempVer; IntSize = $TempBit;Path = $TestProg;Browser = "Microsoft Edge"})
               }
            $GlobalHash.Browsers = $Temp
        }
        function GetLocalPrinters() {
            $GlobalHash.LocalPrinter=@()
            foreach ($printer in $GlobalHash.Printers) {
                if(-Not $GlobalHash.PrinterLocalExclude.Containskey($printer.Name) -and -not ($GlobalHash.PrintersScript | Where-Object {$_.name -eq $printer.Name})){
                    $GlobalHash.LocalPrinter += $printer.Name
                }
            }
        }
        function RecordLogon () {
            If (!(Test-Path ($GlobalHash.LogUNC + "\" +$GlobalHash.UserID + ".csv"))) {
	               Add-Content -Path $GlobalHash.LogUNC -Value "UserName,ComputerName,Date,Time,Status"
               }

            If (!(Test-Path $GlobalHash.UserHomeShare)) {
	               new-item -Path $GlobalHash.UserHomeShare -ItemType directory 
                }

                # Log the time of the logon to csv
                $strLogon= $([Environment]::UserName)+","+$env:ComputerName+","+$($StrDateTime.ToShortDateString())+","+$($StrDateTime.ToShortTimeString()+",Logged on")

                Add-Content -Path ($GlobalHash.LogUNC + "\" +$GlobalHash.UserID + ".csv") -Value $strLogon 
        }
        function PrinterGroupMapping () {
             param(
                [Parameter(Mandatory=$False)][string]$OU=$GlobalHash.PrintGroupsOU
                )
            #Cleanup the input OU
            $OU.Replace("LDAP://","")
            #Search thru all AD User Groups
            foreach ( $Group in $GlobalHash.UserADObject.memberOf) {
                $SetDefault=$False
                If ($Group.ToUpper().Contains($OU.ToUpper())) {
                    $ADGroup =  New-Object System.DirectoryServices.DirectoryEntry("LDAP://" + $Group)                  
                    $ArrDescription = @()
                    $ArrDescription = $ADGroup.description.ToString().split(" ")
                    If ( $ArrDescription) {
                        switch ($ArrDescription.GetUpperBound(0)) {
                            0 {
                                #Just has Printer Server
                                $PrintServer=$ArrDescription[0]
                                $PrintQueue=$ADGroup.sAMAccountName
                               }

                            1 {
                                #Printer Server and Default
                                $PrintServer=$ArrDescription[0]
                                if ($ArrDescription[1].ToString().ToUpper() = "DEFAULT") { $SetDefault=$true}
                                $PrintQueue=$ADGroup.sAMAccountName
                                }
                            2 {
                                #Printer Server, Default and Print Queue Name
                                $PrintServer=$ArrDescription[0]
                                if ($ArrDescription[1].ToString().ToUpper() = "DEFAULT") { $SetDefault=$true}
                                $PrintQueue=$ArrDescription[2]

                                }
                            default {
                                 LogtoGUI -Message ("Could not understand print group description: " + $ADGroup.description.ToString()) -Color "Red" -FontWeight "Bold"
                                }
                            }
                        }else{
                             LogtoGUI -Message ("Could not understand print group description: " + $ADGroup.description.ToString()) -Color "Red" -FontWeight "Bold"
                        }
                    AddPrinter -PrintServer $PrintServer -PrintQueue $PrintQueue -SetDefault:$true 
                }
            }
        }
        function AddPrinter () {
             param(
                [Parameter(Mandatory=$true)][string]$PrintQueue,
                [Parameter(Mandatory=$false)][string]$PrintServer = $GlobalHash.PrintServer,
                [Parameter(Mandatory=$false)][switch]$SetDefault=$false
            )
            $PrinterSharePath = ("\\" + $PrintServer + "\" + $PrintQueue)
            $PrinterMapped = $false
            
            $ws_net = New-Object -COM WScript.Network

            If ($GlobalHash.PrintSpooler) {
                If($GlobalHash.Printers | Where-Object { $_.Name -eq $PrinterSharePath}) {
                     #LogtoGUI -Message ("Printer: " + $PrinterSharePath  + " Already mapped.") -Color "Blue" 
                     If($SetDefault) {
                        $ws_net.SetDefaultPrinter($PrinterSharePath)
                        LogtoGUI -Message ("Default Printer: " + $PrinterSharePath) -Color "Green"
                     }
                }else{
                    try {
                        #Map Printer
                        $ws_net.AddWindowsPrinterConnection($PrinterSharePath)
                        If($SetDefault) {
                            $ws_net.SetDefaultPrinter($PrinterSharePath)
                            LogtoGUI -Message ("Default Printer Added: " + $PrinterSharePath) -Color "Green"
                        }else{
                            LogtoGUI -Message ("Added Printer: " + $PrinterSharePath)
                        }
                    } catch {
                        LogtoGUI -Message ("Could not map printer: " + $PrinterSharePath  + ". Issue mapping Printer.") -Color "Red" -FontWeight "Bold"
                    }
                    #Update printer Cache
                    $GlobalHash.Printers = (Get-Printer | Select-Object Name,ComputerName,Type,Portname,ShareName)
                    $GlobalHash.PrintersScript += (Get-Printer $PrinterSharePath | Select-Object Name,ComputerName,Type,Portname,ShareName)
                    
                }
            }else{
                LogtoGUI -Message ("Could not map printer: " + $PrinterSharePath  + ". Printer Spooler is not running.") -Color "Red" -FontWeight "Bold"
            }
            $ws_net = $null
        }
        function MapDrive () {
             param(
                [Parameter(Mandatory=$true)][string]$Drive,
                [Parameter(Mandatory=$true)][string]$Share,
                [Parameter(Mandatory=$false)][string]$FileServer = $GlobalHash.FileServer,
                [Parameter(Mandatory=$false)][string]$DisplayName = ""
            )
            #Create Object for drive renameing
            $shapp=New-Object -com Shell.Application
            #Test Share formatting
            If ($Share.StartsWith("\"))
            {
                $strPath = ("\\" + $FileServer + $Share)
            }else{
                $strPath = ("\\" + $FileServer + "\" + $Share)
            }
            #Test drive letter formatting
            $Drive = $Drive.Replace("\","")
            $Drive = $Drive.Replace(":","")
            $Drive = $Drive.ToUpper()
            #Test to see if drve is mapped
            If(Test-Path($Drive + ":\")) {
                if((Get-PSDrive -Name $Drive).DisplayRoot -eq $strPath) {
                    #Update Display name
                    if($DisplayName -ne ($shapp.NameSpace($Drive + ":").Self.Name).substring(0,($shapp.NameSpace($Drive + ":").Self.Name).lastindexof("(")).trim() -and $DisplayName -ne "" ) 
                    {
                        $shapp.NameSpace($Drive + ":").Self.Name= $DisplayName
                        LogtoGUI -Message ("Renaming drive " + $Drive + ":\ to " + $DisplayName) 
                    }
                }else{ 
                    LogtoGUI -Message ("Re-mapping network drive: " + $Drive  + ":\. to correct share: " + $strPath) -Color "Yellow"
                    Remove-PSDrive -Name $Drive
                    try {
                        New-PSDrive -Name $Drive -PSProvider FileSystem -Persist -Root $strPath
                        #Update Display name
                       $shapp.NameSpace($Drive + ":").Self.Name= $DisplayName
                    } catch {
                         LogtoGUI -Message ("Could not map drive: " + $Drive  + ". Issue mapping drive to share:" + $strPath) -Color "Red" -FontWeight "Bold"
                    }
                }
            }else{
                try {
                    New-PSDrive -Name $Drive -PSProvider FileSystem -Persist -Root $strPath
                    LogtoGUI -Message ("Mapping network drive: " + $Drive  + ". to share: " + $strPath)
                    #Update Display name
                    $shapp.NameSpace($Drive + ":").Self.Name= $DisplayName
                } catch {
                         LogtoGUI -Message ("Could not map drive: " + $Drive  + ". Issue mapping drive to share:" + $strPath) -Color "Red" -FontWeight "Bold"
                    }
            }
            $shapp =$null
        }
        function Rename-Drive {
            #.Source: http://www.powershellmagazine.com/2014/07/22/pstip-rename-a-local-or-a-mapped-drive-using-shell-application/
            param([Parameter(Mandatory=$true)][string]$DriveLetter, 
                  [Parameter(Mandatory=$true)][string]$DriveName)
            $DriveLetter = $DriveLetter.Replace("\","")
            If (!($DriveLetter.Contains(":")))
            {
                $DriveLetter = $DriveLetter.Substring(1,2) + ":"
            }
            $ShellObject = New-Object –ComObject Shell.Application
            $DriveMapping = $ShellObject.NameSpace( $DriveLetter)
            If( $DriveName -ne ($ShellObject.NameSpace($DriveLetter).Self.Name).substring(0,($ShellObject.NameSpace($DriveLetter).Self.Name).lastindexof("(")).trim())
              {
                $ShellObject.NameSpace($DriveLetter).Self.Name = $DriveName
                LogtoGUI -Message ("Renaming drive " + $DriveLetter + " to " + $DriveName) 
             }
             $ShellObject = $null
        }
        function Set-DefaultWallpaper {
            param([Parameter(Mandatory=$true)][string]$DefaultWallpaper, 
                  [Parameter(Mandatory=$false)][string]$WallpaperStyle = 2)
            $DefaultWallpaper = $DefaultWallpaper.ToUpper()
            If(Test-Path -Path $DefaultWallpaper ) {
            switch((Get-ItemProperty -Path "HKCU:\Control Panel\Desktop" -Name "Wallpaper").WallPaper.ToUpper()) {

                "C:\PROGRAM FILES\CITRIX\ENHANCEDDESKTOPEXPERIENCE\CITRIX_LOGO.JPG"{
                    #XenApp Default Wallpaper
                     New-ItemProperty -Path "HKCU:\Control Panel\Desktop" -Name "Wallpaper" -Value $DefaultWallpaper -PropertyType String -Force | Out-Null
                     New-ItemProperty -Path "HKCU:\Control Panel\Desktop" -Name "WallpaperStyle" -Value $WallpaperStyle -PropertyType DWORD -Force | Out-Null
                     Restart-Explorer
                    }
                 "C:\WINDOWS\WEB\WALLPAPER\WINDOWS\IMG0.JPG" {
                    #Windows Default Wallpaper
                    New-ItemProperty -Path "HKCU:\Control Panel\Desktop" -Name "Wallpaper" -Value $DefaultWallpaper -PropertyType String -Force | Out-Null
                    New-ItemProperty -Path "HKCU:\Control Panel\Desktop" -Name "WallpaperStyle" -Value $WallpaperStyle -PropertyType DWORD -Force | Out-Null
                    Restart-Explorer
                    }       
                 $DefaultWallpaper {
                    #Company Default
                    New-ItemProperty -Path "HKCU:\Control Panel\Desktop" -Name "Wallpaper" -Value $DefaultWallpaper -PropertyType String -Force | Out-Null
                    New-ItemProperty -Path "HKCU:\Control Panel\Desktop" -Name "WallpaperStyle" -Value $WallpaperStyle -PropertyType DWORD -Force | Out-Null
                    Restart-Explorer
                   }     
                default {

                    }
                }
            }
        }
        function Restart-Explorer {
            Stop-Process -ProcessName Explorer
            If (!(Get-Process -Name explorer)) {Start-Process -FilePath Explorer}
        }
        ############################ End Main Functions ##############################
        #endregion
        #region Main Settings
        ############################ Start Main Settings ##########################################
        $GlobalHash.UserADObject = GetADObject -objectName $GlobalHash.UserID -objectType "User"

        $GlobalHash.ComputerADObject = GetADObject -objectName $GlobalHash.RealWorkstation -objectType "Computer"
        #Create Array of OU that the Computer is in.
        $Temp= @()
        ((($GlobalHash.ComputerADObject.Parent).Replace("LDAP://","")).split(",")) | ForEach-Object {
           $Temp += $_.split("=")[1]
        }
        $GlobalHash.ComputerOU = $Temp 
        $Temp = $null

        #$GlobalHash.MappedDrives = (Get-PSDrive) | Where-Object { $_.DisplayRoot -ne $null}
        $GlobalHash.Printers = (Get-Printer | Select-Object Name,ComputerName,Type,Portname,ShareName)
        #Get Print Spooler Status
        If ((Get-Service spooler).Status -eq "Running") {
            $GlobalHash.PrintSpooler = $True
        } else {
            $GlobalHash.PrintSpooler = $False
        }
        ############################ End Main Settings ##########################################
        #endregion
        #region Main Do Work
            ############################ Start Main Work ##########################################
            LogtoGUI -Message ("Welcome " + $GlobalHash.UserADObject.givenName.ToString() + " " + $GlobalHash.UserADObject.sn.ToString() + " please don't close this window..." ) -FontWeight "UltraBold"
            LogtoGUI -Message ("You are logging on to: ") -NoNewLine -FontWeight "Bold"
            LogtoGUI -Message ($GlobalHash.RealWorkstation + ".")
            LogtoGUI -Message ("Current Date is: ") -NoNewLine -FontWeight "Bold"
            LogtoGUI -Message ((Get-Date -Format F).ToString())
            LogtoGUILine

            #Get installed Browsers
            GetBrowsers

            #Add Entry into CSV log
            #RecordLogon

            #region Map Printers via AD Groups
                If ( !($GlobalHash.ComputerOU.Contains("UIC Servers")) -and !($GlobalHash.ComputerOU.Contains("Managed"))) {
                    PrinterGroupMapping -OU $GlobalHash.PrintGroupsOU
                }
            #endregion
            #region Map People with local printers.
                If ( !($GlobalHash.ComputerOU.Contains("wwt-SetLocalPrinterDefault")) -and $GlobalHash.RealWorkstation -eq $GlobalHash.Workstation) {
                    GetLocalPrinters
                    if ($GlobalHash.LocalPrinter.Count -gt 1 -and $GlobalHash.LocalPrinter.Count -ne 0) {
                        #Need to Add to White List
                    }elseif ($GlobalHash.LocalPrinter.Count -eq 1) {
                        $ws_net = New-Object -COM WScript.Network
                        $ws_net.SetDefaultPrinter($GlobalHash.LocalPrinter)
                        LogtoGUI -Message ("Default Printer Set: " + $GlobalHash.LocalPrinter) -Color "Green"
                        $ws_net = $null
                    }
                }
            #endregion
            #region Map based on computer name
                switch ($GlobalHash.Workstation) {
                    #"workstation" {
                    #    AddPrinter -PrintServer $GlobalHash.PrintServer -PrintQueue "test" -SetDefault:$true 
                    #  }


                    default{}
                }

            #endregion
            #region Map based on UserID
                switch ($GlobalHash.UserID) {
                    #"User" {
                    #    AddPrinter -PrinterServer $GlobalHash.PrintServer -PrintQueue "test" -SetDefault:$true 
                    #  }

                    default{}
                }
            #endregion
            #region Map base on Computer OU


            #endregion
            
            #region Preform actions based group membership
                # External S Drive users
                If ($GlobalHash.UserADObject.memberOf.contains((GetADObject -objectType "Group" -objectName "Share").distinguishedName.tostring())) {
                    #MapDrive -Drive "S" -Share "Shared" -FileServer $GlobalHash.FileServer -DisplayName "Shared"
                }

                #Settings for WWT Users
                If ($GlobalHash.UserADObject.memberOf.contains((GetADObject -objectType "Group" -objectName "Domain Users").distinguishedName.tostring())) {
                    #Make sure Favorites folder is created
                    If (!(Test-Path ($GlobalHash.UserHomeShare + "\Favorites")))
                    {
                        New-Item -ItemType directory -Path ($GlobalHash.UserHomeShare + "\Favorites")
                    }
                    #Maps Main Network Drives
                    #MapDrive -Drive "S" -Share "Shared" -FileServer $GlobalHash.FileServer -DisplayName "Shared"
                    #Maps Main Printers
                    If (-not $GlobalHash.ComputerOU.Contains("Servers") ){
                        #AddPrinter -PrintServer $GlobalHash.PrintServer -PrintQueue "printer1" -SetDefault:$false
                        #AddPrinter -PrintServer $GlobalHash.PrintServer -PrintQueue "printer2" -SetDefault:$false 

                    }
                }

		                               
            #endregion 

            #region Other Settings
               # foreach ($browser in $GlobalHash.Browsers.Keys) {
               #      $looptemp = [string] ($browser + " " + $GlobalHash.Browsers.$browser.IntSize + " " + $GlobalHash.Browsers.$browser.Version + " " + $GlobalHash.Browsers.$browser.Path)
               #      LogtoGUI -Message ("{0}" -f $looptemp)
               #     Write-Verbose ("Type: {0}" -f $looptemp) -Verbose
               # }
               LogtoGUI -Message ("Setting Up Other Settings")
               
               #Hide VMWareTools Icon
               If (!((Get-ItemProperty -Path "HKCU:\Software\VMware, Inc.\VMware Tools" -Name "ShowTray").ShowTray)) {
                 New-ItemProperty -Path "HKCU:\Software\VMware, Inc.\VMware Tools" -Name "ShowTray" -Value 0 -PropertyType DWORD -Force | Out-Null
               }
               #Sets Username in Office to AD Displayname
               If (!((Get-ItemProperty -Path "HKCU:\Software\Microsoft\Office\Common\UserInfo" -Name "UserName").UserName)) {
                 New-ItemProperty -Path "HKCU:\Software\Microsoft\Office\Common\UserInfo" -Name "UserName" -Value $GlobalHash.UserADObject.displayName -PropertyType String -Force | Out-Null
               }
               #
               Set-DefaultWallpaper -DefaultWallpaper $GlobalHash.CompanyDefaultWallpaper -WallpaperStyle $GlobalHash.CompanyDefaultWallpaperStyle
            #endregion Other Settings

            ############################ End Main Work ##########################################
           LogtoGUI -Message ("Done . . . ")  -FontWeight "Bold"         
           $uiHash.Window.Title = ($GlobalHash.CompanyName + " - Done ...")
           CloseGUI
        #endregion
  #region Main thread End
    })
    
    $MainpsCmd.Powershell.Runspace = $MainRunspace
    $MainpsCmd.Handle = $MainpsCmd.Powershell.BeginInvoke()
    $jobs.add($MainpsCmd)
    $Script:running = $True
  #endregion
#region Show GUI and Start

     $uiHash.Window.ShowDialog() | Out-Null 

})
$psCmd.Runspace = $newRunspace
$data = $psCmd.BeginInvoke()
#endregion


#region SIG Block
