# Assembly Start
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
# Assembly End	 

# XML file Start
$XMLFile = "PATH TO THE XML FILE";
[XML]$Script:XML = Get-Content $XMLFile
# XML file End

# Title Settings Start
$VersionNumber = "Version 1.0"
$TitleSoftware = "Title"
# Title Settings End

#-------------------------------------------------------------------------------------------
# Mainform Start

# Objects
    $MainForm = New-Object System.Windows.Forms.Form #Mainform Object
    $MainFormExitButton = New-Object System.Windows.Forms.Button # Exitbutton Object
  
   # Settings
   # Mainform 
   $MainForm.Text = "$TitleSoftware $VersionNumber" # Mainform Settings # Title
   $MainForm.StartPosition = "CenterScreen"# Mainform Settings
   $MainForm.Topmost = $false # Mainform Settings
   $MainForm.MinimumSize = "1500,1300"# Mainform Settings
   $Mainform.ShowIcon = $false # Mainform Settings
   # Exit Button
   $MainFormExitButton.Location = "1400,1220" # Mainform Settings
   $MainFormExitButton.Size = "75,23" # Mainform Settings
   $MainFormExitButton.Text = "Exit" # Mainform Settings

   # Controls.Add For the Tab
   $MainForm.Controls.Add($MainFormExitButton)# Exitbutton Controls

   # Buttons Add_Click
   $MainFormExitButton.Add_Click({$MainForm.Close()}) # Exitbutton Add_click
   
  

                                    # Mainform End
          #----------------------------------------------------------------------#

#--------------------------------------------------------------------------------------------
# MainTab Settings
   # Objects
   $MainFormTabControl = New-object System.Windows.Forms.TabControl # Maintab Object
   # Settings
   $MainFormTabControl.Size = "1350,1190" # Maintab Settings size
   $MainFormTabControl.Location = "15,50" # Maintab Settings location
   # Controls.Add For the Tab
   $MainForm.Controls.Add($MainFormTabControl)# Maintab Controls
   



   
                                    # MainTab End
          #----------------------------------------------------------------------#

#--------------------------------------------------------------------------------------------
# Childtab Active Directory   ( TAD = Tab Active Directory ) 
   # Objects
   $TabActiveDirectory = New-Object System.Windows.Forms.TabPage # Tab Active Directory Object
   $TADLabel = New-Object System.Windows.Forms.Label # Enter Username Label Object
   $TADLabelADFunctions = New-Object System.Windows.Forms.Label # Label Quickfix Active Directory Object
   $TADTextbox = New-Object System.Windows.Forms.TextBox # Textbox to search Object in Active Directory Tab
   $TADUnlockUserButton = New-Object System.Windows.Forms.Button # Unlock User Button Object
   $TADSearchButton = New-Object System.Windows.Forms.Button # Search button Object
   $TADRichTextbox = New-Object System.Windows.Forms.RichTextbox # RichTextbox Object
   

 # Settings Start
   # Binding, collection
   $TabActiveDirectory.DataBindings.DefaultDataSourceUpdateMode = 0 # Tab - ActiveDirectory Settings
   # Name
   $TabActiveDirectory.Name = "TabActiveDirectory" # Tab - ActiveDirectory Settings
   #Visual
   $TabActiveDirectory.UseVisualStyleBackColor = $True # Tab - ActiveDirectory Settings
   # Location
   $TADLabel.Location = "75,25" # Enter Username Label Settings
   $TADLabelADFunctions.Location = "1195,100" # Label Quickfix Active Directory Settings
   $TADRichTextbox.Location = "25,100" # RichTextbox Settings
   $TADSearchButton.Location = "250,50" # Search button Settings
   $TADTextbox.Location = "25,50"# Textbox Settings
   $TADUnlockUserButton.Location = "1220,130" # Unlock User Button Settings
   # Text
   $TabActiveDirectory.Text = "Active Directory” # Title on the tab
   $TADLabel.Text = "Enter Username" # Enter Username Label Settings
   $TADLabelADFunctions.Text = "Quickfix - Active Directory" # Label Quickfix Active Directory Settings
   $TADSearchButton.Text = "Search" # Search button Settings
   $TADUnlockUserButton.Text = "Unlock" # Unlock User Button Settings
   # Size
   $TADTextbox.Size = "200,20"# Textbox  Settings
   $TADRichTextbox.Size = "1150,1020" # RichTextbox Settings
   $TADSearchButton.Size = "80,20" # Search button Settings
   $TADLabelADFunctions.Size = "250,20" # Label Quickfix Active Directory Settings
   $TADUnlockUserButton.Size = "80,20" # Unlock User Button Settings
   # Scrollbars
   $TADRichTextbox.ScrollBars = "Vertical" # RichTextbox Settings
   # Font 
   $TADRichTextbox.font ="lucida console" # RichTextbox Settings

   # Controls.Add For the Tab
   $MainFormTabControl.Controls.Add($TabActiveDirectory) # ActiveDirectory Control
   $TabActiveDirectory.Controls.Add($TADLabel) # Enter Username Label Control
   $TabActiveDirectory.Controls.Add($TADTextbox) # Textbox Control
   $TabActiveDirectory.Controls.Add($TADRichTextbox) # RichTextbox Control
   $TabActiveDirectory.Controls.Add($TADSearchButton) # Search button Control
   $TabActiveDirectory.Controls.Add($TADLabelADFunctions) # Ad functions label Control
   $TabActiveDirectory.Controls.Add($TADUnlockUserButton)  # Unlock User Button Control
 # Settings End
   
   # Buttons Add_Click
   $TADSearchButton.Add_Click({ActiveDirectoryUserSearch}) # Search for user Addclick
   $TADUnlockUserButton.Add_Click({UnlockUser})
   
   # Functions
   function ClearRichTextBox { # Clears out the richtextbox for every query.
   $TADRichTextbox.Clear() 
   }
   
   function ActiveDirectoryUserSearch { # Search after chosen attributes for the account in the AD
   #ClearRichTextBox
   Import-Module activedirectory
   $ADUserName = $TADTextbox.Text
   $QueryGetAdUser= Get-ADUser -Identity $ADUserName -Properties Name, whenCreated, DisplayName, CanonicalName, lsPersonnummer, lsPrivatEpost, mobile, telephoneNumber, lsQuotaAdm, lsQuotaPed, lsQuotaMail, Manager, lsAnsvar, lsAvtal, Office, lsChef, lsUppdateradAv, title, lsSITHS, LockedOut, PasswordLastSet, PasswordExpired, msDS-UserPasswordExpiryTimeComputed, LastBadPasswordAttempt | 
   Format-List Name, whenCreated, DisplayName, CanonicalName, lsPersonnummer, lsPrivatEpost, mobile, telephoneNumber, lsQuotaAdm, lsQuotaPed, lsQuotaMail, Manager, lsAnsvar, lsAvtal, Office, lsChef, lsUppdateradAv, title, lsSITHS, LockedOut, PasswordLastSet, PasswordExpired, @{Name="ExpiryDate";Expression={[datetime]::FromFileTime($_."msDS-UserPasswordExpiryTimeComputed")}}, LastBadPasswordAttempt | out-string;
   $TADRichTextbox.text=$QueryGetAdUser}   

  
   function UnlockUser { # Unlocks the Account in the textbox
   $ADUserName = $TADTextbox.Text
   Unlock-ADAccount -Identity $ADUserName | out-string;
   $PopupUnlockAccount = New-Object -ComObject Wscript.Shell
   $PopupUnlockAccount.Popup("Account $ADUserName is now unlocked",0,"Done",0x1)}

                              
                              
                              
                              # Active Directory End
          #----------------------------------------------------------------------#

#--------------------------------------------------------------------------------------------

# Childtab Computer Tab 
# Tab Settings
$TabComputer = New-Object System.Windows.Forms.TabPage # The ChildTab, Computer
$MainFormTabControl.Controls.Add($TabComputer) # The Tab Computer 
$TabComputer.DataBindings.DefaultDataSourceUpdateMode = 0 # The Tab Computer
$TabComputer.UseVisualStyleBackColor = $True # The Tab Computer
$TabComputer.Name = "TabComputer" # The Tab Computer
$TabComputer.Text = "Computer” # Title on the tab
   
 #Child-Objects
  # Buttons
    $ButtonAdminDameware = New-Object System.Windows.Forms.Button # Button to start Dameware
    $ButtonAdminEventviewer = New-Object System.Windows.Forms.Button # Button to start Eventviewer 
    $ButtonAdminHarddriveC = New-Object System.Windows.Forms.Button # Button to se remote pc harddrive
    $ButtonAdminService = New-Object System.Windows.Forms.Button # Button to start Services
    $ButtonAdminUserAndGroup = New-Object System.Windows.Forms.Button # Button to start User And Groups
    $ButtonKillProcess = New-Object System.Windows.Forms.Button # Button to kill the chosen process
    $ButtonOtherSendMessagePc = New-Object System.Windows.Forms.Button # Sends a message to the chosen pc
    $ButtonOtherRestartPc = New-Object System.Windows.Forms.Button # Restarts the chosen pc, send a messeage before
    $ButtonQuickfixAdSearch = New-Object System.Windows.Forms.Button # Button to search the AD
    $ButtonQuickfixApplications = New-Object System.Windows.Forms.Button # Downloads and generates a list of installed apps, can uninstall apps
    $ButtonQuickfixProcess = New-Object System.Windows.Forms.Button # Downloads and generates a list of process thats runs can kill process
    $ButtonQuickfixService = New-Object System.Windows.Forms.Button # Downloads and generates a list of services, can stop, restart process
    $ButtonQuickfixStartUp = New-Object System.Windows.Forms.Button # Downloads and generates of startup, can disable startup items
    $ButtonQuickfixSysteminfo = New-Object System.Windows.Forms.Button # Downloads and generates a list of systeminfo
    # Labels
    $LabelAdminTools = New-Object System.Windows.Forms.Label # Label AdminTools, Text = " - Admin Tools - "
    $LabelEnterHostnameOr  = New-Object System.Windows.Forms.Label # Label EnterHostnameOr, Text = " - Enter Pcname Or Ip-Adress - "
    $LabelKillProcess = New-Object System.Windows.Forms.Label # Label KillProcess, Text = " Enter ProcessID "
    $LabelSendMessage = New-Object System.Windows.Forms.Label # Label Send Message, Text = " Enter Message "
    $LabelOtherFunctions = New-Object System.Windows.Forms.Label # Label Send, Text = " - Other - "
    $LabelStatusboxComputer = New-Object System.Windows.Forms.Label # Label Status , Text = " - Status -  " 
    $LabelSend = New-Object System.Windows.Forms.Label # Label QuickFixTool, Text = " - Send Commands - "
    $LabelQuickFixTool = New-Object System.Windows.Forms.Label
    # ListView
    $lvMain = New-Object System.Windows.Forms.ListView # Listview 
    # Textbox
    $TextboxKillProcess =  New-Object System.Windows.Forms.TextBox # Textbox to kill the chosen processId
    $TextboxSendMessage =  New-Object System.Windows.Forms.TextBox # Textbox to send message to the chosen PC 
    $TextboxSearchbarComputers = New-Object System.Windows.Forms.TextBox # Textbox to search Object in Computer Tab
    $TextboxStatusboxComputer = New-Object System.Windows.Forms.RichTextBox # Statusbox
    # Lv Main Extra Settings 
    $lvMain.FullRowSelect = $True  # Listview 
    $lvMain.GridLines = $True  # Listview 
    $lvMain.Name = "lvMain" # Listview 
    $lvMain.View = 'Details' # Listview 
 
  # Text Data
    # Buttons
    $ButtonAdminDameware.text = " DameWare " # Button to start Dameware
    $ButtonAdminEventviewer.text = "Eventview" # Button to start Eventviewer
    $ButtonAdminHarddriveC.text = "C:\" # Button to se remote pc harddrive
    $ButtonAdminService.text = "Services" # Button to start Services
    $ButtonAdminUserAndGroup.text = "User-Group" # Button to start User And Groups
    $ButtonKillProcess.text = "Kill Process" # Button to kill the chosen process
    $ButtonOtherSendMessagePc.text = "Send Mess" # Sends a message to the chosen pc
    $ButtonOtherRestartPc.text = "Restart" # Restarts the chosen pc, send a messeage before
    $ButtonQuickfixAdSearch.text= "Search AD" # button to search the AD
    $ButtonQuickfixApplications.text = "Apps" # Downloads and generates a list of installed apps, can uninstall apps
    $ButtonQuickfixProcess.text = "Process" # Downloads and generates a list of process thats runs can kill process
    $ButtonQuickfixService.text = "Service" # Downloads and generates a list of services, can stop, restart process
    $ButtonQuickfixStartUp.text = "StartUp" # Downloads and generates of startup, can disable startup items
    $ButtonQuickfixSysteminfo.text = "SystemInfo" # Downloads and generates a list of systeminfo
    # Labels
    $LabelAdminTools.text = "Admin Tools" # Label AdminTools, Text = " - Admin Tools - "
    $LabelEnterHostnameOr.text = " - Enter Pcname Or Ip-Adress - " # Label EnterHostnameOr, Text = " - Enter Pcname Or Ip-Adress - "
    $LabelKillProcess.text = "- Enter Process-ID - "  # Label KillProcess, Text = " Enter ProcessID "
    $LabelSendMessage.text = "- Enter Message" # Label Send Message, Text = " Enter Message "
    $LabelOtherFunctions.text = "Other-Fix"  # Label Send, Text = " - Other - "
    $LabelStatusboxComputer.text  = " - Status -" # Label Status , Text = " - Status -  "
    $LabelQuickFixTool.text = "Quick-Fix " # Label Quickfix Tool, Text = " - Quickfix "
    # Textbox
    $TextboxKillProcess.text =  "" # Textbox to kill the chosen processId
    $TextboxSendMessage.text = "" # Textbox to send message to the chosen PC
    $TextboxStatusboxComputer.text = "" # Statusbox
    $TextboxSearchbarComputers.text = ""  # Textbox to search Object in Computer Tab
    

  # Location
    # Buttons
    $ButtonAdminDameware.location = "1220,125" # Button to start Dameware
    $ButtonAdminEventviewer.location = "1220,155" # Button to start Eventviewer 
    $ButtonAdminHarddriveC.location = "1220,185" # Button to se remote pc harddrive
    $ButtonAdminService.location = "1220,215" # Button to start Services
    $ButtonAdminUserAndGroup.location = "1220,245" # Button to start User And Groups 
    $ButtonKillProcess.location = "190,250"  # Button to kill the chosen process
    $ButtonOtherSendMessagePc.location = "190,175"  # Sends a message to the chosen pc
    $ButtonOtherRestartPc.location  = "1220,780" # Restarts the chosen pc, send a messeage before
    $ButtonQuickfixAdSearch.location = " 1220,525 " # button to search the AD 
    $ButtonQuickfixApplications.location = "1220,375 " # Downloads and generates a list of installed apps, can uninstall apps
    $ButtonQuickfixProcess.location = "1220,405"  # Downloads and generates a list of process thats runs can kill process
    $ButtonQuickfixService.location = "1220,465"  # Downloads and generates a list of services, can stop, restart process
    $ButtonQuickfixStartUp.location = "1220,435"  # Downloads and generates of startup, can disable startup items
    $ButtonQuickfixSysteminfo.location = "1220,495"  # Downloads and generates a list of systeminfo
    # Labels
    $LabelAdminTools.location = "1225,100" # Label AdminTools, Text = " - Admin Tools - "
    $LabelEnterHostnameOr.location = " 25,25 " # Label EnterHostnameOr, Text = " - Enter Pcname Or Ip-Adress - "
    $LabelKillProcess.location = " 55,225 " # Label KillProcess, Text = " Enter ProcessID "
    $LabelSendMessage.location = " 55, 150" # Label Send Message, Text = " Enter Message "
    $LabelOtherFunctions.Location = "1225,750"  # Label Send, Text = " - Other - "
    $LabelStatusboxComputer.location = " 775,25" # Label Status , Text = " - Status -  "
    $LabelQuickFixTool.Location = "1225,350" # Label Quickfix Tool, Text = " - Quickfix Tool - "
    # Listview
    $lvMain.Location = '25, 285' # Listview 
    # Textbox
    $TextboxKillProcess.location = "25,250" # Textbox to kill the chosen processId
    $TextboxSendMessage.location = "25,175" # Textbox to send message to the chosen PC
    $TextboxSearchbarComputers.location = "25,50" # Textbox to search Object in Computer Tab
    $TextboxStatusboxComputer.location = "500, 50" # Statusbox

  # Size
    # Buttons 
    $ButtonAdminDameware.size = "75, 20" # Button to start Dameware
    $ButtonAdminEventviewer.size = "75, 20" # Button to start Eventviewer 
    $ButtonAdminHarddriveC.size = "75, 20" # Button to se remote pc harddrive
    $ButtonAdminService.size = "75, 20" # Button to start Services
    $ButtonAdminUserAndGroup.size = "75, 20" # Button to start User And Groups
    $ButtonKillProcess.size = "75, 20" # Button to kill the chosen process
    $ButtonOtherRestartPc.size = "75, 20" # Restarts the chosen pc, send a messeage before
    $ButtonOtherSendMessagePc.size = "75, 20" # Sends a message to the chosen pc
    $ButtonQuickfixAdSearch.size = " 75,20 " # button to search the AD
    $ButtonQuickfixApplications.size = "75, 20" # Downloads and generates a list of installed apps, can uninstall apps
    $ButtonQuickfixProcess.size = "75, 20" # Downloads and generates a list of process thats runs can kill process
    $ButtonQuickfixService.size = "75, 20" # Downloads and generates a list of services, can stop, restart process
    $ButtonQuickfixStartUp.size = "75, 20" # Downloads and generates of startup, can disable startup items
    $ButtonQuickfixSysteminfo.size = "75, 20" # Downloads and generates a list of systeminfo
    # Labels
    $LabelAdminTools.size = "200,25" # Label AdminTools, Text = " - Admin Tools - "
    $LabelEnterHostnameOr.size = " 200,25 " # Label EnterHostnameOr, Text = " - Enter Pcname Or Ip-Adress - "
    $LabelKillProcess.size = " 200,25 " # Label KillProcess, Text = " Enter ProcessID "
    $LabelSendMessage.size = " 200,25 " # Label Send Message, Text = " Enter Message "
    $LabelOtherFunctions.size = "200,25"  # Label Send, Text = " - Other - "
    $LabelStatusboxComputer.size  = " 200,20" # Label Status , Text = " - Status -  "
    $LabelQuickFixTool.size = "200,25" # Quickfix Tool Label
    # Listview 
    $lvMain.Size = '1075, 850' # Listview 
    # Textbox
    $TextboxSendMessage.size = " 150,20 "# # Textbox to send message to the chosen PC
    $TextboxKillProcess.size = " 150,20 "# Textbox to kill the chosen processId
    $TextboxSearchbarComputers.size = " 150,20 "# Textbox to search Object in Computer Tab
    $TextboxStatusboxComputer.size = "600,225" # Statusbox

  # Scrollbars
    # Textbox  
    $TextboxStatusboxComputer.ScrollBars = "Vertical" 
          
  # Font 
    $TextboxStatusboxComputer.font ="lucida console"
  
  # Controls.Add For the Tab
    # Buttons
    $TabComputer.Controls.Add($ButtonAdminDameware)  # Button to start Dameware
    $TabComputer.Controls.Add($ButtonAdminEventviewer)  # Button to start Eventviewer
    $TabComputer.Controls.Add($ButtonAdminHarddriveC)  # Button to se remote pc harddrive
    $TabComputer.Controls.Add($ButtonAdminService)  # Button to start Services
    $TabComputer.Controls.Add($ButtonAdminUserAndGroup)  # Button to start User And Groups
    $TabComputer.Controls.Add($ButtonKillProcess)  # Button to kill the chosen process
    $TabComputer.Controls.Add($ButtonOtherRestartPc)  # Restarts the chosen pc, send a messeage before
    $TabComputer.Controls.Add($ButtonOtherSendMessagePc)  # Sends a message to the chosen pc
    $TabComputer.Controls.Add($ButtonQuickfixAdSearch)  # Button to search the AD  
    $TabComputer.Controls.Add($ButtonQuickfixApplications)  # Downloads and generates a list of installed apps, can uninstall apps
    $TabComputer.Controls.Add($ButtonQuickfixProcess)  # Downloads and generates a list of process thats runs can kill process
    $TabComputer.Controls.Add($ButtonQuickfixService)# Downloads and generates a list of services, can stop, restart process
    $TabComputer.Controls.Add($ButtonQuickfixStartUp)# Downloads and generates of startup, can disable startup items
    $TabComputer.Controls.Add($ButtonQuickfixSysteminfo)  # Downloads and generates a list of systeminfo
    # Labels
    $TabComputer.Controls.Add($LabelAdminTools)  # Label AdminTools, Text = " - Admin Tools - "
    $TabComputer.Controls.Add($LabelEnterHostnameOr)  # Label EnterHostnameOr, Text = " - Enter Pcname Or Ip-Adress - "
    $TabComputer.Controls.Add($LabelKillProcess) # Label KillProcess, Text = " Enter ProcessID "
    $TabComputer.Controls.Add($LabelSendMessage)  # Label Send Message, Text = " Enter Message "
    $TabComputer.Controls.Add($LabelOtherFunctions) # Label Send, Text = " - Other - "
    $TabComputer.Controls.Add($LabelStatusboxComputer) #  Label Status , Text = " - Status -  "
    $TabComputer.Controls.Add($LabelQuickFixTool) # Label Quickfix Tool, Text = " - Quickfix Tool - "
    # Listview
    $TabComputer.Controls.Add($lvMain) #  lvMain
    # Textbox
    $TabComputer.Controls.Add($TextboxKillProcess) # Textbox to kill the chosen processId
    $TabComputer.Controls.Add($TextboxSendMessage) # Textbox to send message to the chosen PC
    $TabComputer.Controls.Add($TextboxSearchbarComputers) # Textbox to search Object in Computer Tab
    $TabComputer.Controls.Add($TextboxStatusboxComputer) #  Statusbox
 

                                        # Settings End
   
   # Buttons Add_Click
    $ButtonAdminDameware.Add_Click({ AdminOpenDamweare}) #Button to start Dameware
    $ButtonAdminEventviewer.Add_Click({ AdminOpenEventViewer}) # Button to start Eventviewer 
    $ButtonAdminHarddriveC.Add_Click({ AdminOpenHarddriveC}) # Button to se remote pc harddrive
    $ButtonAdminService.Add_Click({ AdminOpenService}) # Button to start Services
    $ButtonAdminUserAndGroup.Add_Click({ AdminOpenUserAndGroup}) # Button to start User And Groups
    $ButtonKillProcess.Add_Click({ AdminKillProcess}) # Button to start User And Groups
    $ButtonOtherRestartPc.Add_Click({ OtherOpenRestartPc}) # Restarts the chosen pc, send a messeage before
    $ButtonOtherSendMessagePc.Add_Click({ OtherOpenSendMessagePc}) # Sends a message to the chosen pc
    $ButtonQuickfixAdSearch.Add_Click({activedirectorysearch})  # button to search the AD 
    $ButtonQuickfixApplications.Add_Click({ QuickFixGetApplications}) # Downloads and generates a list of installed apps, can uninstall apps
    $ButtonQuickfixProcess.Add_Click({ QuickFixGetProcess}) # Downloads and generates a list of process thats runs can kill process
    $ButtonQuickfixService.Add_Click({ QuickFixGetService}) # Downloads and generates a list of services, can stop, restart process
    $ButtonQuickfixSysteminfo.Add_Click({ QuickFixGetSysteminfo}) # Downloads and generates a list of systeminfo
    $ButtonQuickfixStartUp.Add_Click({ QuickFixGetStartUp}) # Downloads and generates of startup, can disable startup items

                                       # Buttons Add_CLick End

   #Functions
   function AdminOpenDamweare { # Button,function to start Dameware
   $StartDamewarePath = "C:\Program Files (x86)\DameWare Development\DameWare Mini Remote Control\DWRCC.exe"
   Start-Process $StartDamewarePath
   $TextboxStatusboxComputer.text = " Starting Dameware Remote"}
    
   function AdminKillProcess { # Button,function kill process
   $QueryComputerKillProcess = $TextboxSearchbarComputers.Text
   $ProcessID = $TextboxKillProcess.text
   $StopProcess = (Get-WmiObject Win32_Process -ComputerName $QueryComputerKillProcess | ?{ $_.ProcessID -match "$ProcessID" }).Terminate()
   $TextboxStatusboxComputer.text = " Stopped $ProcessId on $QueryComputerKillProcess"
   }    
   
   function AdminOpenEventViewer { # Button,function to start Eventviwer 
   $QueryOpenEventViewer = $TextboxSearchbarComputers.Text
   $PingRequest = Test-Connection -ComputerName $QueryOpenEventViewer -Count 1 -Quiet 
   If ($PingRequest -eq $true) 
   { EventVwr $QueryOpenEventViewer
   $TextboxStatusboxComputer.text = " Starting EventViewer On ", $QueryOpenEventViewer } else {
   $TextboxStatusboxComputer.text = " can not contact $QueryOpenEventViewer to start EventViewer"
   }}
     
   function AdminOpenHarddriveC { # Button,function to remote open and see c:\
   $QueryOpenOpenHardriveC = $TextboxSearchbarComputers.Text
   $PingRequest = Test-Connection -ComputerName $QueryOpenOpenHardriveC -Count 1 -Quiet 
   If ($PingRequest -eq $true) 
   { Explorer.exe \\$QueryOpenOpenHardriveC\c$ 
   $TextboxStatusboxComputer.text = " Starting Explorer C:\ On", $QueryOpenOpenHardriveC } else {
   $TextboxStatusboxComputer.text = " can not contact $QueryOpenOpenHardriveC to se Disk C:\ "
   }}
   
   
    
   function AdminOpenService { #Button,function to start Services
   $QueryOpenServices = $TextboxSearchbarComputers.Text
   $PingRequest = Test-Connection -ComputerName $QueryOpenServices -Count 1 -Quiet 
   If ($PingRequest -eq $true) 
   { Services.msc /Computer=$QueryOpenServices
   $TextboxStatusboxComputer.text = " Starting Services On $QueryOpenServices " } else {
   $TextboxStatusboxComputer.text = " can not contact $QueryOpenServices to start Services "
   }}
   
   
   function AdminOpenUserAndGroup { # Button to start User And Groups 
   $QueryOpenUsersAndGroups = $TextboxSearchbarComputers.Text
   $PingRequest = Test-Connection -ComputerName $QueryOpenUsersAndGroups -Count 1 -Quiet
   If ($PingRequest -eq $true) 
   { LUsrMgr.msc /Computer=$QueryOpenUsersAndGroups
   $TextboxStatusboxComputer.text = " Starting Local Users And Groups On $QueryOpenUsersAndGroups " } else { 
   $TextboxStatusboxComputer.text = " can not contact $QueryOpenUsersAndGroups to start Local Users And Groups "
   }}
   
   function OtherOpenRestartPc { # Restarts the chosen pc, send a confirmation box before the reboot 
   $QueryRestartPc = $TextboxSearchbarComputers.Text
   $confirmation = Read-Host "Are you Sure You Want To Proceed:"
   if ($confirmation -eq 'y') { # proceed to the restart
   }
   Restart-Computer -Computername $QueryRestartPc -Force
   $TextboxStatusboxComputer.text = " $QueryRestartPc Is Rebooting "}  
   
  function OtherOpenSendMessagePc { # Sends a message to the chosen pc 
  $QuerySendMessagePc = $TextboxSearchbarComputers.Text
  $msg = $TextboxSendMessage.text
  Invoke-WmiMethod -Path Win32_Process -Name Create -ArgumentList "msg * $msg" -ComputerName $QuerySendMessagePc
  $TextboxStatusboxComputer.text = " Done, Your message was submitted successfully to $QuerySendMessagePc "
  }   
  
  function QuickFixGetApplications { # Downloads and generates a list of installed apps,
  $QueryGetApplicationsPc = $TextboxSearchbarComputers.Text
  $TextboxStatusboxComputer.text = " Downloading Application-Information From $QueryGetApplicationsPc, takes a few seconds"
  ClearListview
  $XML.Options.Applications.Property | %{GetColumn $_}
  ResizeColumns
  $Col0 = $lvMain.Columns[0].Text
  $GetWmiObjectwin32Product = Get-WmiObject win32_Product -ComputerName $QueryGetApplicationsPc | Sort Name,Vendor
  Start-Sleep -m 250
  $GetWmiObjectwin32Product | %{
  $Item = New-Object System.Windows.Forms.ListViewItem($_.$Col0)
  $TextboxStatusboxComputer.text = " Done, ApplicationsInformation-List Is Now Made From $QueryGetApplicationsPc"
  ForEach ($Col in ($lvMain.Columns | ?{$_.Index -ne 0})){$Field = $Col.Text;$Item.SubItems.Add($_.$Field)}
  $lvMain.Items.Add($Item)}}
   
   function QuickFixGetProcess { # Downloads and generates a list of process thats runs 
   $QueryGetProcessPc = $TextboxSearchbarComputers.Text
   $TextboxStatusboxComputer.text = " Downloading Processes-Information From $QueryGetProcessPc , takes a few seconds"
   Start-Sleep -S 3
   ClearListview
   $XML.Options.Processes.Property | %{GetColumn $_}
   ResizeColumns
   $Col0 = $lvMain.Columns[0].Text
   $GetWmiObjectwin32process = Get-WmiObject win32_process -ComputerName $QueryGetProcessPc | Sort Name
   $GetWmiObjectwin32process | %{
   $Item = New-Object System.Windows.Forms.ListViewItem($_.$Col0)
   $TextboxStatusboxComputer.text = "Done, ProcessInformation-List Is Now Made From $QueryGetProcessPc " 
   ForEach ($Col in ($lvMain.Columns | ?{$_.Index -ne 0})){
   $Field = $Col.Text
   $SubItem = $_.$Field
   if($SubItem -ne $null){$Item.SubItems.Add($SubItem)}
   else{$Item.SubItems.Add("")}}
   $lvMain.Items.Add($Item)}}
   
   function QuickFixGetService { # Downloads and generates a list of services, can stop, restart process 
   $QueryGetServicePc = $TextboxSearchbarComputers.Text
   $TextboxStatusboxComputer.text = " Downloading Services-Information From $QueryGetServicePc , takes a few seconds"
   Start-Sleep -S 3
   ClearListview
   $XML.Options.Services.Property | %{GetColumn $_}
   ResizeColumns
   $Col0 = $lvMain.Columns[0].Text
   $GetWmiObjectWin32Service = Get-WmiObject Win32_Service -ComputerName $QueryGetServicePc | Sort Name
   $GetWmiObjectWin32Service | %{
   $Item = New-Object System.Windows.Forms.ListViewItem($_.$Col0)
   $TextboxStatusboxComputer.text = " Done, ServicesInformation-List Is Now Made From $QueryGetServicePc "
   ForEach ($Col in ($lvMain.Columns | ?{$_.Index -ne 0})){$Field = $Col.Text;$Item.SubItems.Add($_.$Field)}
   $lvMain.Items.Add($Item)}}  
   
   function QuickFixGetSysteminfo { # Downloads and generates a list of the system 
   $QueryGetSysteminfoPc = $TextboxSearchbarComputers.Text
   $TextboxStatusboxComputer.text = " Downloading SystemInformation From $QueryGetSysteminfoPc , takes a few seconds"
   Start-Sleep -S 3
   ClearListview
   $GetWmiObjectWin32ComputerSystem = Get-WmiObject Win32_ComputerSystem -ComputerName $QueryGetSysteminfoPc
   $GetWmiObjectWin32ComputerSystemProduct = Get-WmiObject Win32_ComputerSystemProduct -ComputerName $QueryGetSysteminfoPc
   $GetWmiObjectWin32OperatingSystem = Get-WmiObject Win32_OperatingSystem -ComputerName $QueryGetSysteminfoPc
   $GetWmiObjectWin32BIOS = Get-WmiObject Win32_BIOS -ComputerName $QueryGetSysteminfoPc
   $GetWmiObjectWin32Processor = Get-WmiObject Win32_Processor -ComputerName $QueryGetSysteminfoPc
   $GetWmiObjectWin32PhysicalMemory = Get-WmiObject Win32_PhysicalMemory -ComputerName $QueryGetSysteminfoPc
   $GetWmiObjectWin32NetworkAdapter = Get-WmiObject Win32_NetworkAdapterConfiguration -ComputerName $QueryGetSysteminfoPc -Filter "IPEnabled='True'"
   $GetWmiObjectWin32VideoController = Get-WmiObject Win32_VideoController -ComputerName $QueryGetSysteminfoPc
   $GetWmiObjectWin32CDROMDrive = Get-WmiObject Win32_CDROMDrive -ComputerName $QueryGetSysteminfoPc
   $GetWmiObjectWin32LogicalDisk = Get-WmiObject Win32_LogicalDisk -ComputerName $QueryGetSysteminfoPc
   $GetWmiObjectWin32Process  = Get-WmiObject Win32_Process -ComputerName $QueryGetSysteminfoPc
   $TextboxStatusboxComputer.text = " Done, A SystemInformation-List Is Now Made From  $QueryGetSysteminfoPc " 	
   "Property","Value" | %{GetColumn $_}
	if ($XML.Options.SystemInfo.General.Enabled -eq $true){
    $Item = New-Object System.Windows.Forms.ListViewItem("General")
    $Item.BackColor = "Black"
    $Item.ForeColor = "White"
    $lvMain.Items.Add($Item)}    

    if($XML.Options.SystemInfo.General.ComputerName.Enabled -eq $true){
    $Item = New-Object System.Windows.Forms.ListViewItem("Computer Name")
    $Item.SubItems.Add($GetWmiObjectWin32ComputerSystem.Name)
    $lvMain.Items.Add($Item)}
	
    if($XML.Options.SystemInfo.General.CurrentUser.Enabled -eq $true){
    $Item = New-Object System.Windows.Forms.ListViewItem("User")
    
    if($GetWmiObjectWin32ComputerSystem.UserName -ne $null){$Item.SubItems.Add($GetWmiObjectWin32ComputerSystem.UserName)}
    else{$Item.SubItems.Add("")}					
    $lvMain.Items.Add($Item)
				
    if($XML.Options.SystemInfo.General.LogonTime.Enabled -eq $true){
    $Item = New-Object System.Windows.Forms.ListViewItem("User Logon")}
    
    if($GetWmiObjectWin32Process | ?{$_.Name -eq "explorer.exe"}){
    $UserLogonDT = ($GetWmiObjectWin32Process | ?{$_.Name -eq "explorer.exe"} | Sort CreationDate | Select-Object -First 1).CreationDate
    $UserLogon = [System.Management.ManagementDateTimeconverter]::ToDateTime($UserLogonDT).ToString()
    $Item.SubItems.Add($UserLogon)
    }else{
    $Item.SubItems.Add("N/A")}
    $lvMain.Items.Add($Item)}
    
	if($XML.Options.SystemInfo.General.LastRestart.Enabled -eq $true){
	$LastBootUpTime = [System.Management.ManagementDateTimeconverter]::ToDateTime($GetWmiObjectWin32OperatingSystem.LastBootUpTime).ToString()
	$Item = New-Object System.Windows.Forms.ListViewItem("Last Restart")
	$Item.SubItems.Add($LastBootUpTime)
	$lvMain.Items.Add($Item)}
			
	if ($XML.Options.SystemInfo.Build.Enabled -eq $true){
    $Item = New-Object System.Windows.Forms.ListViewItem("Build")
    $Item.BackColor = "Black"
    $Item.ForeColor = "White"
    $lvMain.Items.Add($Item)}
    
    if($XML.Options.SystemInfo.Build.Manufacturer.Enabled -eq $true){
    $Item = New-Object System.Windows.Forms.ListViewItem("Manufacturer")
    $Item.SubItems.Add($GetWmiObjectWin32ComputerSystem.Manufacturer)
    $lvMain.Items.Add($Item)}
			
    if($XML.Options.SystemInfo.Build.Model.Enabled -eq $true){
    $Item = New-Object System.Windows.Forms.ListViewItem("Model")
    $Item.SubItems.Add($GetWmiObjectWin32ComputerSystem.Model)
    $lvMain.Items.Add($Item)}
    
    if($XML.Options.SystemInfo.Build.Serial.Enabled -eq $true){
    $Item = New-Object System.Windows.Forms.ListViewItem("Serial")
    $Item.SubItems.Add($GetWmiObjectWin32BIOS.SerialNumber)
    $lvMain.Items.Add($Item)}
			
    if ($XML.Options.SystemInfo.Hardware.Enabled -eq $true){
    $Item = New-Object System.Windows.Forms.ListViewItem("Hardware")
    $Item.BackColor = "Black"
    $Item.ForeColor = "White"
    $lvMain.Items.Add($Item)}
    
    if($XML.Options.SystemInfo.Hardware.CPU.Enabled -eq $true){
    $GetWmiObjectWin32Processor | %{
    $Item = New-Object System.Windows.Forms.ListViewItem("CPU")
    $Item.SubItems.Add($GetWmiObjectWin32Processor.Name.Trim())
    $lvMain.Items.Add($Item)}
  
    if($XML.Options.SystemInfo.Hardware.RAM.Enabled -eq $true){
    $tRAM = "{0:N2} GB Usable - " -f $($GetWmiObjectWin32ComputerSystem.TotalPhysicalMemory / 1GB)
    $GetWmiObjectWin32PhysicalMemory | %{$tRAM += "[$($_.Capacity / 1GB)] "}
    $Item = New-Object System.Windows.Forms.ListViewItem("RAM")
    $Item.SubItems.Add($tRAM)
    $lvMain.Items.Add($Item)}
				
    if($XML.Options.SystemInfo.Hardware.HD.Enabled -eq $true){
    $GetWmiObjectWin32LogicalDisk | ?{$_.DriveType -eq 3} | %{
    $HDInfo = "{0:N1} GB Free / {1:N1} GB Total" -f ($_.FreeSpace / 1GB), ($_.Size / 1GB)}
    $Item = New-Object System.Windows.Forms.ListViewItem("HD")
    $Item.SubItems.Add($HDinfo)
    $lvMain.Items.Add($Item)}
    
    if($XML.Options.SystemInfo.Hardware.OpticalDrive.Enabled -eq $true){
    $GetWmiObjectWin32CDROMDrive | %{
    $Item = New-Object System.Windows.Forms.ListViewItem("Optical Drive")
    $Item.SubItems.Add("[$($GetWmiObjectWin32CDROMDrive.Drive)] $($GetWmiObjectWin32CDROMDrive.Caption)")}
    $lvMain.Items.Add($Item)}
    
    if($XML.Options.SystemInfo.Hardware.GPU.Enabled -eq $true){
    $GetWmiObjectWin32VideoController | ?{$_.AdapterRAM -gt 0} | %{
    $Item = New-Object System.Windows.Forms.ListViewItem("GPU")
    $Item.SubItems.Add($_.Name)}
    $lvMain.Items.Add($Item)}

    if ($XML.Options.SystemInfo.OS.Enabled -eq $true){
    $Item = New-Object System.Windows.Forms.ListViewItem("Operating System")
    $Item.BackColor = "Black"
    $Item.ForeColor = "White"
    $lvMain.Items.Add($Item)}
    
    if($XML.Options.SystemInfo.OS.OS.Enabled -eq $true){
    $Item = New-Object System.Windows.Forms.ListViewItem("OS Name")
    $Item.SubItems.Add($GetWmiObjectWin32OperatingSystem.Caption)
    $lvMain.Items.Add($Item)}
				
    if($XML.Options.SystemInfo.OS.ServicePack.Enabled -eq $true){
    $Item = New-Object System.Windows.Forms.ListViewItem("Service Pack")
    $Item.SubItems.Add($GetWmiObjectWin32OperatingSystem.CSDVersion)
    $lvMain.Items.Add($Item)}
    
    if($XML.Options.SystemInfo.OS.Architecture.Enabled -eq $true){
    $Item = New-Object System.Windows.Forms.ListViewItem("OS Architecture")
    $Item.SubItems.Add($GetWmiObjectWin32ComputerSystem.SystemType)
    $lvMain.Items.Add($Item)}
    
    if($XML.Options.SystemInfo.OS.ImageDate.Enabled -eq $true){
    $InstallDate = [System.Management.ManagementDateTimeconverter]::ToDateTime($GetWmiObjectWin32OperatingSystem.InstallDate).ToString()
    $Item = New-Object System.Windows.Forms.ListViewItem("Install Date")
    $Item.SubItems.Add($InstallDate)
    $lvMain.Items.Add($Item)}
			
    if ($XML.Options.SystemInfo.IPConfig.Enabled -eq $true){
    $Item = New-Object System.Windows.Forms.ListViewItem("Network Adapters")
    $Item.BackColor = "Black"
    $Item.ForeColor = "White"
    $lvMain.Items.Add($Item)}
    
    $GetWmiObjectWin32NetworkAdapter | %{
    if($XML.Options.SystemInfo.IPConfig.Description.Enabled -eq $true){
    $Item = New-Object System.Windows.Forms.ListViewItem("Description")
    $Item.SubItems.Add($_.Description)
    $lvMain.Items.Add($Item)}
    
    if($XML.Options.SystemInfo.IPConfig.IPAddress.Enabled -eq $true){
    $IPinfo = $null
    ForEach ($IP in $_.IPAddress){$IPinfo += "$IP "}
    $Item = New-Object System.Windows.Forms.ListViewItem("IP Address")
    $Item.SubItems.Add($IPinfo)
    $lvMain.Items.Add($Item)}
					
    if($XML.Options.SystemInfo.IPConfig.MACAddress.Enabled -eq $true){
    $Item = New-Object System.Windows.Forms.ListViewItem("MAC Address")
    $Item.SubItems.Add($_.MACAddress)
    $lvMain.Items.Add($Item)}
    
    if($XML.Options.SystemInfo.IPConfig.DHCPEnabled.Enabled -eq $true){
    $Item = New-Object System.Windows.Forms.ListViewItem("DHCP Enabled")
    $Item.SubItems.Add($_.DHCPEnabled.ToString())
    $lvMain.Items.Add($Item)}
    
    if($XML.Options.SystemInfo.IPConfig.DHCPServer.Enabled -eq $true){
    $Item = New-Object System.Windows.Forms.ListViewItem("DHCP Server")
    $Item.SubItems.Add($_.DHCPServer)
    $lvMain.Items.Add($Item)}
    
    if($XML.Options.SystemInfo.IPConfig.DNSDomain.Enabled -eq $true){
    $Item = New-Object System.Windows.Forms.ListViewItem("DNS Domain")
    $Item.SubItems.Add($_.DNSDomain)
    $lvMain.Items.Add($Item)}}}


    $lvMain.Columns[0].Width = "240"
    $lvMain.Columns[1].Width = ($lvMain.Width - ($lvMain.Columns[0].Width + 22))}
   
   function QuickFixGetStartUp { 
   # Downloads and generates of startup, can disable startup items 
   $QueryGetStartUp = $TextboxSearchbarComputers.Text
   $TextboxStatusboxComputer.text = " Downloading Startup-Information From $QueryGetStartUp , takes a few seconds"
   Start-Sleep -S 3
   ClearListview
   $XML.Options.StartupItems.Property | %{GetColumn $_}
   ResizeColumns
   $Col0 = $lvMain.Columns[0].Text
   $GetWmiObjectwin32StartupCommand = Get-WmiObject win32_StartupCommand -ComputerName $QueryGetStartUp | Sort Name
   $GetWmiObjectwin32StartupCommand | %{
   $Item = New-Object System.Windows.Forms.ListViewItem($_.$Col0)
   $TextboxStatusboxComputer.text = " Done, A Startup-Information-List Is Now Made From  $QueryGetStartUp " 
   ForEach ($Col in ($lvMain.Columns | ?{$_.Index -ne 0})){$Field = $Col.Text;$Item.SubItems.Add($_.$Field)}
   $lvMain.Items.Add($Item)}}
   
   function activedirectorysearch { 
   $TextboxStatusboxComputer.text = " Searching AD, takes a few seconds " 
   Start-Sleep -S 3
   # button to search the AD 
   import-module activedirectory
   $QuerySearchHostnameIp = $TextboxSearchbarComputers.Text
   $TextboxSearchbarComputers = $TextboxSearchbarComputers.Text
   $GetAdComputer = Get-ADComputer -Identity $TextboxSearchbarComputers -Properties Name, Cn, DNSHostName, CanonicalName, SamAccountName, whenChanged, whenCreated, Enabled, OperatingSystem, OperatingSystemServicePack, OperatingSystemVersion, IPv4Address, LastLogonDate, LastLogon | 
   Format-list Name, Cn, DNSHostName, CanonicalName, SamAccountName, whenChanged, whenCreated, Enabled, OperatingSystem, OperatingSystemServicePack, OperatingSystemVersion, IPv4Address, LastLogonDate, @{N='LastLogon'; E={[DateTime]::FromFileTime($_.LastLogon)}} | out-string;    
   $TextboxStatusboxComputer.text = $GetAdComputer
   } 
   
 
   
   function ClearListview{
   $lvMain.Items.Clear()
   $lvMain.Columns.Clear()}   
   
   
   function GetColumn{
   Param([String]$Column)
   # Adding $Column from XML file 
   $lvMain.Columns.Add($Column)
   }
   
   function ResizeColumns{
   # Resizing columns based on column count
   $ColWidth = (($lvMain.Width / ($lvMain.Columns).Count) - 11)
   $lvMain.Columns | %{$_.Width = $ColWidth}
   }
                               # Function End
                               # Computer Tab End
          #----------------------------------------------------------------------#   
                               # Computer Tab End
          #----------------------------------------------------------------------#   

   

#----------------------------------------------
# Object Gui - Mainform - Generate Form
#----------------------------------------------
$MainForm.Add_Shown({$MainForm.Activate()})
[void] $MainForm.ShowDialog() 