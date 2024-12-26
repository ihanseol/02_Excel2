Attribute VB_Name = "Module_ChromeDriver"
Function Is64BitOS() As Boolean
    Dim osInfo As String
    osInfo = Application.OperatingSystem
    
    If InStr(osInfo, "64") > 0 Then
        Is64BitOS = True
    Else
        Is64BitOS = False
    End If
End Function


Sub OpenSpecificChromeVersion()
  Const chromePath = "c:\ProgramData\00_chrome\chrome.exe"
  Dim objShell As Object
  Set objShell = CreateObject("Wscript.Shell")

  objShell.Run Chr(34) & chromePath & Chr(34)
  Set objShell = Nothing
End Sub


Sub RunChrome_SpecificLocationOfChrome()
    Dim driver As New Selenium.chromeDriver
    Dim url As String
    
    ' 2024/12/26
    ' this is the key of run specific version of chrome
    ' driver.SetBinary "c:\ProgramData\00_chrome\chrome.exe" ' Update this path

    If Is64BitOS() Then
        Set SystemProperties = CreateObject("WScript.Shell").Environment("System")
        SystemProperties("webdriver.chrome.driver") = driverPath  ' For 64-bit systems
    Else
        SystemProperties("webdriver.chrome.driver") = driverPath  ' For 32-bit systems
    End If
    
    driver.AddArgument "--remote-debugging-port=9222"  ' Specify the port number
    'driver.AddArgument "--user-data-dir=C:\Users\minhwasoo\AppData\Local\Google\Chrome\User Data"  ' Specify the user data directory
    driver.AddArgument "--disable-gpu"  ' Disable GPU acceleration
    driver.AddArgument "--window-size=1920,1080"  ' Specify the window size
        
    driver.SetBinary "c:\ProgramData\00_chrome\chrome.exe" ' Update this path
    url = "https://www.google.com" ' Update this URL
   
    driver.Start
    driver.Get url
  
    Sleep (7 * 1000)
    driver.Quit
End Sub


Sub RunChromeWithDownloadDirectory()
   
   ' https://stackoverflow.com/questions/63424206/selenium-chromedriver-excel-vba-change-the-download-directory-multiple-time
    
    Dim driver As New Selenium.chromeDriver
    
    driver.SetPreference "download.default_directory", "C:\Folder\"
    driver.SetPreference "download.directory_upgrade", True
    driver.SetPreference "download.prompt_for_download", False

    driver.Get "https://www.google.com/"

End Sub




   





