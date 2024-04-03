Attribute VB_Name = "M2_Exam_GetWellData"
Sub ExamGetWellData()
    Dim chromePath As String
    Dim ChromeOptions As New selenium.ChromeOptions
    Dim driver As New selenium.ChromeDriver
    
    chromePath = "C:\Path\To\chromedriver.exe" ' Update with the path to your chromedriver.exe
    
    ' Set Chrome options
    ChromeOptions.AddArgument "remote-debugging-port=9222"
    
    ' Start Chrome driver with options
    driver.Start "chrome", chromePath, ChromeOptions
    
    ' Navigate to the webpage
    driver.Get "about:blank" ' Update with the URL of the webpage you want to navigate to
    
    ' Wait for the page to load (you may need to implement a wait condition)
    Application.Wait Now + TimeValue("0:00:05") ' Wait for 5 seconds (adjust as needed)
    
    ' Find the elements and interact with the webpage here
    
    ' Close the browser
    driver.Quit
End Sub

