# Export-ADUserPropertyFrequencies_To-Html.ps1

---

### SYNOPSIS
Analyze all Active Directory users by frequency of assigned properties. <br>
Export Analysis results to HTML report. <br>
Find out how many AD users are assigned a specific property. <br>
Find out how many unique property assignments exist for a specific property.

### DESCRIPTION
Analyze all Active Directory users by frequency of assigned properties. <br>
Export Analysis results to HTML report.

### PARAMETER Server
Specifies the Active Directory server domain to query.

### PARAMETER Enabled
Specifies to only return results from enabled users.

### EXAMPLE
PS C:\> Export-ADUserPropertyFrequencies_To-Html.ps1 -Server CyberCondor.local -Enabled $true

## Input
Template.html
- Template for HTML layout 
styles.css
- Styles to display HTML layout properly
Users
- Contains list of all users and their properties

## Output
~\_FrequenyAnalysisReports\ADUserPropertiesFrequenyAnalysisReport_$Date
- Users.html
- styles.css

---

# Details

This PowerShell script analyzes all Active Directory (AD) users by frequency of assigned properties. The script exports the analysis results to an HTML report. The script takes two mandatory parameters:

1.  `Server`: Specifies the Active Directory server domain to query.
2.  `Enabled`: Specifies to only return results from enabled users.

The script:

1.  Writes a message to the console, indicating that it is attempting to query the Active Directory.
    
2.  Tries to run the `Get-ADUser` command to verify that the Active Directory module is installed and the specified server is reachable. If the module is not installed or the server is not reachable, it returns a warning message indicating the cause of the error.
    
3.  Defines the following functions:
    
    -   `Get-ExistingUsers_AD`: Queries the Active Directory to get a list of all existing users and their properties.
    -   `Get-UserRunningThisProgram`: Gets the user who is running the script.
    -   `SanitizeManagerPropertyFormat`: Sanitizes the format of the `Manager` property of each existing user.
    -   `Create-NewHtmlReportFromTemplate`: Creates a new HTML report from a template file.
4.  Calls the `Get-ExistingUsers_AD` function to get the list of existing users and their properties.
    
5.  Calls the `Get-UserRunningThisProgram` function to get the user who is running the script.
    
6.  Calls the `SanitizeManagerPropertyFormat` function to sanitize the format of the `Manager` property of each existing user.
    
7.  Calls the `Create-NewHtmlReportFromTemplate` function to create a new HTML report from a template file.
    
8.  Writes a message to the console indicating the location of the generated HTML report.

---

# Get-PropertyFrequencies

This function, named "Get-PropertyFrequencies", is a PowerShell script that takes two input arguments: `$Property` and `$Object`. The purpose of the function is to determine the frequency of each unique value of a specified property within the given object.

The first steps of the function involve initializing variables and obtaining a list of all unique values for the specified property within the object. If the property is of type "DateTime", the function sets a flag `$isDate` to `$true`.

Next, the function loops through each unique value, creating a new object for each value with properties "Property" (the unique value), "Count" (initialized to 0), and "Frequency" (initialized to "100%").

The function then determines if the property is of type "DateTime". If it is, the function formats the unique values as "yyyy-MM" strings.

Afterwards, the function loops through each value of the specified property in the object. If the value is equal to a unique value, the corresponding "Count" property of the object is incremented. If the property is of type "DateTime", the function also checks for equality after converting both the value and the unique value to strings in the "yyyy-MM" format.

At the end, the function computes the frequency for each unique value as the count divided by the total number of values and returns an array of objects sorted by count and property.

Note: The function also includes a "Write-Progress" command which outputs a progress bar during the execution of the script.
