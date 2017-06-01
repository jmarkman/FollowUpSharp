# FollowUpSharp

Reimplementation of [qfuAuto](https://github.com/jmarkman/qfuAuto) in C#.
Introduction of GUI to [test project](https://github.com/jmarkman/FollowUpSharp_Test), shifting from "test project" to a project that will ultimately be deployed

## Updates

### 3/21 - 3/22/2017
- Green light given to utilize System.Net.Mail aka SmtpClient & MailMessage (finally) and scrap Mail Merge plan, SMTP functionality implemented and tested
- Refactored several classes to reduce the number of methods used since the [test project](https://github.com/jmarkman/FollowUpSharp_Test); some methods were rendered redundant and were subsequently deleted, overall codebase has been minimized
- Added "Insured" object to represent the row of entries in the SQL DB
- Made the default filepath a private class variable in ExcelWriter (TODO: Ask someone if this is a bad thing to do)

### 3/23 - 4/3/2017
- SMTP canned by IT company (didn't know that that 90s-era spam via SMTP was still a problem). Switching to Outlook Interop, which might've been the best choice from the start
- Removed progress bar. Interop usage is actually REALLY fast for some reason
- Cleaned up codebehind
- Made legacy SMTP branch
- Made MVVM branch for more-than-likely design shift from "whatever just type it" to something realistic
- Made "attach file" input section for future user-specified document attaching

### 4/3/2017 - 4/13/2017
- Progress Bar re-added to UI via understanding of BackgroundWorker parameters despite speed of Interop in this case
- Created error logging code snippets for debugging purposes in case users run into program-halting errors
- Removed MVVM branch following advisement of what kind of scale that design patterns like MVVM effectively support
- Reviewed protection levels of variables and classes
- **[Interesting]** Downgraded .NET Framework version from 4.X+ to 3.5 to immediately support Windows 7 without having to install a 4.X+ version of the .NET Framework on the Windows 7 computers, which are the majority of the office (only two computers run Windows 10, and only one of those are active)
- Added DocComments to more methods and various comments, will have to scale back as I get closer to a deployable product
- Removed unnecessary "using" statements

###  4/13/2017 - 4/24/2017
- Changed Insured class to IMSEntry to more accurately represent what exactly was being pulled from the DB
- Added more elements to the IMSEntry class for CYA purposes
- Added connection string to App.config
- Added encrypt/decrypt method for connection string

### 4/24/2017 - 5/17/2017
- Final once-over before initial deployment
- Updated ExcelWriter to reflect the information to be returned from the query
- Updated ExcelWriter to add a bit more formatting (colored header)
- Added TODOs for future
- Left stored procedure placeholders in place of queries; live version will unfortunately have the queries hardcoded in
- Created installer via inno setup (not on git)
- Added program icon
- Program no longer CC's underwriters (only commented out because desire to do so from company might rise)
- First live test in the wild tomorrow!

### 5/17/2017 - 5/18/2017
Went live! Based on feedback:
- Pulling data from another column in the database
- Minor fix to Excel sheet record keeping printout
- (Not programming) Update to follow-up procedure noted  
Need to prepare usage documentation!

### 5/18/2017 - 6/1/2017
![alt text][doh] 
Remade Github project, discovered instances of commits with sensitive connection info and not just ActiveDirectory connect string 
- Fixed issue with program halting without specifying failure when given blank emails, program continues if email provided is blank/null
- Made directory paths as relative as possible with Environment.GetFolderPath()
- Added minor instructional help file

[doh]: https://board.en.ogame.gameforge.com/wcf/images/avatars/8e/1650-8eb373f55056138a90628514d78fd58bd3ad24bd-128.gif "D'oh!"