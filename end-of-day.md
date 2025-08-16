## 🏭 End-of-Day Workflow

This document outlines the steps taken by the system at the end of each business day to ensure inventory levels are accurate and replenishment is triggered when necessary.

---

## 📚 KMDS D1 Menu (`Form_EndOfDay.Form_Open`)

- **Initial Setup**
    - Creates an ADODB recordset to read from the `qryCashRegisters` query.
        ```sql
        SELECT CashSys.ID, CashSys.ComputerName, CashSys.Description, Sysfile.Cr1Installed, Sysfile.Cr2Installed, Sysfile.Cr3Installed, Sysfile.Cr4Installed, Sysfile.Cr5Installed, Sysfile.Cr1CheckOut, Sysfile.Cr2CheckOut, Sysfile.Cr3CheckOut, Sysfile.Cr4CheckOut, Sysfile.Cr5CheckOut
        FROM CashSys, Sysfile
        ORDER BY CashSys.ID;
        ```
- **Register Check**
  - Loops through registers 1 to 5.
  - Checks if any register is checked out (Cr1CheckOut, Cr2CheckOut, etc.).
  - Sets `bAtLeastOneRegisterCheckedOut` to True if any are checked out.
    ```vb
    Select Case !ID
        Case 1
        If !Cr1CheckOut = True Then bAtLeastOneRegisterCheckedOut = True
    ```
- **End of Day Trigger**
  - If at least one register is checked out:
    ```vb
    If bAtLeastOneRegisterCheckedOut Then
    ```
    - Calls **EndOFDayStream** (starts the end of day process).
    - Runs a logout batch file (`kmdslogout.bat`).
        ```vb
        Shell "c:\kmds\kmdslogout.bat", vbNormalFocus
        ```
    - Quits Access (`DoCmd.Quit`).

```mermaid
flowchart LR
    Start["START <br> (D1)"] --> CheckCondition 
    CheckCondition{Checked out?}  
        CheckCondition -- Yes --> EndOFDayStream["Process EOD <br> (EndOFDayStream)"]
        click EndOFDayStream "#EndOFDayStream" "Go to EndOFDayStream"
            EndOFDayStream --> LogOut[Logout]
                LogOut --> Quit[Quit KMDS]
        CheckCondition -- No --> ContinueProcess["Cancel <br> (There are no registers checked out!)"]
```

---

## 🌉 EndOFDayStream (`modEOD.EndOFDayStream`) {#EndOFDayStream}

#### 📝Logging and Setup (`modEOD.EOLog`) {#EOLog}

  - Loads default settings if `gstrMailDir` is empty.
      ```vb
      If gstrMailDir = "" Then LoadDefaults
      ```
      The **`modKMDS.LoadDefaults`** function initializes global variables and system settings for the application.
      ```vb
      rstemp.Open "SysFile", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
      gstrMailDir = !NetworkMailDir
      ```
  - Builds a folder path for the log file using `gstrMailDir` and the current date (year/month/day).
  - Opens (or creates) the log file `EOLog.Txt` in append mode.
      ```vb
      c:\KMDS\reports\2025\08\01\EOLog.Txt
      ```
  - Logs the start of the EOD process.
      ```vb
      EOLog "========================================================="
      EOLog "End of Day - Start"
      EOLog "========================================================="
      ```
  - Loads default configuration values (`gstrMailDir`, `gstrNetworkDatabase`) if missing.
      
      ```vb
      modKMDS.LoadDefaults
      ```
  - Logs the mail directory and network database paths.


```mermaid
flowchart TD
    Start["Start <br> (modEOD.EndOFDayStream)"] --> LoadDefaults

    subgraph LoadDefaults[modEOD.LoadDefaults]
        direction LR
        SysFile --> gstrMailDir[gstrMailDir = !NetworkMailDir]
        SysFile --> gstrNetworkDatabase[gstrNetworkDatabase = !NetworkDatabase] 
    end

    Start --> EOLog
    subgraph EOLog[modEOD.EOLog]
        direction LR
        CheckFolder{Folder Created?}
            CheckFolder -- NO --> CreateFolder[Create Folder]
            CreateFolder --> CheckFile
            CheckFolder -- YES --> CheckFile{File Created?}
                CheckFile -- NO --> CreateFile[Create File]
                CheckFile -- YES --> UpdateFile[Update File]
                CreateFile --> UpdateFile
    end

    Start --> CreateBackup[Create SQL Backup]
    CreateBackup --> MakeSQLBackup
    subgraph MakeSQLBackup[modEOD.MakeSQLBackup]
        direction LR
        A --> B[gstrMailDir = !NetworkMailDir]
        B --> C[gstrNetworkDatabase = !NetworkDatabase] 
    end

    CreateBackup --> DeleteFiles

    DeleteFiles --> G

    G --> F

```







---

# Below are sample layout

```mermaid
flowchart TD
    A[Start of Day-End Process] --> B{Is inventory below threshold?}
    B -- Yes --> C[Reorder stock]
    B -- No --> D[Skip reorder]
    C --> E[Notify supplier]
    D --> E
    E --> F[Log action]
```

```mermaid
sequenceDiagram
    participant User
    participant Access
    User->>Access: Click EndOfDay 
```

```mermaid
flowchart TD
    A[Start] --> B{Is it raining?}
    B -- Yes --> C[Take an umbrella]
    B -- No --> D[Go outside]
    C --> E[Continue day]
    D --> E

    %% Simulated note
    Note1["Note: Check weather app before deciding"]
    B --> Note1
    style Note1 fill:#f9f,stroke:#333,stroke-width:1px
```

```mermaid 
sequenceDiagram
    participant User
    participant System

    User->>System: Request data
    Note right of System: If data is valid
    System-->>User: Return result

    Note right of System: Else show error
    System-->>User: Display error message

```