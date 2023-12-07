#include <Windows.h>
#include <wuapi.h>
#include <iostream>
#include <fstream>
#include <string>
#include <ctime>
#include <chrono>
#include <sstream>
#include <locale>
#include <cstdio>
#include <array>

//#include "json11.hpp"

//using namespace json11;

/*Json GetUpdateJson(IUpdate* update)
{
    Json updateJson = Json::object{};

    BSTR title;
    update->get_Title(&title);
    updateJson["Title"] = L"" + title;
    SysFreeString(title);

    

    return updateJson;
}*/

std::wstring getTimeString() {
  std::time_t currentTime;
  std::tm localTime;
  char timeString[26];

  std::time(&currentTime);
  localtime_s(&localTime, &currentTime);
  asctime_s(timeString, sizeof(timeString), &localTime);
  timeString[strlen(timeString) - 1] = '\0';
  std::wstringstream wss;
  wss << L"[" << timeString << L"] ";

  return wss.str();
}

int main()
{
    std::wstring filename = L"LogSearchUpdateWindow.txt";
    std::wofstream file(filename, std::ios::app);
    if (file.is_open()) {
        HRESULT hr = CoInitialize(NULL);
        IUpdateSession* updateSession = NULL;
        IUpdateSearcher* updateSearcher = NULL;
        BSTR newServiceID = NULL;
        ISearchResult* searchResult = NULL;
        BSTR criteria = NULL;
        IUpdateCollection* updateCollection = NULL;

        IStringCollection* KBID = NULL;
        BSTR bstrKbId = NULL;
        BSTR title = NULL;
        VARIANT_BOOL autoSelectOnWebSites = NULL;
        VARIANT_BOOL canRequireSource = NULL;
        VARIANT_BOOL deltaCompressedContentAvailable = NULL;
        VARIANT_BOOL deltaCompressedContentPreferred = NULL;
        BSTR description = NULL;
        VARIANT_BOOL eulaAccepted = NULL;
        VARIANT_BOOL isBeta = NULL;
        VARIANT_BOOL isDownloaded = NULL;
        VARIANT_BOOL isHidden = NULL;
        VARIANT_BOOL isInstalled = NULL;
        VARIANT_BOOL isMandatory = NULL;
        VARIANT_BOOL isUninstallable = NULL;
        DATE lastDeploymentChangeTime = NULL;
        BSTR size = NULL;
        BSTR supportUrl = NULL;
        IUpdate5* update5 = NULL;
        VARIANT_BOOL rebootRequired = NULL;
        VARIANT_BOOL isPresent = NULL;
        VARIANT_BOOL browseOnly = NULL;
        VARIANT_BOOL perUser = NULL;

        std::wstring serviceIDString = L"9482F4B4-E343-43B6-B170-9A65BC822C77";
        std::wstring criteriaString = L"(IsInstalled=0 AND IsHidden=1) OR (IsInstalled=0 AND IsHidden=0)";
        std::wstring a;

        std::wcout.imbue(std::locale(""));
        file.imbue(std::locale(""));

        std::wstring command;
        std::wcout << getTimeString() << L"Run command: Get Window version" << std::endl;
        file << getTimeString() << L"Run command: Get Window version" << std::endl;
        command = L"powershell -Command \"& { systeminfo | findstr /B /C:'OS Name' /B /C:'OS Version' }\" 2>&1";
        FILE* pipe4 = _wpopen(command.c_str(), L"rt, ccs=UTF-8");
        if (pipe4 == nullptr) {
          throw std::runtime_error("Error opening pipe to PowerShell command");
        }
        std::array<wchar_t, 128> buffer4;
        while (fgetws(buffer4.data(), static_cast<int>(buffer4.size()), pipe4) != nullptr) {
          std::wcout << buffer4.data();
          file << buffer4.data();
        }
        if (_pclose(pipe4) == -1) {
          throw std::runtime_error("Error closing pipe to PowerShell command");
        }
        std::wcout << getTimeString() << L"Done" << std::endl;
        file << getTimeString() << L"Done" << std::endl;


        /*while (1)
        {
            std::wcout << L"Type Service ID: ";
            std::wcin >> serviceIDString;*/

            /*std::wcout << L"Type query: ";
            std::wcin.ignore();
            std::getline(std::wcin, criteriaString);*/

            std::wcout << L"\n\n" << getTimeString() << "========================Starting find missing patchs========================\n\n";
            file << L"\n\n" << getTimeString() << "=======================Starting find missing patchs========================\n\n";

            hr = CoCreateInstance(__uuidof(UpdateSession), NULL, CLSCTX_INPROC_SERVER, __uuidof(IUpdateSession), (LPVOID*)&updateSession);

            
            hr = updateSession->CreateUpdateSearcher(&updateSearcher);

            BSTR serviceID = SysAllocString(serviceIDString.c_str());// SysAllocString(L"3DA21691-E39D-4da6-8A4B-B43877BCB1B7");
            hr = updateSearcher->put_ServiceID(serviceID);
            std::wcout << getTimeString() << "Service ID  entered: " << serviceID << std::endl;
            file << getTimeString() << "Service ID entered: " << serviceID << std::endl;

            
            hr = updateSearcher->get_ServiceID(&newServiceID);
            if (serviceID != NULL)
            {
                std::wcout << getTimeString() << "Service ID applied: " << newServiceID << std::endl;
                file << getTimeString() << "Service ID applied: " << newServiceID << std::endl;
            }
            else
            {
                std::wcout << getTimeString() << "Service ID applied: " << std::endl;
                file << getTimeString() << "Service ID applied: " << std::endl;
            }

            
            if (criteriaString.empty()) {
              std::wcout << getTimeString() << "Use default query: (IsInstalled=0 AND IsHidden=1) OR (IsInstalled=0 AND IsHidden=0) OR (IsInstalled=0 AND DeploymentAction=*)" << std::endl;
              file << getTimeString() << "Use default query: (IsInstalled=0 AND IsHidden=1) OR (IsInstalled=0 AND IsHidden=0) OR (IsInstalled=0 AND DeploymentAction=*)" << std::endl;
              criteria = SysAllocString(L"(IsInstalled=0 AND IsHidden=1) OR (IsInstalled=0 AND IsHidden=0) OR (IsInstalled=0 AND DeploymentAction=*)");
            }
            else {
              std::wcout << getTimeString() << "Use custorm query: " << criteriaString << std::endl;
              file << getTimeString() << "Use custorm query: " << criteriaString << std::endl;
              criteria = SysAllocString(criteriaString.c_str());
            }
            //BSTR criteria = SysAllocString(L"IsInstalled=0 and DeploymentAction='Installation' or IsInstalled=0 and DeploymentAction='OptionalInstallation' or IsPresent=1 and DeploymentAction='Uninstallation' or IsInstalled=1 and DeploymentAction='Installation' and RebootRequired=1 or IsInstalled=0 and DeploymentAction='Uninstallation' and RebootRequired=1");
            hr = updateSearcher->Search(criteria, &searchResult);
            if (hr != 0) {
              std::wcout << getTimeString() << "Invalid query!!!" << std::endl;
              file << getTimeString() << "Invalid query!!!" << std::endl;
              std::wcout << getTimeString() << "Use default query: (IsInstalled=0 AND IsHidden=1) OR (IsInstalled=0 AND IsHidden=0) OR (IsInstalled=0 AND DeploymentAction=*)" << std::endl;
              file << getTimeString() << "Use default query: (IsInstalled=0 AND IsHidden=1) OR (IsInstalled=0 AND IsHidden=0) OR (IsInstalled=0 AND DeploymentAction=*)" << std::endl;
              criteria = SysAllocString(L"(IsInstalled=0 AND IsHidden=1) OR (IsInstalled=0 AND IsHidden=0) OR (IsInstalled=0 AND DeploymentAction=*)");
              hr = updateSearcher->Search(criteria, &searchResult);
            }
            
            hr = searchResult->get_Updates(&updateCollection);
            LONG count;
            updateCollection->get_Count(&count);
            std::wcout << getTimeString() << L"Found " << count << " missing patch(s)\n" << std::endl;
            file << getTimeString() << "Found " << count << " missing patch(s)\n" << std::endl;

            if (count > 0) {
                for (long i = 0; i < count; ++i) {

                    std::wcout << L"\n" << getTimeString() << "========================Patch: " << i + 1 << "========================" << std::endl;
                    file << L"\n" << getTimeString() << "========================Patch: " << i +1 << "========================" << std::endl;
                    IUpdate5* update;
                    hr = updateCollection->get_Item(i,(IUpdate**)&update);
                    // Get update information
                    hr = update->get_KBArticleIDs(&KBID);
                    hr = KBID->get_Item(0, &bstrKbId);
                    if (bstrKbId != NULL)
                    {
                      std::wcout << getTimeString() << "KB ID: KB" << bstrKbId << std::endl;
                      file << getTimeString() << "KB ID: KB" << bstrKbId << std::endl;
                    }
                    else
                    {
                      std::wcout << getTimeString() << "KB ID: Unknown" << std::endl;
                      file << getTimeString() << "KB ID: Unknown" << std::endl;
                    }

                    hr = update->get_Title(&title);
                    if (title != NULL)
                    {
                        std::wcout << getTimeString() << "Title: " << title << std::endl;
                        file << getTimeString() << "Title: " << title << std::endl;
                    }
                    else
                    {
                        std::wcout << getTimeString() << "Title: " << std::endl;
                        file << getTimeString() << "Title: " << std::endl;
                    }
                                        
                    hr = update->get_AutoSelectOnWebSites(&autoSelectOnWebSites);
                    std::wcout << getTimeString() << "AutoSelectOnWebSites: " << (autoSelectOnWebSites ? "True" : "False") << std::endl;
                    file << getTimeString() << "AutoSelectOnWebSites: " << (autoSelectOnWebSites ? "True" : "False") << std::endl;

                    hr = update->get_CanRequireSource(&canRequireSource);
                    std::wcout << getTimeString() << "CanRequireSource: " << (canRequireSource ? "True" : "False") << std::endl;
                    file << getTimeString() << "CanRequireSource: " << (canRequireSource ? "True" : "False") << std::endl;
                                        
                    hr = update->get_DeltaCompressedContentAvailable(&deltaCompressedContentAvailable);
                    std::wcout << getTimeString() << "DeltaCompressedContentAvailable: " << (deltaCompressedContentAvailable ? "True" : "False") << std::endl;
                    file << getTimeString() << "DeltaCompressedContentAvailable: " << (deltaCompressedContentAvailable ? "True" : "False") << std::endl;

                    hr = update->get_DeltaCompressedContentPreferred(&deltaCompressedContentPreferred);
                    std::wcout << getTimeString() << "DeltaCompressedContentPreferred: " << (deltaCompressedContentPreferred ? "True" : "False") << std::endl;
                    file << getTimeString() << "DeltaCompressedContentPreferred: " << (deltaCompressedContentPreferred ? "True" : "False") << std::endl;

                    hr = update->get_Description(&description);
                    if (description != NULL)
                    {
                        std::wcout << getTimeString() << L"Description: " << description << std::endl;
                        file << getTimeString() << L"Description: " << description << std::endl;
                    }
                    else
                    {
                        std::wcout << getTimeString() << "Description: " << std::endl;
                        file << getTimeString() << "Description: " << std::endl;
                    }

                    hr = update->get_EulaAccepted(&eulaAccepted);
                    std::wcout << getTimeString() << "EulaAccepted: " << (eulaAccepted ? "True" : "False") << std::endl;
                    file << getTimeString() << "EulaAccepted: " << (eulaAccepted ? "True" : "False") << std::endl;

                    hr = update->get_IsBeta(&isBeta);
                    std::wcout << getTimeString() << "IsBeta: " << (isBeta ? "True" : "False") << std::endl;
                    file << getTimeString() << "IsBeta: " << (isBeta ? "True" : "False") << std::endl;

                    hr = update->get_IsDownloaded(&isDownloaded);
                    std::wcout << getTimeString() << "IsDownloaded: " << (isDownloaded ? "True" : "False") << std::endl;
                    file << getTimeString() << "IsDownloaded: " << (isDownloaded ? "True" : "False") << std::endl;

                    hr = update->get_IsHidden(&isHidden);
                    std::wcout << getTimeString() << "IsHidden: " << (isHidden ? "True" : "False") << std::endl;
                    file << getTimeString() << "IsHidden: " << (isHidden ? "True" : "False") << std::endl;

                    hr = update->get_IsInstalled(&isInstalled);
                    std::wcout << getTimeString() << "IsInstalled: " << (isInstalled ? "True" : "False") << std::endl;
                    file << getTimeString() << "IsInstalled: " << (isInstalled ? "True" : "False") << std::endl;

                    hr = update->get_IsMandatory(&isMandatory);
                    std::wcout << getTimeString() << "IsMandatory: " << (isMandatory ? "True" : "False") << std::endl;
                    file << getTimeString() << "IsMandatory: " << (isMandatory ? "True" : "False") << std::endl;

                    hr = update->get_IsUninstallable(&isUninstallable);
                    std::wcout << getTimeString() << "IsUninstallable: " << (isUninstallable ? "True" : "False") << std::endl;
                    file << getTimeString() << "IsUninstallable: " << (isUninstallable ? "True" : "False") << std::endl;

                    hr = update->get_LastDeploymentChangeTime(&lastDeploymentChangeTime);
                    SYSTEMTIME sysTime;
                    VariantTimeToSystemTime(lastDeploymentChangeTime, &sysTime);

                    // Format the date portion
                    wchar_t dateFormat[50];
                    GetDateFormatW(LOCALE_USER_DEFAULT, 0, &sysTime, L"MMMM dd, yyyy", dateFormat, 50);

                    // Format the time portion
                    wchar_t timeFormat[50];
                    GetTimeFormatW(LOCALE_USER_DEFAULT, 0, &sysTime, L"hh:mm:ss tt", timeFormat, 50);
                    if (lastDeploymentChangeTime != NULL)
                    {
                        std::wcout << getTimeString() << "LastDeploymentChangeTime: " << dateFormat << " " << timeFormat << std::endl;
                        file << getTimeString() << "LastDeploymentChangeTime: " << dateFormat << " " << timeFormat << std::endl;
                    }
                    else
                    {
                        std::wcout << getTimeString() << "LastDeploymentChangeTime: " << std::endl;
                        file << getTimeString() << "LastDeploymentChangeTime: " << std::endl;
                    }

                    DECIMAL maxDownloadSize;
                    hr = update->get_MaxDownloadSize(&maxDownloadSize);
                    VarBstrFromDec(&maxDownloadSize, LOCALE_USER_DEFAULT, 0, &size);
                    std::wcout << getTimeString() << "MaxDownloadSize: " << size << std::endl;
                    file << getTimeString() << "MaxDownloadSize: " << size << std::endl;

                    DECIMAL minDownloadSize;
                    hr = update->get_MinDownloadSize(&minDownloadSize);
                    VarBstrFromDec(&minDownloadSize, LOCALE_USER_DEFAULT, 0, &size);
                    std::wcout << getTimeString() << "MinDownloadSize: " << size << std::endl;
                    file << getTimeString() << "MinDownloadSize: " << size << std::endl;

                    LONG recommendedCpuSpeed;
                    hr = update->get_RecommendedCpuSpeed(&recommendedCpuSpeed);
                    std::wcout << getTimeString() << "RecommendedCpuSpeed: " << recommendedCpuSpeed << std::endl;
                    file << getTimeString() << "RecommendedCpuSpeed: " << recommendedCpuSpeed << std::endl;

                    LONG recommendedHardDiskSpace;
                    hr = update->get_RecommendedHardDiskSpace(&recommendedHardDiskSpace);
                    std::wcout << getTimeString() << "RecommendedHardDiskSpace: " << recommendedHardDiskSpace << std::endl;
                    file << getTimeString() << "RecommendedHardDiskSpace: " << recommendedHardDiskSpace << std::endl;

                    LONG recommendedMemory;
                    hr = update->get_RecommendedMemory(&recommendedMemory);
                    std::wcout << getTimeString() << "RecommendedMemory: " << recommendedMemory << std::endl;
                    file << getTimeString() << "RecommendedMemory: " << recommendedMemory << std::endl;

                    hr = update->get_SupportUrl(&supportUrl);
                    if (supportUrl != NULL)
                    {
                        std::wcout << getTimeString() << "SupportURL: " << supportUrl << std::endl;
                        file << getTimeString() << "SupportURL: " << supportUrl << std::endl;
                    }
                    else
                    {
                        std::wcout << getTimeString() << "SupportURL: " << std::endl;
                        file << getTimeString() << "SupportURL: " << std::endl;
                    }

                    UpdateType type;
                    hr = update->get_Type(&type);
                    std::wcout << getTimeString() << "Type: " << type << std::endl;
                    file << getTimeString() << "Type: " << type << std::endl;

                    DeploymentAction deploymentAction = daNone;
                    hr = update->get_DeploymentAction(&deploymentAction);
                    std::wcout << getTimeString() << L"DeploymentAction: " << deploymentAction << std::endl;
                    file << getTimeString() << L"DeploymentAction: " << deploymentAction << std::endl;

                    DownloadPriority downloadPriority = dpNormal;
                    hr = update->get_DownloadPriority(&downloadPriority);
                    std::wcout << getTimeString() << "DownloadPriority: " << downloadPriority << std::endl;
                    file << getTimeString() << "DownloadPriority: " << downloadPriority << std::endl;

                    hr = update->QueryInterface(IID_IUpdate5, reinterpret_cast<LPVOID*>(&update5));

                    hr = update5->get_RebootRequired(&rebootRequired);
                    std::wcout << getTimeString() << "RebootRequired: " << (rebootRequired ? "True" : "False") << std::endl;
                    file << getTimeString() << "RebootRequired: " << (rebootRequired ? "True" : "False") << std::endl;

                    hr = update5->get_IsPresent(&isPresent);
                    std::wcout << getTimeString() << "IsPresent: " << (isPresent ? "True" : "False") << std::endl;
                    file << getTimeString() << "IsPresent: " << (isPresent ? "True" : "False") << std::endl;

                    hr = update5->get_BrowseOnly(&browseOnly);
                    std::wcout << getTimeString() << "BrowseOnly: " << (browseOnly ? "True" : "False") << std::endl;
                    file << getTimeString() << "BrowseOnly: " << (browseOnly ? "True" : "False") << std::endl;

                    hr = update5->get_PerUser(&browseOnly);
                    std::wcout << getTimeString() << "PerUser: " << (perUser ? "True" : "False") << std::endl;
                    file << getTimeString() << "PerUser: " << (perUser ? "True" : "False") << std::endl;

                    AutoSelectionMode autoSelection = asLetWindowsUpdateDecide;
                    hr = update5->get_AutoSelection(&autoSelection);
                    std::wcout << getTimeString() << "AutoSelection: " << autoSelection << std::endl;
                    file << getTimeString() << "AutoSelection: " << autoSelection << std::endl;

                    AutoDownloadMode autoDownload = adLetWindowsUpdateDecide;
                    hr = update5->get_AutoDownload(&autoDownload);
                    std::wcout << getTimeString() << "AutoDownload: " << autoDownload << std::endl;
                    file << getTimeString() << "AutoDownload: " << autoDownload << std::endl;

                    SysFreeString(title);
                    SysFreeString(description);
                    SysFreeString(supportUrl);

                    update->Release();
                    update5->Release();

                    update = NULL;
                    KBID = NULL;
                    title = NULL;
                    autoSelectOnWebSites = NULL;
                    canRequireSource = NULL;
                    deltaCompressedContentAvailable = NULL;
                    deltaCompressedContentPreferred = NULL;
                    description = NULL;
                    eulaAccepted = NULL;
                    isBeta = NULL;
                    isDownloaded = NULL;
                    isHidden = NULL;
                    isInstalled = NULL;
                    isMandatory = NULL;
                    isUninstallable = NULL;
                    lastDeploymentChangeTime = NULL;
                    size = NULL;
                    supportUrl = NULL;
                    update5 = NULL;
                    rebootRequired = NULL;
                    isPresent = NULL;
                    browseOnly = NULL;
                    perUser = NULL;
                    
                }
            }
            else {
                // No updates available
                std::wcout << getTimeString() << "No updates available" << std::endl;
                file << getTimeString() << "No updates available" << std::endl;
            }
            searchResult->Release();
            updateCollection->Release();

            updateSearcher->Release();
            updateSession->Release();
            CoUninitialize();

            hr = CoInitialize(NULL);
            updateSession = NULL;
            updateSearcher = NULL;
            newServiceID = NULL;
            searchResult = NULL;
            criteria = NULL;
            updateCollection = NULL;

            std::wcout << std::endl << std::endl;
            std::wcout << getTimeString() << L"=====================Start run power shell command=====================" << std::endl;
            file << getTimeString() << L"=====================Start run power shell command=====================" << std::endl;


            std::wcout << getTimeString() << L"Run command: Install-Module -Name PSWindowsUpdate -Force" << std::endl;
            file << getTimeString() << L"Run command: Install-Module -Name PSWindowsUpdate -Force" << std::endl;

            command = L"powershell -Command \"& { Install-Module -Name PSWindowsUpdate -Force }\" 2>&1";
            FILE* pipe = _wpopen(command.c_str(), L"rt, ccs=UTF-8");
            if (pipe == nullptr) {
              throw std::runtime_error("Error opening pipe to PowerShell command");
            }
            std::array<wchar_t, 128> buffer;
            while (fgetws(buffer.data(), static_cast<int>(buffer.size()), pipe) != nullptr) {
              std::wcout << buffer.data();
              file << buffer.data();
            }
            if (_pclose(pipe) == -1) {
              throw std::runtime_error("Error closing pipe to PowerShell command");
            }
            std::wcout << getTimeString() << L"Done" << std::endl;
            file << getTimeString() << L"Done" << std::endl;

            std::wcout << getTimeString() << L"Run command: Set-ExecutionPolicy Unrestricted -Force" << std::endl;
            file << getTimeString() << L"Run command: Set-ExecutionPolicy Unrestricted -Force" << std::endl;
            command = L"powershell -Command \"& { Set-ExecutionPolicy Unrestricted -Force }\" 2>&1";
            FILE* pipe2 = _wpopen(command.c_str(), L"rt, ccs=UTF-8");
            if (pipe2 == nullptr) {
              throw std::runtime_error("Error opening pipe to PowerShell command");
            }
            std::array<wchar_t, 128> buffer2;
            while (fgetws(buffer2.data(), static_cast<int>(buffer2.size()), pipe2) != nullptr) {
              std::wcout << buffer2.data();
              file << buffer2.data();
            }
            if (_pclose(pipe2) == -1) {
              throw std::runtime_error("Error closing pipe to PowerShell command");
            }
            std::wcout << getTimeString() << L"Done" << std::endl;
            file << getTimeString() << L"Done" << std::endl;

            std::wcout << getTimeString() << L"Run command: Get-WindowsUpdate" << std::endl;
            file << getTimeString() << L"Run command: Get-WindowsUpdate" << std::endl;
            command = L"powershell -Command \"& { Get-WindowsUpdate }\" 2>&1";
            FILE* pipe3 = _wpopen(command.c_str(), L"rt, ccs=UTF-8");
            if (pipe3 == nullptr) {
              throw std::runtime_error("Error opening pipe to PowerShell command");
            }
            std::array<wchar_t, 128> buffer3;
            while (fgetws(buffer3.data(), static_cast<int>(buffer3.size()), pipe3) != nullptr) {
              std::wcout << buffer3.data();
              file << buffer3.data();
            }
            if (_pclose(pipe3) == -1) {
              throw std::runtime_error("Error closing pipe to PowerShell command");
            }
            std::wcout << getTimeString() << L"Done" << std::endl;
            file << getTimeString() << L"Done" << std::endl;

            file << std::endl << std::endl;
            std::wcout << getTimeString() << L"========================================DONE========================================" << std::endl;
            file << getTimeString() << L"========================================DONE========================================" << std::endl;
            std::wcout << getTimeString() << L"====================================================================================" << std::endl;
            file << getTimeString() << L"====================================================================================" << std::endl;
            std::wcout << std::endl << std::endl;

        /*    std::wcout << L"Press Enter 2 time to run again";
            std::wcin.ignore();
            std::getline(std::wcin, a);
        }*/
        file.close();
    }
}


