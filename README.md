# monitoring-os-missing-patches
Monitoring for OS missing patches

How to Build?
Open Visual Studio 2019 > Open file WindowsUpdateAPI.sln
Right click to project WindowsUpdateAPI in Solution Explorer > Properties
In Configuration Properties > Linker > Manifest File > UAC Execution Level > choose requireAdministrator (/level='requireAdministrator') > Apply > OK

Choose Release mode and Run

Popup require Administrator right appear > choose Restart under different credentials

After Visual Studio open again > Run
