# rdp_alive

Small VB.net (gross) project that will generate an exe that keeps RDP sessions (mstsc.exe) unlocked/alive from the client side.

Tested in Visual Studio 2019.

## Steps to compile

Option A (With [VS2019](https://visualstudio.microsoft.com/vs/community)):

- Open the .sln in VS and select Build

Option B (with [dotnet 5.0 SDK](https://dotnet.microsoft.com/download) (or VS2019, but CLI only))

- From the root of the project, execute: `dotnet build -c release`

Original source [here](https://www.codeproject.com/Tips/1174585/RDP-Session-Keep-Alive). I cleaned up the code and made some small modifications. Designed to run as a console application.
