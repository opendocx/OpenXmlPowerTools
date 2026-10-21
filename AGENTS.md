# Build and Test Instructions

## Prerequisites
- .NET 10 SDK (`dotnet-sdk-10.0`) must be installed.
  ```bash
  apt-get update && apt-get install -y dotnet-sdk-10.0
  ```

## Build
From the repository root, run:
```bash
dotnet build
```
This restores NuGet packages and compiles the library and example projects.

## Test
Execute the unit tests with:
```bash
dotnet test --project OpenXmlPowerTools.Tests/OpenXmlPowerTools.Tests.csproj
```
The command builds required projects and runs the test suite.
