setlocal

set PackageVersion=%1
set MSBuildPath=%2

%MSBuildPath% ..\Source\IntelliSense.sln /t:restore,build /p:Configuration=Release /p:ContinuousIntegrationBuild=true
@if errorlevel 1 goto end

cd ..\NuGet
call package.cmd %PackageVersion%
@if errorlevel 1 goto end

:end
