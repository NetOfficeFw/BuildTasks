NetOffice build tasks help your with development of Microsoft Office add-ins.

Use `NetOfficeFw.Build` to automatically register and cleanup your add-ins
when developing Office extensions.

Add-ins are registered to your user profile and the registration does not require
administrator rights.


### Supported Projects

NetOffice build tasks support .NET Framework projects in classic and modern format.  
You can use `packages.config` or `<PackageReference>` in your project.  

Supported .NET Frameworks are 4.6.2 and 4.8.

_The .NET Core assemblies are not supported yet._


### Usage in modern .NET SDK-style projects

1. Add reference to `NetOfficeFw.Build` package to your project file.
2. Set the `<OfficeAppAddin>` property with list of Office applications your add-in supports:
    ```xml
    <Project Sdk="Microsoft.NET.Sdk">
      <PropertyGroup>
        <AssemblyName>AcmeAddin</AssemblyName>
        <TargetFrameworks>net48</TargetFrameworks>
      </PropertyGroup>

      <PropertyGroup>
        <OfficeAppAddin>Word;Excel;PowerPoint;Outlook</OfficeAppAddin>
      </PropertyGroup>

      <ItemGroup>
        <PackageReference Include="NetOfficeFw.Build" Version="2.1.0" />
      </ItemGroup>
    </Project>
    ```


### Usage in classic projects

1. Use **Nuget Packages Manager** to install `NetOfficeFw.Build` package
2. Right click your project and choose **Unload project**
3. Right click your project and choose **Edit Project File**
4. Add the `<OfficeAppAddin>` property with list of Office applications your add-in supports:
    ```xml
    <Project ToolsVersion="15.0" xmlns="http://schemas.microsoft.com/developer/msbuild/2003">
      ...
      <PropertyGroup>
        <OfficeAppAddin>Word;Excel;PowerPoint;Outlook</OfficeAppAddin>
      </PropertyGroup>
      ...
    </Project>
    ```
5. Right click your project and choose **Reload Project**




### License

Source code is licensed under [MIT License](https://github.com/NetOfficeFw/BuildTasks/blob/main/LICENSE.txt).


_Project icon [Card Exchange][icon] is licensed from Icons8 service
under Universal Multimedia Licensing Agreement for Icons8._  
_See <https://icons8.com/license> for more information_

[1]: https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-ovba/4742b896-b32b-4eb0-8372-fbf01e3c65fd
[2]: https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-ovba/575462ba-bf67-4190-9fac-c275523c75fc#intellectual-property-rights-notice-for-open-specifications-documentation
[icon]: https://icons8.com/icon/Mqi1QKj3zbfY/card-exchange
