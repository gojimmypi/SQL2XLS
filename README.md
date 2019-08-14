# SQL2XLS Version 1.2

Given SQL statement(s) in a config file, convert the resultant dataset(s) to individual Excel XLSX or CSV file(s).

Example:

With a config file app.settings like this:

```
  <appSettings>
    <add key="FILE_FORMAT"        value="XLSX"/> <!-- valid values are CSV and XLSX. Default is XLSX -->
    <add key="OVERWRITE_EXISTING" value="TRUE"/>
    <add key="FIELD_HEADER"       value="TRUE"/> <!-- include field names in first row (CSV only)? -->
    <add key="SQL_COMMAND"       value="select 1 as batch_id, * from master.sys.databases; select 1 as batch_id,* from master.dbo.syslogins"/>
    <add key="FILE_PATH"         value="c:\temp\"/>
    <add key="FILE_ROOT_0"       value="Databases - Batch"/>
    <add key="FILE_ROOT_1"       value="Users - Batch"/>
  </appSettings>
```
and invoked like this:

```
sql2xls /SERVER:myserver
```
There should be a file called ```Databases - Batch 1.xlsx``` and another called ```Users - Batch 1.xlsx``` created in c:\temp with [myserver] databases and users listed respectively in each file.

This version also has the [SQLHelper class](SQL2XLS/SQLHelper.cs) included, rather than referenced in [Patterns and Practices](https://github.com/gojimmypi/PatternsAndPractices).


To change source control in Visual Studio:

Remove all existing bindings. If using VSS and switching to git, delete:

```mssccprj.scc
SQL2XLS.csproj.vspscc
vssver2.scc (hidden, read-only, system)```

Then in Visual Studio:

Tools - Options - Source Control - Plug-in Selection

