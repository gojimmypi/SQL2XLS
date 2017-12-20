# SQL2XLS Version 0.1

Given a SQL statement in a config file, convert the resultant dataset(s) to individual Excel XLSX file(s).

Example:

With a config file app.settings like this:

```
  <appSettings>
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


To change source control in Visual Studio:

Remove all existing bindings. If using VSS and switching to git, delete:

```mssccprj.scc
SQL2XLS.csproj.vspscc
vssver2.scc (hidden, read-only, system)```

Then in Visual Studio:

Tools - Options - Source Control - Plug-in Selection

