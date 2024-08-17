# CDExplorer
Change command line directory to the Explorer window's directory.  
## Background
Under Windows, I often switch between GUI and CLI. When I want to switch to CLI from the Explorer, I need to either right-click and choose "Open Windows Terminal here" or copy the path, activate an terminal window, type `cd` or `pushd` and paste the path. I always found it too troublesome, so I wrote this tool.  

I'm used to keeping a terminal running, so with this tool, I just activate the terminal window and type a short word. Then, it changes to the directory where the topmost Explorer in.  
## Usage
- Download and decompress the release package.
- Add ExplorerDir to the PATH environment variable.
- Create a new cmd or powershell prompt, then you can use `cde` and `pushde`.

It looks like:  
```cmd
PS ~ > pushde
PS ~\repo\CDExplorer > popd
PS ~ > cde
PS ~\repo\CDExplorer >
```
Or you can use `ExplorerDir` to get the directory for other commands.  
```cmd
PS ~ > gi (ExplorerDir)

    Directory: C:\Windows

Mode                 LastWriteTime         Length Name
----                 -------------         ------ ----
d----          2024-08-09     8:47                System32
```