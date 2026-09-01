# CDExplorer
Change command line directory to the Explorer window's directory.  
## Background
Under Windows, I often switch between GUI and CLI. When I want to switch to CLI from the Explorer, I need to either right-click and choose "Open Windows Terminal here" or copy the path, activate an terminal window, type `cd` or `pushd` and paste the path. I always found it too troublesome, so I wrote this tool.  

I'm used to keeping a terminal running, so with this tool, I just activate the terminal window and type a short word. Then, it changes to the directory where the topmost Explorer in.  
## Shell Wrapper
### CMD
save this as `cde.bat`
```CMD
@echo off
for /f "delims=" %%i in ('ExplorerDir.exe') do cd /d %%i
```
save this as `pde.bat`
```
@echo off
for /f "delims=" %%i in ('ExplorerDir.exe') do pushd %%i
```
### Powershell
```powershell
function cde {
    $targetDir = & "ExplorerDir.exe"
    if ($LASTEXITCODE -eq 0) {
        Set-Location $targetDir
    }
    else {
        Write-Host $LASTEXITCODE
    }
}

function pde {
    $targetDir = & "ExplorerDir.exe"
    if ($LASTEXITCODE -eq 0) {
        Push-Location $targetDir
    }
    else {
        Write-Host $LASTEXITCODE
    }
}
```
### WSL bash:
```bash
if [ -n "$WSL_DISTRO_NAME" ]; then
    # Helper function to get the WSL path from Windows Explorer
    _get_explorer_path() {

        if ! command -v ExplorerDir.exe > /dev/null 2>&1 || ! command -v wslpath > /dev/null 2>&1; then
            return 2
        fi

        local win_path
        win_path=$(ExplorerDir.exe | tr -d '\r\n')

        if [ -z "$win_path" ]; then
            return 3
        fi

        local wsl_path
        wsl_path=$(wslpath -u "$win_path")
        local status=$?

        if [ $status -ne 0 ]; then
            return 4
        fi

        echo "$wsl_path"
        return 0
    }

    cde() {
        local target_path
        target_path=$(_get_explorer_path) || return $?
        cd "$target_path"
    }

    pde() {
        local target_path
        target_path=$(_get_explorer_path) || return $?
        pushd "$target_path"
    }
fi
```
## Usage
Add the shell wrapper into your shell config file. Then we can:  
cd to current Explorer's Directory:
```
cde
```
or pushd:
```
pde
```