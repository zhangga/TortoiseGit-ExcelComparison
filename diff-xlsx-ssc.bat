@echo off

CHCP 65001

:: 系统环境变量
set ENV_PATH=%PATH%
@echo ====current environment：
@echo %ENV_PATH%

:: 启用命令扩展
setlocal enabledelayedexpansion
set svnStr=SVN
set gitStr=Git
:: 调用这个方法，传入字符串ENV_PATH和要查找的字符串svnStr。lens是它的返回值
call :getSubIndex ENV_PATH svnStr lensSvn
if "%lensSvn%"=="" (
	echo "没有找到SVN"
	goto :notSetSVN
) else (
	echo "找到TortoiseSVN环境变量"
)

:: 替换svn的js文件
call :getLastIndex ENV_PATH lensSvn svnPath
set svnDiffPath="%svnPath%\Diff-Scripts\diff-xls.js"
echo %svnDiffPath%
call :writeJSFile svnDiffPath
echo "成功设置SVN"

:notSetSVN

call :getSubIndex ENV_PATH gitStr lensGit
if "%lensGit%"=="" (
	echo "没有找到Git"
) else (
	echo "找到TortoiseGit环境变量"
)

:: 替换git的js文件
call :getLastIndex ENV_PATH lensGit gitPath
set gitDiffPath="%gitPath%\Diff-Scripts\diff-xls.js"
echo %gitDiffPath%
call :writeJSFile gitDiffPath

echo "成功设置Git"

:notSetGit

pause

exit /b
:getLastIndex
setlocal enabledelayedexpansion
set /A len+=%2
set value=
:strLen_LoopIndex
	set /A num=len-1
	if not "!%1:~%num%,1!"=="" (
		if "!%1:~%num%,1!"==";" (
			echo "%value%"
			endlocal & set %3=%value%
		) else (
			set /A len=len-1
			set value=!%1:~%num%,1!%value%
			goto :strLen_LoopIndex
		)
	) else (
		endlocal & set %3=%value%
	)
exit /b

exit /b
:getSubIndex
setlocal enabledelayedexpansion
:strLen_Loop
    set /A len+=1
    set /A len1+=0
    set /A num=len-1
    ::判断传入第二个参数要查找的字符是否已经遍历到了结尾，如果结尾了就说明匹配到了
    if not "!%2:~%len1%!"=="" (
    ::判断第一个传入的字符串是否已经遍历到了结尾
    if not "!%1:~%num%!"=="" (
        if not "!%2:~%len1%!"=="" (
            if "!%1:~%num%,1!"=="!%2:~%len1%,1!" (
				set /A len1=len1+1
            ) else (
                set /A len1=0
            )
            goto :strLen_Loop
        ) else (
			endlocal & set %3=%num%
          )
        )
    ) else (
        endlocal & set %3=%num%
    )
exit /b

exit /b
:writeJSFile
setlocal enabledelayedexpansion

more +109 %~dp0\diff-xlsx-ssc.bat > !%1!
exit /b

:: js比对脚本



var objArgs = WScript.Arguments;
if (objArgs.length < 2)
{
    Abort("Usage: [CScript | WScript] diff-xls.js base.xls new.xls", "Invalid arguments");
}

var sBaseDoc = objArgs(0);
var sNewDoc = objArgs(1);

var objScript = new ActiveXObject("Scripting.FileSystemObject");

if (objScript.GetBaseName(sBaseDoc) === objScript.GetBaseName(sNewDoc))
{
    Abort("File '" + sBaseDoc + "' and '" + sNewDoc + "' is same file name.\nCannot compare the documents.", "Same file name");
}

if (!objScript.FileExists(sBaseDoc))
{
    Abort("File '" + sBaseDoc + "' does not exist.\nCannot compare the documents.", "File not found");
}

if (!objScript.FileExists(sNewDoc))
{
    Abort("File '" + sNewDoc + "' does not exist.\nCannot compare the documents.", "File not found");
}

sBaseDoc = objScript.GetAbsolutePathName(sBaseDoc);
sNewDoc = objScript.GetAbsolutePathName(sNewDoc);
var sTempFolder = objScript.GetSpecialFolder(2)
var sTempFile = "D:\\temp.txt"
objScript = null;

var fs = new ActiveXObject("Scripting.FileSystemObject");
var f = fs.CreateTextFile(sTempFile, 2, true)
f.WriteLine(sBaseDoc)
f.WriteLine(sNewDoc)
f.close()
fs = null
f = null

WScript
var objShell = new ActiveXObject("WScript.Shell");
objShell.run('"C:\\Program Files\\Microsoft Office\\root\\Client\\AppVLP.exe" "C:\\Program Files (x86)\\Microsoft Office\\Office16\\DCF\\SPREADSHEETCOMPARE.EXE" "D:\\temp.txt"', 0, true)
objShell = null
