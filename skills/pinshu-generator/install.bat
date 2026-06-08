@echo off
chcp 65001 >nul
echo === 聘书生成器 一键安装 ===
echo.

:: 检查 Python
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [X] 未找到 Python，请先安装 python.org
    pause
    exit /b 1
)

:: 安装依赖
echo [*] 安装依赖...
pip install python-docx lxml -q

:: 复制到 skills 目录
set SKILLDIR=%USERPROFILE%\.claude\skills\pinshu-generator
if not exist "%SKILLDIR%" mkdir "%SKILLDIR%"
copy /Y "%~dp0SKILL.md" "%SKILLDIR%\SKILL.md" >nul
copy /Y "%~dp0generate.py" "%SKILLDIR%\generate.py" >nul
copy /Y "%~dp0聘书模板.docx" "%SKILLDIR%\聘书模板.docx" >nul

echo [√] 安装完成！
echo.
echo 下次跟 Claude 说"帮我生成聘书"就会自动调用。
echo 模板在 %SKILLDIR%
pause
