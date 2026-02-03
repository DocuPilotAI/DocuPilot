@echo off
REM 将 DocuPilot/manifest.xml 拷贝到本地 Office wef 目录（仅 Windows）
REM 用法: copy-manifest-to-wef.bat

setlocal enabledelayedexpansion

REM 设置控制台编码为 UTF-8
chcp 65001 >nul

REM 脚本所在目录即项目根目录（DocuPilot）
set "SCRIPT_DIR=%~dp0"
set "MANIFEST_SRC=%SCRIPT_DIR%manifest.xml"

REM 检查是否为 Windows 系统
ver >nul 2>&1
if errorlevel 1 (
    echo [错误] 此脚本仅支持 Windows。
    exit /b 1
)

REM 检查 manifest.xml 是否存在
if not exist "%MANIFEST_SRC%" (
    echo [错误] 未找到 manifest.xml: %MANIFEST_SRC%
    exit /b 1
)

REM Windows Office 共享文件夹路径
set "WEF_BASE=%USERPROFILE%\AppData\Local\Microsoft\Office\16.0\Wef"

echo [信息] 正在将 manifest.xml 拷贝到 Office wef 目录...
echo   源文件: %MANIFEST_SRC%
echo.

REM 创建 wef 目录（如果不存在）
if not exist "%WEF_BASE%" (
    echo [信息] 创建目录: %WEF_BASE%
    mkdir "%WEF_BASE%" 2>nul
    if errorlevel 1 (
        echo [警告] 创建目录失败，可能权限不足
    )
)

REM 复制 manifest.xml
copy "%MANIFEST_SRC%" "%WEF_BASE%\manifest.xml" >nul 2>&1
if errorlevel 1 (
    echo [错误] 拷贝失败，请检查权限或 Office 安装路径
    echo.
    echo 提示：
    echo   1. 确保 Office 已正确安装
    echo   2. 如果使用较旧版本的 Office，可能需要修改脚本中的版本号（16.0 → 15.0）
    echo   3. 请以管理员权限运行此脚本
    exit /b 1
) else (
    echo [成功] ✓ 已复制到: %WEF_BASE%
    echo.
)

echo [完成] 请在 Excel/Word/PowerPoint 中执行以下步骤加载 DocuPilot：
echo   1. 打开任意 Office 应用（Excel、Word 或 PowerPoint）
echo   2. 点击「插入」选项卡
echo   3. 点击「获取加载项」或「我的加载项」
echo   4. 在弹出窗口中选择「共享文件夹」选项卡
echo   5. 找到并点击 DocuPilot 加载项
echo.
echo 注意：如果未看到 DocuPilot，请：
echo   - 重启 Office 应用
echo   - 检查开发服务器是否已启动（npm run dev:https）
echo   - 使用「上传我的加载项」手动选择 manifest.xml 文件
echo.

pause
