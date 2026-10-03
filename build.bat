@echo off
chcp 65001 >nul
echo ============================================
echo   安居客郑州租房爬虫 - 打包成单文件 exe
echo ============================================
echo.

where pyinstaller >nul 2>nul
if errorlevel 1 (
    echo [错误] 没找到 pyinstaller，请先执行：
    echo        pip install -r requirements.txt
    pause
    exit /b 1
)

echo 正在清理旧产物...
if exist build rmdir /s /q build
if exist dist  rmdir /s /q dist
if exist AnjukeZZSpider.spec del /q AnjukeZZSpider.spec

echo 正在打包，请稍候...
pyinstaller --onefile --windowed --noconfirm --clean ^
    --name AnjukeZZSpider ^
    --icon app.ico ^
    --add-data "app.ico;." ^
    --hidden-import openpyxl ^
    anjuke_gui.py

if errorlevel 1 (
    echo.
    echo [失败] 打包出错，请把上面的报错信息发出来。
    pause
    exit /b 1
)

echo.
echo 正在把产物改成中文名...
copy /y "dist\AnjukeZZSpider.exe" "dist\安居客郑州租房爬虫.exe" >nul
if errorlevel 1 (
    echo [提示] 中文名复制失败，直接用 dist\AnjukeZZSpider.exe 也行。
) else (
    del /q "dist\AnjukeZZSpider.exe"
)

echo.
echo [完成] 产物：dist\安居客郑州租房爬虫.exe
echo        可以直接把它复制走或重命名，不影响运行。
echo.
pause
