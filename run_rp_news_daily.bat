@echo off
setlocal

cd /d "%~dp0"
if not exist logs mkdir logs

echo ==================================================>> logs\task.log
echo [%date% %time%] Inicio>> logs\task.log

git pull origin main >> logs\task.log 2>&1
if errorlevel 1 goto :error

"C:\Users\2822\pyenv311\Scripts\python.exe" convert.py >> logs\task.log 2>&1
if errorlevel 1 goto :error

git add data/news.json index.html >> logs\task.log 2>&1

git diff --cached --quiet
if errorlevel 1 (
    git commit -m "actualizacion automatica diaria" >> logs\task.log 2>&1
    git push origin main >> logs\task.log 2>&1
    if errorlevel 1 goto :error
) else (
    echo [%date% %time%] Sin cambios para commit>> logs\task.log
)

echo [%date% %time%] Fin OK>> logs\task.log
exit /b 0

:error
echo [%date% %time%] ERROR>> logs\task.log
exit /b 1