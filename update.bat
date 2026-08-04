@echo off
chcp 65001 >nul
echo === Обновление ФП-Контроль ===
echo.

cd /d "%~dp0"

echo [1/3] Получаем обновления из GitHub...
git pull origin main
if errorlevel 1 (
    echo ОШИБКА: git pull завершился с ошибкой
    pause
    exit /b 1
)

echo.
echo [2/3] Перезапускаем контейнер...
docker compose down
docker compose up -d
if errorlevel 1 (
    echo ОШИБКА: docker compose завершился с ошибкой
    pause
    exit /b 1
)

echo.
echo [3/3] Готово! Портал обновлён и запущен.
echo.
pause
