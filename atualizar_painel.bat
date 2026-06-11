@echo off
chcp 65001 >nul
REM ============================================================
REM  Atualiza o Painel de Descontos e publica no GitHub Pages.
REM  Rode este arquivo (duplo clique) sempre que quiser atualizar.
REM ============================================================
cd /d "%~dp0"

echo.
echo [1/3] Gerando o painel a partir da base mais recente...
python gen_dashboard_v2.py
if errorlevel 1 (
  echo ERRO ao gerar o painel. Verifique se o Python e a base CSV estao no lugar.
  pause
  exit /b 1
)

echo.
echo [2/3] Registrando alteracoes...
git add dist gen_dashboard_v2.py .github .gitignore README.md atualizar_painel.bat
git commit -m "Atualiza painel de descontos"

echo.
echo [3/3] Enviando para o GitHub (Pages publica em ~1-2 min)...
git push

echo.
echo Concluido. O link do painel permanece o mesmo.
pause
