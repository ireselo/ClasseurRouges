@echo off
chcp 65001 >nul
title Classeurs Rouges - Suivi des Bulletins Scolaires (Hors-Ligne)
echo ============================================================================
echo   CLASSEURS ROUGES - Cycle d'Orientation (9CO / 10CO / 11CO)
echo   Mode 100%% Hors-Ligne et Confidentiel - Demarrage en cours...
echo ============================================================================
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0src\Server.ps1"
