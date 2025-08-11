#!/usr/bin/env pwsh
# Script de purge du cache Office pour les compléments
# Usage: ./clear-office-cache.ps1

Write-Host "🧹 Purge du cache Office en cours..." -ForegroundColor Yellow

# 1. Arrêter tous les processus Office
Write-Host "⏹️  Arrêt des processus Office..." -ForegroundColor Cyan
Get-Process | Where-Object {$_.ProcessName -match "(office|word|excel|powerpoint)"} | Stop-Process -Force -ErrorAction SilentlyContinue

# 2. Vider le cache des compléments Office
Write-Host "🗂️  Vidage du cache Wef (16.0)..." -ForegroundColor Cyan
Remove-Item -Path "$env:LOCALAPPDATA\Microsoft\Office\16.0\Wef\*" -Recurse -Force -ErrorAction SilentlyContinue

# 3. Vider le cache alternatif Wef
Write-Host "🗂️  Vidage du cache Wef (général)..." -ForegroundColor Cyan
Remove-Item -Path "$env:LOCALAPPDATA\Microsoft\Office\Wef\*" -Recurse -Force -ErrorAction SilentlyContinue

# 4. Vider le cache Edge WebView2
Write-Host "🌐 Vidage du cache EdgeWebView..." -ForegroundColor Cyan
Remove-Item -Path "$env:LOCALAPPDATA\Microsoft\EdgeWebView\User Data\Default\*" -Recurse -Force -ErrorAction SilentlyContinue

# 5. Vider cache temporaire Office
Write-Host "🗑️  Vidage du cache temporaire..." -ForegroundColor Cyan
Remove-Item -Path "$env:TEMP\*office*" -Recurse -Force -ErrorAction SilentlyContinue

Write-Host "✅ Cache Office purgé avec succès!" -ForegroundColor Green
Write-Host "💡 Vous pouvez maintenant relancer Word et votre complément" -ForegroundColor White