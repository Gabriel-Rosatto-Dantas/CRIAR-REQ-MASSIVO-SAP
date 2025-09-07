#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Script para criar executável do APP_GUI.py usando PyInstaller
"""

import os
import sys
import subprocess
import shutil

def install_requirements():
    """Instala as dependências necessárias."""
    print("📦 Instalando dependências...")
    try:
        subprocess.check_call([sys.executable, "-m", "pip", "install", "-r", "requirements.txt"])
        print("✅ Dependências instaladas com sucesso!")
        return True
    except subprocess.CalledProcessError as e:
        print(f"❌ Erro ao instalar dependências: {e}")
        return False

def create_executable():
    """Cria o executável usando PyInstaller."""
    print("🔨 Criando executável...")
    
    # Comando PyInstaller
    cmd = [
        "pyinstaller",
        "--onefile",                    # Arquivo único
        "--windowed",                   # Sem console (GUI)
        "--name=SAP_Automation",        # Nome do executável
        "--icon=icone.ico",             # Usar ícone personalizado
        "--add-data=requirements.txt;.", # Incluir requirements.txt
        "--add-data=icone.ico;.",       # Incluir arquivo de ícone
        "--hidden-import=win32com.client",
        "--hidden-import=gspread",
        "--hidden-import=pandas",
        "--hidden-import=PIL",
        "--hidden-import=tkinter",
        "--hidden-import=configparser",
        "--hidden-import=threading",
        "--hidden-import=subprocess",
        "--hidden-import=datetime",
        "--hidden-import=re",
        "--hidden-import=os",
        "--hidden-import=sys",
        "--hidden-import=time",
        "--hidden-import=io",
        "--hidden-import=base64",
        "--hidden-import=pywintypes",
        "--hidden-import=pythoncom",
        "APP_GUI.py"
    ]
    
    try:
        subprocess.check_call(cmd)
        print("✅ Executável criado com sucesso!")
        return True
    except subprocess.CalledProcessError as e:
        print(f"❌ Erro ao criar executável: {e}")
        return False

def cleanup():
    """Remove arquivos temporários."""
    print("🧹 Limpando arquivos temporários...")
    
    # Diretórios e arquivos a serem removidos
    cleanup_items = [
        "build",
        "SAP_Automation.spec",
        "__pycache__"
    ]
    
    for item in cleanup_items:
        if os.path.exists(item):
            if os.path.isdir(item):
                shutil.rmtree(item)
            else:
                os.remove(item)
            print(f"🗑️ Removido: {item}")

def main():
    """Função principal."""
    print("🚀 Iniciando processo de criação do executável...")
    print("=" * 50)
    
    # Verificar se o arquivo principal existe
    if not os.path.exists("APP_GUI.py"):
        print("❌ Arquivo APP_GUI.py não encontrado!")
        return False
    
    # Instalar dependências
    if not install_requirements():
        return False
    
    # Criar executável
    if not create_executable():
        return False
    
    # Limpeza
    cleanup()
    
    print("=" * 50)
    print("🎉 Processo concluído com sucesso!")
    print("📁 O executável está em: dist/SAP_Automation.exe")
    print("💡 Você pode compartilhar este arquivo com seus colegas!")
    
    return True

if __name__ == "__main__":
    success = main()
    if not success:
        sys.exit(1)
