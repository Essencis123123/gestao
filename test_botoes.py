#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Script de teste para verificar erros ao iniciar a aplicação."""

import sys
import os

# Adicionar diretório atual ao path
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

try:
    from main import MainWindow
    from PyQt5.QtWidgets import QApplication
    
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    
    # Teste: Tentar chamar os métodos diretamente
    print("✅ Aplicação iniciada com sucesso!")
    print("\nTestando métodos:")
    
    # Verificar se métodos existem
    if hasattr(window, 'exportar_relatorio_analise_precos'):
        print("✅ Método exportar_relatorio_analise_precos existe")
    else:
        print("❌ Método exportar_relatorio_analise_precos NÃO EXISTE")
    
    if hasattr(window, 'abrir_dialog_adicionar_ipca'):
        print("✅ Método abrir_dialog_adicionar_ipca existe")
    else:
        print("❌ Método abrir_dialog_adicionar_ipca NÃO EXISTE")
    
    # Verificar WebBridge
    if hasattr(window, 'bridge'):
        print("✅ WebBridge existe")
        
        if hasattr(window.bridge, 'abrirDialogAdicionarIPCA'):
            print("✅ Método WebBridge.abrirDialogAdicionarIPCA existe")
        else:
            print("❌ Método WebBridge.abrirDialogAdicionarIPCA NÃO EXISTE")
        
        if hasattr(window.bridge, 'exportarRelatorioAnalisePrecos'):
            print("✅ Método WebBridge.exportarRelatorioAnalisePrecos existe")
        else:
            print("❌ Método WebBridge.exportarRelatorioAnalisePrecos NÃO EXISTE")

        # tentar invocar os métodos para garantir que não causam erro
        try:
            print("➡️ Chamando window.bridge.abrirDialogAdicionarIPCA()")
            window.bridge.abrirDialogAdicionarIPCA()
            print("   (sem exceção)")
        except Exception as e:
            print(f"   ❌ Erro ao chamar abrirDialogAdicionarIPCA: {e}")

        try:
            print("➡️ Chamando window.bridge.exportarRelatorioAnalisePrecos()")
            window.bridge.exportarRelatorioAnalisePrecos()
            print("   (sem exceção)")
        except Exception as e:
            print(f"   ❌ Erro ao chamar exportarRelatorioAnalisePrecos: {e}")
    else:
        print("❌ WebBridge NÃO EXISTE")
    
    sys.exit(app.exec_())
    
except Exception as e:
    print(f"❌ ERRO AO INICIAR APLICAÇÃO: {e}")
    import traceback
    traceback.print_exc()
    sys.exit(1)
