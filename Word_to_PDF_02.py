import os
import win32com.client as win32
import win32api
import pythoncom
from tkinter import Tk, filedialog
import tempfile

def converter_word_para_pdf(pasta_origem, pasta_destino):
    # Inicializa o COM para esta thread
    pythoncom.CoInitialize()
    
    word = None
    try:
        word = win32.Dispatch("Word.Application")
        word.Visible = False

        # Verifica se a pasta de destino existe e tem permissão de escrita
        if not os.path.exists(pasta_destino):
            try:
                os.makedirs(pasta_destino)
                print(f"✅ Pasta de destino criada: {pasta_destino}")
            except Exception as e:
                print(f"❌ Não foi possível criar a pasta de destino: {e}")
                return
        
        # Testa permissão de escrita na pasta de destino
        try:
            teste_arquivo = os.path.join(pasta_destino, "_teste_escrita.tmp")
            with open(teste_arquivo, 'w') as f:
                f.write('teste')
            os.remove(teste_arquivo)
            print("✅ Pasta de destino tem permissão de escrita")
        except Exception as e:
            print(f"❌ Sem permissão de escrita na pasta de destino: {e}")
            return

        for arquivo in os.listdir(pasta_origem):
            if arquivo.lower().endswith((".docx", ".doc")):
                caminho_word = os.path.abspath(os.path.join(pasta_origem, arquivo))
                nome_base = os.path.splitext(arquivo)[0]
                
                # CORREÇÃO: Criar um nome de arquivo seguro (sem caracteres especiais)
                nome_base_seguro = "".join(c for c in nome_base if c.isalnum() or c in (' ', '-', '_')).rstrip()
                
                # Caminho do PDF final
                caminho_pdf_final = os.path.join(pasta_destino, nome_base_seguro + ".pdf")
                
                # CORREÇÃO: Usar um arquivo temporário primeiro
                temp_pdf = os.path.join(tempfile.gettempdir(), nome_base_seguro + "_temp.pdf")

                print("\n------------------------------")
                print(f"Arquivo encontrado pelo Python : {arquivo}")
                print(f"Caminho completo Word          : {caminho_word}")
                print(f"Caminho PDF temporário         : {temp_pdf}")
                print(f"Caminho PDF final              : {caminho_pdf_final}")
                print(f"Existe no disco?                : {os.path.exists(caminho_word)}")

                if not os.path.exists(caminho_word):
                    print("⚠ Python não encontrou o arquivo nesse caminho. Pulando...")
                    continue

                # Pega o caminho curto (8.3) para evitar problemas com acentos/símbolos
                try:
                    caminho_word_curto = win32api.GetShortPathName(caminho_word)
                    # Garantir que o caminho use barras duplas invertidas para o COM
                    caminho_word_curto = caminho_word_curto.replace("/", "\\")
                    print(f"Caminho curto (8.3) Word      : {caminho_word_curto}")
                except Exception as e:
                    print(f"Não consegui gerar caminho curto, usando o normal. Erro: {e}")
                    caminho_word_curto = caminho_word.replace("/", "\\")

                try:
                    print("→ Abrindo no Word...")
                    doc = word.Documents.Open(caminho_word_curto)
                    
                    print("→ Salvando como PDF temporário...")
                    
                    # CORREÇÃO: Usar caminho temporário sem caracteres especiais
                    temp_pdf_com = temp_pdf.replace("/", "\\")
                    
                    try:
                        # Tenta ExportAsFixedFormat primeiro
                        doc.ExportAsFixedFormat(OutputFileName=temp_pdf_com, ExportFormat=17)  # 17 = wdExportFormatPDF
                        print("✅ ExportAsFixedFormat funcionou para o arquivo temporário")
                    except Exception as e1:
                        print(f"⚠ ExportAsFixedFormat falhou: {e1}")
                        # Tenta SaveAs como fallback
                        doc.SaveAs(temp_pdf_com, FileFormat=17)
                        print("✅ SaveAs funcionou para o arquivo temporário")
                    
                    doc.Close()
                    
                    # CORREÇÃO: Verificar se o PDF temporário foi criado
                    if os.path.exists(temp_pdf):
                        print(f"✅ PDF temporário criado: {temp_pdf}")
                        
                        # CORREÇÃO: Mover o arquivo temporário para o destino final
                        try:
                            # Se o arquivo final já existir, remove primeiro
                            if os.path.exists(caminho_pdf_final):
                                os.remove(caminho_pdf_final)
                            
                            import shutil
                            shutil.move(temp_pdf, caminho_pdf_final)
                            print(f"✅ PDF movido com sucesso para: {caminho_pdf_final}")
                        except Exception as e_move:
                            print(f"❌ Erro ao mover PDF para destino final: {e_move}")
                            # Tenta copiar como fallback
                            try:
                                shutil.copy2(temp_pdf, caminho_pdf_final)
                                os.remove(temp_pdf)
                                print(f"✅ PDF copiado com sucesso para: {caminho_pdf_final}")
                            except Exception as e_copy:
                                print(f"❌ Erro ao copiar PDF: {e_copy}")
                    else:
                        print(f"❌ PDF temporário não foi criado em: {temp_pdf}")
                    
                except Exception as e:
                    print(f"❌ Erro ao converter {arquivo}: {e}")

    except Exception as e:
        print(f"❌ Erro geral na conversão: {e}")
    finally:
        if word:
            try:
                word.Quit()
            except:
                pass
        # Finaliza o COM para esta thread
        pythoncom.CoUninitialize()
    
    print("\nConversão concluída!")

# ------------------------- #
# Janela para escolher pastas
# ------------------------- #

# Esconde a janela principal do Tkinter
Tk().withdraw()

print("Selecione a pasta onde estão os arquivos Word:")
pasta_origem = filedialog.askdirectory()

if not pasta_origem:
    print("Nenhuma pasta de origem selecionada. Encerrando...")
    exit()

print("Selecione a pasta onde os PDFs serão salvos:")
pasta_destino = filedialog.askdirectory()

if not pasta_destino:
    print("Nenhuma pasta de destino selecionada. Encerrando...")
    exit()

# Executa a conversão
print(f"\n📁 Pasta de origem: {pasta_origem}")
print(f"📁 Pasta de destino: {pasta_destino}")
print("=" * 50)

converter_word_para_pdf(pasta_origem, pasta_destino)