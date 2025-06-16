import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from tkinter import filedialog, messagebox, scrolledtext
import threading
import os
import pandas as pd
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.edge.service import Service as EdgeService
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC


def iniciar_extracao(caminho_arquivo, log_output, tipo_extracao):
    def log(msg):
        log_output.insert("end", msg + "\n")
        log_output.see("end")
        app.update()

    dados_extraidos = []

    try:
        log("\n📂 Lendo planilha...")
        df_doc = pd.read_excel(caminho_arquivo)
        df_doc.columns = df_doc.columns.str.strip()
        numeros = df_doc["DOCUMENTO"].dropna().astype(str).tolist()

        driver = webdriver.Edge()
        wait = WebDriverWait(driver, 20)
        driver.get("https://neoenergia.portaldeassinaturas.com.br")

        messagebox.showinfo("Login", "Faça login manualmente no site e resolva o CAPTCHA.\nClique OK quando estiver na página inicial.")

        meus_doc = wait.until(EC.element_to_be_clickable((By.LINK_TEXT, "Meus documentos")))
        meus_doc.click()
        time.sleep(1)

        def expandir_secao(tipo):
            try:
                painel_expandido = driver.find_elements(By.CSS_SELECTOR, "div.panel-collapse.in")
                if painel_expandido:
                    return

                if tipo == "pendentes":
                    seletor = "div.panel-heading.status.status-waiting-others"
                else:
                    seletor = "div.panel-heading.status.status-done"

                secao = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, seletor)))
                driver.execute_script("arguments[0].scrollIntoView(true);", secao)
                time.sleep(1)
                secao.click()
                time.sleep(2)
            except Exception:
                log(f"⚠️ Falha ao expandir seção '{tipo}'.")
                messagebox.showinfo("Ação necessária", f"Expanda manualmente a seção '{tipo.upper()}' e clique OK para continuar.")

        expandir_secao(tipo_extracao)

        for numero in numeros:
            try:
                log(f"\n🔍 Processando documento: {numero}")

                campo_busca = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "input[placeholder*='procur']")))
                campo_busca.clear()
                campo_busca.send_keys(numero)
                campo_busca.send_keys(Keys.RETURN)
                time.sleep(2)

                try:
                    link_doc = wait.until(EC.element_to_be_clickable((By.XPATH, f"//a[contains(., '{numero}')]")))
                    link_doc.click()
                except Exception:
                    log(f"❌ Documento '{numero}' não encontrado.")
                    continue

                time.sleep(2)

                try:
                    status_menu = wait.until(EC.element_to_be_clickable((By.XPATH,
                        "//div[@id='action-status']//div[contains(@class, 'panel-heading')]/h4[contains(@class, 'panel-title')]")))
                    driver.execute_script("arguments[0].scrollIntoView(true);", status_menu)
                    time.sleep(1)
                    status_menu.click()
                    time.sleep(2)
                except Exception:
                    log("⚠️ Falha ao expandir 'Status das ações'.")

                etapas = driver.find_elements(By.CSS_SELECTOR, "div.etapa")

                for etapa in etapas:
                    try:
                        titulo_div = etapa.find_element(By.XPATH, ".//div[contains(@class, 'panel-heading')]//div[contains(@class, 'col-lg-9')]")
                        texto_completo = titulo_div.text.strip()
                        nome_etapa = texto_completo.split("-")[0].strip()
                    except:
                        nome_etapa = ""

                    membros_finalizados = etapa.find_elements(By.CSS_SELECTOR, "div.member.ended")
                    for membro in membros_finalizados:
                        nome = membro.find_element(By.CLASS_NAME, "name").text.strip() if membro.find_elements(By.CLASS_NAME, "name") else ""
                        email = membro.find_element(By.CLASS_NAME, "email").text.strip() if membro.find_elements(By.CLASS_NAME, "email") else ""
                        data_assinatura = membro.find_element(By.CLASS_NAME, "action-date").text.strip() if membro.find_elements(By.CLASS_NAME, "action-date") else ""
                        parte = membro.find_element(By.CLASS_NAME, "action-title").text.strip() if membro.find_elements(By.CLASS_NAME, "action-title") else ""

                        if tipo_extracao == "finalizados":
                            dados_extraidos.append({
                                "Documento": numero,
                                "Nome": nome,
                                "Email": email,
                                "Data da Assinatura": data_assinatura,
                                "Status": "Finalizado",
                                "Etapa": nome_etapa,
                                "Parte": parte
                            })

                    if tipo_extracao == "pendentes":
                        membros_pendentes = etapa.find_elements(By.CSS_SELECTOR, "div.member.pending")
                        for membro in membros_pendentes:
                            nome = membro.find_element(By.CLASS_NAME, "name").text.strip() if membro.find_elements(By.CLASS_NAME, "name") else ""
                            email = membro.find_element(By.CLASS_NAME, "email").text.strip() if membro.find_elements(By.CLASS_NAME, "email") else ""
                            data_assinatura = membro.find_element(By.CLASS_NAME, "action-date").text.strip() if membro.find_elements(By.CLASS_NAME, "action-date") else ""
                            parte = membro.find_element(By.CLASS_NAME, "action-title").text.strip() if membro.find_elements(By.CLASS_NAME, "action-title") else ""

                            dados_extraidos.append({
                                "Documento": numero,
                                "Nome": nome,
                                "Email": email,
                                "Data da Assinatura": data_assinatura,
                                "Status": "Pendente",
                                "Etapa": nome_etapa,
                                "Parte": parte
                            })

                            try:
                                botao_compartilhar = etapa.find_element(By.XPATH, ".//button[.//i[contains(@class, 'fa-share-alt')]]")
                                driver.execute_script("arguments[0].scrollIntoView(true);", botao_compartilhar)
                                time.sleep(0.5)
                                driver.execute_script("arguments[0].click();", botao_compartilhar)

                                input_box = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "input#ShareLink0.form-control")))
                                link_compartilhamento = input_box.get_attribute("value")
                                dados_extraidos[-1]["Link de Compartilhamento"] = link_compartilhamento

                                try:
                                    botao_fechar = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "button.close")))
                                    botao_fechar.click()
                                    time.sleep(0.5)
                                except:
                                    log("⚠️ janela do link não fechou automaticamente, sem problemas ao código.")
                            except Exception as e:
                                log(f"⚠️ Erro ao extrair link de compartilhamento: {e}")
                                dados_extraidos[-1]["Link de Compartilhamento"] = ""

                driver.back()
                time.sleep(2)
                expandir_secao(tipo_extracao)

            except Exception as e:
                log(f"❌ Erro ao processar {numero}: {e}")
                continue

        df_resultado = pd.DataFrame(dados_extraidos)
        df_resultado["Data (Somente Data)"] = df_resultado["Data da Assinatura"].str.extract(r'(\d{2}/\d{2}/\d{4})')

        salvar_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel Files", "*.xlsx")], title="Salvar como")
        if salvar_path:
            df_resultado.to_excel(salvar_path, index=False)
            log(f"\n✅ Extração finalizada. Arquivo salvo em: {salvar_path}")
            messagebox.showinfo("Concluído", "Extração finalizada com sucesso!")
        else:
            log("⚠️ Salvamento cancelado pelo usuário.")

        driver.quit()

    except Exception as e:
        messagebox.showerror("Erro", str(e))

# --- Interface com ttkbootstrap ---
app = ttk.Window(themename="darkly")
app.title("Gestão de assinatura de minutas")
app.geometry("900x650")

caminho_var = ttk.StringVar()
tipo_extracao_var = ttk.StringVar(value="pendentes")

frame_top = ttk.Frame(app)
frame_top.pack(pady=10)

btn_arquivo = ttk.Button(frame_top, text="Selecionar Planilha Excel", command=lambda: caminho_var.set(filedialog.askopenfilename(filetypes=[("Planilhas Excel", "*.xlsx")])) )
btn_arquivo.pack(side="left", padx=10)

entry_caminho = ttk.Entry(frame_top, textvariable=caminho_var, width=80)
entry_caminho.pack(side="left", padx=10)

frame_opcoes = ttk.Frame(app)
frame_opcoes.pack(pady=5)

ttk.Label(frame_opcoes, text="Tipo de extração:").pack(side="left", padx=5)
ttk.Radiobutton(frame_opcoes, text="Pendentes", variable=tipo_extracao_var, value="pendentes").pack(side="left")
ttk.Radiobutton(frame_opcoes, text="Finalizados", variable=tipo_extracao_var, value="finalizados").pack(side="left")

log_output = scrolledtext.ScrolledText(app, width=110, height=30)
log_output.pack(padx=10, pady=10)

def iniciar_thread():
    if not caminho_var.get():
        messagebox.showwarning("Atenção", "Selecione uma planilha antes de iniciar.")
        return
    threading.Thread(target=iniciar_extracao, args=(caminho_var.get(), log_output, tipo_extracao_var.get()), daemon=True).start()

btn_iniciar = ttk.Button(app, text="Iniciar Extração", bootstyle="success", command=iniciar_thread)
btn_iniciar.pack(pady=10)

app.mainloop()
