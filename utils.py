from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from selenium.common.exceptions import StaleElementReferenceException, TimeoutException
import subprocess
import time
import locale
import os
from bs4 import BeautifulSoup
import html as html_lib
import uo_dict
import re
import tkinter as tk
from tkinter import filedialog
import pandas as pd
import shutil
import requests
import json
from openpyxl import load_workbook

locale.setlocale(locale.LC_TIME, 'pt_BR.UTF-8')

def iniciar_navegador():
    """
    Inicializa e conecta o Selenium a uma instância do navegador Chrome já aberta em modo de depuração remota.

    Returns:
        tuple: (webdriver.Chrome, subprocess.Popen) - O WebDriver e o processo do Chrome.
    """
    try:
        chrome_path = r"C:\Program Files\Google\Chrome\Application\chrome.exe"
        
        # Verificar se o Chrome está instalado
        if not os.path.exists(chrome_path):
            raise FileNotFoundError(f"Chrome não encontrado em: {chrome_path}")
        
        # Iniciar o Chrome em modo debug
        chrome_proc = subprocess.Popen([chrome_path,"--remote-debugging-port=9222",r'--user-data-dir=C:\temp\chromeprofile'])
        
        # Aguardar o Chrome inicializar
        time.sleep(5)
        
        # Configurar opções do Chrome
        options = Options()
        options.add_experimental_option("debuggerAddress", "127.0.0.1:9222")
        options.add_argument("--no-sandbox")
        options.add_argument("--disable-dev-shm-usage")
        
        # Tentar diferentes abordagens para o ChromeDriver
        try:
            # Primeiro: tentar com ChromeDriverManager
            service = Service(ChromeDriverManager().install())
            navegador = webdriver.Chrome(service=service, options=options)
        except Exception as e:
            try:
                navegador = webdriver.Chrome(options=options)
            except Exception as e2:
                # Terceiro: reinstalar ChromeDriver
                # Limpar cache do ChromeDriverManager
                cache_dir = os.path.expanduser("~/.wdm")
                if os.path.exists(cache_dir):
                    shutil.rmtree(cache_dir)
                
                service = Service(ChromeDriverManager().install())
                navegador = webdriver.Chrome(service=service, options=options)
        
        navegador.maximize_window()
        navegador.get("https://fusion.fiemg.com.br/fusion/portal")
        return navegador, chrome_proc
        
    except Exception as e:
        print(f"Erro ao inicializar navegador: {e}")
        # Fechar processo do Chrome se foi criado
        try:
            chrome_proc.terminate()
        except:
            pass
        raise

def clicar_elemento(nav, elemento, tipo,automacao_fusion_instance):
    """
    Localiza e clica em um elemento na página web utilizando JavaScript.

    Args:
        nav (WebDriver): O navegador (WebDriver) usado para interagir com a página.
        elemento (str): O seletor do elemento a ser localizado na página.
        tipo (By): O tipo de seletor (e.g., By.ID, By.CLASS_NAME, etc.) usado para localizar o elemento.

    Functionality:
        - Tenta localizar o elemento especificado utilizando `WebDriverWait`, aguardando até 30 segundos para ele aparecer.
        - Clica no elemento utilizando um comando JavaScript para garantir a execução do clique.
        - Se o elemento não for encontrado após o tempo de espera, exibe um alerta através de uma janela Tkinter, informando o usuário para interagir manualmente.

    Returns:
        None: A função tenta localizar e clicar no elemento, e em caso de falha, exibe uma mensagem de alerta ao usuário.

    Raises:
        Exibe um alerta ao usuário se o elemento não for encontrado dentro do tempo de espera.
    """
    while True:
        try:
            obj = WebDriverWait(nav, 10).until(EC.presence_of_element_located((tipo, elemento)))  # Aguarda 10 segundos até o elemento carregar
            nav.execute_script("arguments[0].click();", obj)  # Clica no objeto utilizando JavaScript
            break  # Sai do loop se o comando for bem-sucedido
        except Exception:
            if not automacao_fusion_instance.handle_custom_messagebox_response():
                break

def acessar_iframe(nav, tempo_espera,automacao_fusion_instance, timeout=10):
    """
    Retorna ao conteúdo principal da página (fora de qualquer iframe) e acessa novamente um iframe.

    Args:
        nav (WebDriver): O navegador (WebDriver) usado para interagir com a página.
        tempo_espera (int, float): Tempo, em segundos, para aguardar antes de retornar ao conteúdo padrão.
        timeout (int, optional): Tempo máximo, em segundos, para aguardar o iframe aparecer. O padrão é 10 segundos.

    Functionality:
        - Aguarda o tempo especificado (tempo_espera) antes de trocar o contexto para o conteúdo padrão da página.
        - Muda o contexto do WebDriver para o conteúdo principal (fora de qualquer iframe).
        - Espera até o `timeout` para que o iframe esteja presente no DOM.
        - Muda o contexto do WebDriver para o iframe localizado.

    Returns:
        None: A função realiza a troca de contexto para o iframe, sem retornar um valor.
    """
    while True:
        try:
            time.sleep(tempo_espera)  # Espera antes de mudar para o conteúdo padrão
            nav.switch_to.default_content()  # Volta ao conteúdo principal da página
            iframe = WebDriverWait(nav, timeout).until(EC.presence_of_element_located((By.TAG_NAME, 'iframe')))  # Espera até que o iframe esteja presente
            nav.switch_to.frame(iframe)  # Troca para o iframe
            break  # Sai do loop se o comando for bem-sucedido
        except Exception:
            if not automacao_fusion_instance.handle_custom_messagebox_response():
                break

def clicar_entidade_por_nome(nav, nome, automacao_fusion_instance=None, timeout=15):
    """
    Clica no <span> que contém o texto (pode ter HTML interno devido a ng-bind-html).
    Usa contains(normalize-space(.), nome) para casar mesmo com marcação interna.
    Retorna True se clicou, False caso contrário.
    """
    xpath = f"//span[contains(@class,'ng-binding') and contains(normalize-space(.), \"{nome}\")]"
    while True:
        try:
            el = WebDriverWait(nav, timeout).until(EC.presence_of_element_located((By.XPATH, xpath)))
            nav.execute_script("arguments[0].scrollIntoView({block: 'center'}); arguments[0].click();", el)
            return True
        except Exception:
            if automacao_fusion_instance is not None:
                if not automacao_fusion_instance.handle_custom_messagebox_response():
                    return False
            else:
                resp = input(f"Elemento com texto '{nome}' não encontrado/clicável. Clique manualmente e pressione Enter para tentar novamente (ou 's' para sair): ")
                if resp.strip().lower() == 's':
                    return False
                # tenta novamente

def enviarkey_elemento(nav, elemento, tipo, texto, automacao_fusion_instance):
    """
    Localiza um elemento na página web, limpa o campo e envia um texto para ele.
    """
    while True:
        try:
            # aguarda até o elemento estar clicável e obtém a referência
            elem = WebDriverWait(nav, 10).until(EC.element_to_be_clickable((tipo, elemento)))
            # tenta limpar via Selenium; se falhar, limpa via JS
            try:
                elem.clear()
            except Exception:
                nav.execute_script("arguments[0].value = '';", elem)
                nav.execute_script("arguments[0].dispatchEvent(new Event('input', { bubbles: true }));", elem)
                nav.execute_script("arguments[0].dispatchEvent(new Event('change', { bubbles: true }));", elem)
            # envia o texto
            elem.send_keys(texto)
            # dispara evento change/input para frameworks reativos
            try:
                nav.execute_script("arguments[0].dispatchEvent(new Event('input', { bubbles: true })); arguments[0].dispatchEvent(new Event('change', { bubbles: true }));", elem)
            except Exception:
                pass
            break  # sucesso
        except Exception:
            if not automacao_fusion_instance.handle_custom_messagebox_response():
                break

def esperar_e_pegar_texto(nav, elemento_id, timeout=10, automacao_fusion_instance=None):
    """
    Espera até que um elemento com o ID especificado apareça e retorna seu texto visível.
    Para inputs hidden que têm o texto visível como nó de texto do pai, retorna o texto do pai
    (removendo valores de inputs filhos). Retorna None se o usuário cancelar via automacao_fusion_instance.
    """
    while True:
        try:
            el = WebDriverWait(nav, timeout).until(EC.presence_of_element_located((By.ID, elemento_id)))
            if el is None:
                return " "
            # garantir referência atual e tratar StaleElementReference
            try:
                el = nav.find_element(By.ID, elemento_id)
            except StaleElementReferenceException:
                time.sleep(0.05)
                el = nav.find_element(By.ID, elemento_id)

            # tentar obter o texto visível que pode estar no pai (ex.: <input hidden> + texto no parent)
            try:
                text = nav.execute_script("""
                    var el = arguments[0];
                    if (!el) return '';
                    var p = el.parentNode;
                    if (p) {
                        var txt = p.textContent || '';
                        // remover valores de inputs/textarea/select dentro do pai para não retornar ids/values
                        var controls = p.querySelectorAll('input, textarea, select');
                        controls.forEach(function(c){ if(c.value) txt = txt.replace(c.value, ''); });
                        return txt.trim();
                    }
                    return (el.textContent || el.value || '').trim();
                """, el)
                if text:
                    return text
            except Exception:
                pass

            # fallback: priorizar atributo 'value' apenas se for realmente informativo
            try:
                val = el.get_attribute("value")
                if val is not None and str(val).strip() != "":
                    # se value for só um id e não queremos, não retornar aqui — mas manter como fallback
                    return str(val).strip()
            except Exception:
                pass

            # último recurso: texto do próprio elemento
            try:
                return el.text.strip()
            except Exception:
                return " "
        except Exception as e:
            # se timeout -> retornar espaço em branco conforme pedido
            if isinstance(e, TimeoutException):
                return " "
            if automacao_fusion_instance is not None:
                if not automacao_fusion_instance.handle_custom_messagebox_response():
                    return None
                # se True, repete o loop
            else:
                # comportamento anterior: raise, mas para evitar travar em busca simples,
                # retornar espaço em branco quando não houver handler (conforme requisito de timeout)
                return " "
            
def pegar_texto_com_quebras(nav, elemento_id, timeout=1, automacao_fusion_instance=None):

    """
    Retorna o texto do elemento preservando quebras de linha (<br>) como '\n'
    e decodificando entidades HTML (ex.: &amp; -> &).
    """
    while True:
        try:
            el = WebDriverWait(nav, timeout).until(EC.presence_of_element_located((By.ID, elemento_id)))
            if el is None:
                return " "
            # garantir referência atual e tratar StaleElementReference
            try:
                el = nav.find_element(By.ID, elemento_id)
            except StaleElementReferenceException:
                time.sleep(0.05)
                el = nav.find_element(By.ID, elemento_id)

            # tentar obter o texto visível que pode estar no pai (ex.: <input hidden> + texto no parent)
            try:
                text = nav.execute_script("""
                    var el = arguments[0];
                    if (!el) return '';
                    var p = el.parentNode;
                    if (p) {
                        var txt = p.textContent || '';
                        // remover valores de inputs/textarea/select dentro do pai para não retornar ids/values
                        var controls = p.querySelectorAll('input, textarea, select');
                        controls.forEach(function(c){ if(c.value) txt = txt.replace(c.value, ''); });
                        return txt.trim();
                    }
                    return (el.textContent || el.value || '').trim();
                """, el)
                if text:
                    return text
            except Exception:
                pass

            # fallback: priorizar atributo 'value' apenas se for realmente informativo
            try:
                val = el.get_attribute("value")
                if val is not None and str(val).strip() != "":
                    return str(val).strip()
            except Exception:
                pass

            # último recurso: texto do próprio elemento
            try:
                return el.text.strip()
            except Exception:
                return " "
        except Exception as e:
            # se timeout -> retornar espaço em branco conforme pedido
            if isinstance(e, TimeoutException):
                return " "
            if automacao_fusion_instance is not None:
                if not automacao_fusion_instance.handle_custom_messagebox_response():
                    return None
                # se True, repete o loop
            else:
                # comportamento anterior: raise, mas para evitar travar em busca simples,
                # retornar espaço em branco quando não houver handler (conforme requisito de timeout)
                return " "
            
def extrair_linhas_tabela(nav, tabela_id, timeout=10, automacao_fusion_instance=None):
    """
    Extrai as linhas de uma tabela (identificada pelo id) e retorna lista de strings,
    cada string com os valores das células separados por ';'.
    Ignora a 4ª coluna (index 3) conforme solicitado.
    Em caso de TimeoutException retorna [" "].
    """
    from selenium.common.exceptions import TimeoutException, StaleElementReferenceException

    while True:
        try:
            tabela = WebDriverWait(nav, timeout).until(EC.presence_of_element_located((By.ID, tabela_id)))
            if tabela is None:
                return [" "]

            js = """
                var table = arguments[0];
                var rows = [];
                if (!table) return rows;
                var trs = table.querySelectorAll('tr');
                trs.forEach(function(tr){
                    var style = window.getComputedStyle(tr);
                    if (style && (style.display === 'none' || style.visibility === 'hidden')) return;
                    var cells = tr.querySelectorAll('td, th');
                    if (!cells.length) return;
                    var parts = [];
                    for (var i = 0; i < cells.length; i++) {
                        // ignorar a 4ª coluna (index 3)
                        if (i === 3) continue;
                        var c = cells[i];
                        // prioridade: se houver span com title, usar title (pega texto completo do tooltip)
                        var spanWithTitle = c.querySelector('span[title]');
                        if (spanWithTitle && spanWithTitle.getAttribute('title')) {
                            parts.push(spanWithTitle.getAttribute('title').trim().replace(/\\s+/g, ' '));
                            continue;
                        }
                        var html = c.innerHTML || '';
                        // preservar <br> como quebra de linha
                        html = html.replace(/<br\\s*\\/?>/gi, '\\n');
                        var tmp = document.createElement('div');
                        tmp.innerHTML = html;
                        // remover valores de inputs/textarea/select dentro da célula
                        var controls = tmp.querySelectorAll('input, textarea, select');
                        controls.forEach(function(ctrl){
                            if (ctrl.value) {
                                tmp.innerHTML = tmp.innerHTML.split(ctrl.value).join('');
                            }
                        });
                        var text = (tmp.textContent || tmp.innerText || '').replace(/\\s+/g, ' ').trim();
                        parts.push(text);
                    }
                    // se todas as colunas foram ignoradas (ex.: só 4 cols e 4ª era a única válida), ainda retorna string vazia
                    rows.push(parts.join(';'));
                });
                return rows;
            """
            linhas = nav.execute_script(js, tabela)
            if not linhas:
                return [" "]
            return linhas

        except Exception as e:
            if isinstance(e, TimeoutException):
                return [" "]
            if automacao_fusion_instance is not None:
                if not automacao_fusion_instance.handle_custom_messagebox_response():
                    return None
            else:
                return [" "]
            
def salvar_lista_historico_xlsx(lista_historico, caminho_arquivo, sheet_name='Planilha1'):
    """
    Salva lista_historico em XLSX com colunas definidas.
    Cada item de lista_historico deve ser uma lista na ordem:
    [numero_chamado,data_inicial,responsavel,area_setor,detalhes_solicitacao,
     urgencia_demanda,justificativa_demanda,data_atual_supervisor,prazo_final,
     historico_nucleo,encaminhamento_supervisor,responsavel_nucleo,acao_nucleo]
    Campos que são listas (ex.: historico_nucleo, encaminhamento_supervisor) serão
    concatenados em uma string separada por " | ".
    """
    import pandas as pd
    from openpyxl import load_workbook
    cols = [
        'numero_chamado','setor','data_inicial','responsavel','uo','detalhes_solicitacao',
        'urgencia_demanda','justificativa_demanda','data_atual_supervisor','prazo_final',
        'encaminhamento_supervisor','acao_supervisor','historico_nucleo','responsavel_nucleo','acao_nucleo', 'sc', 'pedido_protheus'
    ]

    def _cell_to_str(v):
        if v is None:
            return ""
        if isinstance(v, (list, tuple)):
            # juntar cada linha/valor; usar ' | ' para separar registros múltiplos
            return " | ".join(str(x) for x in v)
        return str(v)

    rows = []
    for linha in lista_historico:
        rows.append([_cell_to_str(c) for c in linha])

    df = pd.DataFrame(rows, columns=cols)

    # salvar usando openpyxl engine (mantém novoslines se existirem)
    df.to_excel(caminho_arquivo, index=False, sheet_name=sheet_name, engine='openpyxl')

def fechar_navegador(navegador, chrome_proc):
    """
    Fecha o navegador e o processo do Chrome iniciado em modo debug.
    """
    try:
        if navegador:
            navegador.quit()
        if chrome_proc:
            chrome_proc.terminate()
            chrome_proc.wait()
    except Exception as e:
        print(f"Erro ao fechar o navegador: {e}")

def selecionar_planilha_excel(titulo="Selecionar arquivo Excel", tipo_arquivo=(("Arquivos Excel", "*.xlsx"), ("Todos os arquivos", "*.*"))):
    """
    Abre uma janela de diálogo para o usuário selecionar um arquivo Excel.
    
    Args:
        titulo (str): O título da janela de diálogo.
        tipo_arquivo (tuple): Tupla de tipos de arquivo para o filtro.

    Returns:
        str: O caminho completo do arquivo selecionado pelo usuário.
    """
    root = tk.Tk()
    root.withdraw()  # Oculta a janela principal
    caminho_arquivo = filedialog.askopenfilename(
        title=titulo,
        filetypes= tipo_arquivo
    )
    root.destroy()  # Fecha a janela principal
    return caminho_arquivo

def selecionar_caminho_para_salvar(titulo="Salvar arquivo", tipo_arquivo=(("Arquivos Excel", "*.xlsx"), ("Todos os arquivos", "*.*"))):
    """
    Abre uma janela de diálogo para o usuário selecionar o caminho para salvar um arquivo.
    
    Args:
        titulo (str): O título da janela de diálogo.
        tipo_arquivo (tuple): Tupla de tipos de arquivo para o filtro.

    Returns:
        str: O caminho completo do arquivo selecionado pelo usuário.
    """
    root = tk.Tk()
    root.withdraw()  # Oculta a janela principal
    caminho_arquivo = filedialog.asksaveasfilename(
        title=titulo,
        defaultextension=".xlsx",
        filetypes=tipo_arquivo
    )
    root.destroy()  # Fecha a janela principal
    return caminho_arquivo

def capturar_cookies_e_headers(navegador):
    """
    Captura cookies e headers do navegador logado para usar em requests.
    """
    # Pegar todos os cookies da sessão atual
    selenium_cookies = navegador.get_cookies()
    
    # Converter para formato do requests
    session = requests.Session()
    for cookie in selenium_cookies:
        session.cookies.set(cookie['name'], cookie['value'])
    
    # Headers comuns para o Fusion
    headers = {
        'User-Agent': navegador.execute_script("return navigator.userAgent;"),
        'Accept': 'application/json, text/plain, */*',
        'Accept-Language': 'pt-BR,pt;q=0.9',
        'Referer': 'https://fusion.fiemg.com.br/fusion/portal',
        'X-Requested-With': 'XMLHttpRequest'
    }
    
    return session, headers

def salvar_resposta_em_txt(resultado, caminho_arquivo='resposta_fusion.txt'):
    """
    Salva o conteúdo completo da resposta em um arquivo TXT para análise.
    
    Args:
        resultado: dict retornado pela função fazer_requisicao_fusion
        caminho_arquivo: caminho do arquivo TXT a ser criado
    """
    with open(caminho_arquivo, 'w', encoding='utf-8') as f:
        f.write("=" * 80 + "\n")
        f.write("ANÁLISE COMPLETA DA RESPOSTA FUSION\n")
        f.write("=" * 80 + "\n\n")
        
        # Informações básicas
        f.write(f"Status Code: {resultado.get('status_code')}\n")
        f.write(f"Success: {resultado.get('success')}\n")
        
        # Se for JSON
        if resultado.get('data'):
            f.write("-" * 80 + "\n")
            f.write("TIPO: JSON\n")
            f.write("-" * 80 + "\n\n")
            f.write(json.dumps(resultado['data'], indent=2, ensure_ascii=False))
        
        # Se for HTML/Texto
        elif resultado.get('full_html'):
            f.write("-" * 80 + "\n")
            f.write("TIPO: HTML/TEXTO COMPLETO\n")
            f.write("-" * 80 + "\n\n")
            f.write(resultado['full_html'])
        
        # Se houver erro
        elif resultado.get('error'):
            f.write("-" * 80 + "\n")
            f.write("ERRO\n")
            f.write("-" * 80 + "\n\n")
            f.write(str(resultado['error']))
        
        f.write("\n\n" + "=" * 80 + "\n")
        f.write("FIM DA ANÁLISE\n")
        f.write("=" * 80 + "\n")
    
    print(f"✅ Resposta salva em: {caminho_arquivo}")
    print(f"📄 Tamanho do arquivo: {os.path.getsize(caminho_arquivo):,} bytes")
