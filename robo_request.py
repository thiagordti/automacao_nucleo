from tkinter import messagebox
from utils import *
from concurrent.futures import ThreadPoolExecutor, as_completed
from threading import Semaphore
import time
import random


class RoboRequest:
    def __init__(self, max_workers=3, requests_per_second=2):
        self.max_workers = max_workers
        self.rate_limiter = Semaphore(requests_per_second)
        self.request_interval = 1.0 / requests_per_second

    def iniciar_navegador(self):
        navegador, chrome_proc = iniciar_navegador() # Inicia o navegador e conecta ao Chrome em modo debug
        # Utiliza Tkinter para informar a necessidade de login ao usuario e so continua apos o login ser realizado
        messagebox.showinfo("Login Necessário", "Realize o login no portal e clique em OK para continuar...")
        return navegador, chrome_proc # Retorna o objeto do navegador e o processo do Chrome em modo debug
    
    def finalizar_navegador(self, navegador, chrome_proc):
        fechar_navegador(navegador, chrome_proc) # Fecha o navegador e o processo do Chrome em modo debug

    def selecionar_planilha(self):
        planilha = selecionar_planilha_excel(titulo="Selecionar arquivo Excel/CSV", tipo_arquivo=(("Arquivos Excel", "*.xlsx"), ("Arquivos CSV", "*.csv"), ("Todos os arquivos", "*.*")))
        df = pd.read_csv(planilha, dtype=str, sep=';', encoding='utf-8') if planilha.lower().endswith('.csv') else pd.read_excel(planilha)
        return df

    def caminho_salvar_arquivo(self):
        return selecionar_caminho_para_salvar() # Abre uma janela para selecionar o caminho de salvamento do arquivo
    
    def fazer_requisicao_fusion(self,session, headers, form_id):
        """
        Faz requisição AJAX para o Fusion e retorna o response.
        
        Args:
            session: requests.Session com cookies
            headers: dict com headers
            form_id: ID do formulário (ex: 363321568)
        
        Returns:
            dict com response_data e status_code
        """
        url = f"https://fusion.fiemg.com.br/fusion/portal/ajaxRender/Form"
        
        params = {
            'content': 'true',
            'id': form_id,
            'showContainer': 'false',
            'disableFooter': 'true',
            'full': 'true',
            'edit': 'false',
            'type': 'DSSNAFPDadosDaSolicitacao'
        }
        
        try:
            response = session.get(url, params=params, headers=headers, timeout=30)
            
            print(f"Status Code: {response.status_code}")
            print(f"Content-Type: {response.headers.get('Content-Type')}")
            
            # Tentar parsear como JSON
            try:
                data = response.json()
                return {
                    'success': True,
                    'status_code': response.status_code,
                    'data': data,
                    'raw_text': response.text[:500]  # primeiros 500 chars
                }
            except:
                # Se não for JSON, retornar HTML/texto
                return {
                    'success': True,
                    'status_code': response.status_code,
                    'data': None,
                    'raw_text': response.text[:500],
                    'full_html': response.text
                }
        
        except Exception as e:
            return {
                'success': False,
                'error': str(e)
            }

    def fazer_requisicao_wfprocess(self, navegador, setor, data_busca, offset=0, range_size=100):
        """
        Faz requisição direta ao endpoint WFProcess usando sessão do navegador.
        
        Args:
            navegador: Instância do Selenium WebDriver (navegador logado)
            setor: Nome do setor ("Compras", "Financeiro", "Patrimônio" ou "Regularidade")
            data_busca: Data no formato DD/MM/YYYY
            offset: Posição inicial (paginação)
            range_size: Quantidade de itens por requisição
        
        Returns:
            dict: Dicionário {code: id} ou None em caso de erro
        """
        import requests
        from datetime import datetime
        
        # ✅ CONFIGURAÇÕES POR SETOR
        config_setores = {
            "Compras": {
                "neoId": 361179933,
                "filter_neoId": 366063102,
                "filter_ids": [366063103, 366063104, 366063105],
                "loperand2": 80823850
            },
            "Financeiro": {
                "neoId": 361180049,
                "filter_neoId": 366063106,
                "filter_ids": [366063107, 366063108, 366063109],
                "loperand2": 145984
            },
            "Patrimônio": {
                "neoId": 361180097,
                "filter_neoId": 366063110,
                "filter_ids": [366063111, 366063112, 366063113],
                "loperand2": 44516
            },
            "Regularidade": {
                "neoId": 361180230,
                "filter_neoId": 366063134,
                "filter_ids": [366063135, 366063136, 366063137],
                "loperand2": 40886
            }
        }
        
        # Validar setor
        if setor not in config_setores:
            print(f"❌ Setor '{setor}' inválido. Use: Compras, Financeiro, Patrimônio ou Regularidade")
            return None
        
        config = config_setores[setor]
        
        # ✅ CAPTURAR COOKIES E HEADERS DO NAVEGADOR
        selenium_cookies = navegador.get_cookies()
        
        session = requests.Session()
        for cookie in selenium_cookies:
            session.cookies.set(cookie['name'], cookie['value'])
        
        headers = {
            'User-Agent': navegador.execute_script("return navigator.userAgent;"),
            'Accept': 'application/json, text/plain, */*',
            'Accept-Language': 'pt-BR,pt;q=0.9',
            'Content-Type': 'application/json',
            'Referer': 'https://fusion.fiemg.com.br/fusion/portal',
            'X-Requested-With': 'XMLHttpRequest'
        }
        
        # ✅ CONVERTER DATA DD/MM/YYYY -> ISO 8601
        data_obj = datetime.strptime(data_busca, "%d/%m/%Y")
        data_iso = data_obj.strftime("%Y-%m-%dT03:00:00.000Z")
        
        # ✅ URL E PAYLOAD DINÂMICO
        url = "https://fusion.fiemg.com.br/fusion/services/process/advancedSearch/WFProcess"
        
        payload = {
            "offset": offset,
            "range": range_size,
            "entityFilter": {
                "neoId": config["neoId"],
                "entityName": "WFProcess",
                "name": f"Núcleo - {setor}",
                "description": None,
                "userDefault": False,
                "filter": {
                    "neoType": "Fpersist.QLGroupFilter",
                    "neoId": config["filter_neoId"],
                    "@id": config["filter_neoId"],
                    "operator": "and",
                    "filterList": [
                        {
                            "neoType": "Fpersist.QLEqualsFilter",
                            "neoId": config["filter_ids"][0],
                            "@id": config["filter_ids"][0],
                            "operand1": "model.versionControl.proxy",
                            "operand2Type": "NeoObject",
                            "operator": "=",
                            "loperand2": 315083955
                        },
                        {
                            "neoType": "Fpersist.QLOpFilter",
                            "neoId": config["filter_ids"][1],
                            "operand1": "startDate",
                            "operator": ">=",
                            "operand2Type": "Date",
                            "doperand2": data_iso
                        },
                        {
                            "neoType": "Fpersist.QLEqualsFilter",
                            "neoId": config["filter_ids"][2],
                            "@id": config["filter_ids"][2],
                            "operand1": "taskSet.taskList.user",
                            "operand2Type": "NeoObject",
                            "operator": "=",
                            "loperand2": config["loperand2"]
                        }
                    ]
                },
                "@type": "EntityFilter"
            },
            "entityTypeName": "WFProcess",
            "groupBy": "STARTDATE",
            "order": "desc",
            "formFieldSelected": [],
            "processSelected": 315083955,
            "haveFormFieldSelected": False,
            "haveActivityModelSelected": False,
            "haveManagerSelected": False,
            "haveParticipantsSelected": False,
            "showSubProcess": False,
            "userCode": "trduarte",
            "ignoreFiltersEmpty": False,
            "managedUsers": [205270532],
            "allUsersSelected": False
        }
        
        # ✅ FAZER REQUISIÇÃO
        try:
            print(f"📤 [{setor}] Requisição: offset={offset}, range={range_size}")
            
            response = session.post(url, json=payload, headers=headers, timeout=60)
            
            print(f"📥 Status: {response.status_code}")
            
            if response.status_code == 200:
                data = response.json()
                
                # ✅ CRIAR DICIONÁRIO {CODE: ID}
                dict_processos = {}
                
                for data_grupo, processos_grupo in data.items():
                    for process_id_str, process_data in processos_grupo.items():
                        code = process_data.get('code')
                        process_id = process_data.get('id')
                        
                        if code and process_id:
                            dict_processos[code] = process_id
                
                print(f"✅ Obtidos {len(dict_processos)} processos")
                return dict_processos
            else:
                print(f"❌ Erro HTTP {response.status_code}")
                print(f"Resposta: {response.text[:300]}")
                return None
                
        except Exception as e:
            print(f"❌ Erro na requisição: {e}")
            return None

    def extracao_dados_chamados(self, navegador, setor, data_busca, planilha, planilha_referencia=None):
        navegador.switch_to.default_content()
        clicar_elemento(navegador, '//*[@id="menu_item_small_30"]/div/div', By.XPATH, automacao_fusion_instance=None)
        clicar_elemento(navegador, '#menu_item_expanded_301 > div > div.menu-inner-cell.menu-inner-cell-middle > p', By.CSS_SELECTOR, automacao_fusion_instance=None)
        acessar_iframe(navegador, 2, automacao_fusion_instance=None, timeout=15)
        clicar_entidade_por_nome(navegador, f"Núcleo - {setor}", automacao_fusion_instance=None)
        
        # Criar diretório do setor
        base_dir = r"C:\AutomacaoFusion"
        setor_dir = os.path.join(base_dir, setor)
        os.makedirs(setor_dir, exist_ok=True)
        self.setor_dir = setor_dir
        
        # Capturar sessão
        session, headers = capturar_cookies_e_headers(navegador)
        
        numero_chamados = len(planilha)
        
        dict_processos_raw = {}  # {code: process_id} - IDs brutos da API
        offset = 0
        range_size = 100
        
        print(f"🔄 Iniciando coleta de processos via API...")
        print(f"📊 Meta: {numero_chamados} processos")
        print("-" * 60)
        
        while len(dict_processos_raw) < numero_chamados:
            print(f"\n📤 Requisição: offset={offset}, range={range_size}")
            
            resultado = self.fazer_requisicao_wfprocess(
                navegador=navegador,
                setor=setor,
                data_busca=data_busca,
                offset=offset,
                range_size=range_size
            )
            
            if not resultado:
                print("❌ Falha na requisição. Encerrando coleta.")
                break
            
            processos_novos = 0
            for code, process_id in resultado.items():
                if code not in dict_processos_raw:
                    dict_processos_raw[code] = process_id
                    processos_novos += 1
            
            print(f"✅ Novos processos obtidos: {processos_novos}")
            print(f"📊 Total acumulado: {len(dict_processos_raw)}/{numero_chamados}")
            
            if processos_novos == 0:
                print("⚠️ Não há mais processos disponíveis.")
                break
            
            offset += range_size
            tempo_espera = random.uniform(2.0, 4.0)
            print(f"⏳ Aguardando {tempo_espera:.1f}s antes da próxima requisição...")
            time.sleep(tempo_espera)
        
        print("\n" + "=" * 60)
        print(f"✅ Coleta concluída: {len(dict_processos_raw)} processos")
        
        # ✅ FILTRAR PROCESSOS QUE PRECISAM DE ENTITY_ID
        dict_processos_filtrados = {}  # {code: process_id} - Processos que precisam ser baixados
        
        print("\n🔍 Filtrando processos...")
        print("-" * 60)
        
        for idx, (process_code, process_id) in enumerate(dict_processos_raw.items(), 1):
            # Extrair apenas o número do código (ex: SSNA.004077/2025 -> 004077)
            match = re.search(r'\.(\d+)/', process_code)
            numero_limpo = match.group(1) if match else process_code
            
            # ✅ VERIFICAR SE DEVE OBTER ENTITY_ID
            deve_obter = False
            
            if planilha_referencia is not None and not planilha_referencia.empty:
                processo_existe = planilha_referencia['numero_chamado'].str.contains(numero_limpo, na=False).any()
                
                if not processo_existe:
                    deve_obter = True
                    print(f"[{idx}/{len(dict_processos_raw)}] 📋 {numero_limpo}: Não encontrado -> Marcado para download")
                else:
                    linha_processo = planilha_referencia[planilha_referencia['numero_chamado'].str.contains(numero_limpo, na=False)]
                    if not linha_processo.empty:
                        acao_nucleo_atual = linha_processo['acao_nucleo'].iloc[0]
                        if acao_nucleo_atual != "Concluir Atendimento":
                            deve_obter = True
                            print(f"[{idx}/{len(dict_processos_raw)}] 📋 {numero_limpo}: Status '{acao_nucleo_atual}' -> Marcado para download")
                        else:
                            print(f"[{idx}/{len(dict_processos_raw)}] ⏭️ {numero_limpo}: Já concluído -> Pulando")
            else:
                deve_obter = True
                print(f"[{idx}/{len(dict_processos_raw)}] 📋 {numero_limpo}: Sem planilha referência -> Marcado para download")
            
            if deve_obter:
                dict_processos_filtrados[process_code] = process_id
        
        print("\n" + "=" * 60)
        print(f"✅ Processos filtrados: {len(dict_processos_filtrados)}/{len(dict_processos_raw)}")
        
        # ✅ OBTER ENTITY_IDs EM PARALELO
        print("\n🔄 Obtendo EntityIDs em paralelo...")
        print("-" * 60)
        
        dict_processos = self.obter_entity_ids_batch(session, headers, dict_processos_filtrados)
        
        print("\n" + "=" * 60)
        print(f"✅ EntityIDs obtidos: {len(dict_processos)}")
        
        # ✅ BAIXAR DADOS DOS PROCESSOS EM PARALELO
        print("\n📥 Baixando dados dos processos em paralelo...")
        print("-" * 60)
        
        erros_totais = self.baixar_dados_processos_batch(session, headers, dict_processos, setor_dir)
        
        print("\n" + "=" * 60)
        print(f"📊 Resumo Final:")
        print(f"   Processos coletados: {len(dict_processos_raw)}")
        print(f"   Processos filtrados: {len(dict_processos_filtrados)}")
        print(f"   EntityIDs obtidos: {len(dict_processos)}")
        print(f"   Arquivos salvos: {len(dict_processos) - erros_totais}")
        print(f"   Erros: {erros_totais}")
        print("=" * 60)
        
        return setor_dir

    def extrair_dados_do_txt(self, setor, nome_arquivo):
        """
        Lê um arquivo TXT e extrai os dados usando as mesmas lógicas do Selenium.
        """
        try:
            print(f"📄 Iniciando extração: {nome_arquivo}")
            
            caminho_arquivo = os.path.join(getattr(self, "setor_dir", ""), nome_arquivo)
            
            # ✅ Ler arquivo
            print("  ⏳ Lendo arquivo...")
            with open(caminho_arquivo, 'r', encoding='utf-8') as f:
                conteudo = f.read()
            
            # ✅ Extrair HTML
            print("  ⏳ Extraindo HTML...")
            if 'TIPO: HTML/TEXTO COMPLETO' in conteudo:
                html_content = conteudo.split('TIPO: HTML/TEXTO COMPLETO')[1]
                html_content = html_content.split('=' * 80)[0]
            else:
                html_content = conteudo
            
            # ✅ Parsear HTML UMA VEZ com parser mais rápido
            print("  ⏳ Parseando HTML...")
            soup = BeautifulSoup(html_content, 'lxml')  # 'lxml' é mais rápido que 'html.parser'
            
            # ✅ CACHE: Criar dicionário de elementos por ID (parsing único)
            print("  ⏳ Criando cache de elementos...")
            elementos_cache = {el.get('id'): el for el in soup.find_all(id=True) if el.get('id')}
            
            # Função otimizada com cache
            def get_text_cached(element_id, buscar_em_pai=False):
                """Versão otimizada com cache de elementos"""
                el = elementos_cache.get(element_id)
                if not el:
                    return " "
                
                if buscar_em_pai:
                    text_wrapper = el.find('div', class_='text-wrapper')
                    if text_wrapper:
                        html_interno = str(text_wrapper)
                        html_interno = html_interno.replace('<br>', '\n').replace('<br/>', '\n').replace('<br />', '\n')
                        temp_soup = BeautifulSoup(html_interno, 'lxml')
                        texto = temp_soup.get_text()
                        texto = html_lib.unescape(texto)
                        return texto.strip()
                
                if el.name == 'label':
                    parent = el.parent
                    if parent:
                        # Clonar para não modificar o original
                        parent_copy = BeautifulSoup(str(parent), 'lxml').find()
                        for label_tag in parent_copy.find_all('label'):
                            label_tag.decompose()
                        for tag in parent_copy.find_all(['script', 'style', 'input', 'select', 'textarea', 'img']):
                            tag.decompose()
                        
                        html_interno = str(parent_copy)
                        html_interno = html_interno.replace('<br>', '\n').replace('<br/>', '\n').replace('<br />', '\n')
                        temp_soup = BeautifulSoup(html_interno, 'lxml')
                        texto = temp_soup.get_text()
                        texto = html_lib.unescape(texto)
                        return texto.strip()
                
                if el.name == 'input' and el.get('type') == 'hidden':
                    parent = el.parent
                    if parent:
                        html_interno = str(parent)
                        html_interno = html_interno.replace('<br>', '\n').replace('<br/>', '\n').replace('<br />', '\n')
                        temp_soup = BeautifulSoup(html_interno, 'lxml')
                        parent_text = temp_soup.get_text()
                        input_value = el.get('value', '')
                        if input_value and input_value in parent_text:
                            parent_text = parent_text.replace(input_value, '').strip()
                        parent_text = html_lib.unescape(parent_text)
                        return parent_text.strip()
                
                # Clonar elemento para não modificar o original
                el_copy = BeautifulSoup(str(el), 'lxml').find()
                for tag in el_copy.find_all(['script', 'style', 'input', 'select', 'textarea']):
                    tag.decompose()
                
                html_interno = str(el_copy)
                html_interno = html_interno.replace('<br>', '\n').replace('<br/>', '\n').replace('<br />', '\n')
                temp_soup = BeautifulSoup(html_interno, 'lxml')
                texto = temp_soup.get_text()
                texto = html_lib.unescape(texto)
                return texto.strip()
            
            # ✅ Manter parse_data_textarea_soup otimizado
            def parse_data_textarea_soup(textarea_id):
                """Versão otimizada com cache"""
                try:
                    ta = elementos_cache.get(textarea_id)
                    if not ta:
                        return [" "]

                    xml_content = (ta.string or ta.get_text() or "").strip()
                    if not xml_content:
                        return [" "]
                    
                    xml_content = html_lib.unescape(xml_content)
                    inner = BeautifulSoup(xml_content, "lxml")

                    responsavel = ""
                    data = ""
                    mensagem = ""

                    cdata_sections = inner.find_all(string=True)
                    
                    for cdata in cdata_sections:
                        cdata_text = str(cdata).strip()
                        
                        if 'responsavel' in cdata_text and 'input' in cdata_text:
                            decoded_cdata = html_lib.unescape(cdata_text)
                            temp_soup = BeautifulSoup(decoded_cdata, 'lxml')
                            
                            input_el = temp_soup.find('input', type='hidden')
                            if input_el:
                                parent_text = temp_soup.get_text()
                                input_value = input_el.get('value', '')
                                if input_value:
                                    parent_text = parent_text.replace(input_value, '').strip()
                                responsavel = parent_text.strip()
                                break

                    for cdata in cdata_sections:
                        cdata_text = str(cdata).strip()
                        
                        if 'data__' in cdata_text and 'span' in cdata_text:
                            decoded_cdata = html_lib.unescape(cdata_text)
                            temp_soup = BeautifulSoup(decoded_cdata, 'lxml')
                            
                            date_span = temp_soup.find('span', id=lambda x: x and 'data__' in x)
                            if date_span:
                                data = date_span.get_text(strip=True)
                                break

                    for cdata in cdata_sections:
                        cdata_text = str(cdata).strip()
                        
                        if 'tooltip' in cdata_text and 'title' in cdata_text:
                            decoded_cdata = html_lib.unescape(cdata_text)
                            temp_soup = BeautifulSoup(decoded_cdata, 'lxml')
                            
                            msg_span = temp_soup.find('span', title=True)
                            if msg_span:
                                mensagem = msg_span.get('title', '')
                                mensagem = html_lib.unescape(mensagem)
                                break

                    if not mensagem:
                        overview_tag = inner.find('overview')
                        if overview_tag:
                            overview_content = overview_tag.get_text(strip=True)
                            if overview_content:
                                mensagem = html_lib.unescape(overview_content)

                    responsavel = responsavel.replace(textarea_id.replace('data_', ''), '').strip()
                    data = data.replace('Data.:', '').replace('Data:', '').strip()
                    
                    def _collapse_double(s):
                        s = s.strip()
                        n = len(s)
                        for mid in range(1, n):
                            left = s[:mid].strip()
                            right = s[mid:].strip()
                            if left and left == right:
                                return left
                        return s

                    mensagem = _collapse_double(mensagem)

                    if data and responsavel:
                        combined = f"{responsavel} ; {data} ; {mensagem}"
                        return [combined]
                    else:
                        return [" "]
                        
                except Exception as e:
                    print(f"    ⚠️ Erro ao parsear textarea {textarea_id}: {e}")
                    return [" "]
            
            def get_data_textarea_ids_from_soup(container_div_id=None, allow_global_fallback=True):
                """Versão otimizada com cache"""
                ids = []
                
                if container_div_id:
                    container = elementos_cache.get(container_div_id)
                    if container:
                        ids.extend(
                            ta.get('id')
                            for ta in container.find_all('textarea', id=lambda v: v and v.startswith('data_'))
                        )
                
                if not ids and allow_global_fallback:
                    ids.extend(
                        el_id for el_id in elementos_cache.keys()
                        if el_id.startswith('data_') and elementos_cache[el_id].name == 'textarea'
                    )
                
                return [i for i in ids if i]

            def parse_all_textareas(textarea_ids):
                combined, seen = [], set()
                for ta_id in textarea_ids:
                    parsed = parse_data_textarea_soup(ta_id)
                    for entry in parsed:
                        entry = entry.strip()
                        if entry and entry not in seen:
                            seen.add(entry)
                            combined.append(entry)
                return combined or [" "]

            # ✅ Extração com prints de progresso

            numero_chamado = get_text_cached("div_Codigo__", buscar_em_pai=True)

            data_inicial = get_text_cached("var_DadosDaSolicitacao__Responsavel__data__").replace('Data.:', '').replace('Data:', '').strip()

            responsavel = get_text_cached("var_DadosDaSolicitacao__Responsavel__responsavel__")

            detalhes_solicitacao = get_text_cached("var_DadosDaSolicitacao__DescricaoDaDemanda___view_textarea")
            uo = ''
            if detalhes_solicitacao:
                for uo_unidade in uo_dict.uo_dict:
                    pattern = rf'(?<!\d){uo_unidade}(?!\d)'
                    if re.search(pattern, detalhes_solicitacao):
                        uo = str(uo_unidade)
                        break
            
            if detalhes_solicitacao == "":
                detalhes_solicitacao = get_text_cached("var_DadosDaSolicitacao__NecessidadeDeCompras__JustificativaDaDefinicaoDeUrgencia___view_textarea", buscar_em_pai=True)
            
            print("  ⏳ Extraindo: Urgência da demanda...")
            urgencia_demanda = get_text_cached("label_DadosDaSolicitacao__UrgenciaDemanda__").replace('Urgência da Demanda:', '').strip()
            if urgencia_demanda == "":
                urgencia_demanda = get_text_cached("label_DadosDaSolicitacao__VariacoesDaDemanda__").replace('Urgência da Demanda:', '').strip()
            
            print("  ⏳ Extraindo: Justificativa da demanda...")
            justificativa_demanda = get_text_cached("var_DadosDaSolicitacao__JustificativaDaDefinicaoDeUrgencia___view_textarea")
            if justificativa_demanda == "":
                justificativa_demanda = get_text_cached("var_DadosDaSolicitacao__NecessidadeDeCompras__JustificativaDaNecessidadeDeCompra___view_textarea", buscar_em_pai=True)
            
            print("  ⏳ Extraindo: Datas do supervisor...")
            data_atual_supervisor = get_text_cached("var_SupervisorAnalise__DataAtual__").replace('Data Atual:', '').strip()
            prazo_final = get_text_cached("var_SupervisorAnalise__PrazoDeAtendimento__").replace('Prazo de Atendimento:', '').strip()
            
            print("  ⏳ Extraindo: IDs de textareas do supervisor...")
            supervisor_ids = get_data_textarea_ids_from_soup('dlist_SupervisorAnalise__HistoricoDeAtendimentoNucleoAdministrativo__')
            supervisor_ids_alt = get_data_textarea_ids_from_soup('dlist_SupervisorAnalise__HistoricoDeAtendimento__')
            
            print("  ⏳ Extraindo: IDs de textareas do núcleo...")
            nucleo_ids = get_data_textarea_ids_from_soup('dlist_NucleoAdministrativoAnalise__Historico__', allow_global_fallback=False)
            
            print("  ⏳ Parseando: Histórico do núcleo...")
            historico_nucleo = parse_all_textareas(nucleo_ids)
            
            print("  ⏳ Parseando: Encaminhamento do supervisor...")
            encaminhamento_supervisor = parse_all_textareas(supervisor_ids)
            if encaminhamento_supervisor == [" "]:
                encaminhamento_supervisor = parse_all_textareas(supervisor_ids_alt)
            
            print("  ⏳ Extraindo: Ações...")
            acao_supervisor = get_text_cached("label_SupervisorAnalise__Acao__").replace('Ação:', '').strip()
            responsavel_nucleo = get_text_cached("var_NucleoAdministrativoAnalise__Responsavel__responsavel__")
            acao_nucleo = get_text_cached("label_NucleoAdministrativoAnalise__Acoes__").replace('Ação:', '').strip()
            
            print(f"✅ Extração concluída: {nome_arquivo}")
            
            return [
                numero_chamado, setor, data_inicial, responsavel, uo, detalhes_solicitacao,
                urgencia_demanda, justificativa_demanda, data_atual_supervisor, prazo_final,
                encaminhamento_supervisor, acao_supervisor, historico_nucleo, responsavel_nucleo, acao_nucleo, '', ''
            ]
            
        except Exception as e:
            print(f"❌ Erro na extração de {nome_arquivo}: {e}")
            return None
        
    def obter_entity_id_do_processo(self, session, headers, process_id):
        """
        Faz requisição para o endpoint BPM e extrai o entityId.
        """
        url = f"https://fusion.fiemg.com.br/fusion/rest/bpm/entity/process/wfProcessVo/{process_id}"
        
        try:
            response = session.get(url, headers=headers, timeout=30)
            
            if response.status_code == 200:
                try:
                    data = response.json()
                    entity_id = data.get('entityId')
                    
                    if entity_id:
                        return str(entity_id)
                    else:
                        print(f"⚠️ EntityId não encontrado para processo {process_id}")
                        return None
                        
                except Exception as e:
                    print(f"❌ Erro ao parsear JSON para processo {process_id}: {e}")
                    return None
            else:
                print(f"❌ Erro HTTP {response.status_code} para processo {process_id}")
                return None
                
        except Exception as e:
            print(f"❌ Erro na requisição para processo {process_id}: {e}")
            return None
    
    def obter_entity_ids_batch(self, session, headers, process_ids_dict):
        """Obtém múltiplos EntityIDs em paralelo com rate limiting"""
        results = {}
        total = len(process_ids_dict)
        processados = 0
        
        def fetch_with_limit(idx, process_code, process_id):
            self.rate_limiter.acquire()
            time.sleep(self.request_interval)
            try:
                print(f"  [{idx}/{total}] 🔍 Buscando EntityID para {process_code} (ProcessID: {process_id})...")
                entity_id = self.obter_entity_id_do_processo(session, headers, str(process_id))
                
                if entity_id:
                    print(f"  [{idx}/{total}] ✅ {process_code} → EntityID: {entity_id}")
                    return process_code, entity_id
                else:
                    print(f"  [{idx}/{total}] ❌ {process_code} → EntityID não encontrado")
                    return process_code, None
            except Exception as e:
                print(f"  [{idx}/{total}] ❌ {process_code} → Erro: {e}")
                return process_code, None
            finally:
                self.rate_limiter.release()
        
        with ThreadPoolExecutor(max_workers=self.max_workers) as executor:
            futures = {
                executor.submit(fetch_with_limit, idx, code, pid): code 
                for idx, (code, pid) in enumerate(process_ids_dict.items(), 1)
            }
            
            for future in as_completed(futures):
                code, entity_id = future.result()
                if entity_id:
                    results[code] = int(entity_id)
                
                processados += 1
                # Print de progresso a cada 10 itens
                if processados % 10 == 0 or processados == total:
                    print(f"\n📊 Progresso: {processados}/{total} processos ({(processados/total)*100:.1f}%)")
                    print(f"   ✅ EntityIDs obtidos: {len(results)}")
                    print(f"   ❌ Falhas: {processados - len(results)}\n")
        
        return results
    
    def baixar_dados_processos_batch(self, session, headers, dict_processos, setor_dir):
        """Baixa dados de múltiplos processos em paralelo"""
        erros = 0
        total = len(dict_processos)
        
        def fetch_processo(idx, processo, entity_id):
            self.rate_limiter.acquire()
            time.sleep(self.request_interval)
            try:
                print(f"[{idx}/{total}] 📥 {processo} (EntityID: {entity_id})")
                resultado = self.fazer_requisicao_fusion(session, headers, str(entity_id))
                
                html_completo = resultado.get('full_html', '')
                if html_completo:
                    salvar_resposta_em_txt(resultado, os.path.join(setor_dir, f"resposta_fusion_{entity_id}.txt"))
                    print(f"✅ Arquivo salvo")
                    return True
                return False
            finally:
                self.rate_limiter.release()
        
        with ThreadPoolExecutor(max_workers=self.max_workers) as executor:
            futures = {
                executor.submit(fetch_processo, idx, proc, eid): proc
                for idx, (proc, eid) in enumerate(dict_processos.items(), 1)
            }
            
            for future in as_completed(futures):
                if not future.result():
                    erros += 1
        
        return erros
    
    def processar_arquivos_batch(self, setor, lista_arquivos):
        """Processa múltiplos arquivos em paralelo"""
        
        def processar_arquivo(nome_arquivo):
            return self.extrair_dados_do_txt(setor, nome_arquivo)
        
        with ThreadPoolExecutor(max_workers=10) as executor:
            resultados = list(executor.map(processar_arquivo, lista_arquivos))
        
        return resultados