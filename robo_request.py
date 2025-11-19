from tkinter import messagebox
from utils import *
import random


class RoboRequest:
    def __init__(self):
        pass

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
            
            # Fazer requisição
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
            
            # Adicionar novos processos ao dicionário
            processos_novos = 0
            for code, process_id in resultado.items():
                if code not in dict_processos_raw:
                    dict_processos_raw[code] = process_id
                    processos_novos += 1
            
            print(f"✅ Novos processos obtidos: {processos_novos}")
            print(f"📊 Total acumulado: {len(dict_processos_raw)}/{numero_chamados}")
            
            # Se não obteve nenhum processo novo, significa que chegou ao fim
            if processos_novos == 0:
                print("⚠️ Não há mais processos disponíveis.")
                break
            
            # Incrementar offset para próxima página
            offset += range_size
            
            # ✅ SIMULAR TEMPO DE ESPERA HUMANO
            tempo_espera = random.uniform(2.0, 4.0)
            print(f"⏳ Aguardando {tempo_espera:.1f}s antes da próxima requisição...")
            time.sleep(tempo_espera)
        
        print("\n" + "=" * 60)
        print(f"✅ Coleta concluída: {len(dict_processos_raw)} processos")
        
        # ✅ FILTRAR PROCESSOS E OBTER ENTITY_IDs
        dict_processos = {}  # {code: entity_id} - IDs finais para requisição
        
        print("\n🔍 Filtrando processos e obtendo EntityIDs...")
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
                    print(f"[{idx}/{len(dict_processos_raw)}] 📋 {numero_limpo}: Não encontrado na planilha -> Obtendo EntityID")
                else:
                    linha_processo = planilha_referencia[planilha_referencia['numero_chamado'].str.contains(numero_limpo, na=False)]
                    if not linha_processo.empty:
                        acao_nucleo_atual = linha_processo['acao_nucleo'].iloc[0]
                        if acao_nucleo_atual != "Concluir Atendimento":
                            deve_obter = True
                            print(f"[{idx}/{len(dict_processos_raw)}] 📋 {numero_limpo}: Status '{acao_nucleo_atual}' -> Obtendo EntityID")
                        else:
                            print(f"[{idx}/{len(dict_processos_raw)}] ⏭️ {numero_limpo}: Já concluído -> Pulando")
            else:
                deve_obter = True
                print(f"[{idx}/{len(dict_processos_raw)}] 📋 {numero_limpo}: Sem planilha referência -> Obtendo EntityID")
            
            # ✅ OBTER ENTITY_ID
            if deve_obter:
                entity_id = self.obter_entity_id_do_processo(session, headers, str(process_id))
                
                if entity_id:
                    dict_processos[process_code] = int(entity_id)
                else:
                    print(f"⚠️ EntityId não obtido para {process_code}")
                
                # Tempo de espera entre requisições
                time.sleep(random.uniform(0.5, 2.0))
        
        print("\n" + "=" * 60)
        print(f"✅ Processos com EntityId: {len(dict_processos)}")
        
        # ✅ FAZER REQUISIÇÕES DOS FORMULÁRIOS
        print("\n📥 Baixando dados dos processos...")
        print("-" * 60)
        
        erros_totais = 0
        
        for idx, (processo, entity_id) in enumerate(dict_processos.items(), 1):
            form_id = str(entity_id)
            
            print(f"[{idx}/{len(dict_processos)}] 📥 {processo} (EntityID: {form_id})")
            
            resultado = fazer_requisicao_fusion(session, headers, form_id)
            html_completo = resultado.get('full_html', '')
            
            if not html_completo:
                erros_totais += 1
                print(f"❌ Erro ao obter dados")
                continue
            
            salvar_resposta_em_txt(resultado, os.path.join(setor_dir, f"resposta_fusion_{form_id}.txt"))
            print(f"✅ Arquivo salvo")
            
            time.sleep(random.uniform(0.5, 2.0))
        
        print("\n" + "=" * 60)
        print(f"📊 Resumo Final:")
        print(f"   Processos coletados: {len(dict_processos_raw)}")
        print(f"   EntityIDs obtidos: {len(dict_processos)}")
        print(f"   Arquivos salvos: {len(dict_processos) - erros_totais}")
        print(f"   Erros: {erros_totais}")
        print("=" * 60)
        
        return setor_dir

    def extrair_dados_do_txt(self, setor, nome_arquivo):
        """
        Lê um arquivo TXT e extrai os dados usando as mesmas lógicas do Selenium.
        
        Args:
            caminho_arquivo: Caminho para o arquivo TXT
        
        Returns:
            Lista com os dados extraídos na mesma ordem que o Selenium
        """
        try:
            caminho_arquivo = os.path.join(getattr(self, "setor_dir", ""), nome_arquivo)
            # Ler o arquivo
            with open(caminho_arquivo, 'r', encoding='utf-8') as f:
                conteudo = f.read()
            
            # Extrair apenas o HTML (após "TIPO: HTML/TEXTO COMPLETO")
            if 'TIPO: HTML/TEXTO COMPLETO' in conteudo:
                html_content = conteudo.split('TIPO: HTML/TEXTO COMPLETO')[1]
                html_content = html_content.split('=' * 80)[0]
            else:
                html_content = conteudo
            
            # Parsear HTML
            soup = BeautifulSoup(html_content, 'html.parser')
            
            # Função auxiliar MELHORADA para buscar por ID ou div pai
            def get_text(element_id, buscar_em_pai=False):
                """
                Busca texto por ID preservando quebras de linha (<br>) como '\n'.
                """
                el = soup.find(id=element_id)
                if not el:
                    return " "
                
                # Se precisa buscar no elemento pai (ex: div_Codigo__ -> div.text-wrapper)
                if buscar_em_pai:
                    text_wrapper = el.find('div', class_='text-wrapper')
                    if text_wrapper:
                        # ✅ CONVERTER <br> EM \n NO HTML INTERNO
                        html_interno = str(text_wrapper)
                        html_interno = html_interno.replace('<br>', '\n').replace('<br/>', '\n').replace('<br />', '\n')
                        # Parsear novamente para pegar o texto
                        temp_soup = BeautifulSoup(html_interno, 'html.parser')
                        texto = temp_soup.get_text()
                        texto = html_lib.unescape(texto)
                        return texto.strip()
                    

                if el.name == 'label':
                    parent = el.parent
                    if parent:
                        # Remover o próprio label e outros labels do parent
                        for label_tag in parent.find_all('label'):
                            label_tag.decompose()
                        # Remover tags indesejadas
                        for tag in parent.find_all(['script', 'style', 'input', 'select', 'textarea', 'img']):
                            tag.decompose()
                        # Converter <br> em \n
                        html_interno = str(parent)
                        html_interno = html_interno.replace('<br>', '\n').replace('<br/>', '\n').replace('<br />', '\n')
                        temp_soup = BeautifulSoup(html_interno, 'html.parser')
                        texto = temp_soup.get_text()
                        texto = html_lib.unescape(texto)
                        return texto.strip()
                
                # Para inputs hidden, buscar texto no parent
                if el.name == 'input' and el.get('type') == 'hidden':
                    parent = el.parent
                    if parent:
                        html_interno = str(parent)
                        html_interno = html_interno.replace('<br>', '\n').replace('<br/>', '\n').replace('<br />', '\n')
                        temp_soup = BeautifulSoup(html_interno, 'html.parser')
                        parent_text = temp_soup.get_text()
                        input_value = el.get('value', '')
                        if input_value and input_value in parent_text:
                            parent_text = parent_text.replace(input_value, '').strip()
                        parent_text = html_lib.unescape(parent_text)
                        return parent_text.strip()
                
                # Remover scripts, styles, inputs
                for tag in el.find_all(['script', 'style', 'input', 'select', 'textarea']):
                    tag.decompose()
                
                # ✅ CONVERTER <br> EM \n NO HTML INTERNO DO ELEMENTO
                html_interno = str(el)
                html_interno = html_interno.replace('<br>', '\n').replace('<br/>', '\n').replace('<br />', '\n')
                
                # Parsear novamente para pegar o texto com as quebras de linha
                temp_soup = BeautifulSoup(html_interno, 'html.parser')
                texto = temp_soup.get_text()
                texto = html_lib.unescape(texto)
                return texto.strip()
            # Função para tabelas (sem mudanças)
            def parse_data_textarea_soup(soup, textarea_id):
                """
                Parseia o conteúdo de uma textarea específica identificada por seu ID,
                extraindo informações de responsável, data e mensagem.
                Args:
                    soup: Objeto BeautifulSoup do HTML completo
                    textarea_id: ID da textarea a ser parseada
                Returns:
                    Lista com as entradas extraídas no formato "Responsável ; Data ; Mensagem"
                """
                try:
                    ta = soup.find(id=textarea_id)
                    if not ta:
                        return [" "]

                    xml_content = (ta.string or ta.get_text() or "").strip()
                    if not xml_content:
                        return [" "]
                    
                    # ✅ PRIMEIRO UNESCAPE
                    xml_content = html_lib.unescape(xml_content)
                    inner = BeautifulSoup(xml_content, "html.parser")

                    # ✅ BUSCAR E PROCESSAR SEÇÕES CDATA
                    responsavel = ""
                    data = ""
                    mensagem = ""

                    # 1. RESPONSÁVEL: Buscar no primeiro CDATA que contém input hidden
                    cdata_sections = inner.find_all(string=True)
                    for cdata in cdata_sections:
                        cdata_text = str(cdata).strip()
                        
                        # Se contém input hidden com responsavel
                        if 'responsavel' in cdata_text and 'input' in cdata_text:
                            # ✅ SEGUNDO UNESCAPE para decodificar &lt; &gt; &#39; etc.
                            decoded_cdata = html_lib.unescape(cdata_text)
                            temp_soup = BeautifulSoup(decoded_cdata, 'html.parser')
                            
                            # Buscar input hidden
                            input_el = temp_soup.find('input', type='hidden')
                            if input_el:
                                # Pegar texto após o input
                                parent_text = temp_soup.get_text()
                                input_value = input_el.get('value', '')
                                if input_value:
                                    parent_text = parent_text.replace(input_value, '').strip()
                                responsavel = parent_text.strip()
                                break

                    # 2. DATA: Buscar span com id contendo 'data__'
                    for cdata in cdata_sections:
                        cdata_text = str(cdata).strip()
                        
                        if 'data__' in cdata_text and 'span' in cdata_text:
                            # ✅ SEGUNDO UNESCAPE
                            decoded_cdata = html_lib.unescape(cdata_text)
                            temp_soup = BeautifulSoup(decoded_cdata, 'html.parser')
                            
                            # Buscar span de data
                            date_span = temp_soup.find('span', id=lambda x: x and 'data__' in x)
                            if date_span:
                                data = date_span.get_text(strip=True)
                                break

                    # 3. MENSAGEM: Buscar span com title (tooltip)
                    for cdata in cdata_sections:
                        cdata_text = str(cdata).strip()
                        
                        if 'tooltip' in cdata_text and 'title' in cdata_text:
                            # ✅ SEGUNDO UNESCAPE
                            decoded_cdata = html_lib.unescape(cdata_text)
                            temp_soup = BeautifulSoup(decoded_cdata, 'html.parser')
                            
                            # Buscar span com title
                            msg_span = temp_soup.find('span', title=True)
                            if msg_span:
                                mensagem = msg_span.get('title', '')
                                # ✅ TERCEIRO UNESCAPE para a mensagem (pode estar triplo-escapada)
                                mensagem = html_lib.unescape(mensagem)
                                break

                    # 4. BUSCAR NA TAG <overview> como fallback para mensagem
                    if not mensagem:
                        overview_tag = inner.find('overview')
                        if overview_tag:
                            overview_content = overview_tag.get_text(strip=True)
                            if overview_content:
                                # ✅ SEGUNDO UNESCAPE
                                mensagem = html_lib.unescape(overview_content)

                    # Limpeza final
                    responsavel = responsavel.replace(textarea_id.replace('data_', ''), '').strip()
                    data = data.replace('Data.:', '').replace('Data:', '').strip()
                    
                    # Remover duplicações na mensagem
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
                    return [" "]
            
            def get_data_textarea_ids_from_soup(soup, container_div_id=None, allow_global_fallback=True): # Obtém IDs de textareas data_ dentro de um container específico ou globalmente
                ids = [] # Lista para armazenar os IDs encontrados
                if container_div_id: # Se um ID de container for fornecido, buscar dentro dele
                    container = soup.find(id=container_div_id) # Encontrar o container pelo ID
                    if container: # Se o container for encontrado
                        ids.extend( # Extrair IDs de textareas data_ dentro do container
                            ta.get('id') # Obter o ID do textarea
                            for ta in container.find_all('textarea', id=lambda v: v and v.startswith('data_')) # Filtrar textareas com ID começando com 'data_'
                        )
                if not ids and allow_global_fallback: # Se nenhum ID foi encontrado e o fallback global é permitido
                    ids.extend( # Extrair IDs de textareas data_ globalmente
                        ta.get('id') # Obter o ID do textarea
                        for ta in soup.find_all('textarea', id=lambda v: v and v.startswith('data_')) # Filtrar textareas com ID começando com 'data_'
                    ) 
                return [i for i in ids if i] # Retornar a lista de IDs encontrados

            def parse_all_textareas(textarea_ids): # Parseia todas as textareas dadas pelos IDs e combina os resultados, removendo duplicatas
                combined, seen = [], set() # Listas para resultados combinados e conjunto para rastrear vistos
                for ta_id in textarea_ids: # Iterar sobre cada ID de textarea
                    parsed = parse_data_textarea_soup(soup, ta_id) # Parsear o conteúdo da textarea
                    for entry in parsed: # Iterar sobre cada entrada parseada
                        entry = entry.strip() # Remover espaços extras
                        if entry and entry not in seen: # Se a entrada não for vazia e não tiver sido vista antes
                            seen.add(entry) # Marcar como visto
                            combined.append(entry) # Adicionar à lista combinada
                return combined or [" "] # Retornar a lista combinada ou uma lista com espaço se vazia

            # Extrair dados na mesma ordem do Selenium
            numero_chamado = get_text("div_Codigo__", buscar_em_pai=True)  # Número do chamado
            data_inicial = get_text("var_DadosDaSolicitacao__Responsavel__data__").replace('Data.:', '').replace('Data:', '').strip() # Data inicial
            responsavel = get_text("var_DadosDaSolicitacao__Responsavel__responsavel__") # Responsável pela solicitação
            detalhes_solicitacao = get_text("var_DadosDaSolicitacao__DescricaoDaDemanda___view_textarea") # Detalhes da solicitação
            uo = '' # Unidade Organizacional
            if detalhes_solicitacao: # Tentar extrair UO dos detalhes da solicitação
                for uo_unidade in uo_dict.uo_dict:
                    pattern = rf'(?<!\d){uo_unidade}(?!\d)'
                    if re.search(pattern, detalhes_solicitacao):
                        uo = str(uo_unidade)
                        break
            if detalhes_solicitacao == "": # Tentar extrair detalhes da solicitação alternativos
                detalhes_solicitacao = get_text("var_DadosDaSolicitacao__NecessidadeDeCompras__JustificativaDaDefinicaoDeUrgencia___view_textarea", buscar_em_pai=True) # Detalhes da solicitação alternativos
            urgencia_demanda = get_text("label_DadosDaSolicitacao__UrgenciaDemanda__").replace('Urgência da Demanda:', '').strip() # Urgência da demanda
            if urgencia_demanda == "": # Tentar extrair urgência da demanda alternativos
                urgencia_demanda = get_text("label_DadosDaSolicitacao__VariacoesDaDemanda__").replace('Urgência da Demanda:', '').strip() # Urgência da demanda alternativos
            justificativa_demanda = get_text("var_DadosDaSolicitacao__JustificativaDaDefinicaoDeUrgencia___view_textarea") # Justificativa da demanda
            if justificativa_demanda == "": # Tentar extrair justificativa da demanda alternativos
                justificativa_demanda = get_text("var_DadosDaSolicitacao__NecessidadeDeCompras__JustificativaDaNecessidadeDeCompra___view_textarea", buscar_em_pai=True) # Justificativa da demanda alternativos
            data_atual_supervisor = get_text("var_SupervisorAnalise__DataAtual__").replace('Data Atual:', '').strip() # Data atual do supervisor
            prazo_final = get_text("var_SupervisorAnalise__PrazoDeAtendimento__").replace('Prazo de Atendimento:', '').strip() # Prazo final
            supervisor_ids = get_data_textarea_ids_from_soup(
                soup, 'dlist_SupervisorAnalise__HistoricoDeAtendimentoNucleoAdministrativo__'
            ) # IDs das textareas do supervisor
            supervisor_ids_alt = get_data_textarea_ids_from_soup(
                soup, 'dlist_SupervisorAnalise__HistoricoDeAtendimento__'
            ) # IDs das textareas do supervisor alternativos
            nucleo_ids = get_data_textarea_ids_from_soup(
                soup, 'dlist_NucleoAdministrativoAnalise__Historico__', allow_global_fallback=False
            ) # IDs das textareas do núcleo administrativo
            historico_nucleo = parse_all_textareas(nucleo_ids) # Histórico do núcleo administrativo
            encaminhamento_supervisor = parse_all_textareas(supervisor_ids) # Encaminhamento do supervisor
            if encaminhamento_supervisor == [" "]: # Tentar extrair encaminhamento do supervisor alternativos
                encaminhamento_supervisor = parse_all_textareas(supervisor_ids_alt) # Encaminhamento do supervisor alternativos
            acao_supervisor = get_text("label_SupervisorAnalise__Acao__").replace('Ação:', '').strip() # Ação do supervisor
            responsavel_nucleo = get_text("var_NucleoAdministrativoAnalise__Responsavel__responsavel__") # Responsável do núcleo administrativo
            acao_nucleo = get_text("label_NucleoAdministrativoAnalise__Acoes__").replace('Ação:', '').strip() # Ação do núcleo administrativo
            
            return [
                numero_chamado, setor ,data_inicial, responsavel, uo, detalhes_solicitacao,
                urgencia_demanda, justificativa_demanda, data_atual_supervisor, prazo_final,
                encaminhamento_supervisor, acao_supervisor, historico_nucleo, responsavel_nucleo, acao_nucleo, '', ''
            ] # Retorna os dados extraídos em uma lista na ordem correta
            
        except Exception as e:
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