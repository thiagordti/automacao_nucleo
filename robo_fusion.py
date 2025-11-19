
from utils import *

class RoboFusion:
    def __init__(self):
        pass
    
    def iniciar_navegador(self):
        navegador, chrome_proc = iniciar_navegador() # Inicia o navegador e conecta ao Chrome em modo debug
        input("Realize o login no portal e pressione Enter para continuar...") # Pausa para o login manual
        return navegador, chrome_proc
    
    def finalizar_navegador(self, navegador, chrome_proc):
        fechar_navegador(navegador, chrome_proc) # Fecha o navegador e o processo do Chrome em modo debug

    def caminho_salvar_arquivo(self):
        return selecionar_caminho_para_salvar() # Abre uma janela para selecionar o caminho de salvamento do arquivo

    def extrair_historico_chamados(self, navegador, data_automacao, setor):
        navegador.switch_to.default_content()  # Volta ao conteúdo principal da página
        clicar_elemento(navegador, '//*[@id="menu_item_small_30"]/div/div', By.XPATH, automacao_fusion_instance=None) # Clica em Processos
        clicar_elemento(navegador, '#menu_item_expanded_301 > div > div.menu-inner-cell.menu-inner-cell-middle > p', By.CSS_SELECTOR, automacao_fusion_instance=None) # Clica em Consultas
        acessar_iframe(navegador, 2, automacao_fusion_instance=None, timeout=15) # Acessa o iframe
        clicar_entidade_por_nome(navegador, f"Núcleo - {setor}", automacao_fusion_instance=None) # Clica na entidade Núcleo - XXXX
        enviarkey_elemento(navegador, "input.form-control.ng-valid-date", By.CSS_SELECTOR, data_automacao, automacao_fusion_instance=None) # Define a data inicial
        time.sleep(2)  # Pequena pausa para garantir que o campo foi preenchido
        clicar_elemento(navegador, 'advancedSearchBtn', By.ID, automacao_fusion_instance=None) # Clica em Consultar

        timeout = 120 # tempo máximo para rolar e carregar itens
        start = time.time() # marca o tempo de início

        ul = WebDriverWait(navegador, 10).until(EC.presence_of_element_located((By.XPATH, "//ul[contains(@class,'groupeditemlist') and contains(@class,'pull-left')]")))  # Aguarda até o elemento UL estar presente
        li_name = "noHoverSpan" # classe dos LIs dentro do UL

        max_no_progress = 100         # número de tentativas sem crescimento até parar
        no_progress = 0         # contador de tentativas sem crescimento    
        prev_count = 0      # contador anterior de itens

        while True:  
            items = ul.find_elements(By.CLASS_NAME, li_name) # Pega todos os elementos LI dentro do UL
            cur_count = len(items) # Conta quantos itens existem atualmente

            # rolar para o último item visível (ou para o container caso não exista item)
            if items: 
                navegador.execute_script("arguments[0].scrollIntoView(false);", items[-1]) # Rola para o último item
            else:
                navegador.execute_script("arguments[0].scrollTop = arguments[0].scrollHeight;", ul) # Rola para o final do container UL
            navegador.execute_script("window.scrollBy(0, 200);") # Pequeno ajuste para garantir visibilidade
            time.sleep(0.6)  # ajuste se necessário para permitir carregamento

            items = ul.find_elements(By.CLASS_NAME, li_name) # Pega novamente todos os elementos LI dentro do UL
            new_count = len(items) # Conta quantos itens existem agora

            if new_count > cur_count: # houve progresso
                no_progress = 0 # reseta o contador de tentativas sem progresso
            else: # sem progresso
                no_progress += 1 # incrementa o contador de tentativas sem progresso

            # condição de saída: não houve progresso por várias tentativas ou timeout global
            if no_progress >= max_no_progress: # muitas tentativas sem progresso
                break # sair do loop
            if time.time() - start > timeout: # timeout global
                break # sair do loop

        total_items = len(ul.find_elements(By.CLASS_NAME, li_name)) # Conta o total de itens encontrados

        lista_historico = [] # Lista para armazenar os dados extraídos
        for i in range(total_items):
            acessar_iframe(navegador, 2, automacao_fusion_instance=None, timeout=15) # Acessa o iframe
            ul = WebDriverWait(navegador, 10).until(EC.presence_of_element_located((By.XPATH, "//ul[contains(@class,'groupeditemlist') and contains(@class,'pull-left')]"))) # Aguarda até o elemento UL estar presente
            items = ul.find_elements(By.CLASS_NAME, "noHoverSpan") # Pega todos os elementos LI dentro do UL
            if not items:
                raise Exception("Nenhum item com class 'noHoverSpan' encontrado") # erro crítico se não encontrar itens
            chamado = items[i] # Pega o i-ésimo elemento LI
            # garantir que está visível e tentar clicar
            navegador.execute_script("arguments[0].scrollIntoView({block:'center'});", chamado) # Rola para o item
            try: # tenta clicar diretamente
                chamado.click() # Clica no chamado
            except Exception:
                navegador.execute_script("arguments[0].click();", chamado) # Clica no chamado via JS se falhar
            iframe2 = WebDriverWait(navegador, 10).until(EC.presence_of_element_located((By.TAG_NAME, "iframe"))) # Aguarda até o iframe estar presente
            navegador.switch_to.frame(iframe2) # Troca para o iframe do chamado
            time.sleep(2)  # Espera para garantir que o conteúdo carregou
            numero_chamado = esperar_e_pegar_texto(navegador, "label_Codigo__", timeout=5, automacao_fusion_instance=None).replace('Código:', '') # Pega o número do chamado
            data_inicial = pegar_texto_com_quebras(navegador, "var_DadosDaSolicitacao__Responsavel__data__", timeout=1, automacao_fusion_instance=None).replace('Data.:', '') # Pega a data inicial
            responsavel = esperar_e_pegar_texto(navegador, "var_DadosDaSolicitacao__Responsavel__responsavel__", timeout=1, automacao_fusion_instance=None) # Pega o responsável
            area_setor = esperar_e_pegar_texto(navegador, "label_DadosDaSolicitacao__Responsavel__areaSetor__", timeout=1, automacao_fusion_instance=None).replace('Área / Setor:', '') # Pega a área/setor
            detalhes_solicitacao = pegar_texto_com_quebras(navegador, "var_DadosDaSolicitacao__DescricaoDaDemanda___view_textarea", timeout=1, automacao_fusion_instance=None).replace('Descrição da Demanda:function textElementChangedvar_DadosDaSolicitacao__DescricaoDaDemanda__(){var targetEl;targetEl = textareavar_DadosDaSolicitacao__DescricaoDaDemanda__;var remaining = 2000 - targetEl.value.length;if(remaining < 0){targetEl.value = targetEl.value.substring(0, 2000);remaining = 2000 - targetEl.value.length;}}', '')
            urgencia_demanda = esperar_e_pegar_texto(navegador, "label_DadosDaSolicitacao__UrgenciaDemanda__", timeout=1, automacao_fusion_instance=None).replace('Urgência da Demanda:', '')
            justificativa_demanda = pegar_texto_com_quebras(navegador, "var_DadosDaSolicitacao__JustificativaDaDefinicaoDeUrgencia___view_textarea", timeout=1, automacao_fusion_instance=None).replace('Justificativa da Definição de Urgência:function textElementChangedvar_DadosDaSolicitacao__JustificativaDaDefinicaoDeUrgencia__(){var targetEl;targetEl = textareavar_DadosDaSolicitacao__JustificativaDaDefinicaoDeUrgencia__;var remaining = 2000 - targetEl.value.length;if(remaining < 0){targetEl.value = targetEl.value.substring(0, 2000);remaining = 2000 - targetEl.value.length;}}', '')
            data_atual_supervisor = pegar_texto_com_quebras(navegador, "var_SupervisorAnalise__DataAtual__", timeout=1, automacao_fusion_instance=None).replace('Data Atual:', '')
            prazo_final = pegar_texto_com_quebras(navegador, "var_SupervisorAnalise__PrazoDeAtendimento__", timeout=1, automacao_fusion_instance=None).replace('Prazo de Atendimento:', '')
            historico_nucleo = extrair_linhas_tabela(navegador, "tblist_NucleoAdministrativoAnalise__Historico__", timeout=1)
            encaminhamento_supervisor = extrair_linhas_tabela(navegador, 'tblist_SupervisorAnalise__HistoricoDeAtendimentoNucleoAdministrativo__', timeout=1)
            if encaminhamento_supervisor == [" "]:
                encaminhamento_supervisor = extrair_linhas_tabela(navegador, 'tblist_SupervisorAnalise__HistoricoDeAtendimento__', timeout=1)
            acao_supervisor = esperar_e_pegar_texto(navegador, "label_SupervisorAnalise__Acao__", timeout=1, automacao_fusion_instance=None).replace('Ação:', '')
            responsavel_nucleo = esperar_e_pegar_texto(navegador, "var_NucleoAdministrativoAnalise__Responsavel__responsavel__", timeout=1, automacao_fusion_instance=None)
            acao_nucleo = esperar_e_pegar_texto(navegador, "label_NucleoAdministrativoAnalise__Acoes__", timeout=1, automacao_fusion_instance=None).replace('Ação:', '')
            setor_selecionado = setor  # Adiciona o setor atual
            lista_historico.append([numero_chamado,data_inicial,responsavel,area_setor,detalhes_solicitacao,urgencia_demanda,justificativa_demanda,data_atual_supervisor,prazo_final,historico_nucleo,encaminhamento_supervisor,acao_supervisor,responsavel_nucleo,acao_nucleo,setor_selecionado])
            clicar_elemento(navegador, 'task_back_btn', By.CLASS_NAME, automacao_fusion_instance=None) # Clica no botão Voltar para a lista de chamados

        salvar_lista_historico_xlsx(lista_historico, self.caminho_salvar_arquivo())