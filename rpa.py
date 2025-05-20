import time
import re
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException
from webdriver_manager.chrome import ChromeDriverManager
import openpyxl # Para criar a planilha Excel

# --- Configurações ---
# Deixe em branco para o script perguntar, ou defina aqui:
GROUP_NAME = ""  # Ex: "Amigos do Futebol"
EXCEL_FILE_NAME = "contatos_whatsapp_grupo.xlsx"
TIMEOUT_SECONDS = 30  # Tempo máximo de espera para elementos carregarem

# --- Funções Auxiliares ---
def setup_driver():
    """Configura e retorna o WebDriver do Chrome."""
    chrome_options = Options()
    # Remove a barra "O Chrome está sendo controlado por software de teste automatizado"
    chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
    chrome_options.add_experimental_option('useAutomationExtension', False)
    # Inicia maximizado
    chrome_options.add_argument("--start-maximized")
    # Opcional: Para tentar manter a sessão do WhatsApp (requer configuração manual do profile_path)
    # profile_path = 'SEU/CAMINHO/PARA/PERFIL/DO/CHROME' # Ex: C:\\Users\\SeuUsuario\\AppData\\Local\\Google\\Chrome\\User Data\\ProfileWhatsappAutomacao
    # chrome_options.add_argument(f"user-data-dir={profile_path}")

    try:
        # Usa webdriver-manager para baixar e configurar o ChromeDriver automaticamente
        service = Service(ChromeDriverManager().install())
        driver = webdriver.Chrome(service=service, options=chrome_options)
    except Exception as e_webdriver_manager:
        print(f"Erro ao configurar o ChromeDriver com webdriver-manager: {e_webdriver_manager}")
        print("Tentando especificar o caminho do chromedriver manualmente se o erro persistir.")
        print("Certifique-se de ter o chromedriver.exe no PATH ou especifique o caminho abaixo.")
        # Exemplo de como especificar o caminho manualmente (descomente e ajuste se necessário):
        # Lembre-se de baixar o chromedriver compatível com sua versão do Chrome.
        # CHROMEDRIVER_PATH = "C:/caminho/para/seu/chromedriver.exe"
        # service = Service(executable_path=CHROMEDRIVER_PATH)
        # driver = webdriver.Chrome(service=service, options=chrome_options)
        raise  # Re-levanta a exceção se não houver fallback manual configurado
    return driver

def filter_bmp_chars(text):
    """Filtra a string para manter apenas caracteres do Basic Multilingual Plane (BMP)."""
    if not text:
        return ""
    return "".join(c for c in text if ord(c) <= 0xFFFF)

def clean_phone_number(text_input):
    """
    Limpa e formata o número de telefone para o formato internacional (sem '+', apenas dígitos).
    Tenta extrair de formatos como '1234567890@c.us' ou texto com caracteres especiais.
    """
    if not text_input:
        return None

    # Tenta extrair número de um JID (ex: "5511999998888@c.us")
    match_jid = re.search(r"(\d+)@c\.us", text_input)
    if match_jid:
        return match_jid.group(1)

    # Remove todos os caracteres não numéricos
    # Isso removerá '+', '-', '(', ')', espaços, etc.
    # Para wa.me, queremos apenas os dígitos, começando com o código do país.
    # Ex: "5511999998888"
    cleaned_number = re.sub(r"\D", "", text_input)

    return cleaned_number if cleaned_number else None


def extract_contacts_from_group(driver, group_name_to_find):
    """Localiza o grupo, abre informações e extrai contatos."""
    contacts_data = []
    
    # Filtra caracteres não-BMP (como emojis) do nome do grupo para a busca
    group_name_for_search = filter_bmp_chars(group_name_to_find)
    
    if not group_name_for_search.strip():
        print(f"Erro: O nome do grupo '{group_name_to_find}' resultou em uma string vazia após filtrar caracteres especiais/emojis.")
        print("Por favor, forneça um nome de grupo com caracteres de texto válidos para a busca.")
        return []

    if group_name_for_search != group_name_to_find:
        print(f"Aviso: O nome do grupo original '{group_name_to_find}' contém caracteres especiais/emojis.")
        print(f"Procurando pelo grupo usando o nome filtrado: '{group_name_for_search}'")
    else:
        print(f"Procurando pelo grupo: '{group_name_for_search}'")

    try:
        # 1. Encontrar o campo de busca de conversas e digitar o nome do grupo
        search_box_xpath = "//div[@contenteditable='true'][@data-tab='3']"
        try:
            search_box = WebDriverWait(driver, TIMEOUT_SECONDS).until(
                EC.presence_of_element_located((By.XPATH, search_box_xpath))
            )
        except TimeoutException:
            search_box_xpath_alt = "//div[@role='textbox'][@title='Caixa de texto de pesquisa']"
            search_box = WebDriverWait(driver, TIMEOUT_SECONDS).until(
                EC.presence_of_element_located((By.XPATH, search_box_xpath_alt))
            )

        search_box.clear()
        search_box.send_keys(group_name_for_search) # Usa o nome filtrado
        print(f"Digitado '{group_name_for_search}' na busca.")
        time.sleep(3) 

        # 2. Clicar no grupo nos resultados da busca
        # O seletor para o item do grupo precisa corresponder ao nome que aparece na UI,
        # que pode ser o nome original com emojis. Se a busca funcionar com o nome filtrado,
        # o WhatsApp pode ainda exibir o nome completo no resultado.
        # Tentaremos clicar usando o nome filtrado, mas se o WhatsApp exibir o nome completo,
        # pode ser necessário ajustar este seletor ou a estratégia.
        # A forma mais segura é que `group_name_for_search` seja suficiente para o WhatsApp
        # listar o grupo desejado como primeiro ou único resultado clicável.
        
        # Tentativa de encontrar o grupo pelo nome filtrado (se o WhatsApp o exibir assim na busca)
        # ou pelo nome original (se a busca pelo nome filtrado ainda resultar no nome original sendo exibido)
        # Esta parte é delicada. O ideal é que `group_name_for_search` seja o que aparece no `title` do resultado.
        
        # Primeiro, tenta clicar no grupo usando o nome filtrado (que foi usado na busca)
        group_chat_xpath_filtered = f"//span[@dir='auto'][@title='{group_name_for_search}']"
        try:
            group_chat_element = WebDriverWait(driver, TIMEOUT_SECONDS).until(
                EC.element_to_be_clickable((By.XPATH, group_chat_xpath_filtered))
            )
            print(f"Tentando clicar no grupo com title='{group_name_for_search}'")
        except TimeoutException:
            print(f"Não foi possível encontrar o grupo com title='{group_name_for_search}'.")
            # Se o nome original (com emojis) for curto o suficiente para o atributo title,
            # e a busca pelo nome filtrado o retornar, tentamos o nome original.
            # No entanto, o atributo title também pode ter problemas com não-BMP.
            # A melhor aposta é que o `group_name_for_search` seja o que o WhatsApp usa para o `title`
            # ou que seja o texto principal visível.
            # Se o nome do grupo na lista de chats (após a busca) for o original com emojis,
            # e o title também, esta parte pode falhar.
            # O mais seguro é que o `group_name_for_search` seja suficiente para o WhatsApp exibir o grupo
            # e o XPATH abaixo seja genérico o suficiente para pegar o primeiro resultado relevante.
            # Ex: (//span[@dir='auto'][contains(@title,'TEXTO_DO_GRUPO')])[1]
            # Por simplicidade, vamos assumir que o title corresponderá ao `group_name_for_search`
            # ou que o usuário fornecerá um `group_name_for_search` que funcione.
            
            # Se o nome filtrado não funcionar para o clique, e o nome original for diferente,
            # tentamos com o nome original, mas isso pode reintroduzir o problema se o `title` tiver emojis.
            # A melhor abordagem aqui é que o XPATH para clicar no grupo seja robusto.
            # Vamos tentar um XPATH que procure pelo texto visível se o title falhar.
            print(f"Tentando encontrar o grupo pelo texto visível: '{group_name_for_search}'")
            group_chat_xpath_visible_text = f"//span[@dir='auto'][normalize-space()='{group_name_for_search}']"
            try:
                group_chat_element = WebDriverWait(driver, TIMEOUT_SECONDS).until(
                    EC.element_to_be_clickable((By.XPATH, group_chat_xpath_visible_text))
                )
            except TimeoutException:
                 # Se o nome do grupo original for diferente e a busca pelo nome filtrado o trouxe,
                 # pode ser que o elemento clicável ainda tenha o nome original.
                 if group_name_for_search != group_name_to_find:
                    print(f"Tentando encontrar o grupo pelo texto visível original: '{group_name_to_find}' (pode falhar com emojis)")
                    group_chat_xpath_original_text = f"//span[@dir='auto'][normalize-space()='{group_name_to_find}']"
                    group_chat_element = WebDriverWait(driver, TIMEOUT_SECONDS).until(
                        EC.element_to_be_clickable((By.XPATH, group_chat_xpath_original_text))
                    )
                 else:
                    raise # Re-levanta a exceção se todas as tentativas falharem

        group_chat_element.click()
        print(f"Grupo selecionado (usando critério de busca/clique com base em '{group_name_for_search}' ou '{group_name_to_find}').")
        time.sleep(2)

        # 3. Clicar no cabeçalho do grupo para abrir as informações do grupo
        # O cabeçalho DEVE corresponder ao nome exibido na conversa, que é o nome original.
        group_header_xpath = f"//header//span[@dir='auto'][@title='{group_name_to_find}']"
        try:
            group_header = WebDriverWait(driver, TIMEOUT_SECONDS).until(
                EC.element_to_be_clickable((By.XPATH, group_header_xpath))
            )
        except TimeoutException:
            print("Não foi possível encontrar o cabeçalho do grupo pelo nome exato (com emojis). Tentando seletor alternativo para o header do chat ativo...")
            alternative_header_xpath = "(//div[@id='main']//header)[1]"
            group_header = WebDriverWait(driver, TIMEOUT_SECONDS).until(
                EC.element_to_be_clickable((By.XPATH, alternative_header_xpath))
            )
        group_header.click()
        print("Informações do grupo abertas.")
        time.sleep(3) 

        # 4. Encontrar a lista de participantes e extrair
        group_info_panel_xpath = "//div[@data-testid='chat-info-drawer']//section"
        try:
            group_info_panel = WebDriverWait(driver, TIMEOUT_SECONDS).until(
                EC.presence_of_element_located((By.XPATH, group_info_panel_xpath))
            )
        except TimeoutException:
            print(f"Painel de informações do grupo não encontrado com o seletor principal. Verifique o XPath: {group_info_panel_xpath}")
            return [] 

        print("Painel de informações do grupo encontrado. Tentando extrair participantes...")
        
        last_height = driver.execute_script("return arguments[0].scrollHeight", group_info_panel)
        processed_titles = set() 

        while True:
            participant_elements_xpath = "//div[@data-testid='chat-info-drawer']//div[@role='listitem']"
            time.sleep(1)
            participants_elements = group_info_panel.find_elements(By.XPATH, participant_elements_xpath)
            # print(f"Encontrados {len(participants_elements)} elementos de participantes visíveis nesta rolagem...")

            if not participants_elements and not contacts_data: 
                print("Nenhum elemento de participante encontrado com o seletor. Verifique o XPath.")
                break

            new_contacts_found_this_scroll = False
            for p_element in participants_elements:
                name_or_number_text = ""
                phone_attribute_title = "" 
                
                try:
                    try:
                        name_span = p_element.find_element(By.XPATH, ".//span[@aria-label]")
                        name_or_number_text = name_span.get_attribute("aria-label").strip()
                    except NoSuchElementException:
                        name_span = p_element.find_element(By.XPATH, ".//span[@dir='auto'][1]")
                        name_or_number_text = name_span.text.strip()
                    try:
                        title_span = p_element.find_element(By.XPATH, ".//span[@title]")
                        phone_attribute_title = title_span.get_attribute("title").strip()
                    except NoSuchElementException:
                        phone_attribute_title = name_or_number_text 
                except NoSuchElementException:
                    continue
                except Exception as e_inner:
                    print(f"  Erro ao processar um participante: {e_inner}")
                    continue

                if "Você" in name_or_number_text or "You" in name_or_number_text:
                    pass 
                    continue

                if phone_attribute_title in processed_titles:
                    continue
                processed_titles.add(phone_attribute_title)
                new_contacts_found_this_scroll = True

                cleaned_phone = clean_phone_number(phone_attribute_title)
                if not cleaned_phone: 
                    cleaned_phone = clean_phone_number(name_or_number_text)

                display_name = name_or_number_text
                if display_name == cleaned_phone and cleaned_phone:
                    display_name = f"Contato ({cleaned_phone})"
                elif not display_name and cleaned_phone: 
                    display_name = f"Contato ({cleaned_phone})"
                elif not display_name and not cleaned_phone:
                    display_name = "Nome/Número Desconhecido"

                contact_entry = {"name": display_name, "phone": cleaned_phone}
                contacts_data.append(contact_entry)
                # print(f"  Extraído: Nome='{display_name}', Tel Limpo='{cleaned_phone}' (Origem Tel: '{phone_attribute_title}')")

            driver.execute_script("arguments[0].scrollTop = arguments[0].scrollHeight", group_info_panel)
            time.sleep(2.5) 

            new_height = driver.execute_script("return arguments[0].scrollHeight", group_info_panel)
            if new_height == last_height and not new_contacts_found_this_scroll : 
                print("Fim da lista de participantes alcançado (altura não mudou e sem novos contatos).")
                break
            if new_height == last_height and new_contacts_found_this_scroll:
                # print("Altura não mudou, mas novos contatos foram encontrados. Tentando mais uma rolagem para garantir.")
                last_height = new_height 
                time.sleep(1) 
                current_scroll_top = driver.execute_script("return arguments[0].scrollTop", group_info_panel)
                if group_info_panel.size['height'] > 0 : # Evita divisão por zero se o painel não tiver altura
                    max_scroll_top = new_height - group_info_panel.size['height']
                    if current_scroll_top >= max_scroll_top - 5: 
                        # print("Já está no final do scroll.")
                        break
                else: # Se não puder determinar a altura do cliente, assume que terminou se a altura do scroll não mudar
                    # print("Não foi possível determinar a altura do painel cliente, assumindo fim do scroll.")
                    break


            last_height = new_height
            # print("Rolando para carregar mais participantes...")

    except TimeoutException:
        print(f"Elemento não encontrado ou tempo esgotado ao tentar interagir com o grupo (usando nome para busca: '{group_name_for_search}').")
        print("Verifique os seletores XPath e se o nome do grupo filtrado ainda permite encontrá-lo.")
    except Exception as e:
        print(f"Ocorreu um erro inesperado durante a extração: {e}")
        import traceback
        traceback.print_exc()

    final_contacts = []
    seen_contacts = set()
    for contact in contacts_data:
        identifier = (contact.get("name", ""), contact.get("phone", ""))
        if identifier not in seen_contacts:
            final_contacts.append(contact)
            seen_contacts.add(identifier)
    print(f"Total de contatos únicos extraídos: {len(final_contacts)}")
    return final_contacts


def save_to_excel(contacts, filename):
    """Salva os contatos em uma planilha Excel."""
    if not contacts:
        print("Nenhum contato para salvar na planilha.")
        return

    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.title = "Contatos WhatsApp"

    sheet["A1"] = "Nome Registrado / Exibido"
    sheet["B1"] = "Número de Telefone (Limpo)"
    sheet["C1"] = "Link Direto (wa.me)"

    row_num = 2
    for contact in contacts:
        sheet[f"A{row_num}"] = contact["name"]
        phone_number = contact["phone"]
        sheet[f"B{row_num}"] = phone_number

        if phone_number:
            wa_link = f"https://wa.me/{phone_number}"
            sheet[f"C{row_num}"] = wa_link
            sheet[f"C{row_num}"].hyperlink = wa_link
            sheet[f"C{row_num}"].style = "Hyperlink"
        else:
            sheet[f"C{row_num}"] = "Número não disponível"
        row_num += 1

    try:
        workbook.save(filename)
        print(f"Dados dos contatos salvos com sucesso em '{filename}'")
    except Exception as e:
        print(f"Erro ao salvar o arquivo Excel: {e}")
        print(f"Verifique se o arquivo '{filename}' não está aberto em outro programa.")

if __name__ == "__main__":
    driver = None
    try:
        if not GROUP_NAME:
            GROUP_NAME_INPUT = input("Digite o nome EXATO do grupo do WhatsApp que você quer processar: ")
            if not GROUP_NAME_INPUT:
                print("Nome do grupo não fornecido. Saindo.")
                exit()
        else:
            GROUP_NAME_INPUT = GROUP_NAME

        driver = setup_driver()
        driver.get("https://web.whatsapp.com/")
        
        print("-" * 50)
        print("ATENÇÃO: Por favor, escaneie o QR Code no navegador para fazer login no WhatsApp Web.")
        print("Após o WhatsApp Web carregar COMPLETAMENTE, volte a esta janela do terminal.")
        input("Pressione Enter AQUI no terminal após o login e carregamento completo do WhatsApp Web...")
        print("-" * 50)

        print("Aguardando alguns segundos para a interface do WhatsApp Web estabilizar...")
        try:
            WebDriverWait(driver, 60).until(
                EC.presence_of_element_located((By.XPATH, "//div[@contenteditable='true'][@data-tab='3'] | //div[@role='textbox'][@title='Caixa de texto de pesquisa']"))
            )
            print("Interface principal do WhatsApp Web parece carregada.")
        except TimeoutException:
            print("A interface principal do WhatsApp não carregou a tempo. O script pode falhar.")
        
        time.sleep(5) 

        extracted_contacts = extract_contacts_from_group(driver, GROUP_NAME_INPUT)

        if extracted_contacts:
            print(f"\nForam extraídos {len(extracted_contacts)} contatos únicos do grupo '{GROUP_NAME_INPUT}'.")
            save_to_excel(extracted_contacts, EXCEL_FILE_NAME)
        else:
            print(f"Nenhum contato foi extraído do grupo '{GROUP_NAME_INPUT}'. Verifique o nome do grupo e os seletores no script se necessário.")

    except Exception as e_main:
        print(f"Ocorreu um erro geral no script: {e_main}")
        import traceback
        traceback.print_exc()
    finally:
        if driver:
            print("\nScript finalizado. O navegador será fechado em 15 segundos...")
            time.sleep(15)
            driver.quit()
