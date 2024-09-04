import streamlit as st
import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders
from streamlit_option_menu import option_menu



# Define o layout da página
st.set_page_config(
    page_title="Envio de email",
    page_icon=":email:", 
    initial_sidebar_state="expanded",
    layout="wide"
)


# Função para cadastrar o email do remetente
def cadastrar_remetente():
    st.title("Cadastro de Email do Remetente")
    col1, col2, col3 = st.columns([1,20,1])
    with col2:
        st.error("O e-mail e senha colocados não são salvo em nenhum lugar. É totalmente seguro. Ficam salvos apenas em cash para realizar os disparos. basta um F5 na pagina e já são apagados")
    # Formulário para o usuário inserir as credenciais
    with st.form(key="form_remetente"):
        email = st.text_input("Email")
        senha = st.text_input("Senha", type="password")

        # Checkbox para escolher o provedor de email
        provedor = st.radio("Selecione o provedor de email", options=["Outlook/Hotmail", "Gmail"])
        
        # Botão para submeter o formulário
        submit_button = st.form_submit_button(label="Cadastrar")
        
        if submit_button:
            # Verifica se ambos os campos estão preenchidos
            if not email or not senha:
                st.warning("Email e senha são obrigatórios.")
            else:
                # Armazena as informações na sessão
                st.session_state["email"] = email
                st.session_state["senha"] = senha
                st.session_state["provedor"] = provedor
                col1, col2, col3 = st.columns([1,0.9,1])
                with col2:
                    st.success(f"Email cadastrado com sucesso como {provedor}!")



# Função para checar se um email foi cadastrado
def usar_credenciais():
    email = st.session_state.get("email", None)
    senha = st.session_state.get("senha", None)
    provedor = st.session_state.get("provedor", None)
    
    if email and senha:
        col1, col2, col3 = st.columns([1,0.34,1])
        with col2:
            st.success("Remetente cadastrado!")
        col1, col2, col3 = st.columns([1,1,1])
        with col2:
            st.success(f"Email cadastrado: {email}")
            st.success(f"Provedor: {provedor}")
        return True
    else:
        col1, col2, col3 = st.columns([1,2,1])
        with col2:
            st.error("Nenhum email cadastrado. Por favor, cadastre o remetente na página de Cadastro de Remetente.")
        return False


# Função de envio de emails
def enviar_emails():
    st.title("Envio de Emails em Massa")  # Define o título da página do Streamlit.
    
    uploaded_file = st.file_uploader("Escolha um arquivo Excel...", type="xlsx")  
    # Permite o upload de um arquivo Excel pelo usuário.

    if uploaded_file is not None:  # Verifica se um arquivo foi carregado.
        xls = pd.ExcelFile(uploaded_file)  # Carrega o arquivo Excel.
        sheets = xls.sheet_names  # Obtém o nome de todas as abas no Excel.

        if len(sheets) > 1:  # Se houver mais de uma aba, permite que o usuário selecione uma.
            sheet_name = st.selectbox("Selecione a aba do Excel que deseja usar", sheets)
        else:
            sheet_name = sheets[0]  # Caso contrário, usa a única aba disponível.
        
        df = pd.read_excel(uploaded_file, sheet_name=sheet_name)  # Lê os dados da aba selecionada no Excel.
        st.write("Arquivo carregado com sucesso!")  # Exibe uma mensagem de sucesso.
        st.write(df)  # Mostra o conteúdo do arquivo carregado.

        col_email = st.selectbox("Selecione a coluna com os e-mails", df.columns.tolist())  
        # Permite ao usuário selecionar a coluna que contém os e-mails.

        enviar_anexos = st.checkbox("Deseja enviar anexos?", value=False)  
        # Oferece a opção de enviar anexos.

        if enviar_anexos:
            col_arquivo = st.selectbox("Selecione a coluna com os nomes dos arquivos", df.columns.tolist())  
            # Se a opção de anexos for selecionada, permite a seleção da coluna com os nomes dos arquivos.
        else:
            col_arquivo = None

        usar_cc = st.checkbox("Deseja adicionar e-mails em Cópia (CC)?")  
        # Oferece a opção de adicionar e-mails em cópia.

        if usar_cc:
            col_cc = st.selectbox("Selecione a coluna com os e-mails em Cópia", df.columns.tolist())  
            # Se a opção de CC for selecionada, permite a seleção da coluna com os e-mails em cópia.
        else:
            col_cc = None

        incluir_saudacao = st.checkbox("Deseja incluir uma saudação?", value=False)  
        # Oferece a opção de incluir uma saudação personalizada.

        if incluir_saudacao:
            col_nome = st.selectbox("Selecione a coluna com os nomes da pessoas", df.columns.tolist())  
            # Se a saudação for selecionada, permite a seleção da coluna com os nomes das pessoas.
        else:
            col_nome = None

        if col_email and (not enviar_anexos or col_arquivo):  
            # Verifica se a coluna de e-mails está selecionada e, se necessário, a coluna de arquivos.

            selecionar_todos = st.checkbox("Selecionar todos os e-mails", value=True)  
            # Oferece a opção de selecionar todos os e-mails.

            if selecionar_todos:
                selected_emails = df[col_email].dropna().unique().tolist()  
                # Seleciona todos os e-mails únicos, ignorando os valores nulos.
                st.write("Todos os e-mails serão processados.")
            else:
                unique_emails = df[col_email].dropna().unique().tolist()  
                # Obtém uma lista de e-mails únicos.
                selected_emails = st.multiselect("Selecione os e-mails que deseja processar", options=unique_emails)  
                # Permite ao usuário selecionar os e-mails que deseja processar.

            if selected_emails:  # Se pelo menos um e-mail foi selecionado.
                df_selecionado = df[df[col_email].isin(selected_emails)]  
                # Filtra o DataFrame para incluir apenas os e-mails selecionados.
                
                st.header('Configurações definidas:')  # Exibe um cabeçalho para as configurações definidas.
                st.success(f"Coluna de e-mails: {col_email}")  # Exibe a coluna de e-mails selecionada.

                if enviar_anexos:
                    st.success(f"Coluna de arquivos: {col_arquivo}")  # Exibe a coluna de arquivos selecionada.
                if col_cc:
                    st.success(f"Coluna de CC: {col_cc}")  # Exibe a coluna de CC selecionada.

                if enviar_anexos:
                    uploaded_files = st.file_uploader("Escolha os arquivos anexos", type="xlsx", accept_multiple_files=True)  
                    # Permite o upload de múltiplos arquivos anexos.

                    if uploaded_files:  # Verifica se os anexos foram carregados.
                        file_names = [file.name for file in uploaded_files]  # Obtém os nomes dos arquivos carregados.
                        expected_file_names = df_selecionado[col_arquivo].dropna().unique().tolist()  
                        # Obtém os nomes dos arquivos esperados com base na coluna selecionada.

                        if all(file_name in expected_file_names for file_name in file_names):  
                            # Verifica se todos os arquivos carregados correspondem aos esperados.
                            st.success("Todos os anexos estão corretos.")  # Exibe uma mensagem de sucesso.

                            subject = st.text_input("Título do E-mail")  # Permite ao usuário definir o título do e-mail.
                            body = st.text_area("Corpo do E-mail")  # Permite ao usuário definir o corpo do e-mail.
                            cc_emails_global = st.text_input("CC Global...", "").split(',')  
                            # Permite ao usuário adicionar e-mails em cópia globalmente.

                            if st.button("Enviar E-mails"):  # Se o botão "Enviar E-mails" for clicado.
                                for _, row in df_selecionado.iterrows():  
                                    # Itera sobre as linhas do DataFrame filtrado.
                                    email = row[col_email]
                                    nome = row[col_nome] if col_nome else ""
                                    saudacao = obter_saudacao(nome) if incluir_saudacao else ""  
                                    # Adiciona a saudação, se necessário.
                                    corpo_email = f"{saudacao}{body}"  # Cria o corpo do e-mail.
                                    file_name = row[col_arquivo] if enviar_anexos else None
                                    cc_emails_spec = [cc.strip() for cc in row[col_cc].split(',')] if col_cc else []  
                                    # Adiciona e-mails em cópia específica, se necessário.

                                    if enviar_anexos and file_name in file_names:  
                                        # Se houver anexos e o arquivo estiver na lista de nomes esperados.
                                        file = next(file for file in uploaded_files if file.name == file_name)  
                                        # Obtém o arquivo correspondente.
                                        config_email(email, file, subject, corpo_email, cc_emails_global + cc_emails_spec)  
                                        # Envia o e-mail com o anexo.
                                        st.success(f"Email enviado para {email} com o anexo {file_name[:-5]} e em CC para {', '.join(cc_emails_global + cc_emails_spec)}.")  
                                        # Exibe uma mensagem de sucesso para cada e-mail enviado.
                                    elif not enviar_anexos:  
                                        # Se não houver anexos.
                                        config_email(email, None, subject, corpo_email, cc_emails_global + cc_emails_spec)  
                                        # Envia o e-mail sem anexo.
                                        st.success(f"Email enviado para {email} sem anexo e em CC para {', '.join(cc_emails_global + cc_emails_spec)}.")  
                                        # Exibe uma mensagem de sucesso para cada e-mail enviado sem anexo.
                        else:
                            st.warning("Alguns anexos não correspondem aos nomes escolhidos...")  
                            # Exibe um aviso se os anexos não corresponderem aos nomes esperados.
                else:
                    subject = st.text_input("Título do E-mail")  # Permite ao usuário definir o título do e-mail.
                    body = st.text_area("Corpo do E-mail")  # Permite ao usuário definir o corpo do e-mail.
                    cc_emails_global = st.text_input("CC Global...", "").split(',')  
                    # Permite ao usuário adicionar e-mails em cópia globalmente.

                    if st.button("Enviar E-mails"):  # Se o botão "Enviar E-mails" for clicado.
                        for _, row in df_selecionado.iterrows():  
                            # Itera sobre as linhas do DataFrame filtrado.
                            email = row[col_email]
                            nome = row[col_nome] if col_nome else ""
                            saudacao = obter_saudacao(nome) if incluir_saudacao else ""  
                            # Adiciona a saudação, se necessário.
                            corpo_email = f"{saudacao}{body}"  # Cria o corpo do e-mail.
                            cc_emails_spec = [cc.strip() for cc in row[col_cc].split(',')] if col_cc else []  
                            # Adiciona e-mails em cópia específica, se necessário.
                            config_email(email, None, subject, corpo_email, cc_emails_global + cc_emails_spec)  
                            # Envia o e-mail sem anexo.
                            st.success(f"Email enviado para {email} sem anexo e em CC para {', '.join(cc_emails_global + cc_emails_spec)}.")  
                            # Exibe uma mensagem de sucesso para cada e-mail enviado sem anexo.
            else:
                st.error("Por favor, selecione ao menos um e-mail para processar.")  
                # Exibe um erro se nenhum e-mail foi selecionado.
        else:
            st.error("Por favor, selecione todas as colunas necessárias.")  
            # Exibe um erro se as colunas necessárias não forem selecionadas.



# Função para obter a saudação com base na hora do dia e nome do destinatário
def obter_saudacao(nome):
    from datetime import datetime

    hora_atual = datetime.now().hour
    
    if hora_atual < 12:
        saudacao = "Bom dia"
    elif 12 <= hora_atual < 18:
        saudacao = "Boa tarde"
    else:
        saudacao = "Boa noite"
    
    return f"{saudacao}, {nome}.\n\n"




# Função de ajustes de parâmetros para o envio do e-mail
def config_email(to_email, attachment, subject, body, cc_emails):
    from_email = st.session_state["email"]
    password = st.session_state["senha"]
    provedor = st.session_state.get("provedor")

    # Configurações do servidor SMTP de acordo com o provedor de email
    if provedor == "Outlook/Hotmail":
        smtp_server = "smtp.office365.com"
        smtp_port = 587
    elif provedor == "Gmail":
        smtp_server = "smtp.gmail.com"
        smtp_port = 587
    else:
        st.error("Provedor de email desconhecido.")
        return

    msg = MIMEMultipart('html')
    msg['From'] = from_email
    msg['To'] = to_email
    msg['Subject'] = subject
    if cc_emails:
        msg['Cc'] = ', '.join(cc_emails)
    
    msg.attach(MIMEText(body, 'plain'))
    
    # Se houver um anexo, adiciona-o ao email
    if attachment is not None:
        try:
            part = MIMEBase('application', 'octet-stream')
            part.set_payload(attachment.read())
            encoders.encode_base64(part)
            part.add_header('Content-Disposition', f'attachment; filename= {attachment.name}')
            msg.attach(part)
        except Exception as e:
            st.error(f"Erro ao anexar arquivo: {str(e)}")
            return
    
    try:
        with smtplib.SMTP(smtp_server, smtp_port) as server:
            server.starttls()
            server.login(from_email, password)
            server.send_message(msg)
        # st.success(f"Email enviado com sucesso para {to_email}.")
    except Exception as e:
        st.error(f"""Falha ao enviar o email: {str(e)}\n
                 VERIFIQUE SE O PROVEDOR SELECIONADO É O CORRETO!!""")


# ------------------------------------------------------ Menu de navegação ------------------------ #
cols1, cols2, cols3 = st.columns([1, 1.5, 1])
with cols2:
    selected_page = option_menu(
        menu_title=None,
        options=["Envio de E-mail", "Cadastro de Remetente"],
        icons=["bi bi-envelope-at", "gear"],
        menu_icon="cast",
        default_index=0,
        orientation="horizontal"
    )


# Lógica de seleção da página
if selected_page == "Envio de E-mail":
    # Verifica se há credenciais antes de chamar a função de envio de e-mails
    if usar_credenciais():
        enviar_emails()

elif selected_page == "Cadastro de Remetente":
        cadastrar_remetente()

