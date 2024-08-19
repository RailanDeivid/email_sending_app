import streamlit as st
import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders

# Define o layout da página
st.set_page_config(
    page_title="Envio de email",
    page_icon="../img/mail.png", 
    initial_sidebar_state="expanded",
    layout="wide"
)

def usar_credenciais():
    email = st.session_state.get("email", None)
    senha = st.session_state.get("senha", None)
    
    if email and senha:
        st.success("Remetente cadastrado!")
        st.write(f"Email cadastrado: {email}")
        return True
    else:
        st.error("Nenhum email cadastrado. Por favor, cadastre o remetente na página de Cadastro de Remetente.")
        return False

def enviar_emails():
    st.title("Envio de Emails em Massa")

    uploaded_file = st.file_uploader("Escolha um arquivo Excel contendo os e-mails e nomes dos anexos", type="xlsx")
    
    if uploaded_file is not None:
        df = pd.read_excel(uploaded_file)
        st.write("Arquivo carregado com sucesso!")
        st.write(df.head())
        
        col1, col2 = st.columns(2)
        with col1:
            col_email = st.selectbox("Selecione a coluna com os e-mails", df.columns.tolist())
        with col2:
            col_arquivo = st.selectbox("Selecione a coluna com os nomes dos arquivos", df.columns.tolist())
        
        # Checkbox para usar a coluna de CC
        usar_cc = st.checkbox("Deseja adicionar e-mails em Cópia (CC)?")

        if usar_cc:
            col_cc = st.selectbox("Selecione a coluna com os e-mails em Cópia", df.columns.tolist())
        else:
            col_cc = None
        
        if col_email and col_arquivo:
            selecionar_todos = st.checkbox("Selecionar todos os e-mails", value=True)

            if selecionar_todos:
                selected_emails = df[col_email].dropna().unique().tolist()
                st.write("Todos os e-mails serão processados.")
            else:
                unique_emails = df[col_email].dropna().unique().tolist()
                selected_emails = st.multiselect("Selecione os e-mails que deseja processar", options=unique_emails)
            
            if selected_emails:
                df_selecionado = df[df[col_email].isin(selected_emails)]
                
                # Monta a mensagem de sucesso com as configurações definidas
                configuracoes = f"""
                    \nColuna de e-mails: {col_email}
                    \nColuna de arquivos: {col_arquivo}
                    \nE-mails selecionados: {", ".join(selected_emails)}
                """
                
                # Se a coluna de CC foi selecionada, adicionar ao resumo
                if col_cc:
                    configuracoes += f"\nColuna de CC: {col_cc}"
                
                st.header('Configurações definidas:')
                st.success(configuracoes)
                
                st.warning("Agora, faça o upload dos arquivos anexos. Certifique-se de que os arquivos tenham os mesmos nomes listados na coluna selecionada.")
                
                uploaded_files = st.file_uploader("Escolha os arquivos anexos", type="xlsx", accept_multiple_files=True)
                
                if uploaded_files:
                    file_names = [file.name for file in uploaded_files]
                    st.write("Arquivos carregados:")
                    st.write(file_names)

                    # Adiciona a extensão .xlsx aos nomes dos arquivos extraídos da coluna selecionada
                    expected_file_names = [f"{name}.xlsx" for name in df_selecionado[col_arquivo].dropna().unique().tolist()]
                    
                    if all(file_name in expected_file_names for file_name in [name.split('.')[0] for name in file_names]):
                        st.success("Todos os anexos estão corretos.")

                        subject = st.text_input("Título do E-mail")
                        body = st.text_area("Corpo do E-mail")
                        cc_emails_global = st.text_input("CC Global (Separados por vírgula)", "").split(',')

                        if st.button("Enviar E-mails"):
                            for _, row in df_selecionado.iterrows():
                                email = row[col_email]
                                file_name = f"{row[col_arquivo]}.xlsx"  # Adiciona a extensão .xlsx ao nome do arquivo
                                cc_emails_spec = [cc.strip() for cc in row[col_cc].split(',')] if col_cc and pd.notna(row[col_cc]) else []
                                
                                if file_name in file_names:
                                    file = next(file for file in uploaded_files if file.name == file_name)
                                    send_email(email, file, subject, body, cc_emails_global + cc_emails_spec)
                                    st.success(f"Email enviado para {email} com o anexo {file_name} e em CC para {', '.join(cc_emails_global + cc_emails_spec)}.")
                    else:
                        st.warning("Alguns anexos não correspondem aos nomes escolhidos como (nomes dos arquivos). Verifique se os arquivos estão corretos.")
            else:
                st.error("Por favor, selecione ao menos um e-mail para processar.")
        else:
            st.error("Por favor, selecione todas as colunas necessárias.")

def send_email(to_email, attachment, subject, body, cc_emails):
    from_email = st.session_state["email"]
    password = st.session_state["senha"]
    
    smtp_server = "smtp.office365.com"
    smtp_port = 587
    
    msg = MIMEMultipart()
    msg['From'] = from_email
    msg['To'] = to_email
    msg['Subject'] = subject
    if cc_emails:
        msg['Cc'] = ', '.join(cc_emails)
    
    msg.attach(MIMEText(body, 'plain'))
    
    part = MIMEBase('application', 'octet-stream')
    part.set_payload(attachment.read())
    encoders.encode_base64(part)
    part.add_header('Content-Disposition', f'attachment; filename= {attachment.name}')
    msg.attach(part)
    
    with smtplib.SMTP(smtp_server, smtp_port) as server:
        server.starttls()
        server.login(from_email, password)
        server.send_message(msg)

# Verifica se há credenciais antes de chamar a função de envio de e-mails
if usar_credenciais():
    enviar_emails()
