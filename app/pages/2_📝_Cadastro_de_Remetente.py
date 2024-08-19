import streamlit as st


# Define o layout da página
st.set_page_config(
    page_title="Envio de email",
    page_icon="📝", 
    initial_sidebar_state="expanded",
    layout="wide"
)



# Função para cadastrar o email do remetente
def cadastrar_remetente():
    st.title("Cadastro de Email do Remetente")
    
    # Formulário para o usuário inserir as credenciais
    with st.form(key="form_remetente"):
        email = st.text_input("Email")
        senha = st.text_input("senha", type="password")
        
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
                st.success("Email cadastrado com sucesso!")

# Chamada da função
cadastrar_remetente()


