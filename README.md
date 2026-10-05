# 📞 CRM Call Center - Migração de Clientes (Cloud App)

Aplicativo web interativo desenvolvido em Python (Streamlit) para gestão de carteira e atendimento telefônico/WhatsApp de consultores.

## 🚀 Recursos
- **Visualização Única:** Exibe **1 cliente por vez** com progresso da carteira (ex: #3 de 57).
- **Cartão de Informações:** CNPJ, Razão Social, Contato, Telefone (com link direto para WhatsApp), Viabilidade GPON / Banda Larga, Plano Atual e Recomendação de Aparelhos.
- **Trava de Atendimento:** O botão **"Gravar & Avançar para Próximo Cliente"** só é liberado após o consultor selecionar o resultado da chamada e escrever uma observação obrigatória.
- **Botões Rápidos:** "Atendeu Chamada", "Não Atendeu", "Agendar Retorno" e "Sem Interesse".
- **Persistência de Dados:** Salva em tempo real em um banco SQLite interno.
- **Exportação:** Botão na barra lateral para baixar o relatório final consolidado em Excel (`.xlsx`).

---

## 🛠️ Como Executar Localmente

1. Navegue até a pasta do projeto:
   ```bash
   cd C:\Users\VIVO\.gemini\antigravity\scratch\app_migra_celio
   ```

2. Instale as dependências:
   ```bash
   pip install -r requirements.txt
   ```

3. Execute a aplicação:
   ```bash
   streamlit run app.py
   ```

4. Acesse no navegador em `http://localhost:8501`.

---

## ☁️ Como Fazer Deploy na Nuvem (Gratuito)

### Opção 1: Streamlit Community Cloud (Recomendado & Gratuito)
1. Suba esta pasta para um repositório no **GitHub**.
2. Acesse [share.streamlit.io](https://share.streamlit.io).
3. Conecte sua conta GitHub e selecione o repositório.
4. Escolha `app.py` como arquivo principal e clique em **Deploy**.

### Opção 2: Render ou Railway (via Docker)
1. Crie uma conta no [Render.com](https://render.com) ou [Railway.app](https://railway.app).
2. Conecte o repositório Git.
3. Escolha o ambiente **Docker**. O sistema utilizará o `Dockerfile` automaticamente.
