# 📊 Slides Semanais — Banco Bari

Interface web para geração automática dos slides de diretoria.
**Nenhuma instalação necessária** — roda 100% no navegador.

---

## 🚀 Como colocar no ar (uma vez só, ~10 minutos)

### Passo 1: Criar conta no GitHub
1. Acesse [github.com](https://github.com) e crie uma conta (grátis)
2. Clique em **"New repository"** (botão verde)
3. Nome: `slides-bari` (ou o que preferir)
4. Marque **Private** (para ficar só para vocês)
5. Clique em **Create repository**

### Passo 2: Subir os arquivos
1. Na página do repositório, clique em **"uploading an existing file"**
2. Arraste todos os arquivos desta pasta:
   - `app.py`
   - `requirements.txt`
   - `.streamlit/config.toml` (crie a pasta `.streamlit` primeiro)
3. Clique em **Commit changes**

> **Dica:** Para a pasta `.streamlit`, você pode clicar em "Add file" → "Create new file" e digitar `.streamlit/config.toml` como nome do arquivo, colando o conteúdo do config.toml.

### Passo 3: Conectar ao Streamlit Cloud
1. Acesse [share.streamlit.io](https://share.streamlit.io)
2. Faça login com sua conta do GitHub
3. Clique em **"New app"**
4. Selecione:
   - **Repository:** `slides-bari`
   - **Branch:** `main`
   - **Main file path:** `app.py`
5. Clique em **"Deploy!"**
6. Aguarde ~2 minutos para o deploy

### Passo 4: Pronto!
Você vai receber um link tipo: `https://slides-bari.streamlit.app`

Compartilhe esse link com o time. Qualquer pessoa acessa pelo navegador.

---

## 📋 Uso semanal

1. Abra o link no navegador
2. Carregue as bases:
   - **Base do Funil** → `Atualizar Entrada nas Fases.xlsx`
   - **Base Dashboard (Opps)** → `Atualizar Entrada nas Fases Dash.xlsx`
   - **Base Dashboard (Leads)** → `Entradas nas Fases Leads Dash.xlsx`
   - **Apresentação Modelo** → o `.pptx` atual
3. O Planejamento é opcional (só precisa atualizar se mudou)
4. Confira as datas (calcula automaticamente)
5. Clique em **🚀 Gerar Apresentação**
6. Baixe o `.pptx` e suba pro Google Slides

---

## 🔒 Privacidade

- O repositório é **privado** — só quem você convidar vê o código
- Os dados (planilhas) **não ficam salvos** no servidor — são processados em memória e descartados
- O Streamlit Cloud é gratuito para repositórios privados

---

## ❓ Problemas comuns

**"App is sleeping"**
→ Apps gratuitos dormem após inatividade. Basta recarregar a página e esperar ~30 segundos.

**Erro ao processar**
→ Verifique se as planilhas estão no formato correto (mesmo que usa no script original).

**Quero mudar algo no código**
→ Edite o `app.py` no GitHub. O Streamlit Cloud atualiza automaticamente em ~1 minuto.

---

## 📁 Estrutura do repositório

```
slides-bari/
├── app.py                  ← Código principal
├── requirements.txt        ← Dependências Python
├── README.md               ← Este arquivo
└── .streamlit/
    └── config.toml         ← Tema e configurações
```
