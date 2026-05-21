# 🌾 Dashboard Manutenção — Colhedoras

Dashboard Streamlit para acompanhamento de **orçamento x gasto de manutenção** por centro de custo.  
Os dados são lidos diretamente do SharePoint via Microsoft Graph API.

---

## 📋 Funcionalidades

- Orçamento gerado automaticamente: `Toneladas × R$ 2,60`
- KPIs de realizado: toneladas, orçamento, gasto, saldo
- Projeção de safra até novembro com base na taxa atual
- Gráfico Orçamento x Gasto por CC
- Evolução acumulada no tempo
- Tabela com semáforo 🟢🟡🔴 por CC
- Filtros por período e centro de custo

---

## 📁 Estrutura do Projeto

```
dashboard-manutencao/
├── dashboard_manutencao.py        ← app principal
├── requirements.txt               ← dependências Python
├── .gitignore                     ← protege secrets e arquivos desnecessários
├── README.md                      ← este arquivo
└── .streamlit/
    ├── config.toml                ← tema visual (sobe para o GitHub)
    └── secrets.toml.example       ← modelo dos secrets (sobe para o GitHub)
```

> ⚠️ O arquivo `.streamlit/secrets.toml` com as credenciais reais **nunca sobe para o GitHub**.  
> As credenciais são configuradas diretamente no painel do Streamlit Cloud.

---

## ⚙️ Deploy no Streamlit Cloud

### 1. Suba o projeto para o GitHub (repositório privado)

### 2. Acesse [share.streamlit.io](https://share.streamlit.io)
- New app → escolha o repositório
- Main file: `dashboard_manutencao.py`
- Clique em **Advanced settings → Secrets**

### 3. Cole os secrets (com o client_secret real):
```toml
[azure]
client_id     = "32fd9763-64cf-4620-8159-9811f56d8617"
client_secret = "SEU_CLIENT_SECRET_REAL"
tenant_id     = "6487fea7-a142-4721-8aad-35508e51c639"
```

### 4. Clique em **Deploy!**

---

## 💻 Rodar localmente

```bash
# 1. Crie o ambiente virtual
python -m venv venv
venv\Scripts\activate

# 2. Instale as dependências
pip install -r requirements.txt

# 3. Crie o arquivo de secrets local
copy .streamlit\secrets.toml.example .streamlit\secrets.toml
# edite o secrets.toml e cole o client_secret real

# 4. Rode
streamlit run dashboard_manutencao.py
```

---

## 📊 Fonte dos Dados

| Item | Valor |
|---|---|
| Arquivo | `manutencao_powerbi.xlsx` |
| Aba Manutenção | `fManutencao` |
| Aba Toneladas | `COLHEDORAS` |
| Taxa orçamento | R$ 2,60 / tonelada |
| Atualização | a cada 5 minutos |
