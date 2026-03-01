# 📊 Automação de Relatórios e Envio de OnePages

## 📌 Contexto do Projeto

Neste projeto, atuo como integrante da equipe de Análise de Dados de uma empresa fictícia com 25 lojas.

Diariamente, é necessário:

- Gerar um relatório individual (OnePage) para cada loja
- Enviar o OnePage para o respectivo gerente
- Gerar um relatório consolidado para a diretoria
- Enviar automaticamente esse relatório para o e-mail da diretoria

Todo esse processo foi automatizado utilizando Python, eliminando tarefas manuais e reduzindo riscos operacionais.

---

## 🚀 Objetivo

Automatizar completamente a geração e o envio de relatórios diários, garantindo:

- Agilidade no envio de informações
- Padronização dos relatórios
- Redução de erros humanos
- Organização modular do código

---

## ⚙️ Funcionalidades

- 📈 Geração automática de OnePage individual por loja
- 📊 Geração de relatório geral consolidado da diretoria
- 📧 Envio automático de e-mails personalizados para cada gerente
- 🧾 Envio automático do relatório consolidado para a diretoria
- 🧩 Código organizado em funções reutilizáveis

---

## 🛠 Tecnologias Utilizadas

- Python
- Pandas
- SMTP (Gmail)
- HTML
- Manipulação de arquivos Excel e CSV

---

## ⏰ Execução Automática

O projeto está configurado para execução diária utilizando o Agendador de Tarefas do Windows (Windows Task Scheduler).

Dessa forma, os relatórios são gerados e enviados automaticamente todos os dias, sem necessidade de execução manual.

---

## 📂 Estrutura do Projeto

```
📁 Bases de Dados/
│
├── Vendas.xlsx
├── Lojas.csv
├── Emails.csv
│
📄 main.py
📄 gerar_onepage.py
📄 enviar_email.py
📄 requirements.txt
📄 README.md
```

---

## ▶️ Como Executar o Projeto

### 1️⃣ Clonar o repositório

```
git clone <url-do-repositorio>
```

Entrar na pasta:

```
cd nome-do-repositorio
```

---

### 2️⃣ Criar ambiente virtual

```
python -m venv venv
```

---

### 3️⃣ Ativar o ambiente

Windows:

```
venv\Scripts\activate
```

---

### 4️⃣ Instalar dependências

```
pip install -r requirements.txt
```

---

### 5️⃣ Configurar envio de e-mail

O projeto utiliza SMTP do Gmail.

Para funcionar corretamente:

- Crie uma **senha de aplicativo** no Gmail
- Configure no código suas credenciais
- Priorize usar variáveis de ambiente para armazenar seu email e senha
- Não exponha credenciais reais em repositórios públicos

---

### 6️⃣ Executar o projeto

```
python main.py
```

Os relatórios serão gerados automaticamente e os e-mails enviados conforme configurado.

---

## ⚠️ Observações Importantes

- Para testar o projeto, recomenda-se alterar os e-mails presentes na base `Emails.csv` para e-mails próprios.
- Este projeto possui e-mails configurados apenas para fins de demonstração.
- Não é recomendado expor credenciais reais publicamente.

---

## 📈 Possíveis Melhorias Futuras

- Implementação de logs de envio
- Tratamento mais robusto de erros
- Deploy em ambiente cloud
- Interface gráfica para configuração

---

## 🎯 Conclusão

Este projeto demonstra:

- Automação de processos corporativos
- Organização modular e escalável de código
- Manipulação de dados com Pandas
- Integração com serviços externos (SMTP)
- Aplicação prática de Python no contexto empresarial

Projeto desenvolvido para fins de estudo, prática e portfólio.
