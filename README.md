Aqui está o modelo de README adaptado para o seu código:

---

# 📝 Projeto: Automação de Extração de Dados de Boletos

### 📋 Descrição do Projeto

Este projeto implementa uma automação para extrair dados de boletos eletrônicos recebidos via e-mail. O sistema coleta os arquivos em PDF anexados a mensagens de e-mail, processa as informações relevantes como data de vencimento, valor e código de barras, e armazena esses dados em uma planilha Excel. O objetivo é automatizar o processo de leitura de boletos de diversas instituições (Unimed, SEMAE, Nubank, XP, CPFL) e garantir que os dados fiquem acessíveis de forma estruturada.

---

### 🛠️ Tecnologias Utilizadas

- **Python**: Linguagem principal do projeto.
- **openpyxl**: Biblioteca para manipulação de planilhas Excel.
- **pdfplumber**: Biblioteca para extração de conteúdo de arquivos PDF.
- **re (Regular Expressions)**: Utilizada para extração de padrões de texto nos boletos.
- **Imbox**: Biblioteca para conexão e leitura de e-mails no servidor IMAP.
- **datetime**: Biblioteca padrão para manipulação de datas.
- **json**: Utilizada para gerenciamento das credenciais e senhas de PDFs.
- **pdfminer**: Para manipulação avançada de arquivos PDF protegidos por senha.

---

### ✨ Funcionalidades Principais

1. **Conexão com o e-mail**: Conecta automaticamente à caixa de entrada usando credenciais armazenadas em um arquivo JSON.
2. **Leitura de anexos PDF**: Faz o download e processa arquivos PDF anexados a e-mails de fontes pré-determinadas.
3. **Extração de dados**: Extrai informações como data de vencimento, valor da fatura e código de barras dos boletos.
4. **Manipulação de planilhas Excel**: Insere os dados extraídos em uma planilha Excel, salvando o histórico das extrações.
5. **Gerenciamento de senhas de PDF**: Lê arquivos PDF protegidos por senha, quando necessário, utilizando senhas armazenadas em um arquivo JSON.
6. **Monitoramento da Caixa de Entrada**: Exibe a quantidade de e-mails lidos e não lidos na caixa de entrada após a extração.

---

### 📋 Requisitos

- **Python 3.7+**
- **Bibliotecas**: `openpyxl`, `pdfplumber`, `Imbox`, `re`, `datetime`, `json`, `pdfminer.six`
- **Credenciais de e-mail**: Um arquivo JSON contendo as credenciais do e-mail e senhas de arquivos PDF.

---

### 📂 Estrutura do Projeto

```bash
.
├── BD-Boletos.xlsx          # Planilha de destino para armazenar os dados extraídos
├── credenciais_gmail.json   # Arquivo JSON contendo credenciais de e-mail e senhas de PDFs
├── anexos/                  # Diretório onde os anexos baixados dos e-mails serão armazenados
└── script.py                # Script Python com toda a lógica de extração e manipulação
```

---

### 🚀 Como Executar

1. Clone o repositório e instale as dependências.
   
   ```bash
   pip install openpyxl pdfplumber imbox
   ```

2. Configure o arquivo `credenciais_gmail.json` com as informações de e-mail e senhas de PDFs.

3. Execute o script principal.

   ```bash
   python script.py
   ```

---

### 🙏 Agradecimentos

Agradecemos ao time da Empowerdata pelo suporte contínuo e sempre comprometido.
