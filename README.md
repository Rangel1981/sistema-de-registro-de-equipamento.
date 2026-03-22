# 🛠️ Sistema de Inventário T.I (CRUD Desktop)

Este é um sistema de gerenciamento de ativos de hardware desenvolvido para otimizar o controle de equipamentos (Notebooks, Desktops e periféricos) em um ambiente corporativo de suporte técnico.

O projeto utiliza **Programação Orientada a Objetos (POO)** e uma interface gráfica moderna para facilitar o uso por profissionais de TI.

## 🚀 Funcionalidades
- **Cadastro de Ativos:** Registro de patrimônio, marca, tipo e detalhes técnicos.
- **Listagem em Tempo Real:** Visualização dinâmica dos dados através de uma tabela (Treeview).
- **Exclusão Inteligente:** Remoção de registros com confirmação de segurança (UX).
- **Exportação para Excel:** Geração de relatórios `.xlsx` automáticos para auditoria.
- **Banco de Dados Persistente:** Armazenamento local utilizando SQLite.

## 💻 Tecnologias Utilizadas
- **Python 3.x**
- **CustomTkinter:** Interface gráfica de usuário (GUI) com tema Dark/Light.
- **SQLite3:** Banco de dados relacional leve e embutido.
- **Pandas & Openpyxl:** Processamento de dados e geração de planilhas.
- **Pathlib:** Gerenciamento inteligente de diretórios e arquivos.

## 📁 Estrutura do Projeto
- `sistema_ti.py`: Código principal com a lógica de interface e banco.
- `/arquivos_setor`: Pasta gerada automaticamente para armazenar o banco `.db`.
- `/arquivos_setor/backups`: Local de armazenamento de cópias de segurança.

## 🔧 Como Rodar o Projeto
1. Certifique-se de ter o Python instalado.
2. Instale as dependências necessárias:
   ```bash
   pip install customtkinter pandas openpyxl
