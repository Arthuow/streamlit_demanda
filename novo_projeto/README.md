# Novo Projeto

Aplicação web desenvolvida com Flask para processamento de dados.

## Funcionalidades

- Upload de arquivos CSV e Excel
- Processamento de dados
- Visualização dos resultados em tabela
- Interface moderna e responsiva

## Requisitos

- Python 3.8 ou superior
- pip (gerenciador de pacotes Python)

## Instalação

1. Clone o repositório:
```bash
git clone https://github.com/seu-usuario/novo-projeto.git
cd novo-projeto
```

2. Crie um ambiente virtual (opcional, mas recomendado):
```bash
python -m venv venv
source venv/bin/activate  # Linux/Mac
venv\Scripts\activate     # Windows
```

3. Instale as dependências:
```bash
pip install -r requirements.txt
```

## Executando a Aplicação

1. Inicie o servidor Flask:
```bash
python app.py
```

2. Acesse a aplicação no navegador:
```
http://localhost:5000
```

## Uso

1. Na interface web, clique em "Carregar Dados" e selecione:
   - Arquivo base (CSV)
   - Arquivo de atributos (Excel)

2. Após carregar os arquivos, clique em "Calcular"

3. Os resultados serão exibidos na tabela

## Estrutura do Projeto

```
novo-projeto/
├── app.py              # Aplicação Flask
├── requirements.txt    # Dependências
├── templates/         # Templates HTML
│   └── index.html     # Interface principal
└── uploads/          # Diretório para arquivos temporários
```

## Contribuindo

1. Faça um fork do projeto
2. Crie uma branch para sua feature (`git checkout -b feature/AmazingFeature`)
3. Commit suas mudanças (`git commit -m 'Add some AmazingFeature'`)
4. Push para a branch (`git push origin feature/AmazingFeature`)
5. Abra um Pull Request

## Licença

Este projeto está licenciado sob a licença MIT - veja o arquivo [LICENSE](LICENSE) para detalhes. 