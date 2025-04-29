# Análise de Demanda

Aplicação web para análise de demanda elétrica, desenvolvida com Flask.

## Funcionalidades

- Upload de arquivos de medição (CSV) e atributos (Excel)
- Cálculo automático de demanda
- Visualização dos resultados em tabela
- Exportação dos resultados para Excel
- Interface moderna e responsiva

## Requisitos

- Python 3.8 ou superior
- pip (gerenciador de pacotes Python)

## Instalação

1. Clone o repositório:
```bash
git clone https://github.com/seu-usuario/analise-demanda.git
cd analise-demanda
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
   - Arquivo de medição (CSV)
   - Arquivo de atributos (Excel)

2. Após carregar os arquivos, clique em "Calcular Demanda"

3. Os resultados serão exibidos na tabela

4. Para exportar os resultados, clique em "Exportar"

## Estrutura do Projeto

```
analise-demanda/
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
