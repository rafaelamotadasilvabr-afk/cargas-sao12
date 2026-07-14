# CARGAS SAO12

Painel em Python + Streamlit para controle operacional de cargas.

## O que a versão atual entrega

- upload simultâneo dos relatórios `AWBStatusAtPieceLevel` de SAO12 e TRES1;
- consolidação em uma linha por AWB;
- prioridade operacional quando a mesma AWB possui mais de um status;
- cards para Pendente Embarque, Pendente Entrega, Pendente Desembarque, Missing Cargo e Discrepância;
- filtros por cliente, operação, status, base, destino e SLA;
- ranking das bases ofensoras pelo campo `OPSStation`;
- exportação do controle para Excel e CSV;
- identificação de cliente não mapeado, SLA ausente, base ausente e múltiplos status;
- campos de upload dos dois relatórios complementares.

## Clientes reconhecidos

- Riachuelo;
- Tania Bulhões — incluindo `TB COMERCIO DE PRESENTES SA`;
- Della Via;
- Três Corações — todos os registros com `OriginCode = TRES1`.

## Executar no Windows

1. Instale Python 3.11 ou superior.
2. Durante a instalação, marque `Add Python to PATH`.
3. Extraia a pasta do projeto.
4. Dê dois cliques em `iniciar_dashboard.bat`.
5. O navegador abrirá o painel.
6. Anexe os relatórios e clique em `PROCESSAR RELATÓRIOS`.

Na primeira execução, o arquivo `.bat` cria um ambiente virtual e instala as dependências.

## Executar manualmente

```bash
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
streamlit run app.py
```

## Publicar no GitHub e Streamlit Community Cloud

1. Crie um repositório no GitHub.
2. Envie todos os arquivos deste projeto, exceto a pasta `.venv`.
3. No Streamlit Community Cloud, selecione o repositório.
4. Defina `app.py` como arquivo principal.

Não envie relatórios operacionais para o GitHub.

## Regra de criticidade

1. Missing Cargo
2. Discrepância Criada
3. Pendente Embarque
4. Pendente Desembarque
5. Pendente Entrega
6. Atribuído à Rota
7. Saído para Entrega
8. Entregue
9. Baixado

## Integração dos relatórios complementares

A estrutura já possui campos separados para relatório de embarque e relatório de entrega. Nesta primeira versão, eles são recebidos, validados e exibidos. Para fazer o cruzamento definitivo, é necessário mapear os cabeçalhos reais desses dois arquivos.
