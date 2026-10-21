# Análise de Passagens — ARTESP

Aplicação em Python para processar documentos e imagens relacionados à análise de passagens, evasões de pedágio, isenções judiciais e identificação de veículos. O ponto de entrada é `main.py`, que apresenta um menu e encaminha cada opção ao módulo responsável.

Este README descreve a organização atual do projeto e serve como guia inicial para quem ainda não trabalhou com o código.

## Início rápido

Abra o PowerShell na pasta do projeto e execute:

```powershell
.\.venv\Scripts\python.exe .\main.py
```

O comando chama diretamente o Python do ambiente `.venv`; não é necessário ativar o ambiente antes. Também é possível executar `python main.py`, mas nesse caso será usado o Python que estiver configurado no terminal, que pode ser diferente do ambiente do projeto.

No menu, escolha o processamento e informe uma das formas de entrada oferecidas:

1. Um arquivo;
2. Uma pasta, processando os arquivos diretamente nela;
3. Uma pasta com subpastas, incluindo a busca recursiva.

Algumas rotinas também solicitam uma pasta de imagens associadas. Siga as instruções exibidas no próprio terminal.

## Opções do menu

| Opção | Processamento | Entrada principal | Pasta de resultados |
|---:|---|---|---|
| 1 | Análise de Fraudes | Planilhas Excel | `resultados/Análise de Fraudes/` |
| 2 | Autoban — Isenção Judicial | PDF | `resultados/Autoban — Isenção Judicial/` |
| 3 | SPVias — Evasão de Pedágio | Planilhas Excel | `resultados/SPVias — Evasão de Pedágio/` |
| 4 | Tamoios — Evasão Free Flow | Planilhas Excel | `resultados/Tamoios — Evasão Free Flow/` |
| 5 | Tamoios / Rota das Bandeiras — Isenção | Planilhas Excel e pasta de imagens | `resultados/Tamoios - Rota das Bandeiras — Isenção/` |
| 6 | CNL — Isenção Judicial | Planilhas Excel | `resultados/CNL — Isenção Judicial/` |
| 7 | Via Colinas — Evasão de Pedágio | Planilhas Excel | `resultados/Via Colinas — Evasão de Pedágio/` |
| 8 | Ecovias — Evasão de Pedágio | PDF | `resultados/Ecovias - Evasão de Pedágio/` |
| 9 | SAU — Veículos Administrativos | Imagens JPG, JPEG ou PNG | A rotina SAU informa o destino dos resultados no terminal |
| 10 | SAU — Veículos Operacionais | Imagens JPG, JPEG ou PNG | A rotina SAU informa o destino dos resultados no terminal |

Os nomes das pastas dos processamentos principais são definidos centralmente em `config.py`. A saída é gravada dentro da pasta correspondente ao segmento.

## Processamento em lote

Ao selecionar uma pasta, os arquivos compatíveis são reunidos em um único processamento. Para as rotinas que geram planilhas consolidadas, o resultado é salvo ao final da análise, em um único Excel para aquela seleção de arquivo ou pasta. Quando a opção inclui subpastas, a busca percorre também os diretórios internos.

O terminal acompanha o andamento durante a leitura e apresenta o resumo final depois do processamento. Em caso de falhas em arquivos individuais, consulte as mensagens exibidas no terminal e verifique também o resultado consolidado.

Se já existir um Excel com o mesmo nome no destino, a aplicação procura um nome livre, acrescentando um número ao final (por exemplo, `resultado2.xlsx`) para não substituir o arquivo anterior.

## Leitura e comparação de placas

A lógica reutilizável de placas está em `placas.py`. Ela reúne operações comuns de normalização e comparação de placas, incluindo o mapa de caracteres visualmente confundíveis (`MAPA_DIGITOS`). Os módulos de extração podem usar essas funções para comparar a placa reconhecida pelo OCR com a placa obtida de outra fonte.

As rotinas de OCR e detecção são executadas por processos auxiliares:

- `ocr_process.py`: reconhecimento de texto e placa;
- `detector_process.py`: detecção/localização de placas;
- `common.py`: funções compartilhadas entre os processadores;
- `config.py`: caminhos, modelos, ambientes auxiliares e configuração de saída.

Ao alterar a regra de normalização ou comparação, prefira centralizar a mudança em `placas.py` e revisar os módulos que a consomem. Nem todos os segmentos analisam os mesmos dados ou aplicam exatamente os mesmos critérios de negócio.

## Ambientes Python e dependências

O projeto mantém três ambientes virtuais, cada um com seu arquivo de dependências:

| Ambiente | Uso | Arquivo de dependências |
|---|---|---|
| `.venv` | Aplicação principal e menu (`main.py`) | `requirements.txt` |
| `.venv_ocr` | Reconhecimento de texto e placas | `requirements_ocr.txt` |
| `.venv_yolo` | Detecção de placas | `requirements_yolo.txt` |

`config.py` aponta os processos auxiliares para os executáveis específicos de `.venv_ocr` e `.venv_yolo`. Inicie o menu pelo `.venv` para garantir que os pacotes principais estejam disponíveis:

```powershell
.\.venv\Scripts\python.exe .\main.py
```

Para recriar os ambientes em uma máquina com Python compatível, execute na raiz do projeto:

```powershell
py -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt

py -m venv .venv_ocr
.\.venv_ocr\Scripts\python.exe -m pip install -r requirements_ocr.txt

py -m venv .venv_yolo
.\.venv_yolo\Scripts\python.exe -m pip install -r requirements_yolo.txt
```

As listas foram preparadas a partir do `pip list` de cada ambiente e fixam as versões informadas. O `requirements_ocr.txt` foi montado com a lista enviada porque o executável atual de `.venv_ocr` não iniciou neste caminho. Esses arquivos registram versões, mas não garantem instalação idêntica em qualquer máquina: OCR e detecção usam pacotes de GPU/CUDA, e a compatibilidade depende do Python, do driver e do hardware disponíveis. Valide a instalação e os processos auxiliares no computador de destino.

Não distribua as pastas `.venv`, `.venv_ocr` ou `.venv_yolo` como código-fonte. Recrie-as usando os arquivos de dependências e mantenha os modelos necessários na pasta `models/`.

## Modelos necessários

A pasta `models/` contém arquivos usados pela detecção e pelo processamento de imagens, incluindo:

- `bestnew1024.onnx` — modelo configurado para detecção;
- `best_1024.onnx` e `best_512.onnx` — outros modelos presentes no projeto;
- `FSRCNN_x2.pb` — modelo de super-resolução usado no processamento de imagem.

Mantenha os modelos necessários junto ao projeto e confira os caminhos configurados em `config.py` ao instalar em outra máquina. Os modelos podem ser grandes; não os remova sem confirmar quais rotinas os carregam.

## Organização dos arquivos

| Arquivo ou pasta | Responsabilidade |
|---|---|
| `main.py` | Menu principal, seleção de entradas e encaminhamento aos processadores |
| `config.py` | Configurações compartilhadas, caminhos de ambientes e diretórios de saída |
| `common.py` | Utilitários comuns aos processadores |
| `placas.py` | Limpeza, normalização e comparação de placas |
| `analises_fraudes.py` | Análise de fraudes |
| `extrair_autoban.py` | Extração para Autoban — Isenção Judicial |
| `extrair_spvias.py` | Extração para SPVias — Evasão de Pedágio |
| `extrair_tamoios.py` | Extração para Tamoios — Evasão Free Flow |
| `extrair_tamoios_isencao.py` | Análise de isenções Tamoios / Rota das Bandeiras |
| `cnl_isencao_judicial.py` | Análise de isenções judiciais CNL |
| `via_colinas.py` | Análise de evasão Via Colinas |
| `extrair_ecovias.py` | Extração para Ecovias — Evasão de Pedágio |
| `detector_process.py` | Processo auxiliar de detecção de placas |
| `ocr_process.py` | Processo auxiliar de OCR |
| `SAU/` | Processamentos de imagens de veículos administrativos e operacionais |
| `requirements.txt` | Dependências do ambiente principal `.venv` |
| `requirements_ocr.txt` | Dependências do ambiente `.venv_ocr` |
| `requirements_yolo.txt` | Dependências do ambiente `.venv_yolo` |
| `models/` | Modelos necessários ao processamento de imagens |
| `resultados/` | Planilhas e arquivos gerados pelas análises |

## Diagnóstico rápido

- **O programa não inicia:** execute o comando de início rápido e confira se `.venv\Scripts\python.exe` existe.
- **Erro ao abrir Excel:** confirme que o arquivo não está aberto em outro programa, que a extensão é compatível e que a planilha possui as abas esperadas para aquele segmento.
- **Aba não encontrada:** confira o nome da aba no arquivo de entrada e a mensagem do processador; nomes de abas podem variar entre relatórios.
- **Erro de OCR ou detecção:** confira se `.venv_ocr` e `.venv_yolo` existem, se os modelos estão em `models/` e se os caminhos em `config.py` apontam para eles.
- **Nenhum resultado gerado:** verifique se os arquivos selecionados são do tipo esperado e se contêm os dados necessários. Leia os avisos e erros exibidos durante a execução.
- **Resultado em outra pasta:** consulte o caminho indicado no terminal e confirme a configuração de saída em `config.py`.

## Recomendações para manutenção e repasse

1. Execute o projeto pela raiz com o comando indicado neste README.
2. Ao adicionar um processador, registre a opção e o caminho do script em `main.py`; configure o diretório de saída em `config.py` quando aplicável.
3. Preserve os critérios específicos de cada segmento: a extração e a análise de negócio variam, mesmo quando etapas de placa podem ser compartilhadas.
4. Ao modificar OCR, detector ou comparação de placas, verifique os outros processadores que usam essas funções.
5. Para o repasse a outra equipe, inclua o código, os modelos autorizados e um inventário reproduzível das dependências e versões dos três ambientes.
6. Evite armazenar relatórios reais com dados pessoais no repositório. Use arquivos de exemplo anonimizados para documentar entradas.


