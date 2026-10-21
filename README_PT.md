# HRV BatchExtract

*[Read in English](README.md)*

O HRV BatchExtract é um aplicativo para Windows que automatiza a extração dos resultados de variabilidade da frequência cardíaca (VFC) de curto prazo do **Kubios HRV Scientific**. Ele abre cada arquivo de intervalos RR no Kubios, aplica as mesmas configurações de análise a todos os arquivos, copia as tabelas de resultados do Kubios e as reúne em uma única planilha do Excel.

**O HRV BatchExtract não calcula a VFC.** Todos os índices são calculados pelo Kubios. Os valores da planilha são copiados como texto das tabelas de resultados do Kubios e, portanto, são idênticos aos exibidos pelo programa.

O HRV BatchExtract é um projeto independente. Não tem vínculo com a Kubios Oy, não é endossado nem distribuído por ela, e não inclui nem modifica o Kubios.

## Versões

| Pasta | Idioma da interface | Arquivo para iniciar |
|---|---|---|
| [`en/`](en) | Inglês | `RUN_HRV_BatchExtract.bat` |
| [`pt/`](pt) | Português | `RODAR_HRV_BatchExtract.bat` |

As duas versões executam o mesmo procedimento. Diferem apenas no idioma da interface, das mensagens e dos nomes de saída.

## Requisitos

- Windows 10 ou 11 (64 bits), um monitor, escala de exibição inalterada durante a execução.
- **Kubios HRV Scientific Lite 4.3.0** (gratuito, exige conta Kubios) ou uma licença paga do Kubios HRV Scientific. O aplicativo foi desenvolvido e testado com o Lite 4.3.0 em inglês. Uma versão paga ou outra versão deve ser testada em alguns arquivos antes do uso.
- Python 3.10 ou superior (qualquer distribuição, inclusive Anaconda ou a que acompanha o Spyder).
- Internet na primeira execução, para instalar as bibliotecas Python: pyautogui, pyperclip, pygetwindow, pandas, numpy, openpyxl, pillow e winsdk.

## Arquivos de entrada

- Um arquivo de texto (.txt) por participante, com um intervalo RR por linha, uma única coluna, sem cabeçalho.
- Cada arquivo já recortado no segmento de análise de **5 min** (Task Force, 1996). O HRV BatchExtract não seleciona nem edita segmentos.
- Intervalos RR em segundos ou milissegundos (detectados automaticamente).
- Todos os arquivos em uma mesma pasta. O nome do arquivo é o código do participante (o sufixo `_5min` é removido). Os nomes não devem conter identificadores pessoais.

## Configurações aplicadas a todos os arquivos

- Correção de artefatos: por limiar (threshold), nível **medium**, reaplicada e conferida na tela em todos os arquivos.
- Amostra de análise: 00:00:00 a 00:05:00.
- Demais configurações: padrões do Kubios.

## Como executar

1. Baixe a pasta do seu idioma (`en` ou `pt`) e mantenha os dois arquivos juntos em uma pasta local.
2. Ligue o computador na tomada, desative a suspensão e o bloqueio de tela, feche outros programas e silencie as notificações.
3. Dê dois cliques no arquivo para iniciar (`RODAR_HRV_BatchExtract.bat`). A primeira execução instala as bibliotecas.
4. Na janela, escolha a pasta com os arquivos .txt. Confira o caminho do `launch_KubiosHRV.exe` (o caminho padrão de instalação já vem preenchido).
5. Clique em **RODAR**.
6. **Não use o mouse nem o teclado durante a execução.** O aplicativo os controla.
7. Para parar na hora, leve o mouse ao canto superior esquerdo da tela. **Parar** encerra a execução após o arquivo em andamento. Ao rodar de novo na mesma pasta, o aplicativo continua com os arquivos ainda não processados.

## Saídas

É criada a subpasta `hrv_batchextract_resultados` dentro da pasta de dados:

- `HRV_BatchExtract.xlsx`: aba *HRV* (uma linha por participante: código, 51 índices do Kubios, SD1% e SD2%), aba *Verificacao* (unidade, conferência da média de RR, conferência de SD1 vs RMSSD/√2, rótulo da correção, número e percentual de batimentos corrigidos) e aba *Metodo* (configurações aplicadas).
- `brutos/`: o texto copiado das tabelas de resultados do Kubios para cada arquivo.
- `capturas_correcao/`: uma captura de tela do campo de correção de batimentos do Kubios para cada arquivo.

Os arquivos RR originais nunca são alterados. A entropia multiescala não é extraída (não disponível na versão Lite).

## Limitações

- O aplicativo opera a interface do Kubios por posições na tela. Está validado apenas para o Kubios HRV Scientific Lite 4.3.0, em inglês, no Windows.
- O computador não pode ser usado durante a execução.
- O aplicativo não julga a qualidade do sinal. A seleção do segmento, a inspeção visual e os critérios de exclusão continuam sob responsabilidade do pesquisador.

## Licença

Licença MIT (ver [LICENSE](LICENSE)). Você pode baixar, usar e modificar a sua cópia, mantendo o aviso de direitos autorais. Alterações neste repositório são feitas apenas pelo proprietário.

## Como citar

Um protocolo de estudo que descreve o HRV BatchExtract e sua validação está em preparação. Até sua publicação, cite este repositório.

O código foi escrito com o auxílio de uma ferramenta de inteligência artificial generativa (Claude, Anthropic). O autor definiu os requisitos, testou o aplicativo com dados reais, conferiu sua saída com o Kubios e é responsável pelo código.
