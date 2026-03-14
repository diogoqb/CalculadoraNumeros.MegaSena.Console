Overview
Este repositório contém a aplicação CalculadoraNumeros.MegaSena.Console que lê um arquivo sorteio.csv, calcula várias análises estatísticas e gera um arquivo Excel consolidado com todas as análises (sem gráficos). A pasta Data dentro do diretório do aplicativo é usada tanto para leitura quanto para gravação.

Files Included
- CalculadoraNumeros.MegaSena.Console.exe (ou a pasta publish gerada pelo dotnet publish)
- run_app.cmd (opcional)
- run_app.ps1 (opcional)
- README.md
- README.txt
- Data (pasta; o programa cria automaticamente se não existir)
- sorteio.csv (arquivo de entrada; adicione na pasta Data)

How to Run
Self contained executable (recommended)
1. Descompacte o ZIP.
2. Execute run_app.cmd ou dê duplo clique em CalculadoraNumeros.MegaSena.Console.exe.
3. A pasta Data será criada automaticamente se não existir.

Framework dependent
1. Instale o runtime .NET correspondente (ex.: .NET 8).
2. Execute o .exe na pasta publish ou rode dotnet CalculadoraNumeros.MegaSena.Console.dll.

Execução manual
- PowerShell: .\CalculadoraNumeros.MegaSena.Console.exe
- CMD: CalculadoraNumeros.MegaSena.Console.exe

Where to Put Input File
- Coloque sorteio.csv em Data (mesmo diretório do executável).
- O arquivo de saída analises_consolidado_dd-MM-yyyy_HH-mm.xlsx será gravado em Data.

Menu and Usage
- 1 Números isolados por sorteio
- 2 Números isolados por aparições
- 3–7 Top combinações (duplas a sextetos)
- 13 Gerar arquivo Excel consolidado com todas as análises (sem gráficos)
- 0 Sair

Troubleshooting
- Arquivo não encontrado: verifique Data\sorteio.csv e o formato.
- Permissão negada: execute em pasta com permissão de escrita.
- Erro .NET: se framework dependent, confirme runtime instalado.
- Antivírus: pode sinalizar executáveis novos; libere o arquivo se necessário.

Distribution Notes
- Recomendo distribuir um ZIP com a pasta publish (self contained).
- Inclua run_app.cmd e run_app.ps1.
- Adicione CHANGELOG ou versão no nome do ZIP para atualizações.

Contact
Para bugs ou melhorias envie:
- Descrição do problema
- Mensagem de erro completa do console
- Exemplo do sorteio.csv (se aplicável)