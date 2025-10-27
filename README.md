# Monitor de Contagem de Arquivos (Formulários de Despesa)

Este projeto contém um script Python para automatizar a varredura e a contagem de arquivos em diretórios de rede, consolidando os resultados em uma única planilha Excel.

O script foi desenvolvido para rodar em segundo plano (via Agendador de Tarefas do Windows), fornecendo um status atualizado da entrega de arquivos em pastas de "Enviados" e "Recebidos", com uma lógica de negócio específica para cada um.

## O que o script faz?

O processo principal executa as seguintes etapas:

1.  **Varredura de "Enviados"**: Escaneia recursivamente todas as subpastas dentro do `CAMINHO_ENVIADOS`. Se um arquivo estiver em `.../Enviados/Comercial/relatorio.xlsx`, ele é listado com a Área "Comercial" e Status "Enviado".
2.  **Varredura de "Recebidos" (Lógica V1)**: Escaneia o `CAMINHO_RECEBIDOS`, mas aplica uma regra especial: ele só procura arquivos que estejam *dentro* de uma subpasta chamada `V1`. Por exemplo, para a área "Comercial", ele só listará arquivos de `.../Recebidos/Comercial/V1/`.
3.  **Consolidação**: Todos os arquivos encontrados são compilados em um DataFrame do Pandas com as colunas: `CaminhoCompleto`, `NomeArquivo`, `Area` e `Status`.
4.  **Atualização do Excel**: O script abre um arquivo Excel (`.xlsx` ou `.xlsm`) existente, **remove a aba de "Contagem" antiga** e a recria com os dados atualizados. Isso é feito de forma segura, preservando todas as outras abas, formatações e macros (VBA) que possam existir no arquivo.
