# <a name="quarterly-sales-report-task-pane-add-in-sample-for-excel-2016"></a>Exemplo do Suplemento do Painel de Tarefas de Relatório de Vendas Trimestral para o Excel 2016

_Aplica-se a: Excel 2016_

Esse é um suplemento simples do painel de tarefas que carrega alguns dados em uma planilha e cria um gráfico básico no Excel 2016. Há dois tipos: o editor de código e o Visual Studio.

![Exemplo de Relatório de Vendas Trimestral](../Images/QuarterlySalesReport_report.PNG)

## <a name="try-it-out"></a>Experimente
### <a name="code-editor-version"></a>Versão do editor de código

A maneira mais fácil de implantar e testar o suplemento é copiar os arquivos para um compartilhamento de rede.

1.  Implante os arquivos de origem na pasta Projeto do Editor de Códigos usando seu servidor favorito.
2.  Abra o arquivo de manifesto (QuarterlySalesReportManifest.xml) e substitua https://yourserver com o nome do seu servidor (isto é, https://localhost) Faça isso com o elemento \<SourceLocation\> e o elemento \<URLs\>.
3.  Copie o manifesto (QuarterlySalesReportManifest.xml) para um compartilhamento de rede (por exemplo, \\\MyShare\MyManifests).
4.  Adicione o local de compartilhamento que contém o manifesto como um catálogo de aplicativos confiáveis no Excel.

    a.  Inicie o Excel e abra uma planilha em branco.

    b.  Escolha a guia **Arquivo** e escolha **Opções**.

    c.  Escolha **Central de Confiabilidade**, e escolha o botão **Configurações da Central de Confiabilidade**.

    d.  Escolha **Catálogos de Suplementos Confiáveis**.

    e.  Na caixa **URL de Catálogo**, insira o caminho para o compartilhamento de rede que você criou na etapa 3 e escolha **Adicionar Catálogo**.

   f.  Marque a caixa de seleção **Mostrar no Menu** e escolha **OK**. Será exibida uma mensagem para informá-lo de que suas configurações serão aplicadas na próxima vez que você iniciar o Office.

5.  Teste e execute o suplemento.

    a.  Na **guia Inserir** no Excel 2016, escolha **Meus Suplementos**.

    b.  Na caixa de diálogo **Suplementos do Office**, escolha **Pasta Compartilhada**.

    c.  Escolha **Exemplo de Relatório de Vendas Trimestral**>**Inserir**. O suplemento abre em um painel de tarefas à direita da planilha atual, conforme mostrado na figura a seguir.

  ![Exemplo de Relatório de Vendas Trimestral](../Images/QuarterlySalesReport_taskpane.PNG))

    d.  Clique no botão **Click me!** para processar os dados e o gráfico dentro da planilha, conforme mostrado na figura a seguir.  Para ver o gráfico sendo atualizado dinamicamente, basta alterar os dados no intervalo.

  ![Exemplo de Relatório de Vendas Trimestral](../Images/QuarterlySalesReport_report.PNG)

### <a name="visual-studio-version"></a>Versão do Visual Studio
1.  Copie o projeto para uma pasta local e abra o Excel-Add-in-JS-QuarterlySalesReport.sln no Visual Studio.
2.  Pressione F5 para criar e implantar o suplemento de exemplo. O Excel inicia e o suplemento abre em um painel de tarefas à direita da planilha em branco, conforme mostrado na figura a seguir.

  ![Exemplo de Relatório de Vendas Trimestral](../Images/QuarterlySalesReport_taskpane.PNG)

3. Clique no botão **Click me!** para processar os dados e o gráfico dentro da planilha, conforme mostrado na figura a seguir.  Para ver o gráfico sendo atualizado dinamicamente, basta alterar os dados no intervalo.

  ![Exemplo de Relatório de Vendas Trimestral](../Images/QuarterlySalesReport_report.PNG)

## <a name="code-it"></a>Code it

Se você quiser gravar esse suplemento sozinho desde o início, confira [Criar seu primeiro Suplemento do Excel](https://github.com/OfficeDev/office-js-docs/blob/master/excel/build-your-first-excel-add-in.md).


### <a name="learn-more"></a>Saiba mais


1.  [Visão geral da programação de Suplementos do Excel](https://github.com/OfficeDev/office-js-docs/blob/master/excel/excel-add-ins-programming-overview.md)
2.  [Explorador de Trechos para Excel](http://officesnippetexplorer.azurewebsites.net/#/snippets/excel)
3.  [Exemplos de código de Suplementos do Excel](https://github.com/OfficeDev/office-js-docs/blob/master/excel/excel-add-ins-code-samples.md)
4.  [Referência da API JavaScript de Suplementos do Excel](https://github.com/OfficeDev/office-js-docs/blob/master/excel/excel-add-ins-javascript-reference.md)
5.  [Criar seu primeiro Suplemento do Excel](https://github.com/OfficeDev/office-js-docs/blob/master/excel/build-your-first-excel-add-in.md)