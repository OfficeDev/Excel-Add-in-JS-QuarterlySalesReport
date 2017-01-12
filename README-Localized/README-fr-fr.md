# <a name="quarterly-sales-report-task-pane-add-in-sample-for-excel-2016"></a>Exemple de complément de volet Office - Rapport sur les ventes trimestrielles pour Excel 2016

_S’applique à : Excel 2016_

Il s’agit d’un complément de volet Office Excel simple permettant de charger des données dans une feuille de calcul et de créer un graphique de base dans Excel 2016. Il a deux versions : éditeur de code et Visual Studio.

![Exemple de rapport sur les ventes trimestrielles](../images/QuarterlySalesReport_report.PNG)

## <a name="try-it-out"></a>Essayez !
### <a name="code-editor-version"></a>Version d’éditeur de code

Pour déployer et tester votre complément, le plus simple consiste à copier les fichiers sur un partage réseau.

1.  Déployez les fichiers source dans le dossier du projet d’éditeur de code de projet à l’aide de votre serveur favori.
2.  Ouvrez le fichier manifeste (QuarterlySalesReportManifest.xml) et remplacez https://yourserver par le nom de votre serveur. (c'est-à-dire https://localhost). Réalisez cette modification dans les éléments \<SourceLocation\> et \<URL\>.
3.  Copiez le fichier manifeste (QuarterlySalesReportManifest.xml) dans un partage réseau (par exemple, \\\MyShare\MyManifests).
4.  Ajoutez l’emplacement de partage qui contient le fichier manifeste sous forme de catalogue d’applications approuvées dans Excel.

    a.  Lancez Excel et ouvrez une feuille de calcul vide.

    b.  Choisissez l’onglet **Fichier**, puis choisissez **Options**.

    c.  Choisissez l’onglet **Centre de gestion de la confidentialité**, puis choisissez **Paramètres du Centre de gestion de la confidentialité**.

    d.  Choisissez **Catalogues de compléments approuvés**.

    e.  Dans la zone **URL du catalogue**, entrez le chemin d’accès du partage réseau que vous avez créé à l’étape 3, puis choisissez **Ajouter un catalogue**.

   Activez la case à cocher **Afficher dans le menu**, puis cliquez sur **OK**. Un message s’affiche pour vous informer que vos paramètres seront appliqués la prochaine fois que vous démarrerez Office.

5.  Testez et exécutez le complément.

    a.  Dans l’onglet **Insertion** d’Excel 2016, choisissez **Mes compléments**.

    b.  Dans la boîte de dialogue **Compléments Office**, choisissez **Dossier partagé**.

    c.  Choisissez **Exemple de rapport sur les ventes trimestrielles**>**Insertion**. Le complément s’ouvre dans un volet Office à droite de la feuille de calcul active, comme indiqué dans l’illustration suivante.

  ![Exemple de rapport sur les ventes trimestrielles](../images/QuarterlySalesReport_taskpane.PNG))

    d.  Cliquez sur le bouton **Cliquez ici !** pour afficher les données et le graphique à l’intérieur de la feuille de calcul, comme indiqué dans l’illustration suivante.  Pour mettre à jour le graphique dynamiquement, il vous suffit de modifier les données de la plage.

  ![Exemple de rapport sur les ventes trimestrielles](../images/QuarterlySalesReport_report.PNG)

### <a name="visual-studio-version"></a>Version de Visual Studio
1.  Copiez le projet dans un dossier local et ouvrez le fichier Excel-Add-in-JS-QuarterlySalesReport.sln dans Visual Studio.
2.  Appuyez sur F5 pour créer et déployer l’exemple de complément. Excel démarre et le complément s’ouvre dans un volet Office à droite de la feuille de calcul active, comme indiqué dans l’illustration suivante.

  ![Exemple de rapport sur les ventes trimestrielles](../images/QuarterlySalesReport_taskpane.PNG)

3. Cliquez sur le bouton **Cliquez ici !** pour afficher les données et le graphique à l’intérieur de la feuille de calcul, comme indiqué dans l’illustration suivante. Pour mettre à jour le graphique dynamiquement, il vous suffit de modifier les données de la plage.

  ![Exemple de rapport sur les ventes trimestrielles](../images/QuarterlySalesReport_report.PNG)

## <a name="code-it"></a>Le coder

Si vous souhaitez écrire ce complément vous-même à partir de zéro, voir [Créer son premier complément Excel](https://github.com/OfficeDev/office-js-docs/blob/master/excel/build-your-first-excel-add-in.md).


### <a name="learn-more"></a>En savoir plus


1.  [Présentation de la programmation JavaScript pour les compléments Excel](https://github.com/OfficeDev/office-js-docs/blob/master/excel/excel-add-ins-programming-overview.md)
2.  [Explorateur d’extraits de code pour Excel](http://officesnippetexplorer.azurewebsites.net/#/snippets/excel)
3.  [Exemples de code pour les compléments Excel](https://github.com/OfficeDev/office-js-docs/blob/master/excel/excel-add-ins-code-samples.md)
4.  [Référence de l’API JavaScript pour les compléments Excel](https://github.com/OfficeDev/office-js-docs/blob/master/excel/excel-add-ins-javascript-reference.md)
5.  [Création de votre premier complément Excel](https://github.com/OfficeDev/office-js-docs/blob/master/excel/build-your-first-excel-add-in.md)