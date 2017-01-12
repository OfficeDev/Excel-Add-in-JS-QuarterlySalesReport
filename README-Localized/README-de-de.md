# <a name="quarterly-sales-report-task-pane-add-in-sample-for-excel-2016"></a>Aufgabenbereich-Add-In-Beispiel für Quartalsumsatzbericht für Excel 2016

_Gilt für: Excel 2016_

Dies ist ein einfaches Excel-Aufgabenbereich-Add-In, das einige Daten in ein Arbeitsblatt lädt und ein einfaches Diagramm in Excel 2016 erstellt. Es ist in zwei Versionen verfügbar: Code-Editor und Visual Studio.

![Quartalsumsatzbericht-Beispiel](../Images/QuarterlySalesReport_report.PNG)

## <a name="try-it-out"></a>Probieren Sie es aus
### <a name="code-editor-version"></a>Code-Editor-Version

Am einfachsten können Sie Ihr Add-In bereitstellen und testen, indem Sie die Dateien in eine Netzwerkfreigabe kopieren.

1.  Stellen Sie die Quelldateien im Ordner „Code Editor Proj“ unter Verwendung eines Servers Ihrer Wahl.
2.  Öffnen Sie die Manifestdatei („QuarterlySalesReportManifest.xml“), und ersetzen Sie „https://yourserver“ durch den Namen Ihres Server. (d. h. https://localhost). Führen Sie dies für die beiden Elemente \<SourceLocation\> und \<Urls\> aus.
3.  Kopieren Sie das Manifest (QuarterlySalesReportManifest.xml) in eine Netzwerkfreigabe (z. B. \\\MyShare\MyManifests).
4.  Fügen Sie den Freigabepfad, unter dem das Manifest enthalten ist, als vertrauenswürdigen App-Katalog in Excel hinzu.

    a.  Starten Sie Excel, und öffnen Sie ein leeres Arbeitsblatt.

    b.  Klicken Sie auf die Registerkarte **Datei**, und klicken Sie dann auf **Optionen**.

    c.  Wählen Sie **Sicherheitscenter** aus, und klicken Sie dann auf die Schaltfläche **Einstellungen für das Sicherheitscenter**.

    d.  Wählen Sie **Vertrauenswürdige Add-In-Kataloge** aus.

    e.  Geben Sie im Feld  **Katalog-URL** den Pfad zu der in Schritt 3 erstellten Netzwerkfreigabe ein, und klicken Sie auf **Katalog hinzufügen**.

   f. Aktivieren Sie das Kontrollkästchen **Im Menü anzeigen**, und wählen Sie dann **OK**. Eine Meldung wird angezeigt, dass Ihre Einstellungen angewendet werden, wenn Office das nächste Mal gestartet wird.

5.  Testen und führen Sie das Add-In aus.

    a.  Klicken Sie auf der Registerkarte **Einfügen** in Excel 2016 auf **Meine-Add-Ins**.

    b.  Wählen Sie im Dialogfenster **Office-Add-Ins** die Option **Freigegebener Ordner** aus.

    c.  Wählen Sie **Quartalsumsatzbericht-Beispiel**>**Einfügen** aus. Das Add-In wird in einem Aufgabenbereich rechts neben dem aktuellen Arbeitsblatt geöffnet, wie in der folgenden Abbildung dargestellt.

  ![Quartalsumsatzbericht-Beispiel](../Images/QuarterlySalesReport_taskpane.PNG))

    d.  Klicken Sie auf die Schaltfläche **Klicken Sie hier!**, um die Daten und das Diagramm im Arbeitsblatt anzuzeigen, wie in der folgenden Abbildung dargestellt.  Ändern Sie zum dynamischen Aktualisieren des Diagramms einfach die Daten in dem Bereich.

  ![Quartalsumsatzbericht-Beispiel](../Images/QuarterlySalesReport_report.PNG)

### <a name="visual-studio-version"></a>Visual Studio-Version
1.  Kopieren Sie das Projekt in einen lokalen Ordner, und öffnen Sie die Datei „Excel-Add-in-JS-QuarterlySalesReport.sln“ in Visual Studio.
2.  Drücken Sie F5, um das Beispiel-Add-In zu erstellen und bereitzustellen. Excel wird gestartet und das Add-In wird in einem Aufgabenbereich rechts neben einem leeren Arbeitsblatt geöffnet, wie in der folgenden Abbildung dargestellt.

  ![Quartalsumsatzbericht-Beispiel](../Images/QuarterlySalesReport_taskpane.PNG)

3. Klicken Sie auf die Schaltfläche **Klicken Sie hier!**, um die Daten und das Diagramm im Arbeitsblatt anzuzeigen, wie in der folgenden Abbildung dargestellt. Ändern Sie zum dynamischen Aktualisieren des Diagramms einfach die Daten in dem Bereich.

  ![Quartalsumsatzbericht-Beispiel](../Images/QuarterlySalesReport_report.PNG)

## <a name="code-it"></a>Schreiben des Codes

Wenn Sie dieses Add-in selbst neu schreiben möchten, finden Sie unter [Erstellen Ihres ersten Excel-Add-Ins](https://github.com/OfficeDev/office-js-docs/blob/master/excel/build-your-first-excel-add-in.md) Informationen dazu.


### <a name="learn-more"></a>Weitere Informationen


1.  [Programmierungsübersicht für Excel-Add-Ins](https://github.com/OfficeDev/office-js-docs/blob/master/excel/excel-add-ins-programming-overview.md)
2.  [Codeausschnitt-Explorer für Excel](http://officesnippetexplorer.azurewebsites.net/#/snippets/excel)
3.  [Codebeispiele zu Excel-Add-Ins](https://github.com/OfficeDev/office-js-docs/blob/master/excel/excel-add-ins-code-samples.md)
4.  [JavaScript-API-Referenz zu Excel-Add-Ins](https://github.com/OfficeDev/office-js-docs/blob/master/excel/excel-add-ins-javascript-reference.md)
5.  [Erstellen Ihres ersten Excel-Add-Ins](https://github.com/OfficeDev/office-js-docs/blob/master/excel/build-your-first-excel-add-in.md)