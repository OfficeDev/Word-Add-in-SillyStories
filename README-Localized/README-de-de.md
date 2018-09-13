# <a name="silly-stories-word-add-in-sample-load-files-and-use-content-controls"></a>Word-Add-In-Beispiel „Alberne Geschichten“: Laden von Dateien und Verwenden von Inhaltssteuerelementen

Dieses Word-Add-In veranschaulicht folgende Aktionen:

1. Laden einer Liste mit DOCX-Dateien aus einem Dienst und Auffüllen eines Dropdown-Steuerelements mit den Dateinamen als Optionen
2. Laden einer DOCX-Datei aus dem Dienst und Einfügen der Datei in das Word-Dokument
3. Laden der Inhaltssteuerelement-Auflistung und Erstellen von Eingabefeldern basierend auf den Inhaltssteuerelementen
4. Aktualisieren des Textwerts der Inhaltssteuerelement-Auflistung basierend auf den Werten in den Eingabefeldern
5. Verwenden von Office-UI-Fabric zum Erstellen einer optimalen Word-Benutzererfahrung

> Hinweis: Gekicher ist eine Nebenwirkung der Ausführung dieses Beispiels.

## <a name="prerequisites"></a>Voraussetzungen

Für die Verwendung des Word-Add-In-Beispiels „Alberne Geschichten“ gelten folgende Anforderungen.

* [Node.js](https://nodejs.org) für Unterstützung von DOCX-Dateien
* [npm](https://www.npmjs.com/) für die Installation der Abhängigkeiten
* JQuery für die Office-UI-Fabric-[Dropdown](dev.office.com/fabric/components/dropdown)-Komponente.
* Word 2016 oder beliebiger Client, der die Javascript-API für Word unterstützt In diesem Beispiel wird eine Prüfung der Anforderungen durchgeführt, um zu sehen, ob dieses in einem unterstützten Host ausgeführt wird.

## <a name="start-the-web-application"></a>Starten der Webanwendung

1. Installieren Sie Projektabhängigkeiten mithilfe des Paket-Managers von Node (npm), indem Sie ```npm install``` im Stammverzeichnis des Projekts an der Befehlszeile ausführen.
2. Starten Sie den Entwicklungsserver, indem Sie ```node server.js``` im Stammverzeichnis des Projekts ausführen. Unter 127.0.0.1:8080 wird das Add-In ausgeführt.

### <a name="configure-and-run-on-word-for-mac-2016"></a>Konfigurieren und Ausführen unter Word für Mac 2016

1. Erstellen Sie einen Ordner namens „wef“ unter Users/Library/Containers/com.microsoft.word/Data/Documents/.
2. Legen Sie das Manifest im Ordner „wef“ ab (Users/Library/Containers/com.microsoft.word/Data/Documents/wef).
3. Öffnen Sie Word 2016 auf dem Mac, und klicken Sie auf der Registerkarte „Einfügen“ auf die Dropdownliste „Mein Add-Ins“. Das Add-In sollte in der Dropdownliste angezeigt werden. Wählen Sie es aus, und es wird geladen.

### <a name="configure-and-run-on-word-for-windows-2016"></a>Konfigurieren und Ausführen unter Word für Windows 2016

1. Erstellen Sie eine Netzwerkfreigabe oder [geben Sie einen Ordner im Netzwerk frei](https://technet.microsoft.com/de-de/library/cc770880.aspx), und platzieren Sie die [word-add-in-sillystories.xml](word-add-in-sillystories.xml)-Manifestdatei darin. Sie haben nun Ihr Add-In bereitgestellt. Jetzt müssen Sie Word mitteilen, wo es das Add-In finden kann.
2. Starten Sie Word, und öffnen Sie ein Dokument.
3. Klicken Sie auf die Registerkarte **Datei**, und klicken Sie dann auf **Optionen**.
4. Wählen Sie **Sicherheitscenter** aus, und klicken Sie dann auf die Schaltfläche **Einstellungen für das Sicherheitscenter**.
5. Wählen Sie **Vertrauenswürdige Add-In-Kataloge** aus.
6. Geben Sie in das Feld **Katalog-URL** den Netzwerkpfad zur Ordnerfreigabe an, die die Datei „word-add-in-sillystories.xml“ enthält, und wählen Sie dann **Katalog hinzufügen**.
7. Aktivieren Sie das Kontrollkästchen **Im Menü anzeigen**, und klicken Sie dann auf **OK**.
8. Es wird eine Meldung angezeigt, dass Ihre Einstellungen angewendet werden, wenn Sie Office das nächste Mal starten. 

Nun können Sie es in Word ausführen. 

1. Öffnen Sie ein Word-Dokument. 
2. Klicken Sie auf der Registerkarte **Einfügen** in Word 2016 auf **Meine-Add-Ins**. 
3. Klicken Sie auf die Registerkarte **Freigegebener Ordner**.
4. Wählen Sie das **Add-In „Alberne Geschichten“**, und wählen Sie dann **Einfügen**.
5. Eine neue Gruppe mit dem Namen**Word-Add-In** wird auf der Registerkarte **Start** angezeigt. Die Gruppe verfügt über eine Schaltfläche mit dem Namen **Alberne Geschichten**. (Diese sind nicht im Screenshot sichtbar.) Klicken Sie auf diese Schaltfläche, um den Add-In-Aufgabenbereich zu öffnen.
6. Wählen Sie eine Geschichte aus, um Textbausteine in das Word-Dokument einzufügen.

__Abbildung 1. Das in Word geladene Add-In „Alberne Geschichten“__

![Abbildung der Word-Anwendung mit geladenem Add-In „Alberne Geschichten“](../readme-images/sillystoriesUI.PNG)

## <a name="questions-and-comments"></a>Fragen und Kommentare

Wir schätzen Ihr Feedback hinsichtlich des Word-Add-In-Beispiels „Alberne Geschichten“. Sie können uns Ihre Fragen und Vorschläge über den Abschnitt [Probleme](https://github.com/OfficeDev/Word-Add-in-SIllyStories/issues) dieses Repositorys senden.

Allgemeine Fragen zur Add-In-Entwicklung sollten in [Stack Overflow](http://stackoverflow.com/questions/tagged/Office365+API) gestellt werden. Stellen Sie sicher, dass Ihre Fragen oder Kommentare mit [office-js], [word] und [API] markiert sind.

## <a name="learn-more"></a>Weitere Informationen

Hier noch einige weitere Ressourcen zum Erstellen von Add-Ins auf Basis von Word-JavaScript-APIs:

* [Querladen eines Office-Add-Ins auf einem iPad und einem Mac-Computer](http://dev.office.com/docs/add-ins/testing/sideload-an-office-add-in-on-ipad-and-mac)
* 
  [Office-Add-Ins-Plattformübersicht](https://msdn.microsoft.com/de-de/library/office/jj220082.aspx)
* [Word-Add-Ins](https://github.com/OfficeDev/office-js-docs/blob/master/word/word-add-ins.md)
* [Programmierungsübersicht für Word-Add-Ins](https://github.com/OfficeDev/office-js-docs/blob/master/word/word-add-ins-programming-guide.md)
* [Codeausschnitt-Explorer für Word](http://officesnippetexplorer.azurewebsites.net/#/snippets/word)
* [JavaScript-API-Referenz zu Word-Add-Ins](https://github.com/OfficeDev/office-js-docs/tree/master/word/word-add-ins-javascript-reference)

## <a name="copyright"></a>Copyright
Copyright (c) 2015 Microsoft. Alle Rechte vorbehalten.
