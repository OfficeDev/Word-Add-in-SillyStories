# <a name="silly-stories-word-add-in-sample-load-files-and-use-content-controls"></a>Exemple de complément Word Silly stories : chargement de fichiers et utilisation de contrôles de contenu

Ce complément Word vous indique comment procéder pour :

1. Charger une liste de fichiers docx à partir d’un service et remplir un contrôle de zone de liste déroulante avec les noms de fichiers sous forme d’options.
2. Charger un fichier docx à partir du service et l’insérer dans le document Word.
3. Charger la collection de contrôles de contenu et créer des zones de saisie en fonction des contrôles de contenu.
4. Mettre à jour la valeur de texte de la collection de contrôles de contenu en fonction des valeurs dans les zones de saisie.
5. Utiliser la structure de l’interface utilisateur Office pour créer une expérience utilisateur Word transparente.

> Remarque : le rire est un effet secondaire à l’exécution de cet exemple.

## <a name="prerequisites"></a>Conditions préalables

Pour utiliser l’exemple de complément Word Silly stories, les éléments suivants sont requis.

* [Node.js](https://nodejs.org) pour servir les fichiers docx.
* [npm](https://www.npmjs.com/) pour installer les dépendances.
* JQuery pour le composant [dropdown](dev.office.com/fabric/components/dropdown) de la structure de l’interface utilisateur Office.
* Word 2016 ou tout autre client qui prend en charge l’API JavaScript de Word. Cet exemple effectue une vérification des conditions requises pour voir s’il est exécuté sur un hôte pris en charge.

## <a name="start-the-web-application"></a>Démarrage de l’application Web

1. Installez les dépendances du projet avec le gestionnaire de package de Node (npm) en exécutant ```npm install``` dans le répertoire racine du projet dans la ligne de commande.
2. Démarrez le serveur de développement en exécutant ```node server.js``` dans le répertoire racine du projet. Exécution du complément : 127.0.0.1:8080.

### <a name="configure-and-run-on-word-for-mac-2016"></a>Configuration et exécution sur Word pour Mac 2016

1. Créez un dossier nommé « wef » dans Users/Library/Containers/com.microsoft.word/Data/Documents/
2. Placez le fichier manifeste dans le dossier « wef » (Users/Library/Containers/com.microsoft.word/Data/Documents/wef).
3. Ouvrez Word 2016 sur le Mac et cliquez sur l’onglet Insertion, puis sur la liste déroulante Mes compléments. Vous devez voir le complément dans la liste déroulante. Sélectionnez-le pour le charger.

### <a name="configure-and-run-on-word-for-windows-2016"></a>Configuration et exécution sur Word 2016 pour Windows

1. Créez un partage réseau ou [partagez un dossier sur le réseau](https://technet.microsoft.com/fr-fr/library/cc770880.aspx), puis placez-y le fichier manifeste [word-add-in-sillystories.xml](word-add-in-sillystories.xml). À ce stade, vous avez déployé votre complément. Vous devez maintenant indiquer à Word où trouver le complément.
2. Lancez Word et ouvrez un document.
3. Choisissez l’onglet **Fichier**, puis choisissez **Options**.
4. Choisissez l’onglet **Fichier**, puis choisissez **Options**.
5. Choisissez **Catalogues de compléments approuvés**.
6. Dans la zone **URL du catalogue**, entrez le chemin réseau pour accéder au partage de dossier qui contient le fichier word-add-in-sillystories.xml, puis choisissez **Ajouter un catalogue**.
7. Activez la case à cocher **Afficher dans le menu**, puis cliquez sur **OK**.
8. Un message vous informe que vos paramètres seront appliqués lors du prochain démarrage d’Office. Fermez et redémarrez Word. 

Vous êtes maintenant prêt à exécuter le complément dans Word. 

1. Ouvrez un document Word. 
2. Dans l’onglet **Insertion** de Word 2016, choisissez **Mes compléments**. 
3. Sélectionnez l’onglet **Dossier partagé**.
4. Choisissez **Complément Silly stories**, puis sélectionnez **Insérer**.
5. Un nouveau groupe appelé **Complément Word** apparaîtra sur l’onglet **Accueil**. Le groupe comporte un bouton nommé **Silly Stories**. (Ceci n’est pas visible dans la capture d’écran.) Cliquez sur ce bouton pour ouvrir le volet Office du complément.
6. Sélectionnez une histoire pour entrer du texte réutilisable dans le document Word.

__Figure 1. Complément Silly stories chargé dans Word__

![Image de l’application Word avec le complément Silly stories chargé](../readme-images/sillystoriesUI.PNG)

## <a name="questions-and-comments"></a>Questions et commentaires

Nous serions ravis de connaître votre opinion sur l’exemple de complément Word Silly stories. Vous pouvez nous faire part de vos questions et suggestions dans la rubrique [Problèmes](https://github.com/OfficeDev/Word-Add-in-SIllyStories/issues) de ce référentiel.

Si vous avez des questions générales sur le développement de compléments, envoyez-les sur [Stack Overflow](http://stackoverflow.com/questions/tagged/Office365+API). Posez vos questions avec les balises [office-js], [word] et [API].

## <a name="learn-more"></a>En savoir plus

Voici des ressources supplémentaires pour vous aider à créer des compléments basés sur l’API JavaScript de Word :

* [Charger une version test d’un complément Office sur iPad ou Mac](http://dev.office.com/docs/add-ins/testing/sideload-an-office-add-in-on-ipad-and-mac)
* 
  [Vue d’ensemble de la plateforme des compléments Office](https://msdn.microsoft.com/fr-fr/library/office/jj220082.aspx)
* [Compléments Word](https://github.com/OfficeDev/office-js-docs/blob/master/word/word-add-ins.md)
* [Présentation de la programmation pour les compléments Word](https://github.com/OfficeDev/office-js-docs/blob/master/word/word-add-ins-programming-guide.md)
* [Explorateur d’extraits de code pour Word](http://officesnippetexplorer.azurewebsites.net/#/snippets/word)
* [Référence de l’API JavaScript pour les compléments Word](https://github.com/OfficeDev/office-js-docs/tree/master/word/word-add-ins-javascript-reference)

## <a name="copyright"></a>Copyright
Copyright (c) 2015 Microsoft. Tous droits réservés.
