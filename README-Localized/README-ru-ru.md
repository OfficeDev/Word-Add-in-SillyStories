# <a name="silly-stories-word-add-in-sample-load-files-and-use-content-controls"></a>Пример надстройки Word "Silly stories": загрузка файлов и использование элементов управления контентом

В этой надстройке Word показано, как:

1. скачать список DOCX-файлов из службы и заполнить раскрывающийся список именами файлов;
2. скачать DOCX-файл из службы и вставить его в документ Word;
3. загрузить коллекцию элементов управления контентом и создать поля ввода для них;
4. обновлять текстовое значение коллекции элементов управления контентом в зависимости от значений в полях ввода;
5. создать гармоничный пользовательский интерфейс Word с помощью Office UI Fabric.

> Примечание. Этот пример может вас рассмешить.

## <a name="prerequisites"></a>Необходимые компоненты

Чтобы использовать пример надстройки Word "Silly stories", необходимо следующее:

* [node.js](https://nodejs.org) для подготовки DOCX-файлов.
* [npm](https://www.npmjs.com/) для установки зависимостей.
* JQuery для компонента Office UI Fabric [dropdown](dev.office.com/fabric/components/dropdown).
* Word 2016 или другой клиент, поддерживающий API JavaScript для Word. В этом примере выполняется проверка требований, чтобы определить, запущена ли надстройка в поддерживаемом ведущем приложении.

## <a name="start-the-web-application"></a>Запуск веб-приложения

1. Установите зависимости проекта с помощью диспетчера пакетов Node (npm), выполнив команду ```npm install``` для корневого каталога проекта в командной строке.
2. Запустите сервер разработки, выполнив команду ```node server.js``` для корневого каталога проекта. Надстройка будет запущена по адресу 127.0.0.1:8080.

### <a name="configure-and-run-on-word-for-mac-2016"></a>Настройка и запуск в Word для Mac 2016

1. Создайте папку wef в разделе Users/Library/Containers/com.microsoft.word/Data/Documents/.
2. Поместите манифест в папку wef (Users/Library/Containers/com.microsoft.word/Data/Documents/wef).
3. Откройте Word 2016 на компьютере Mac и выберите раскрывающийся список "Мои надстройки" на вкладке "Вставка". В нем должна появиться надстройка. Чтобы загрузить надстройку, выберите ее.

### <a name="configure-and-run-on-word-for-windows-2016"></a>Настройка и запуск в Word для Windows 2016

1. Создайте сетевую папку или [откройте доступ к папке в сети](https://technet.microsoft.com/ru-ru/library/cc770880.aspx) и поместите в нее файл манифеста [word-add-in-sillystories.xml](word-add-in-sillystories.xml). Вы только что развернули надстройку. Теперь необходимо сообщить Word, где она находится.
2. Запустите Word и откройте документ.
3. Перейдите на вкладку **Файл**, а затем выберите **Параметры**.
4. Выберите **Центр управления безопасностью**, а затем нажмите кнопку **Параметры центра управления безопасностью**.
5. Выберите пункт **Доверенные каталоги надстроек**.
6. В поле **URL-адрес каталога** введите сетевой путь к общей папке, содержащей файл word-add-in-sillystories.xml, а затем выберите элемент **Добавить каталог**.
7. Установите флажок **Показывать в меню** и нажмите кнопку **ОК**.
8. Появится сообщение, информирующее о том, что параметры будут применены при следующем запуске Office. Закройте и перезапустите Word. 

Теперь все готово к запуску надстройки в Word. 

1. Откройте документ Word. 
2. На вкладке **Вставка** в Word 2016 выберите **Мои приложения**. 
3. Перейдите на вкладку **Общая папка**.
4. Выберите элемент **Silly stories add-in** (Надстройка "Silly stories"), а затем нажмите кнопку **Вставить**.
5. На вкладке **Главная** появится новая группа под названием **Надстройка Word** с кнопкой **Silly Stories**. (Они не видны на снимке экрана.) Нажмите эту кнопку, чтобы открыть область задач надстройки.
6. Выберите историю, чтобы добавить стандартный текст в документ Word.

__Рис. 1. Надстройка "Silly stories", загруженная в Word__

![Изображение приложения Word с загруженной надстройкой "Silly stories"](../readme-images/sillystoriesUI.PNG)

## <a name="questions-and-comments"></a>Вопросы и комментарии

Мы будем рады получить от вас отзывы о примере надстройки Word "Silly stories". Вы можете задавать нам вопросы и добавлять предложения в разделе [Issues](https://github.com/OfficeDev/Word-Add-in-SIllyStories/issues) (Проблемы) этого репозитория.

Общие вопросы о разработке надстроек следует задавать на сайте [Stack Overflow](http://stackoverflow.com/questions/tagged/Office365+API). Помечайте свои вопросы и комментарии тегами [office-js], [word] и [API].

## <a name="learn-more"></a>Подробнее

Ниже представлены дополнительные ресурсы, посвященные созданию надстроек на основе API JavaScript для Word.

* [Загрузка неопубликованной надстройки Office на iPad и Mac](http://dev.office.com/docs/add-ins/testing/sideload-an-office-add-in-on-ipad-and-mac)
* 
  [Обзор платформы надстроек Office](https://msdn.microsoft.com/ru-ru/library/office/jj220082.aspx)
* [Надстройки Word](https://github.com/OfficeDev/office-js-docs/blob/master/word/word-add-ins.md)
* [Общие сведения о создании надстроек Word](https://github.com/OfficeDev/office-js-docs/blob/master/word/word-add-ins-programming-guide.md)
* [Обозреватель фрагментов кода для Word](http://officesnippetexplorer.azurewebsites.net/#/snippets/word)
* [Справочник по API JavaScript для надстроек Word](https://github.com/OfficeDev/office-js-docs/tree/master/word/word-add-ins-javascript-reference)

## <a name="copyright"></a>Авторское право
(c) Корпорация Майкрософт (Microsoft Corporation), 2015. Все права защищены.
