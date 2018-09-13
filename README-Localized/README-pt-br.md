# <a name="silly-stories-word-add-in-sample-load-files-and-use-content-controls"></a>Exemplo de suplemento do Word Histórias engraçadas: carregar arquivos e usar controles de conteúdo

Este suplemento do Word mostrará como:

1. Carregar uma lista de arquivos docx a partir de um serviço e preencher um controle de caixa suspensa tendo os nomes de arquivo como opções.
2. Carregar um arquivo docx a partir do serviço e inseri-lo no documento do Word.
3. Carregar o conjunto do controle de conteúdo e criar caixas de entrada com base nos controles de conteúdo.
4. Atualizar o valor do texto do conjunto do controle de conteúdo com base nos valores das caixas de entrada.
5. Usar o Office UI Fabric para criar uma experiência de usuário do Word simplificada.

> Observação: Uma risada é um dos efeitos colaterais da execução deste exemplo.

## <a name="prerequisites"></a>Pré-requisitos

Para usar o Exemplo de suplemento do Word Histórias engraçadas, são necessários.

* [node](https://nodejs.org) para atender os arquivos docx.
* [npm](https://www.npmjs.com/) para instalar as dependências.
* JQuery, para o componente [suspenso](dev.office.com/fabric/components/dropdown) do Office UI Fabric.
* Word 2016 ou qualquer cliente que ofereça suporte à API Javascript do Word. Este exemplo faz uma verificação de requisitos para ver se ele está sendo executado em um host compatível.

## <a name="start-the-web-application"></a>Iniciar o aplicativo Web

1. Instale as dependências do projeto com o NPM (Gerenciador de Pacotes de Nós) executando ```npm install``` no diretório raiz do projeto, na linha de comando.
2. Inicie o servidor de desenvolvimento executando ```node server.js``` no diretório raiz do projeto. O suplemento será executado em 127.0.0.1:8080.

### <a name="configure-and-run-on-word-for-mac-2016"></a>Configurar e executar no Word para Mac 2016

1. Crie uma pasta chamada "wef" em Users/Library/Containers/com.microsoft.word/Data/Documents/
2. Coloque o manifesto na pasta wef (Users/Library/Containers/com.microsoft.word/Data/Documents/wef)
3. Abra o Word 2016 no Mac e clique na guia Inserir > menu suspenso Meus Suplementos. Você deverá ver o suplemento listado no menu suspenso. Selecione-o e ele carregará o suplemento.

### <a name="configure-and-run-on-word-for-windows-2016"></a>Configurar e executar no Word para Windows 2016

1. Crie um compartilhamento de rede ou [compartilhe uma pasta para a rede](https://technet.microsoft.com/pt-br/library/cc770880.aspx) e coloque o arquivo de manifesto [word-add-in-sillystories.xml](word-add-in-sillystories.xml) nele. Nesse momento, você implantou seu suplemento. Agora, você precisa informar ao Word onde encontrar o suplemento.
2. Inicie o Word e abra um documento.
3. Escolha a guia **Arquivo** e escolha **Opções**.
4. Escolha **Central de Confiabilidade**, e escolha o botão **Configurações da Central de Confiabilidade**.
5. Escolha **Catálogos de Suplementos Confiáveis**.
6. Na caixa de diálogo **URL do Catálogo**, digite o caminho de rede para o compartilhamento de pasta que contém word-add-in-sillystories.xml e, em seguida, escolha **Adicionar Catálogo**.
7. Selecione a caixa de seleção **Mostrar no Menu** e, em seguida, escolha **OK**.
8. Será exibida uma mensagem para informá-lo de que suas configurações serão aplicadas na próxima vez que você iniciar o Office. Feche e reinicie o Word. 

Agora você está pronto para executá-lo no Word. 

1. Abra um documento do Word. 
2. Na guia **Inserir** no Word 2016, escolha **Meus Suplementos**. 
3. Selecione a guia **Pasta compartilhada**.
4. Escolha o **suplemento Histórias engraçadas** e selecione **Inserir**.
5. Um novo grupo chamado **Suplemento do Word** aparecerá na guia **Página Inicial**. O grupo tem um botão chamado **Histórias engraçadas**. (Não aparecem na captura de tela.) Clique neste botão para abrir o painel de tarefas do suplemento.
6. Selecione uma história para que o texto clichê seja inserido no documento do Word.

__Figura 1. O suplemento Histórias engraçadas carregado no Word__

![Imagem do aplicativo Word com o suplemento Histórias engraçadas carregado](../readme-images/sillystoriesUI.PNG)

## <a name="questions-and-comments"></a>Perguntas e comentários

Gostaríamos de saber sua opinião sobre o exemplo de suplemento do Word Histórias engraçadas. Você pode enviar perguntas e sugestões na seção [Problemas](https://github.com/OfficeDev/Word-Add-in-SIllyStories/issues) deste repositório.

As perguntas sobre o desenvolvimento de suplementos em geral devem ser postadas no [Stack Overflow](http://stackoverflow.com/questions/tagged/Office365+API). Não deixe de marcar as perguntas ou comentários com [office-js], [word] e [API].

## <a name="learn-more"></a>Learn more

Veja mais recursos para ajudá-lo a criar suplementos baseados na API Javascript do Word:

* [Realizar sideload de um suplemento do Office no iPad e no Mac](http://dev.office.com/docs/add-ins/testing/sideload-an-office-add-in-on-ipad-and-mac)
* 
  [Visão geral da plataforma de Suplementos do Office](https://msdn.microsoft.com/pt-br/library/office/jj220082.aspx)
* [Suplementos do Word](https://github.com/OfficeDev/office-js-docs/blob/master/word/word-add-ins.md)
* [Visão geral da programação de suplementos do Word](https://github.com/OfficeDev/office-js-docs/blob/master/word/word-add-ins-programming-guide.md)
* [Explorador de trecho para Word](http://officesnippetexplorer.azurewebsites.net/#/snippets/word)
* [Referência da API JavaScript para suplementos do Word](https://github.com/OfficeDev/office-js-docs/tree/master/word/word-add-ins-javascript-reference)

## <a name="copyright"></a>Direitos autorais
Copyright © 2015 Microsoft. Todos os direitos reservados.
