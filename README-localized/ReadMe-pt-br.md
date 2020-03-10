---
languages:
- javascript
page_type: sample
description: "The sample implements an Outlook add-in that adds buttons to the Outlook ribbon."
products:
- office
- office-outlook
urlFragment: outlook-ribbon-addin
---

# Exemplo de AttachmentsDemo de suplemento do Outlook

O exemplo implementa um suplemento do Outlook que adiciona botões à Faixa de Opções do Outlook. Ele permite ao usuário salvar todos os anexos no OneDrive. O exemplo ilustra os seguintes conceitos:
 
- Adicionar [botões de comando do suplemento](https://docs.microsoft.com/outlook/add-ins/add-in-commands-for-outlook) à faixa de opções do Outlook ao ler e-mails, incluindo um botão sem IU e um botão que abre um painel de tarefas
- Implementar uma WebAPI para [recuperar anexos por meio de um token de retorno de chamada e da API do Outlook REST](https://dev.office.com/docs/add-ins/outlook/use-rest-api)
- [Usar o token de acesso de SSO](https://docs.microsoft.com/outlook/add-ins/authenticate-a-user-with-an-sso-token) para chamar a API do Microsoft Graph sem perguntar ao usuário
- Se o token de SSO não estiver disponível, autenticar o OneDrive do usuário usando o fluxo implícito OAuth2 por meio da [biblioteca office-js-auxiliares](https://github.com/OfficeDev/office-js-helpers).
- Usar o [API do Microsoft Graph](https://developer.microsoft.com/graph/docs/api-reference/v1.0/resources/onedrive) para criar arquivos no OneDrive.

## Configurar o Exemplo

Antes de executar o exemplo, você precisará fazer algumas coisas para fazê-lo funcionar corretamente.

1. Você deve ter um locatário do Office 365 ou uma conta do Outlook.com. Embora os aplicativos de e-mail funcionem com instalações locais do Exchange, a API do Microsoft Graph exige o Office 365 ou Outlook.com.
2. É necessário registrar o aplicativo de exemplo no [Portal de registro de aplicativo da Microsoft](https://apps.dev.microsoft.com) para obter uma ID de aplicativo para acessar a API do Microsoft Graph.
    1. Vá até o [Portal de Registro de Aplicativos da Microsoft](https://apps.dev.microsoft.com). Caso não seja solicitado a entrada, clique no botão **Acessar lista de aplicativos** e entre com uma conta da Microsoft (Outlook.com) ou com a sua conta corporativa ou de estudante (Office 365). Depois de entrar, clique no botão **Adicionar um aplicativo**. Insira `AttachmentsDemo` para o nome e clique em **Criar aplicativo**.
    1. Localize a seção **Segredos do Aplicativo** e clique no botão **Gerar Nova Senha**. Uma caixa de diálogo será exibida com a senha gerada. Copie esse valor antes de descartar a caixa de diálogo e salve-o.
    1. Localize a seção **Plataformas** e clique em **Adicionar plataforma**. Escolha **Web**e, em seguida, digite `http://localhost:44349/MessageRead.html` em **URIs de redirecionamento**.
        > **Observação:** O número da porta na URI de redirecionamento (`44349`) pode ser diferente na máquina de desenvolvimento. Você pode encontrar o número da porta correto para o seu computador selecionando o projeto AttachmentDemoWeb no **Gerenciador de soluções** observando  a configuração **SSL URL** no**Servidor de Desenvolvimento** na janela Propriedades.
        
    1. Clique em **Adicionar Plataforma**. Escolha **Web API**. Configure esta seção da seguinte maneira:
        - Em **URI da ID do aplicativo**, altere o valor padrão, inserindo o seu host e o número da porta antes da GUID listada. Por exemplo, se o valor padrão é `api://05adb30e-50fa-4ae2-9cec-eab2cd6095b0`e seu aplicativo está sendo executado em `localhost: 44349`, o valor é `api://localhost:44349/05adb30e-50fa-4ae2-9cec-eab2cd6095b0`.
        - Em **Aplicativos pré-qualificados**, digite `d3590ed6-52B3-4102-AEFF-aad2292ab01c` para a **ID do aplicativo**. Clique em **Escopo** no menu suspenso e selecione a única entrada ali. Isso pré-autoriza a área de trabalho do Office (no Windows) a acessar o aplicativo.
        - Em **Aplicativos pré-qualificados**, digite `bc59ab01-8403-45c6-8796-ac3ef710b3e3` para a **ID do aplicativo**. Clique em **Escopo** no menu suspenso e selecione a única entrada ali. Isso pré-autoriza o Outlook na Web a acessar o aplicativo.
    1. Localize a seção **Permissões do Microsoft Graph** no registro do aplicativo. Ao lado de **Permissões delegadas**, clique em **Adicionar**. Selecione **Files.ReadWrite**, **Mail.Read**, **offline\_access**, **openid**e **perfil**. Clique em **OK**.

Clique em **Salvar** para concluir o registro. Copie a **ID de aplicativo** e salve-a no mesmo lugar com a senha de aplicativo que você salvou anteriormente. Precisaremos desses valores em breve.

Aqui está a aparência dos detalhes do registro do seu aplicativo quando você terminar.

![O registro de aplicativo concluído](readme-images/app-registration.PNG)
![O registro de aplicativo concluído parte 2](readme-images/web-api-app-registration.PNG)

Substitua [authconfig.js](AttachmentDemoWeb/Scripts/authconfig.js) pelo valor `ID GERADA DO APLICATIVO AQUI` com a ID do aplicativo gerada como parte do processo de registro do aplicativo.

Substitua [AttachmentDemo.xml](AttachmentDemo/AttachmentDemoManifest/AttachmentDemo.xml) pelo valor `ID GERADA DO APLICATIVO AQUI` com a ID do aplicativo gerada como parte do processo de registro do aplicativo.

> **Observação**: Certifique-se de que o número da porta no elemento `Recurso` corresponda à porta usada pelo seu projeto. Ele também deve corresponder à porta usada durante o registro do aplicativo.

Substitua [Web.config](AttachmentDemoWeb/Web.config) pelo valor `ID GERADA DO APLICATIVO AQUI` com a ID do aplicativo e `SENHA DO APLICATIVO AQUI` com a senha gerada do aplicativo como parte do processo de registro do aplicativo.

## Fornecer consentimento do usuário para o aplicativo

Nesta etapa forneceremos consentimento do usuário para as permissões que acabamos de configurar no aplicativo. Esta etapa **só** é necessária porque carregaremos o suplemento para desenvolvimento e teste. Normalmente, um suplemento de produção será listado na Office Store e os usuários serão instruídos a dar consentimento durante o processo de instalação da loja.

Você tem duas opções para fornecer o consentimento. Você pode usar uma conta de administrador e consentir uma única vez para todos os usuários em sua organização do Office 365 ou você possam usar uma conta qualquer para consentir apenas para um determinado usuário.

### Oferecer consentimento de administrador para todos os usuários

Se você tiver acesso a uma conta de administrador do locatário, este método permitirá dar consentimento a todos os usuários em sua organização. Isso pode ser conveniente se você tiver vários desenvolvedores que precisam desenvolver e testar o suplemento.

1. Vá para `https://login.microsoftonline.com/common/adminconsent?client_id={application_ID}&state=12345`, onde `{application_ID}` é a ID do aplicativo mostrada no seu registro de aplicativo.
1. Entre com sua conta de administrador.1. Analise as permissões e clique em **Aceitar**.

O navegador tentará redirecionar para seu aplicativo, que pode não estar em execução. É provável que você veja um erro "este site não pode ser acessado" depois de clicar em **Aceitar**. Não há problema, ainda assim o consentimento foi gravado.

### Oferecer consentimento para um único usuário

Se você não tiver acesso a uma conta de administrador de locatários ou quiser apenas limitar o consentimento a alguns usuários, com esse método é possível oferecer consentimento para um único usuário.

1. Vá para `https://login.microsoftonline.com/common/oauth2/authorize?client_id={application_ID}&state=12345&response_type=code`, onde `{application_ID}` é a ID do aplicativo mostrada no seu registro de aplicativo.
1. Entre com sua conta.
1. Analise as permissões e clique em **Aceitar**.

O navegador tentará redirecionar para seu aplicativo, que pode não estar em execução. É provável que você veja um erro "este site não pode ser acessado" depois de clicar em **Aceitar**. Não há problema, ainda assim o consentimento foi gravado.

## Execução do Exemplo

> **Observação**: O Visual Studio pode mostrar um aviso ou erro sobre o elemento `WebApplicationInfo` ser inválido. O erro pode não aparecer até você tentar criar a solução. Até o momento deste artigo, o Visual Studio não atualizou os arquivos de esquema para incluir o elemento `WebApplicationInfo`. Para solucionar esse problema, você pode usar o arquivo de esquema atualizado neste repositório: [MailAppVersionOverridesV1\_1.xsd](manifest-schema-fix/MailAppVersionOverridesV1_1.xsd).
>
> 1. Em sua máquina de desenvolvimento, localize o MailAppVersionOverridesV1\_1 existente. Ele deve estar localizado no diretório de instalação do Visual Studio em `./Xml/Schemas/{lcid}`. Por exemplo, em uma instalação típica do VS 2017 32 bits em um sistema em inglês (EUA), o caminho completo seria `C:\Arquivos de programas (x86) \Microsoft Visual Studio\2017\Enterprise\Xml\Schemas\1033`.
> 1. Renomeie o arquivo existente para `MailAppVersionOverridesV1_1.old`.
> 1. Mova a versão do arquivo deste repositório para a pasta.

Você pode executar o exemplo diretamente do Visual Studio. Selecione o projeto **AttachmentDemo** no **Gerenciador de Soluções**, em seguida, escolha o valor **Iniciar ação** (na janela de propriedades do**suplemento**). Você pode escolher qualquer navegador instalado para iniciar o Outlook na Web ou pode escolher **Cliente para área de trabalho do Office** para iniciar o Outlook. Se você escolher **Cliente para área de trabalho do Office**, certifique-se de configurar o Outlook para se conectar ao usuário do Office 365 ou Outlook.com no qual você deseja instalar o suplemento.

> **Observação:** O recurso de token de SSO já está na visualização do Outlook 2016 para Windows no momento.

Pressione **F5** para criar e depurar o projeto. Você deverá solicitar uma conta de usuário e uma senha. Use um usuário em seu locatário do Office 365 ou uma conta do Outlook.com. O suplemento será instalado para esse usuário e o Outlook na Web ou o Outlook serão abertos. Selecione qualquer mensagem, e você verá os botões de suplemento na faixa de opções do Outlook.

**Suplemento no Outlook na área de trabalho**

![Os botões de suplemento da faixa de opções do Outlook na área de trabalho](readme-images/buttons-outlook.PNG)

**Suplemento no Outlook na Web**

![Realizar o sideload de um suplemento do Outlook na Web](readme-images/buttons-owa.PNG)

## Direitos autorais

Copyright (c) Microsoft. Todos os direitos reservados.
