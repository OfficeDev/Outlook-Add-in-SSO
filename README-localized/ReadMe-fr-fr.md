---
languages:
- javascript
page_type: sample
description: "L’exemple implémente un complément Outlook qui ajoute des boutons au ruban Outlook."
products:
- office
- office-outlook
urlFragment: outlook-ribbon-addin
---

# Exemple de Complément Outlook AttachmentsDemo

L’exemple implémente un complément Outlook qui ajoute des boutons au ruban Outlook. L’utilisateur peut ainsi enregistrer toutes les pièces jointes sur son espace OneDrive. L'exemple illustre les concepts suivants :
 
- Ajout de [boutons complémentaires de commande](https://docs.microsoft.com/outlook/add-ins/add-in-commands-for-outlook) au ruban Outlook lors de la lecture d’un message, comprenant un bouton sans interface utilisateur et un bouton pour ouvrir un volet des tâches
- Implémentation d’une WebAPI [pour récupérer des pièces jointes à l’aide d’un jeton de rappel et de l’API REST Outlook](https://dev.office.com/docs/add-ins/outlook/use-rest-api)
- [Utilisation du jeton d’accès SSO](https://docs.microsoft.com/outlook/add-ins/authenticate-a-user-with-an-sso-token) pour appeler l’API Microsoft Graph sans demander confirmation à l’utilisateur
- Si le jeton SSO n’est pas disponible, authentification auprès de l’espace OneDrive de l’utilisateur à l’aide du flux implicite Oauth2 de la [bibliothèque office-js-helpers](https://github.com/OfficeDev/office-js-helpers).
- Utilisation de l’[API Microsoft Graph](https://developer.microsoft.com/graph/docs/api-reference/v1.0/resources/onedrive) pour créer des fichiers dans OneDrive.

## Configuration de l’exemple

Avant d’exécuter l’exemple, vous devez effectuer quelques opérations pour qu’il fonctionne correctement :

1. Disposer d’un compte Office 365 client ou Outlook.com. Bien que les applications de messagerie fonctionnent avec les installations sur site d’Exchange, l’API Microsoft Graph requiert Office 365 ou Outlook.com.
2. Enregistrer l’exemple d’application dans le [Portail d’inscription des applications Microsoft](https://apps.dev.microsoft.com) pour obtenir une ID d’application pour accéder à l’API Microsoft Graph.
    1. Accédez au [Portail d’inscription des applications Microsoft](https://apps.dev.microsoft.com). S’il ne vous est pas demandé de vous connecter, cliquez sur le bouton **Accéder à la liste des applications** et connectez-vous avec votre compte Microsoft (Outlook.com), ou votre compte scolaire ou professionnel (Office 365). Une fois connecté, cliquez sur le bouton **Ajouter une application**. Entrez `AttachmentsDemo` pour le nom et cliquez sur **Créer une application**.
    1. Cherchez la section **Secrets de l'application**, puis cliquez sur le bouton **Générer un nouveau mot de passe**. Une boîte de dialogue s’affiche avec le mot de passe généré. Copiez cette valeur avant de fermer la boîte de dialogue et gardez-le.
    1. Cherchez la section **Plateformes**, puis cliquez sur **Ajouter une plateforme**. Sélectionnez **Web**, puis entrez `http://localhost:44349/MessageRead.html` sous **URI Redirect**.
        > **Remarque :** Le numéro de port dans l’URI de redirection (`44349`) peut être différent sur votre ordinateur de développement. Vous pouvez trouver le numéro de port correct pour votre ordinateur en sélectionnant le projet AttachmentDemoWeb dans l’**Explorateur de solutions**, il s’agit du paramètre URL SSL sous **Serveur de développement** dans la fenêtre de propriétés.
        
    1. Cliquez sur **Ajouter une plateforme**. Choisissez **Web API**. Configurez cette section comme suit :
        - Sous **URI de l’ID d’application**, modifiez la valeur par défaut en insérant votre hôte et votre numéro de port avant le GUID qui y est répertorié. Par exemple, si la valeur par défaut est `api://05adb30e-50fa-4ae2-9cec-eab2cd6095b0`et que votre application est en cours d’exécution sur `localhost:44349`, la valeur doit être`api://localhost:44349/05adb30e-50fa-4ae2-9cec-eab2cd6095b0`.
        - Sous **Applications préalablement autorisées**, entrez d3590ed6-52b3-4102-aeff-aad2292ab01c comme ID d’application. Cliquez sur le menu déroulant **Étendue** puis sélectionnez la seule entrée à cet emplacement. Cela permet de pré-autoriser Bureau Office (sur Windows) à accéder à l’application.
        - Sous **Applications préalablement autorisées**, entrez `bc59ab01-8403-45c6-8796-ac3ef710b3e3` comme **ID d’application**. Cliquez sur le menu déroulant **Étendue** puis sélectionnez la seule entrée à cet emplacement. Cela permet de pré-autoriser Outlook sur le web à accéder à l’application.
    1. Cherchez la section **Autorisations pour Microsoft Graph** dans l’inscription de l’application. À côté d’**Autorisations déléguées**, cliquez sur **Ajouter**. Sélectionnez **Files.ReadWrite**, **Mail.Read**, **offline\_access**, **openid** et **profile**. Cliquez sur **OK**.

Cliquez sur **Enregistrer** pour terminer l’inscription. Copiez l’**ID d’application** et gardez-le au même endroit que le mot de passe d’application précédemment sauvegardé. Ces valeurs seront utiles ultérieurement.

Voici comment doivent se présenter les détails de votre inscription d’application une fois ces étapes effectuées.

![L’inscription de l’application complétée](readme-images/app-registration.PNG)
![L’inscription de l’application complétée (partie 2)](readme-images/web-api-app-registration.PNG)

Modifiez [authconfig.js](AttachmentDemoWeb/Scripts/authconfig.js) en remplaçant la valeur `VOTRE ID D’APPLICATION` par l’ID d’application généré lors du processus d’inscription d’application.

Modifiez [AttachmentDemo.xml](AttachmentDemo/AttachmentDemoManifest/AttachmentDemo.xml) en remplaçant la valeur `VOTRE ID D’APPLICATION` par l’ID d’application généré lors du processus d’inscription d’application.

> **Remarque** : Assurez-vous que le numéro de port dans l’élément `Ressource` correspond au port utilisé par votre projet. Il doit également correspondre au port utilisé lors de l’enregistrement de l’application.

Modifiez [Web.config](AttachmentDemoWeb/Web.config) en remplaçant la valeur `VOTRE ID D’APPLICATION` par l’ID d’application et `VOTRE MOT DE PASSE` par le mot de passe d’application généré lors du processus d’inscription d’application.

## Fournir l’accord de l’utilisateur à l’application

Au cours de cette étape, nous allons fournir l’accord de l’utilisateur pour les autorisations que nous venons de configurer pour l’application. Cette étape est **uniquement** nécessaire parce qu’il s’agit de charger latéralement le complément pour le développement et les tests. En règle générale, un complément de production est répertorié dans Office Store. Les utilisateurs sont invités à fournir leur consentement pendant le processus d’installation via le magasin.

Pour cela, deux possibilités vous sont offertes : vous pouvez utiliser un compte Administrateur pour donner une seule fois votre consentement à tous les utilisateurs de votre organisation Office 365, ou vous pouvez utiliser n’importe quel compte pour donner votre consentement à un seul utilisateur.

### Consentement administrateur à tous les utilisateurs

Si vous avez accès à un compte Administrateur client, cette méthode vous permet de donner votre consentement à tous les utilisateurs de votre organisation. Cette méthode peut s’avérer utile si plusieurs développeurs ont besoin de développer et de tester votre complément en même temps.

1. Accédez à `https://login.microsoftonline.com/common/adminconsent?client_id={application_ID}&state=12345`, où `{application_ID}` est l’ID d’application indiqué dans l’inscription de l’application.
1. Connectez-vous avec votre compte Administrateur. Consultez les autorisations et cliquez sur **Accepter**.

Le navigateur tentera de vous rediriger vers votre application, laquelle n’est peut-être pas en cours d’exécution. Le message d’erreur « Impossible d’accéder à ce site » peut s’afficher quand vous aurez cliqué sur **Accepter**. Ne vous inquiétez pas, le consentement a été enregistré.

### Consentement à un seul utilisateur

Si vous n’avez pas accès à un compte Administrateur client, ou si vous souhaitez simplement donner votre consentement à quelques utilisateurs seulement, cette méthode vous permettra de donner votre consentement à un seul utilisateur.

1. Accédez à `https://login.microsoftonline.com/common/oauth2/authorize?client_id={application_ID}&state=12345&response_type=code` où `{application_ID}` est l’ID d’application indiqué dans l’inscription de l’application.
1. Connectez-vous à votre compte.
1. Consultez les autorisations et cliquez sur **Accepter**.

Le navigateur tentera de vous rediriger vers votre application, laquelle n’est peut-être pas en cours d’exécution. Le message d’erreur « Impossible d’accéder à ce site » peut s’afficher quand vous aurez cliqué sur **Accepter**. Ne vous inquiétez pas, le consentement a été enregistré.

## Exécution de l’exemple

> **Remarque** : Visual Studio est susceptible d’afficher un avertissement ou une erreur indiquant la non-validité de l’élément`WebApplicationInfo`. Il se peut que l’erreur ne s’affiche pas avant la génération de la solution. À date de rédaction de ce texte, Visual Studio n’a pas mis à jour ses fichiers de schéma pour inclure l’élément `WebApplicationInfo`. Pour contourner ce problème, vous pouvez utiliser le fichier de schéma mis à jour dans ce référentiel : [MailAppVersionOverridesV1\_1.xsd](manifest-schema-fix/MailAppVersionOverridesV1_1.xsd).
>
> 1. Sur votre ordinateur de développement, recherchez l’élément MailAppVersionOverridesV1\_1.xsd existant. Il doit se trouver dans le répertoire d’installation Visual Studio sous `./Xml/Schemas/{lcid}`. Par exemple, sur une installation standard de VS 2017 32 bits sur un système anglais (États-Unis), le chemin d’accès complet est `C:\Program Files (x86) \Microsoft Visual Studio\2017\Enterprise\Xml\Schemas\1033`.
> 1. Renommez le fichier existant en `MailAppVersionOverridesV1_1. old`.
> 1. Déplacez la version du fichier de ce référentiel vers le dossier.

Vous pouvez exécuter l’exemple directement à partir de Visual Studio. Sélectionnez le projet AttachmentDemo dans l’Explorateur de solutions, puis sélectionnez la valeur **Action de démarrage** souhaitée (sous Complément dans la fenêtre de propriétés). Vous pouvez choisir n’importe quel navigateur pour démarrer Outlook sur le web ou bien **Client Office Bureau** pour démarrer Outlook. Si vous choisissez **Client Office Bureau**, assurez-vous de configurer Outlook pour le faire connecter au compte utilisateur Office 365 ou Outlook.com sur lequel vous souhaitez installer le complément.

> **Remarque :** La fonctionnalité de jeton SSO est pour le moment en mode aperçu dans Outlook 2016 pour Windows uniquement.

Appuyez sur **F5** pour créer et déboguer l’application. Une boîte de dialogue vous invite à entrer un compte d’utilisateur et un mot de passe. Assurez-vous d’utiliser un compte Office 365 client ou Outlook.com. Le complément est installé pour cet utilisateur, et Outlook ou Outlook sur le web s’ouvre. Sélectionnez un message pour afficher les boutons du complément dans le ruban Outlook.

**Complément dans Outlook sur le bureau**

![Boutons du complément sur le ruban dans Outlook sur le bureau](readme-images/buttons-outlook.PNG)

**Complément dans Outlook sur le web**

![Boutons du complément dans Outlook sur le web](readme-images/buttons-owa.PNG)

## Copyright

Copyright (c) Microsoft. Tous droits réservés.
