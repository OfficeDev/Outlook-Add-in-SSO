---
languages:
- javascript
page_type: sample
description: "El ejemplo es sobre implementar un complemento de Outlook que agrega botones a la cinta de Outlook."
products:
- office
- office-outlook
urlFragment: outlook-ribbon-addin
---

# Complemento de Outlook de ejemplo AttachmentsDemo

El ejemplo es sobre implementar un complemento de Outlook que agrega botones a la cinta de Outlook. Permite que el usuario guarde todos los datos adjuntos en su OneDrive. El ejemplo ilustra los siguientes conceptos:
 
- Agregar [botones de comando de complemento](https://docs.microsoft.com/outlook/add-ins/add-in-commands-for-outlook) a la cinta de opciones de Outlook al leer correo, incluido un botón sin interfaz de usuario y un botón que abre un panel de tareas.
- Implementación de un WebAPI para [recuperar los datos adjuntos a través de un token de devolución de llamada y la API de REST de Outlook](https://dev.office.com/docs/add-ins/outlook/use-rest-api)
- [Usar el token de acceso SSO](https://docs.microsoft.com/outlook/add-ins/authenticate-a-user-with-an-sso-token) para llamar a la API de Microsoft Graph sin tener que preguntar al usuario
- Si el token de SSO no está disponible, autenticándose en OneDrive del usuario utilizando el flujo implícito OAuth2 a través de la [biblioteca office-js-helpers](https://github.com/OfficeDev/office-js-helpers).
- Usar la [API de Microsoft Graph](https://developer.microsoft.com/graph/docs/api-reference/v1.0/resources/onedrive) para crear archivos en OneDrive.

## Configurar el ejemplo

Antes de ejecutar el ejemplo, tendrá que hacer algunas cosas para que funcione correctamente.

1. Necesita un espacio empresarial de Office 365 o una cuenta Outlook.com. Si bien las aplicaciones de correo funcionarán con instalaciones locales de Exchange, la API de Microsoft Graph requiere Office 365 o Outlook.com.
2. Debe registrar la aplicación de ejemplo en el [Portal de registro de aplicaciones de Microsoft](https://apps.dev.microsoft.com) para obtener la Id. de la aplicación para tener acceso a la API de Microsoft Graph.
    1. Vaya al [Portal de registro de aplicaciones de Microsoft](https://apps.dev.microsoft.com). Si no se le pide que inicie sesión, haga clic en el botón **Ir a la lista de aplicaciones** e inicie sesión con su cuenta de Microsoft (Outlook.com) o su cuenta profesional o educativa (Office 365). Una vez que haya iniciado sesión, haga clic en el botón **Agregar una aplicación**. Escriba `AttachmentsDemo` para el nombre y haga clic en **Crear aplicación**.
    1. En la sección **Secretos de aplicación**, haga clic en **Generar nueva contraseña**. Verá un cuadro de diálogo con la contraseña generada. Copie este valor antes de cerrar el cuadro de diálogo y guárdelo.
    1. En la sección **Plataformas**, haga clic en **Agregar plataforma**. Elija **Web** y, a continuación, escriba `https://localhost:44349/MessageRead.html` en **redirigir URI**.
        > **Nota:** El número de puerto en el URI de redireccionamiento (`44349`) puede ser diferente en su equipo de desarrollo. Para buscar el número de puerto correcto para su equipo, seleccione el proyecto **AttachmentDemoWeb** en el **Explorador de soluciones** y, a continuación, busque la **URL de SSL** en la ventana de propiedades **Servidor de desarrollo**.
        
    1. Haga clic en **Agregar plataforma**. Elija **API web**. Configure esta sección de la siguiente manera:
        - En **URI de Id. de aplicación**, cambie el valor predeterminado insertando el host y número de puerto antes del GUID que aparece allí. Por ejemplo, si el valor predeterminado es `api://05adb30e-50fa-4ae2-9cec-eab2cd6095b0`, y la aplicación se está ejecutando en `localhost:44349`, el valor es `api://localhost:44349/05adb30e-50fa-4ae2-9cec-eab2cd6095b0`.
        - En **Aplicaciones preautorizadas**, escriba `d3590ed6-52b3-4102-aeff-aad2292ab01c` para la **Id. de aplicación**. Haga clic en la lista desplegable **Ámbito** y seleccione la única entrada allí. Esto preautoriza al escritorio de Office (en Windows) a acceder a la aplicación.
        - En **Aplicaciones preautorizadas**, escriba `bc59ab01-8403-45c6-8796-ac3ef710b3e3` para la **Id. de aplicación**. Haga clic en la lista desplegable **Ámbito** y seleccione la única entrada allí. Esto preautoriza a Outlook en la Web a acceder a la aplicación.
    1. Busque la sección **permisos de Microsoft Graph** en el registro de aplicación. Junto a **Permisos delegados**, haga clic en **Agregar**. Seleccione **Files.ReadWrite**, **Mail.Read**, **offline\_access**, **OpenID** y **perfil**. Haga clic en **Aceptar**.

Haga clic en **Guardar** para completar el registro. Copie la **Id. de aplicación** y guárdela en el mismo lugar con la contraseña de la aplicación que guardó anteriormente. En breve necesitaremos estos valores.

Cuando termine, este es el aspecto que deberían tener los detalles del registro de aplicación.

![El registro de aplicación finalizado](readme-images/app-registration.PNG)
![El registro de aplicación finalizado parte 2](readme-images/web-api-app-registration.PNG)

Edite [authconfig.js](AttachmentDemoWeb/Scripts/authconfig.js) y reemplace el valor de `la Id. de aplicación aquí` con la Id. de aplicación que generó como parte del proceso de registro de aplicación.

Edite [AttachmentDemo.xml](AttachmentDemo/AttachmentDemoManifest/AttachmentDemo.xml) y reemplace el valor de `la Id. de aplicación aquí` con la Id. de aplicación que generó como parte del proceso de registro de aplicación.

> **Nota**: Asegúrese de que el número de puerto en el elemento `Recurso` coincida con el puerto utilizado por su proyecto. También tiene que coincidir con el puerto que utilizó al registrar la aplicación.

Edite [Web.config](AttachmentDemoWeb/Web.config) y reemplace el valor de `la Id. de aplicación aquí` con la Id. de aplicación y `la contraseña de aplicación aquí` con la contraseña de la aplicación que generó como parte del proceso de registro de aplicación.

## Dar consentimiento al usuario para la aplicación

En este paso, daremos el consentimiento del usuario para los permisos que acabamos de configurar en la aplicación. Este paso **solo** es necesario, porque se cargará el complemento para el desarrollo y las pruebas. Normalmente, un complemento de producción se mostrará en la tienda de Office y se solicitará a los usuarios que den su consentimiento durante el proceso de instalación a través de la tienda.

Tiene dos opciones para proporcionar consentimiento. Puede usar una cuenta de administrador y dar consentimiento una vez para todos los usuarios de su organización de Office 365, o puede usar cualquier cuenta para dar su consentimiento solamente para ese usuario.

### Proporcionar permisos de administrador para todos los usuarios

Si tiene acceso a una cuenta de administrador de espacios empresariales, este método le permite proporcionar consentimiento para todos los usuarios de su organización. Puede ser útil si tiene varios desarrolladores que necesitan desarrollar y probar el complemento.

1. Vaya a `https://login.microsoftonline.com/common/adminconsent?client_id={application_ID}&state=12345`, donde `{application_ID}` es la Id. de aplicación que se muestra en el registro de aplicación.
1. Inicie sesión con su cuenta de administrador. 1. Revise los permisos y haga clic en **Aceptar**.

El explorador intentará redirigirle a la aplicación, que puede no estar ejecutándose. Es posible que vea el error "no se puede localizar este sitio" al hacer clic en **Aceptar**. Esto está bien, el consentimiento queda registrado.

### Proporcionar consentimiento para un único usuario

Si no tiene acceso a una cuenta de administrador de espacios empresariales o desea limitar permisos a unos pocos usuarios, este método le permitirá proporcionar consentimiento a un solo usuario.

1. Vaya a `https://login.microsoftonline.com/common/oauth2/authorize?client_id={application_ID}&state=12345&response_type=code`, donde `{application_ID}` es la Id. de aplicación que se muestra en el registro de aplicación.
1. Inicie sesión con su cuenta.
1. Revise los permisos y haga clic en **Aceptar**.

El explorador intentará redirigirle a la aplicación, que puede no estar ejecutándose. Es posible que vea el error "no se puede localizar este sitio" al hacer clic en **Aceptar**. Esto está bien, el consentimiento queda registrado.

## Ejecutar el ejemplo

> **Nota**: Es posible que Visual Studio muestre una advertencia o un error acerca de que el elemento `WebApplicationInfo` no es válido. Es posible que el error no se muestre hasta que intente crear la solución. Al momento de escribir esto, Visual Studio no ha actualizado sus archivos de esquema para incluir el elemento `WebApplicationInfo`. Para solucionar este problema, puede usar el archivo de esquema actualizado en este repositorio: [MailAppVersionOverridesV1\_1.xsd](manifest-schema-fix/MailAppVersionOverridesV1_1.xsd).
>
> 1. En el equipo de desarrollo, busque el MailAppVersionOverridesV1\_1.xsd existente. Debe estar en el directorio de instalación de Visual Studio en `./Xml/Schemas/{lcid}`. Por ejemplo, en una instalación típica de VS 2017 32 bits en un sistema inglés (Estados Unidos), la ruta de acceso completa sería `C:\Archivos de programa (x86) \Microsoft Visual Studio\2017\Enterprise\Xml\Schemas\1033`.
> 1. Cambie el nombre del archivo existente a `MailAppVersionOverridesV1_1.old`.
> 1. Mueva la versión del archivo de este repositorio a la carpeta.

Puede ejecutar el ejemplo directamente desde Visual Studio. Seleccione el proyecto **AttachmentDemo** en el **Explorador de soluciones**y, a continuación, elija el valor **Acción de inicio** que desee (en **complemento** en la ventana de propiedades). Puede elegir cualquier explorador instalado para iniciar Outlook en la Web o puede elegir **Cliente de escritorio de Office** para iniciar Outlook. Si elige **Cliente de escritorio de Office**, asegúrese de configurar Outlook para conectarse al usuario de Office 365 o Outlook.com para el que desea instalar el complemento.

> **Nota:** La característica de token de SSO se encuentra en versión preliminar en Outlook 2016 para Windows, por ahora.

Presione **F5** para crear y depurar el proyecto. Se le pedirá la cuenta de usuario y contraseña. Asegúrese de utilizar un usuario en el espacio empresarial de Office 365 o una cuenta de Outlook.com. El complemento se instalará para ese usuario y se abrirá Outlook en la Web o Outlook. Seleccione cualquier mensaje, debería ver los botones de complemento en la cinta de Outlook.

**Complemento en Outlook en el escritorio**

![Los botones de complemento en la cinta de opciones en Outlook en el escritorio](readme-images/buttons-outlook.PNG)

**Complemento en Outlook en la Web**

![Los botones de complemento en Outlook en la Web](readme-images/buttons-owa.PNG)

## Derechos de autor

Copyright (c) Microsoft. Todos los derechos reservados.
