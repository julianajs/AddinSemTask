# Deploy Azure Web App - Versão Estática

Esta versão não usa npm, webpack ou build no Azure.

## Como publicar

1. No Azure App Service, vá em Deployment Center ou Kudu.
2. Faça ZIP Deploy deste pacote.
3. Em Configuration > General settings:
   - Startup Command: deixe vazio.
4. Acesse a URL raiz do Web App.

## Importante

Depois de publicar, atualize o `manifest.xml` substituindo as URLs locais pela URL do Web App, por exemplo:

https://SEU-APP.azurewebsites.net/src/commands/commands.html
https://SEU-APP.azurewebsites.net/src/assets/icon-16.png
https://SEU-APP.azurewebsites.net/src/assets/icon-32.png
https://SEU-APP.azurewebsites.net/src/assets/icon-80.png
