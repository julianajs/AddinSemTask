# Outlook Add-in - Monitorar Reunião

Este projeto cria um add-in do Outlook com o botão **Monitorar** na faixa de opções durante a edição de uma reunião.

Ao clicar no botão, o add-in captura dados da reunião e envia para um fluxo do Power Automate sem abrir barra lateral.

## Estrutura

```text
outlook-monitorar-addin/
├─ manifest.xml
├─ package.json
├─ webpack.config.js
├─ src/
│  ├─ commands/
│  │  ├─ commands.html
│  │  └─ commands.js
│  └─ assets/
│     ├─ icon-16.png
│     ├─ icon-32.png
│     └─ icon-80.png
└─ README.md
```

## UX/UI

- Botão: **Monitorar**
- Grupo: **Automação**
- Tooltip: **Envia os dados desta reunião para o fluxo de monitoramento.**
- Feedback no Outlook:
  - Enviando reunião para monitoramento...
  - Reunião enviada para monitoramento.
  - Não foi possível enviar a reunião para monitoramento. Tente novamente.

## Instalação local

```bash
npm install
npm run validate
npm start
```

## Produção

Antes de publicar em produção, altere os domínios no `manifest.xml` para o domínio HTTPS real onde o add-in será hospedado.

Atualmente o projeto usa:

```text
https://localhost:3000
```

## Segurança

O endpoint do Power Automate está no arquivo `src/commands/commands.js`.

Para produção, recomenda-se usar uma camada intermediária:

```text
Outlook Add-in → Azure Function / API Management → Power Automate
```

Isso evita exposição direta da URL do fluxo no JavaScript do cliente.
