# TarjarDoc MRD

**Ocultador de PDFs Assinados** - Uma ferramenta desenvolvida pela **GEI (Gestão da Informação)** da **AEDAS** para o **Projeto MRD (Médio Rio Doce)**.

## 📌 Sobre o Projeto

Muitos softwares convencionais (como o Adobe Acrobat) bloqueiam a edição ou a aplicação de ferramentas de censura (tarja) em documentos que possuem assinatura digital ou certificação (ex: Gov.br). 

O **TarjarDoc MRD** foi criado para preencher essa lacuna institucional. Ele permite que os usuários apliquem tarjas visuais em documentos sensíveis de forma geométrica, achatando (flatten) a estrutura do PDF e gerando um novo documento limpo. 

**Aviso Legal:** O uso desta ferramenta para mascarar informações de um documento previamente assinado digitalmente corrompe a integridade da assinatura. Portanto, os documentos exportados por este software perdem a sua validade jurídica original, servindo exclusivamente para fins de compartilhamento seguro de informações descaracterizadas.

## 🚀 Funcionalidades

- **Processamento 100% Local:** Sem backend. Os arquivos PDF nunca saem do computador do usuário, garantindo sigilo absoluto e conformidade com as diretrizes de Segurança da Informação.
- **Preservação de Layout:** Diferente de ferramentas de IA que extraem texto e destroem a formatação de PDFs híbridos/escaneados, esta ferramenta aplica a censura sobre a camada original do documento.
- **Auditoria Serverless:** Integração nativa com o ecossistema Microsoft 365. Utiliza um Webhook do **Power Automate** para salvar um log de uso (Nome do arquivo, Qtd de páginas, Qtd de Tarjas, Data/Hora) diretamente em uma lista do Microsoft Lists/SharePoint.
- **Integração com Microsoft Teams:** O software foi arquitetado para ser embarcado (via manifest) como um aplicativo pessoal dentro do MS Teams da corporação, capturando via SDK (SSO invisível) a identidade do usuário para o log de auditoria.

## 🛠️ Tecnologias Utilizadas

- **HTML5, CSS3, JavaScript (Vanilla)** - Sem necessidade de Node.js ou empacotadores (bundlers).
- **[PDF.js](https://mozilla.github.io/pdf.js/)** (Mozilla) - Para renderização e visualização nativa das páginas do PDF no navegador.
- **[PDF-lib](https://pdf-lib.js.org/)** - Para a edição, achatamento (flattening) das assinaturas e inserção geométrica vetorial das tarjas pretas no binário do documento exportado.
- **Microsoft Teams JavaScript SDK** - Para extração de contexto de identidade (Nome/Email) do usuário ativo no Teams.

## ⚙️ Configuração do Webhook (Auditoria)

Para que o painel de auditoria funcione, é necessário configurar um fluxo no **Power Automate**:

1. Crie um gatilho **"Quando uma solicitação HTTP é recebida"** configurado para *Alguém* (Anyone).
2. Utilize o seguinte esquema JSON no gatilho:
```json
{
    "type": "object",
    "properties": {
        "usuarioResponsavel": { "type": "string" },
        "emailUsuario": { "type": "string" },
        "nomeDoArquivo": { "type": "string" },
        "numeroDePaginas": { "type": "integer" },
        "quantidadeDeTarjas": { "type": "integer" },
        "dataHora": { "type": "string" }
    }
}
```
3. Conecte o passo seguinte ao **SharePoint (Criar item)** ou Excel Online.
4. Cole a URL gerada pelo Power Automate na variável `WEBHOOK_URL` dentro do arquivo `script.js`.

---
*Desenvolvido internamente para uso corporativo - GEI / AEDAS 2026*
