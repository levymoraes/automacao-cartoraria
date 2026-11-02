# Estudo de caso: automação documental com revisão humana

## Contexto

Uma equipe recebe documentos digitalizados e precisa transferir informações para modelos Word, preservando a estrutura e identificando o que ainda depende de conferência. Os documentos variam: podem faltar dados, existir valores conflitantes ou haver mais participantes do que posições no modelo.

## Entrega implementada

Foi construído um fluxo autônomo com n8n, Linear e Node.js. A fila é consultada por horários configuráveis e na inicialização. Cada tarefa mantém o próprio modelo, documentos, leitura estruturada, resultado e recibos.

A IA propõe os dados e as evidências. A aplicação valida o contrato, verifica os vínculos e preenche campos em uma cópia do DOCX. O funcionário recebe uma minuta preliminar, com lacunas em vermelho e uma lista de pendências. A entrega só é concluída depois de confirmar o arquivo anexado e a mudança para In Review.

## Trabalho de engenharia demonstrado

- Integração de APIs e orquestração de tarefas com n8n.
- Contratos estruturados, JSON Schema e validação local.
- Manipulação de pacotes OOXML e preservação de formatação.
- Auditoria de origem, titularidade e finalidade das informações.
- Tratamento de falhas, checkpoints, idempotência e isolamento entre tarefas.
- Controle persistente de chamadas e orçamento.
- Testes de regressão com falhas de transporte, upload e retomada.
- Separação entre código público e operação com dados privados.

## Autonomia e limites

O sistema processa a fila e entrega resultados sem orientação humana durante cada execução, quando instalado e autorizado. Decisões documentais pendentes ficam explícitas para o funcionário; o sistema não transforma ausência de evidência em confirmação.

Novos tipos de documento podem reutilizar a infraestrutura. Cada novo modelo exige preparação dos campos, conferência e validação do comportamento esperado. Não há promessa de suporte automático a qualquer documento.

## Evidência pública

`npm test` executa as regressões com dados fictícios. `npm run demo` demonstra o caminho de entrega com a IA e o Linear simulados localmente. `npm run workflows:build` gera os workflows manual e agendado.

Os testes integrados do projeto privado incluíram seis modelos e entrega real para revisão. A publicação não reproduz documentos de terceiros nem atribui aprovação jurídica a esses testes.
