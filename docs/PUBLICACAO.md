# Política desta edição pública

Este repositório é uma edição de portfólio separada da instalação operacional. A exportação usa uma lista permitida; não copia a árvore privada nem seu histórico.

## Conteúdo publicado

- Código e testes com dados fictícios.
- Configurações de exemplo sem vínculos reais.
- Geradores de fixtures e workflows.
- Documentação técnica escrita para o portfólio.

## Conteúdo excluído

- Credenciais, certificados, arquivos de ambiente preenchidos e configurações locais.
- Documentos recebidos, modelos institucionais, imagens e minutas reais.
- Cadastros privados de campos e substituições de anonimização.
- Respostas reais da IA, logs, relatórios, recibos e artefatos de execução.
- IDs e links de tarefas, equipes e workspaces usados no desenvolvimento.

Arquivos gerados em `storage/` são ignorados. A demonstração não consulta credenciais e só permite transporte de rede para o simulador local. Os modelos de configuração mantêm as chamadas pagas e os agendamentos bloqueados.

Execute `npm run check` antes de publicar alterações. O verificador é heurístico: nunca substitui a conferência dos arquivos adicionados e do histórico Git. O GitHub Actions não recebe segredos da instalação.
