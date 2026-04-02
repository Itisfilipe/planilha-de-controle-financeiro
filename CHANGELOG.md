# Changelog

## [1.0.0] — 2026-04-02

Primeira versão estável.

### Funcionalidades
- **Abas mensais** com resumo automático de entradas e saídas por categoria
- **Log de transações** com dropdown de categorias, seletor de data e flag "Parcela?"
- **Aba Dívidas** para acompanhar parcelas e financiamentos (status Ativa/Quitada automático)
- **Aba "Como usar"** com instruções completas (criada automaticamente no primeiro uso)
- **Atualizar categorias** reconstrói o resumo sem perder dados do log (com migração automática se o layout mudou)
- **Proteção de células** com fundo cinza e aviso ao editar fórmulas
- **Locale pt_BR** configurado automaticamente (datas, moeda, separador decimal)

### Segurança
- Abas mensais nunca são sobrescritas — se já existem, o script navega até elas
- Confirmação obrigatória antes de ações destrutivas
- Dados do log preservados ao atualizar categorias
