# Upadsfast — README

## O que é?

**Upadsfast** é um sistema de documentação, skills e workflows para gerar anúncios do **Google Ads (Rede de Pesquisa)** de forma rápida e automatizada.

O agente recebe uma **URL ou texto de página de oferta**, extrai automaticamente todas as variáveis da oferta, e gera o pacote completo de anúncios e extensões prontos para **bulk upload** no Google Ads.

---

## Estrutura de Pastas

```
Upadsfast/
├── SKILL.md                          # Skill principal (regras, protocolo, escopo)
├── README.md                         # Este arquivo
│
├── workflows/
│   ├── generate-ads.md               # Workflow: gerar anúncios a partir de URL/texto
│   └── bulk-upload-scripts.md        # Workflow: automação via Google Ads Scripts
│
├── templates/
│   ├── column-headers.md             # Cabeçalhos exatos de cada aba
│   └── formulas.md                   # Fórmulas para Google Sheets
│
├── examples/
│   └── capnos-us.md                  # Exemplo real completo (CAPNOS - US)
│
├── resources/
│   ├── copy-rules.md                 # Regras de copywriting (foco em oferta)
│   └── limits-reference.md           # Referência rápida de limites do Google Ads
│
└── scripts/
    └── bulk-upload-templates.md      # Templates de scripts genéricos
```

---

## Como Usar

### 1. Gerar Anúncios (Modo Rápido)

1. Cole uma URL ou texto de página de oferta
2. Informe: idioma, país, frete (sim/não)
3. O agente gera automaticamente as 5 abas de bulk upload

### 2. Montar Planilha Matriz

1. Crie uma planilha no Google Sheets
2. Use os cabeçalhos de `templates/column-headers.md`
3. Use as fórmulas de `templates/formulas.md` para automatizar
4. Cole os dados gerados pelo agente

### 3. Automatizar com Scripts

1. Suba a planilha para o Google Drive
2. Use os scripts de `scripts/bulk-upload-templates.md`
3. Execute primeiro com `.preview()`, depois com `.apply()`

---

## Regras Fundamentais

- **Foco 100% em oferta direta** (desconto, frete, garantia, preço, urgência)
- **Nunca falar de benefícios do produto**
- **Sempre respeitar limites de caracteres** (30 headlines, 90 descriptions, etc.)
- **Sempre gerar as 5 abas completas** (RSA, Callouts, Sitelinks, Snippets, Promoção)
- **Sempre validar antes de entregar**

---

## Fontes Oficiais

| Fonte | URL |
|-------|-----|
| Templates de bulk upload | <https://support.google.com/google-ads/answer/10702525> |
| Introdução ao bulk upload | <https://support.google.com/google-ads/answer/10702932> |
| Upload de planilhas | <https://support.google.com/google-ads/answer/10702433> |
| Google Ads Scripts (bulk upload) | <https://developers.google.com/google-ads/scripts/docs/concepts/bulk-upload> |
| Limites de conta | <https://support.google.com/google-ads/answer/6372658> |
