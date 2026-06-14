# WhatsApp Manager v1.0

Ferramenta CLI para extrair contatos de grupos do WhatsApp e disparar mensagens em massa com controle de intervalo.

---

## 📁 Arquivos

| Arquivo | Função |
|---|---|
| `main.py` | Programa principal |
| `instalar.bat` | Instala dependências (rodar 1x) |
| `iniciar.bat` | Inicia o programa |
| `gerar_exe.bat` | Gera o `.exe` standalone |

---

## 🚀 Como usar

### 1. Instalação (primeira vez)
1. Instale o Python: https://www.python.org/downloads/
   - ⚠️ Marque **"Add Python to PATH"** na instalação
2. Dê duplo clique em `instalar.bat`
3. Aguarde a instalação

### 2. Executar
- Dê duplo clique em `iniciar.bat`

---

## 📋 Módulos

### Módulo 1 — Extrator de Contatos
1. Abre o WhatsApp Web no Chrome
2. Você escaneia o QR Code
3. Abre o grupo desejado manualmente
4. Pressiona ENTER no programa
5. O programa extrai todos os números e salva em `.xlsx`

> ⚠️ Funciona melhor quando os contatos **não estão salvos na agenda** (números ficam visíveis).

### Módulo 2 — Disparador de Mensagens
1. Carrega a planilha `.xlsx`
2. Você configura 3 mensagens diferentes
3. Configura intervalos e tamanho dos lotes
4. O programa dispara sorteando as mensagens aleatoriamente

**Configurações disponíveis:**
- Intervalo mínimo e máximo entre mensagens
- Tamanho do lote (ex: 50 mensagens)
- Pausa entre lotes (ex: 10 minutos)

**Recursos:**
- ✅ Retoma de onde parou se interrompido
- ✅ Log completo em `log_envios.txt`
- ✅ Marca enviados/falhas na planilha

---

## ⚠️ Avisos importantes

- Use um número com histórico antigo no WhatsApp
- Não exagere nos disparos — respeite os intervalos configurados
- O WhatsApp pode banir números que fazem disparos em massa
- Este programa usa o WhatsApp Web (não oficial) — use por sua conta e risco

---

## 📊 Formato da planilha Excel

| numero | enviado | erro |
|---|---|---|
| +5511999999999 | NAO | |
| +5521988888888 | SIM | |

A coluna `enviado` é atualizada automaticamente durante o disparo.
