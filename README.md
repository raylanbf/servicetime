# ⏱ Service Timer

Aplicativo de controle de tempo por tipo de serviço. Roda como um botão flutuante na tela, sempre visível por cima das outras janelas.

---

## O que faz

- Cronometra o tempo gasto em cada tipo de serviço
- Suporta **pausas** com registro de horário de início, pausa, retorno e fim
- Salva os registros em **JSON** e **DOCX** localmente
- Envia os dados para uma planilha do **Google Sheets** com um clique
- Cada usuário da equipe tem sua própria aba na planilha
- Só envia registros que ainda não foram sincronizados (sem duplicatas)

---

## Requisitos

- **Python 3.8+** — [python.org/downloads](https://www.python.org/downloads/)
- Marcar **"Add Python to PATH"** durante a instalação

---

## Instalação

```bash
# 1. Clone o repositório
git clone https://github.com/raylanbf/servicetime.git
cd servicetime

# 2. Instale a dependência
pip install python-docx
```

Ou simplesmente dê duplo clique em **`instalar.bat`**.

---

## Como usar

Dê duplo clique em **`iniciar.bat`** — ou execute:

```bash
python timer_servico.py
```

Na **primeira abertura** o programa pede seu nome. Esse nome identifica sua aba na planilha do Google Sheets.

---

## Interface

```
┌─ ⏱ Service Timer ───────── — □  ✕ ─┐
│ 👤 Raylan                            │
│ ─────────────────────────────────── │
│ Tipo de serviço               [ ✎ ] │
│ [ Resolução de CSC          ▼     ] │
│ ┌─────────────────────────────────┐ │
│ │         00:00:00                │ │
│ │         Aguardando              │ │
│ └─────────────────────────────────┘ │
│ [ ▶ Iniciar ]    [ ⏸ Pausar ]      │
│ [ ⏹ Finalizar serviço           ]  │
│ ─────────────────────────────────── │
│ [ 📊 Enviar ao Google Sheets ] [ ? ]│
│ 3 registro(s)                  [ ⚙ ]│
└─────────────────────────────────────┘
```

| Ação | Descrição |
|------|-----------|
| **▶ Iniciar** | Inicia o cronômetro e minimiza automaticamente |
| **⏸ Pausar / ▶ Retomar** | Pausa e retoma, registrando os horários |
| **⏹ Finalizar** | Salva o registro com todos os horários |
| **✎** | Abre o editor para adicionar/remover/reordenar serviços |
| **—** | Minimiza para a barra compacta |
| **□** | Restaura a janela completa |
| **📊** | Envia registros pendentes ao Google Sheets |
| **?** | Instruções e script para configurar o Google Sheets |
| **⚙** | Configurações (nome e URL do Web App) |

### Barra compacta (minimizado)

Ao iniciar um serviço a janela recolhe automaticamente para uma barra de **55px**:

```
┌─ ⏱ Service Timer ───────── □  ✕ ─┐
│ ▶  00:12:34   Resolução de CSC    │
└───────────────────────────────────┘
```

- 🟢 **Verde** = em andamento
- 🟡 **Amarelo** = pausado
- ⚫ **Cinza** = aguardando

---

## Arquivos gerados

Os arquivos abaixo são criados automaticamente na mesma pasta do programa:

| Arquivo | Descrição |
|---------|-----------|
| `config.json` | Nome do usuário e URL do Web App |
| `registros.json` | Todos os registros com flag de enviado/pendente |
| `registros.docx` | Documento Word com histórico formatado |

---

## Configurar Google Sheets

> Feito **uma única vez** pelo responsável da equipe. Os demais membros só precisam da URL gerada.

### 1. Criar a planilha

Acesse [sheets.google.com](https://sheets.google.com) e crie uma planilha nova.

### 2. Criar o script

Com a planilha aberta, clique em **Extensões > Apps Script**, apague o conteúdo existente e cole o script abaixo:

```javascript
function doPost(e) {
  try {
    var dados = JSON.parse(e.postData.contents);
    var planilha = SpreadsheetApp.getActiveSpreadsheet();
    var nomeAba  = dados.usuario || "Sem nome";

    var aba = planilha.getSheetByName(nomeAba);
    if (!aba) {
      aba = planilha.insertSheet(nomeAba);
      aba.appendRow([
        "Usuário", "Data", "Tipo de Serviço",
        "Início", "Fim", "Duração", "Pausas"
      ]);
      aba.getRange(1, 1, 1, 7).setFontWeight("bold");
    }

    var registros = dados.registros || [];
    var linhas = [];
    for (var i = 0; i < registros.length; i++) {
      var r = registros[i];
      var pausas = (r.pausas || []).map(function(p) {
        return (p.pausa || "") + " → " + (p.retorno || "-");
      }).join("; ");
      linhas.push([
        r.usuario      || nomeAba,
        r.data         || "",
        r.tipo_servico || "",
        r.inicio       || "",
        r.fim          || "",
        r.tempo_total  || "",
        pausas
      ]);
    }

    if (linhas.length > 0) {
      aba.getRange(aba.getLastRow() + 1, 1, linhas.length, 7)
         .setValues(linhas);
    }

    return ContentService
      .createTextOutput(JSON.stringify({ status: "ok", linhas: linhas.length }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: "erro", erro: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
```

### 3. Publicar como Web App

1. Clique em **Implantar > Nova implantação**
2. Tipo: **App da Web**
3. Executar como: **Eu mesmo**
4. Quem pode acessar: **Qualquer pessoa**
5. Clique em **Implantar**
6. Copie a URL gerada: `https://script.google.com/macros/s/.../exec`

### 4. Configurar no programa

Clique em **⚙** no programa e cole a URL no campo **"URL do Web App"**.

> Cada membro da equipe usa a mesma URL. Os dados de cada um vão para uma aba separada com o nome do usuário.

---

## Estrutura do JSON

```json
{
  "usuario": "Raylan",
  "registros": [
    {
      "usuario": "Raylan",
      "tipo_servico": "Resolução de CSC",
      "data": "2026-05-06",
      "inicio": "09:00:00",
      "pausas": [
        { "pausa": "09:30:00", "retorno": "09:45:00" }
      ],
      "fim": "10:00:00",
      "tempo_total": "0:45:00",
      "tempo_total_segundos": 2700,
      "enviado": false
    }
  ]
}
```
