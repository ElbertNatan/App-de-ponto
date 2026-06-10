# Folha de Ponto

Aplicativo desktop em Python (Tkinter) para controle de horas trabalhadas, com
banco de horas, exportação para Excel e interface moderna em abas.

---

## Recursos

- **Cronômetro ao vivo** com botões **Play / Pause / Finalizar** registrando os
  horários automaticamente no dia atual.
- **Interface em abas**:
  - **Dia** — carrossel com setas para navegar entre dias e editar manualmente
    Início, Início da Pausa, Fim da Pausa e Fim.
  - **Mês atual** — visão consolidada do mês corrente.
  - **Banco geral** — saldo acumulado considerando meta diária de 8h.
  - **Marcações** — log cronológico de todos os eventos registrados.
- **Tipos de dia**: NORMAL, FERIADO, ATESTADO, FOLGA (os três últimos
  neutralizam a meta do dia para fins de banco de horas).
- **Excel** `folha_ponto.xlsx` com uma aba por mês + Resumo + Marcações,
  formatado com cabeçalho, zebra e bordas.
- **Backup em texto**: além do Excel, mantém os arquivos `ponto_AAAA_MM.txt`
  mensais (formato legível à mão).
- **Migração automática** dos arquivos `.txt` antigos na primeira execução.
- **Banco de horas** com meta de 8h em dias úteis e saldo acumulado.

---

## Requisitos

- **Python 3.10+** (testado em 3.14) — disponível em <https://www.python.org/downloads/>.
- **Tkinter** — já vem com a instalação padrão do Python no Windows.
- **openpyxl** — para gravar a planilha `.xlsx`.

### Instalação das dependências

```powershell
pip install openpyxl
```

> Se o `pip` não estiver no PATH, use `python -m pip install openpyxl`.

---

## Como executar

### Opção 1 — Atalho `.bat` (recomendado no Windows)

Dê duplo clique em `executavel.bat`. Ele inicia o app via `pythonw` (sem abrir
janela de console).

### Opção 2 — Linha de comando

```powershell
python folha_ponto.py
```

Na primeira execução, qualquer arquivo `ponto_AAAA_MM.txt` presente na pasta é
migrado automaticamente para o `folha_ponto.xlsx`.

---

## Estrutura do projeto

```
App ponto/
├── folha_ponto.py        # Aplicativo principal (Tkinter)
├── executavel.bat        # Atalho Windows que dispara pythonw
├── folha_ponto.xlsx      # Planilha gerada/atualizada pelo app (não versionada)
├── ponto_AAAA_MM.txt     # Backup mensal em texto (não versionado)
└── README.md
```

> Os arquivos `*.xlsx`, `*.txt` e `__pycache__/` ficam no `.gitignore` — são
> dados pessoais de quem usa o app, não devem entrar no repositório.

---

## Formato do arquivo `.txt`

Cada dia é registrado num bloco delimitado por `==============================`:

```
==============================
Data: 01/06/2026

[10:01:55] INÍCIO DO TRABALHO
[11:49:46] INÍCIO DA PAUSA
[12:14:58] FIM DA PAUSA
[12:14:58] INÍCIO DO TRABALHO
[18:53:58] FINALIZAÇÃO DO DIA

Resumo:
Trabalho total: 08:26:51
Pausa total:    00:25:12
==============================
```

---

## Planilha `folha_ponto.xlsx`

| Aba                 | Conteúdo                                                                |
|---------------------|-------------------------------------------------------------------------|
| `AAAA-MM` (uma/mês) | Data, Dia, Início, Ini Pausa, Fim Pausa, Fim, Trabalho, Pausa, Meta, Saldo, Tipo, Observação |
| `Resumo`            | Dias trabalhados, Trabalho total, Meta total, Saldo do mês, Saldo acumulado por mês |
| `Marcacoes`         | Log cronológico (Data, Tipo, Observação)                                |

---

## Banco de horas

- **Meta diária**: 8 horas em dias úteis.
- **Saldo do dia** = `Trabalho - Meta`.
- Dias do tipo **FERIADO**, **ATESTADO** ou **FOLGA** zeram a meta — saldo do
  dia fica neutro.
- **Saldo acumulado** = soma dos saldos diários, mês a mês, exibido na aba
  *Banco geral* e na aba *Resumo* do Excel.

---

## Dicas

- Editou um horário do dia atual? O cronômetro recalcula na hora.
- Mudou de tipo de dia (ex.: marcou como FERIADO)? O banco de horas é
  recalculado e o saldo do mês refletido na aba *Resumo*.
- Quer auditar um dia antigo? Abra o `.txt` do mês — é texto puro,
  legível, fácil de comparar com o Excel.
