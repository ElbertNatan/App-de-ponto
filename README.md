# Folha de Ponto

Aplicativo desktop em Python (Tkinter + Pillow) para controle de horas
trabalhadas, com banco de horas, exportação para Excel e UI moderna —
cards com cantos arredondados, sombras suaves, botões custom desenhados
em Canvas e ícone próprio na barra de tarefas.

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
- **Pillow** — usado para renderizar cantos arredondados, sombras suaves
  nos cards e o ícone do app em múltiplas resoluções (16-256 px).

### Instalação das dependências

```powershell
pip install openpyxl Pillow
```

> Se o `pip` não estiver no PATH, use `python -m pip install openpyxl Pillow`.
>
> O Pillow é opcional do ponto de vista funcional — sem ele a aba *Dia*
> sobe normalmente, mas os widgets caem para retângulos simples sem
> arredondamento/sombra. O ícone também precisa de Pillow para ser
> gerado (`folha_ponto.ico` já vem no repositório, então só é necessário
> reinstalar a Pillow se quiser regerar o ícone).

---

## Como executar

### Opção 1 — Atalho `.bat` (recomendado no Windows)

Dê duplo clique em `executavel.bat`. Ele inicia o app via `pythonw` (sem abrir
janela de console) apontando para `src\folha_ponto.py`.

### Opção 2 — Linha de comando

```powershell
python src\folha_ponto.py
```

Na primeira execução, qualquer arquivo `ponto_AAAA_MM.txt` presente em `data/`
é migrado automaticamente para `data\folha_ponto.xlsx`.

---

## Estrutura do projeto

```
App ponto/
├── executavel.bat        # Atalho Windows que dispara pythonw src/folha_ponto.py
├── README.md
├── .gitignore
├── src/                  # Codigo-fonte + assets versionados
│   ├── folha_ponto.py    # Aplicativo principal (Tkinter + Pillow)
│   └── folha_ponto.ico   # Icone do app (10 tamanhos, 16-256 px)
└── data/                 # Dados gerados pelo usuario (NAO versionados)
    ├── folha_ponto.xlsx  # Planilha de horas
    └── ponto_AAAA_MM.txt # Backup mensal em texto
```

> A pasta `data/` inteira e `__pycache__/` ficam no `.gitignore` — são
> dados pessoais de quem usa o app, não devem entrar no repositório.
> O app cria `data/` automaticamente na primeira execução.

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
- **Meta do mês**: `8h × (dias úteis do mês inteiro)`, independente de já
  ter ocorrido ou não — assim a Meta exibida já reflete o total esperado
  do mês desde o dia 1.
- **Saldo do dia** = `Trabalho - Meta`.
- Dias do tipo **FERIADO**, **ATESTADO** ou **FOLGA** zeram a meta do dia
  (saldo neutro) e, na soma do mês, contam como 8h "creditadas" para que
  `Trabalho - Meta` continue batendo com a soma dos saldos diários.
- **Saldo acumulado** = soma dos saldos diários, mês a mês, exibido na aba
  *Banco geral* e na aba *Resumo* do Excel.
- **Auto-save**: o estado do dia é gravado no `folha_ponto.xlsx` a cada
  10 minutos após o início do dia, então um desligamento inesperado não
  perde mais do que esse intervalo.

---

## Dicas

- Editou um horário do dia atual? O cronômetro recalcula na hora.
- Mudou de tipo de dia (ex.: marcou como FERIADO)? O banco de horas é
  recalculado e o saldo do mês refletido na aba *Resumo*.
- Quer auditar um dia antigo? Abra o `.txt` do mês — é texto puro,
  legível, fácil de comparar com o Excel.
