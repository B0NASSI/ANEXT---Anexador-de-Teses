# ANEXT — Anexador de Teses

App interno (Tkinter + ttkbootstrap) que gera, divide e junta os documentos de cada segurado da tese num PDF, para a equipe de revisão de insumos do FAP. Usuário: advogado(a) do escritório Rodriguez & Sousa, não necessariamente com bagagem técnica profunda — prefira explicações práticas a jargão quando o assunto sair do código em si.

**Mantenha este arquivo atualizado** conforme o projeto muda — é o que evita ter que reexplicar o contexto do zero em cada conversa nova.

## Onde isso vive

- Monorepo privado `claude-code` (GitHub, B0NASSI) — este repositório, pasta `ANEXT/`. É a fonte de verdade.
- Espelho de código-fonte público `ANEXT---Anexador-de-Teses` (GitHub, B0NASSI) — é dali que o launcher instalado baixa as releases (`GITHUB_OWNER`/`GITHUB_REPO` em `src/launcher.py`). Precisa ser sincronizado manualmente a cada release (não é um subtree/submodule de verdade — ver "Processo de release" abaixo).

## Versão atual: 3.7

Changelog completo em `NOTAS DE ATUALIZAÇÃO/*.txt` (um arquivo por versão — também aparece dentro do próprio app, no menu "Notas de atualização").

## Arquitetura, em resumo

- `src/app.py` — janela principal, 4 abas: All-in-one, Gerar capas, Dividir capas, Juntar PDFs.
- `src/launcher.py` — checa releases no GitHub, baixa, confere SHA256 (`ASSET_NAME + ".sha256"`, ver `installer/gerar_checksum.py`), aplica a atualização e reabre o app. **Atualização é obrigatória** — sem opt-in, todo mundo recebe ao mesmo tempo.
  - **Pegadinha importante**: `apply_update()` só sobrescreve `ANEXT.exe` e o conteúdo de `_internal/` — o `ANEXT Launcher.exe` já instalado NUNCA é atualizado pelo auto-update. Uma correção só dentro de `launcher.py` não chega em quem já tem o app instalado; se o bug precisa alcançar todo mundo, a correção/mitigação tem que rodar no `app.py` (que sim é atualizado a cada release). Ex.: `_limpar_zips_temporarios_orfaos()` em `app.py`, que limpa na inicialização um bug antigo do launcher que deixava `tmp*.zip` (~40MB) esquecidos na pasta de instalação a cada atualização aplicada com sucesso.
- `src/log_setup.py` — log rotativo (`logging.handlers.RotatingFileHandler`) em `logs/<nome>.log`, ao lado do `.exe`. Usado por `app.py` (`anext.log`) e `launcher.py` (`launcher.log`). Cada uma das 4 abas loga sucesso (`logger.info`) e erro com traceback (`logger.exception`) de toda operação — sem isso, um erro dentro de um callback do Tkinter desaparece sem rastro (build usa `console=False`).
- `src/limites_caminho.py` — trunca nomes livres (segurado, pasta da tese) o quanto for preciso pra não estourar o limite de 260 caracteres de caminho do Windows.
- `src/tabela.py`, `src/documento.py`, `src/gerar.py`, `src/separar_capas.py`, `src/juntar_pdfs.py`, `src/compactar_pdf.py`, `src/pdf.py` — lógica de negócio, majoritariamente sem dependência de UI.
- `CaixaColarTabela` (em `app.py`) — componente reutilizado nas abas All-in-one e Gerar capas para colar a tabela do Word (Ctrl+V) com captura automática da formatação (mesclagens/cores/bordas/fontes) via Word invisível, com fallback para texto simples.
- `_avisar_se_desatualizado()` (em `app.py`, chamada logo antes de `root.mainloop()`) — rede de segurança para quando alguém abre o `ANEXT.exe` direto, sem passar pelo launcher (ex.: fixou o ícone errado na barra de tarefas a partir do app já aberto, em vez do atalho da área de trabalho que aponta pro launcher). Reaproveita `get_latest_release`/`is_newer`/`read_local_version` de `launcher.py` via `import launcher` dentro da própria thread de checagem — só avisa com um `messagebox.showinfo` (sem número de versão na mensagem, decisão deliberada) pra fechar e reabrir pelo atalho; nunca baixa nem trava o app, falha em silêncio (só loga) se não conseguir checar. Mesmo padrão já usado no RequerimentoGERID.
  - **Pegadinha que já mordeu no REQUERID**: `app.py` importa `launcher.py` só por essas 3 funções, então o build do `ANEXT.exe` passa a precisar de `requests` de verdade (não só o hidden import — collect_all completo, ver `.spec`) e **não pode excluir `email`/`http`** dos `excludes` do Analysis do app — `requests` depende deles por baixo dos panos (via urllib3) pra abrir a conexão HTTPS. Excluir quebra a checagem em silêncio (sem crash, só nunca encontra atualização), só detectável com log dentro do except. Sempre confira o `.spec` antes de adicionar qualquer coisa que use `requests`/rede em `app.py`.
  - Detalhe de implementação: `log_setup.configurar_logging("launcher.log", BASE_DIR)` em `launcher.py` fica dentro de `main()`, não no nível do módulo — se rodasse na importação, o `app.py` ganharia um segundo handler de log (`launcher.log`) duplicando tudo que ele já loga em `anext.log`.

## Testes

`tests/` (pytest, `pytest.ini` aponta `pythonpath = src`). Cobre `tabela.py`, `separar_capas.py`, `juntar_pdfs.py`, `compactar_pdf.py` — tudo com PDFs sintéticos (nunca usar `ANEXT/Exemplos/` como fixture: é gitignorado por ter dados reais de cliente, não existe em clone novo).

**Não cobertos, propositalmente**: `documento.py`, `gerar.py`, `pdf.py` — dependem do Word via COM (`pywin32`), não dá pra rodar limpo sem Word instalado.

Rodar: `.venv\Scripts\pytest -v` (venv próprio do ANEXT, criado com `uv venv .venv` + `uv pip install -r requirements-dev.txt`).

## Processo de release (manual, sem CI ainda)

1. Bump `versao.txt` + criar `NOTAS DE ATUALIZAÇÃO/X.Y.txt`, commit no monorepo.
2. `pyinstaller "ANEXT - Anexador de Teses.spec"` (gera `dist/ANEXT/` com `ANEXT.exe` + `_internal/`).
3. Zipar o conteúdo de `dist/ANEXT/` (arquivos na raiz do zip, não a pasta) → `ANEXT-app.zip`.
4. `python installer/gerar_checksum.py ANEXT-app.zip` → gera o `.sha256`.
5. `gh release create <versão> ANEXT-app.zip ANEXT-app.zip.sha256 --repo B0NASSI/ANEXT---Anexador-de-Teses`.
6. Sincronizar o código-fonte pro repositório público: **usar um `git worktree` isolado** apontando pra `anext-origin/main`, copiar os arquivos rastreados de `ANEXT/` (via `git ls-files`, sem prefixo), commitar e dar push — nunca faça isso com `cd` solto misturado em vários comandos (já causou um incidente de arquivos apagados no lugar errado; sempre use `git -C <caminho>` explícito).

**Sempre peça confirmação antes de publicar** (`gh release create`), mesmo em correção rotineira — e separe explicitamente "correção" de "feature nova/não testada" nesse pedido, pra não empacotar as duas coisas juntas sem o usuário poder escolher segurar a feature.

## Pendências conhecidas

- Sem CI — testes existem mas dependem de rodar `pytest` manualmente antes de empacotar.
- "Colar texto" (lançado na 3.7) ainda não foi testado com casos reais/tabelas complexas pela equipe.
- Pedido em aberto (não implementado): levar o mesmo `log_setup.py`/logging pro RequerimentoGERID, que hoje só tem um log manual (`registrar_erro` → `requerid_log.txt`), sem rotação e só de erro.
