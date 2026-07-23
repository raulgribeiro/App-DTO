# Dashboard DTO — Diagnóstico de Trabalho Operacional

## Arquivos
- `index.html` — o dashboard pronto (é o que vai para o GitHub Pages)
- `template.html` — o "molde" do dashboard (layout, filtros, gráficos, tema claro/escuro)
- `extract_dto.py` — lê o `.xlsx` exportado do Forms e gera `dto_data.json` limpo
- `build_dashboard.py` — injeta o `dto_data.json` no `template.html` e gera o `index.html`
- `atualizar_dashboard.py` — script único que faz tudo (extrai + gera + sobe pro GitHub)

## Primeira vez (configuração do repositório GitHub)
1. Crie um repositório no GitHub (ex: `dto-dashboard`).
2. Ative o GitHub Pages: Settings → Pages → Branch `main` → pasta `/ (root)`.
3. Clone o repositório na sua máquina e copie estes arquivos para dentro dele.
4. Dê o primeiro commit/push manual:
   ```
   git add .
   git commit -m "Primeira versao do dashboard DTO"
   git push origin main
   ```
5. O site ficará disponível em algo como `https://seu-usuario.github.io/dto-dashboard/`.

## Toda vez que quiser atualizar com um novo export do Forms
1. Exporte as respostas do Forms como `.xlsx` e salve na pasta do repositório.
2. Rode:
   ```
   python atualizar_dashboard.py "nome_do_arquivo_exportado.xlsx"
   ```
3. O script extrai os dados, gera o novo `index.html` e já sobe pro GitHub.
   Em ~1 minuto o site atualiza sozinho (GitHub Pages).

## Personalização
- Cores/tema: editar as variáveis `:root` e `.light` no início do `<style>` do `template.html`.
- Colunas usadas: ver `extract_dto.py` (mapeamento das colunas B, F, J, L, M, O, P, R, Q:AK, AL, AM, AN).
- Filtros/gráficos/KPIs: editar o bloco `<script>` do `template.html` (função `render()`).
