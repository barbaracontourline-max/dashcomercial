# Painel Comercial Contourline — hospedagem (GitHub Pages)

Atualiza sozinho (a cada ~15 min) lendo a planilha do Google e publica num link fixo.

## Passos (uma vez só)

1. **Criar o repositório** em github.com → New repository → nome `dashcomercial` → **Public** → Create.

2. **Subir os arquivos**: na página do repo → "Add file" → "Upload files" →
   arraste `gerar_painel.py`, `requirements.txt` e a pasta `Imagens`. Commit.

3. **Criar o agendador**: aba **Actions** → "set up a workflow yourself" →
   apague o conteúdo e cole o conteúdo do arquivo `.github/workflows/atualizar.yml`
   (está nesta pasta). Nomeie `atualizar.yml` → Commit.

4. **Adicionar o segredo do link da planilha**: Settings → Secrets and variables →
   Actions → New repository secret →
   - Name: `SHEET_ID`
   - Secret: *(o ID da planilha — peça ao Drucker/está no chat)*
   → Add secret.

5. **Ligar o Pages**: Settings → Pages → Source: **GitHub Actions**.

6. **Rodar**: aba Actions → "Atualizar Painel Comercial" → "Run workflow".
   Em ~1-2 min sai o link: `https://SEU-USUARIO.github.io/dashcomercial/`

Pronto. Daí em diante atualiza sozinho a cada ~15 min; F5 mostra a última versão.

## Quando as metas dos vendedores mudarem
As metas ficam embutidas em `gerar_painel.py` (METAS_EMBED). É só avisar o Drucker que ele atualiza.
