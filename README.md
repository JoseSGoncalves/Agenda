# Site unificado (portal + agenda)

Uma pasta para **um único deploy** no Netlify: o **portal** na raiz (`/`) e a **agenda** em `/agenda/`.

A app **Finanças** (`financeiro/` no repositório) corre **só no teu PC** (IndexedDB local); não é copiada para esta pasta.

## Estrutura

| Caminho | Conteúdo |
|---------|----------|
| `/` | Portal (`index.html`, `app.js`, `styles.css`) |
| `/agenda/` | Agenda (`agenda/index.html`, …) |
| `config.js` (raiz) | Supabase: `PORTAL_CONFIG` e `AGENDA_CONFIG` (mesmo projecto) |

## Configurar

1. Edita **`config.js`** na raiz desta pasta (URL e chave publishable/anon do Supabase).
2. No **Supabase → Table Editor → `applications`**, na linha da agenda, define **`external_url`**, por exemplo `https://teu-site.netlify.app/agenda/` (barra final recomendada).
3. Para a entrada «Finanças» no portal: deixa **`external_url`** vazio (ou remove a linha) se a app só existir no teu computador.

## Supabase (Auth + API)

- **Authentication → URL Configuration**
  - **Site URL:** `https://teu-site.netlify.app` (o teu domínio real).
  - **Redirect URLs:** inclui `https://teu-site.netlify.app` e `https://teu-site.netlify.app/**`.
- **API Keys → Publishable → HTTP referrers:** `https://teu-site.netlify.app/*` (e `http://localhost:*/*` para desenvolvimento local).

## Testar localmente

Na pasta `site/`:

```bash
cd site
python3 -m http.server 8080
```

- Portal: **http://localhost:8080/**
- Agenda: **http://localhost:8080/agenda/**

## Publicar no Netlify

1. **Add site → Deploy manually** (ou Git com **base directory** = `site` se o repo tiver mais pastas).
2. Arrasta **esta pasta `site/`** inteira (deve incluir `agenda/`, `netlify.toml`, `config.js`).
3. O URL de produção (ex. `https://algo.netlify.app`) é o que usas em **Site URL** e em **`external_url`**.

## Publicar no GitHub (repositório + GitHub Pages)

1. Na pasta `site/`, copia a configuração: `cp config.example.js config.js` e edita URL e chave do Supabase (ou reutiliza o teu `config.js` local).
2. Inicializa o Git e faz o primeiro commit (o `.gitignore` ignora `config.js` por defeito; para a agenda funcionar online, inclui-o no deploy com `git add -f config.js` depois de preenchido, ou mantém um repo privado).
3. Cria um repositório vazio no GitHub e faz push da pasta `site/` como raiz do repo (ou define **pasta de publicação** no GitHub Pages conforme a estrutura que escolheres).
4. **Settings → Pages → Build and deployment → Branch:** escolhe o branch (ex. `main`) e pasta **`/` (root)** se o conteúdo do repo for exactamente o conteúdo de `site/`.
5. URLs típicas: portal em `https://UTILIZADOR.github.io/NOME-DO-REPO/` e agenda em `https://UTILIZADOR.github.io/NOME-DO-REPO/agenda/`.
6. No **Supabase → Authentication → URL Configuration**, acrescenta o domínio `https://UTILIZADOR.github.io` e em **Redirect URLs** inclui `https://UTILIZADOR.github.io/**`. Em **API → Publishable key → HTTP referrers**, inclui `https://UTILIZADOR.github.io/*`.

## Actualizar código a partir do repositório

Quando alterares `portal-web/` ou `agenda-servicos/`, na raiz do projecto corre:

```bash
./scripts/sync-site.sh
```

Depois volta a fazer deploy da pasta `site/`. O script **não** sobrescreve `site/config.js` — mantém as tuas chaves.

## Dois sites Netlify em vez de um

Ainda podes publicar `portal-web` e `agenda-servicos` em URLs diferentes; vê `DEPLOY.md` na raiz do projecto.
