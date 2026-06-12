# DJ Famosa — Press Kit Oficial

Press Kit profissional da DJ Famosa (@famosasouza), com design refinado, animações sofisticadas e deploy automático via GitHub Pages.

## 🎯 Sobre o Projeto

Este repositório contém o press kit oficial da DJ Famosa, uma DJ e produtora musical brasileira com presença marcante na cena eletrônica e tribal.

**Características:**
- ✨ Design profissional com gradientes e animações suaves
- 🎨 Dark mode + light mode toggle
- 📱 Totalmente responsivo (mobile, tablet, desktop)
- ⚡ Animações de scroll reveal
- 🔗 Links com underline animado
- 🎵 Integração com SoundCloud
- 📊 Estatísticas da carreira

## 📂 Estrutura do Projeto

```
.
├── presskit.html              # Press kit principal (estático)
├── index.html.pages           # Redirect para GitHub Pages
├── server.js                  # Servidor Node.js (desenvolvimento local)
├── IMG_4543.jpg              # Foto principal (hero)
├── IMG_4533.jpg              # Foto bio/galeria
├── LOGO-DJ-FAMOSA-COLORIDA-1.jpg # Logo oficial
├── .github/
│   └── workflows/
│       └── deploy-pages.yml   # Workflow de deploy automático
└── README_PRESSKIT.md         # Este arquivo
```

## 🚀 Deploy

O projeto está configurado com GitHub Pages e faz deploy automático a cada push para:
- `main` branch
- `claude/wizardly-gates-bk99zs` branch

### URL de Produção
https://nederhec.github.io/Projects/presskit.html

### Como Ativar GitHub Pages

1. Acesse: **Settings** → **Pages**
2. Em "Build and deployment":
   - **Source:** Deploy from a branch
   - **Branch:** main (or your target branch)
   - **Folder:** / (root)
3. Salve

O GitHub Actions fará o deploy automaticamente!

## 🛠️ Desenvolvimento Local

### Rodar o servidor

```bash
node server.js
```

Acesse: http://localhost:3000/presskit

### Editar o presskit

O arquivo `presskit.html` contém todo o HTML/CSS/JS. Edite diretamente e recarregue o navegador.

### Atualizar imagens

Substitua os arquivos:
- `IMG_4543.jpg` — foto principal
- `IMG_4533.jpg` — foto bio
- `LOGO-DJ-FAMOSA-COLORIDA-1.jpg` — logo

Faça commit e push para deploy automático.

## 📱 Seções Principais

1. **Hero** — Nome, tagline, foto e CTAs
2. **Stats** — 8+ anos, 200+ shows, 50K seguidores, 12 estados
3. **Bio** — História, trajetória, sonoridade e skills
4. **Música** — Releases & mixes do SoundCloud
5. **Assinatura** — Identidade, narrativa e som
6. **Contato** — Email, Instagram, SoundCloud

## 🎨 Cores & Design

**Paleta Principal:**
- Rosa: `#ec4899`
- Laranja: `#f97316`
- Roxo: `#7c3aed`

**Tipografia:**
- Display: Space Mono
- Body: Space Grotesk

## 📦 Otimizações

- Imagens de alta qualidade (3-4 MB cada)
- CSS crítico inline
- Animações com GPU acceleration
- Lazy loading em scroll reveals
- Mobile-first responsive design

## 🔄 Workflow de Desenvolvimento

```bash
# 1. Editar arquivos localmente
vim presskit.html

# 2. Testar localmente
node server.js
# Acesse http://localhost:3000/presskit

# 3. Commitar e fazer push
git add .
git commit -m "Update presskit content"
git push origin claude/wizardly-gates-bk99zs

# 4. Mergear para main (quando pronto)
# (via PR #6 ou direto)

# 5. Deploy automático!
# GitHub Pages vai atualizar em ~1 minuto
```

## 📞 Contato

- Email: famosasouzaa@gmail.com
- Instagram: @famosasouza
- SoundCloud: dj-famosa

## 📄 Licença

© 2026 DJ Famosa. Todos os direitos reservados.

---

**Desenvolvido com ❤️ para DJ Famosa**
