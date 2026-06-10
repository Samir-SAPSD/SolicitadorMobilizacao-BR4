# Template da Homepage

## Arquivo: `src/pages/index.js`

```jsx
import React from 'react';
import clsx from 'clsx';
import Link from '@docusaurus/Link';
import useDocusaurusContext from '@docusaurus/useDocusaurusContext';
import Layout from '@theme/Layout';
import styles from './index.module.css';

// ── Cards de seções principais ────────────────────────────────────────────────
const FeatureList = [
  {
    title: 'Arquitetura',
    emoji: '🏗️',
    description: 'Visão macro do sistema: diagrama de componentes, decisões de design e fluxo ponta a ponta.',
    link: '/architecture',
  },
  {
    title: 'Arquivos de Código',
    emoji: '📁',
    description: 'Documentação detalhada de cada arquivo: funções, classes, fluxos e decisões de engenharia.',
    link: '/files/server',
  },
  {
    title: 'Módulos & Integrações',
    emoji: '🔗',
    description: 'Mapa de dependências, integrações externas e protocolo de comunicação entre módulos.',
    link: '/modules/overview',
  },
  {
    title: 'Melhorias',
    emoji: '🔧',
    description: 'Problemas identificados, dívidas técnicas e próximos passos priorizados por impacto.',
    link: '/improvements',
  },
];

function Feature({ title, emoji, description, link }) {
  return (
    <div className={clsx('col col--3')}>
      <div className="text--center padding-horiz--md padding-vert--md">
        <div style={{ fontSize: '3rem', marginBottom: '0.5rem' }}>{emoji}</div>
        <h3>{title}</h3>
        <p>{description}</p>
        <Link className="button button--secondary button--sm" to={link}>
          Ver documentação →
        </Link>
      </div>
    </div>
  );
}

function HomepageHeader() {
  const { siteConfig } = useDocusaurusContext();
  return (
    <header className={clsx('hero hero--primary', styles.heroBanner)}>
      <div className="container">
        <h1 className="hero__title">{siteConfig.title}</h1>
        <p className="hero__subtitle">{siteConfig.tagline}</p>
        <div className={styles.buttons}>
          <Link
            className="button button--secondary button--lg"
            to="/intro">
            Começar →
          </Link>
        </div>
      </div>
    </header>
  );
}

export default function Home() {
  return (
    <Layout
      title="Início"
      description="Documentação técnica gerada automaticamente">
      <HomepageHeader />
      <main>
        <section style={{ padding: '2rem 0' }}>
          <div className="container">
            <div className="row">
              {FeatureList.map((props, idx) => (
                <Feature key={idx} {...props} />
              ))}
            </div>
          </div>
        </section>
      </main>
    </Layout>
  );
}
```

---

## Arquivo: `src/pages/index.module.css`

```css
.heroBanner {
  padding: 4rem 0;
  text-align: center;
  position: relative;
  overflow: hidden;
}

.buttons {
  display: flex;
  align-items: center;
  justify-content: center;
  margin-top: 1.5rem;
}

@media screen and (max-width: 996px) {
  .heroBanner {
    padding: 2rem;
  }
}
```

---

## Adaptação Automática dos Cards

A Skill deve ajustar `FeatureList` com base nas categorias detectadas:

| Categoria detectada | Card gerado |
|---|---|
| `architecture.md` existe | Card "Arquitetura" → `/architecture` |
| `docs/files/` existe | Card "Arquivos de Código" → primeiro arquivo de `files/` |
| `docs/modules/` existe | Card "Módulos" → `/modules/overview` ou primeiro `.md` |
| `improvements.md` existe | Card "Melhorias" → `/improvements` |
| Qualquer outro subdiretório | Card com nome capitalizado do diretório |

Se não houver subdiretórios, gerar apenas os cards dos arquivos raiz.
