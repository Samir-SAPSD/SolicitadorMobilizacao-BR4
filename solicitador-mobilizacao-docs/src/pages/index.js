import clsx from 'clsx';
import Link from '@docusaurus/Link';
import useDocusaurusContext from '@docusaurus/useDocusaurusContext';
import Layout from '@theme/Layout';
import styles from './index.module.css';

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
      <div className="text--center padding-horiz--md padding-vert--lg">
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
          <Link className="button button--secondary button--lg" to="/intro">
            Começar →
          </Link>
        </div>
      </div>
    </header>
  );
}

export default function Home() {
  return (
    <Layout title="Início" description="Documentação técnica gerada automaticamente">
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
