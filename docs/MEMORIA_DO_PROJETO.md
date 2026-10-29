# Lista de Campo — Memória Técnica e Histórico do Projeto
> **Criador & Autor:** Lucas Aøsf  
> **Organização / Identidade:** Pesquisa & Soluções Ambientais  
> **Repositório GitHub:** `lucasaosf/caderneta-de-campo`  
> **Documento de Contexto Consolidado:** Gerado a partir do histórico completo de concepção, requisitos e desenvolvimento.

---

## 1. Visão Geral e Propósito

O **Lista de Campo** é um aplicativo voltado para biólogos, ecólogos, consultores ambientais e ornitólogos/mastozoólogos/herpetólogos em atividades de inventário, monitoramento e resgate de fauna.

O aplicativo foi projetado com a filosofia **Offline-First**: funciona perfeitamente sem sinal de internet em áreas remotas de floresta ou cerrado, permitindo o registro de espécies, geolocalização contínua de trilhas e contagem de indivíduos, com sincronização em nuvem e exportação de relatórios analíticos prontos.

---

## 2. Identidade Visual e Design System

- **Paleta de Cores (Pesquisa & Soluções Ambientais):**
  - **Verde Floresta:** `--accent: #173A1E`, `--surface: #132F1C`, `--bg: #0E2515` (tema escuro)
  - **Creme / Linho:** `--bg: #F3EEE3`, `--surface: #FBF8F1` (tema claro)
  - **Sálvia:** `--accent: #B2C2A5`, `--accent-soft: #1E472A`
  - **Bronze / Dourado:** `--bronze: #D4B277`, `--gold: #B08854`
- **Tipografia:**
  - *Cormorant Garamond* (fontes clássicas, elegantes para títulos, capa e numerações com peso editorial)
  - *Jost* (fonte moderna e limpa para leitura em campo, botões técnicos e listas)
- **Modos de Tema:**
  - **Escuro (Dark):** Otimizado para visualização sob luz solar ou à noite sem ofuscamento.
  - **Claro (Light):** Identidade com fundo creme e tipografia contrastante.
  - *(Nota: O modo noturno vermelho foi descontinuado na v16 para simplificar a alternância direta de temas).*
- **Ícones Oficiais dos Grupos Faunísticos:**
  - 🐦 **Aves (Ornitofauna):** Bem-te-vi
  - 🐆 **Mamíferos (Mastofauna):** Onça-pintada
  - 🐸 **Herpetofauna (Répteis e Anfíbios):** Sapo / Cecília / Serpente
  - 🧭 **Geral:** Coleta multi-taxonômica integrada
- **Marca d'água:** `Lista de Campo · Lucas Aøsf`

---

## 3. Arquitetura e Estrutura Técnica

### 3.1. Front-end PWA (HTML5 / Vanilla JS / CSS Moderno)
- **Zero frameworks pesados:** Máxima leveza e velocidade de carregamento em aparelhos celulares modestos em campo.
- **`index.html`:** Aplicação de página única (SPA) modular contendo Capa, Minhas Listas, Tela de Campo Ativa, Tela de Apelidos, Perfil e Modais de Exportação.
- **`sw.js` (Service Worker):** Gerenciamento inteligente de cache offline, permitindo o funcionamento sem rede e atualização atômica de versões via `version.json`.
- **`manifest.webmanifest`:** Configuração PWA nativa para Android e iOS, com ícones adaptativos e modo standalone.

### 3.2. Banco Taxonômico Brasileiro (`especies.js`)
- Banco completo pré-carregado com milhares de espécies da fauna brasileira:
  - Classificação taxonômica: Classe, Ordem, Subordem, Família, Subfamília, Tribo, Gênero, Espécie e Autor.
  - Atributos ecológicos: Nomes populares (PT), Bioma, Origem (Nativa/Exótica), Dieta, Locomoção.
  - Status de Conservação Oficial: Classificações da **IUCN** e **MMA** (CR - Criticamente em Perigo, EN - Em Perigo, VU - Vulnerável, NT, LC), CITES e Endemismo.
  - Alertas visuais destacados durante o preenchimento de espécies ameaçadas ou invasoras.

### 3.3. Sistema de Apelidos Rápidos (Aliases)
- Permite que o pesquisador crie atalhos e gírias individuais de campo (ex: `jussu` para *Bothrops jararacussu*, `jara` para *Bothrops jararaca*, `bentevi` para *Pitangus sulphuratus*).
- Durante o registro, o atalho autocompleta instantaneamente.
- Na exportação final, o apelido é automaticamente convertido na taxonomia formal completa (sem expor a coluna de apelidos na planilha).

### 3.4. Rastreamento e GPS (Modo Bolso / Pocket Mode)
- Gravação das coordenadas de cada espécime avistado (Latitude, Longitude e Precisão do GPS).
- Gravação de trilha contínua do percurso amostral.
- **Pocket Mode:** Mantém a gravação de rota ativa mesmo com a tela do aparelho protegida no bolso, utilizando WakeLock e fallback de geolocalização contínua.

### 3.5. Autenticação e Nuvem (Firebase)
- Arquitetura híbrida:
  - **Com conta (Firebase Auth + Firestore):** Login Google, Apple ou E-mail/Senha. Listas, apelidos e fotos de perfil sincronizam entre múltiplos dispositivos.
  - **Sem login (Modo Convidado):** Operação total em `localStorage` local.

### 3.6. Motor de Análise e Exportação (`planilha.js`)
- **Excel Completo (`.xlsx`):**
  - Matriz de Riqueza e Abundância.
  - Distinção de registros **V** (Visual) e **A** (Auditivo).
  - Curvas de Acúmulo de Espécies.
  - Curva de Rarefação nos padrões do software **EstimateS**: estimadores Jackknife 1, Chao 1, Bootstrap, S(est), com número de randomizações configurável e cálculo de intervalos de confiança (IC = 1,96 × SD).
  - Compilado cumulativo para projetos com múltiplas campanhas consecutivas.
- **Exportações SIG / Ambientais:**
  - **Darwin Core Archive (DwC):** Padrão aceito pelo SiBBr, GBIF, IBAMA e CETESB.
  - **KML:** Visualização direta no Google Earth com marcadores e rotas.
  - **GeoJSON:** Camada geoespacial pronta para QGIS / ArcGIS.
  - **CSV:** Dados brutos tabulados.
  - **Card Infográfico:** Resumo formatado com imagem e dados essenciais para compartilhamento imediato via WhatsApp.

---

## 4. Roadmap e Próximos Passos Agendados

1. **Google Play Store:**
   - Empacotamento via Capacitor / TWA (os ganchos `window.Capacitor` já estão preparados na arquitetura).
2. **Monetização Freemium:**
   - Modelo de teste: Geração gratuita de até 3 a 5 planilhas analíticas completas.
   - Plano Pro/Vitalício: Pagamento único de ~R$ 50, com liberação atrelada diretamente ao e-mail de login do usuário (sem necessidade de chave de licença manual).
3. **Entrada de Registros por Voz (Speech-to-Text):**
   - Web Speech API para ditar contagens em campo com mãos ocupadas (binóculos/gravador).
4. **Vídeo Tutorial e Divulgação:**
   - Atualização contínua de [como-usar.html](file:///c:/Users/Lucas%20Aosf/Desktop/Projeto-Caderneta/como-usar.html) com instruções e demonstração das novas ferramentas.
5. **IndexedDB:**
   - Suporte transparente para trilhas de GPS longas e extensos bancos de listas acumuladas.

---
*Este documento é mantido como base de conhecimento viva para orientar futuras atualizações e sprints de desenvolvimento do aplicativo.*
