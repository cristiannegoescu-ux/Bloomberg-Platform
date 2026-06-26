/**
 * Bloomberg Reporter Platform — PowerPoint Generator
 * Uses pptxgenjs (pure Node) to build the presentation.
 * Run: node generate-pptx.js
 * Output: Bloomberg_Reporter_Platform.pptx (in this directory)
 */

const PptxGenJS = require('pptxgenjs');
const path = require('path');
const fs   = require('fs');

const pptx = new PptxGenJS();

// ── Global theme ──────────────────────────────────────────────────────────────
pptx.layout    = 'LAYOUT_WIDE';
pptx.author    = 'Cristian Negoescu';
pptx.company   = 'Ubisoft — QC Technology Group';
pptx.subject   = 'Bloomberg Reporter Platform — Strategic Vision';
pptx.title     = 'Bloomberg Reporter Platform';

const BG         = '0B0F1A';
const CARD_BG    = '161B22';
const ACCENT     = '4F8EF7';
const GREEN      = '3FB950';
const AMBER      = 'D29922';
const RED        = 'F85149';
const PURPLE     = 'C77DFF';
const CYAN       = '58C9DA';
const TEXT_PRI   = 'E6EDF3';
const TEXT_SEC   = '8B949E';
const TEXT_MUT   = '484F58';
const BORDER     = '303444';

const LOGO_PATH  = path.join(__dirname, 'public', 'Bloomberg_Logotype_White.png');
const BG_IMG     = path.join(__dirname, 'public', 'Bloomberg_Imagery_NEUTRAL_159.jpg');
const QA_IMG     = path.join(__dirname, 'public', 'images', 'qa-companion.jpg');
const QC_IMG     = path.join(__dirname, 'public', 'images', 'qc-assistant.jpg');
const EE_IMG     = path.join(__dirname, 'public', 'images', 'eagle-eye.jpg');
const UB_SVG     = path.join(__dirname, 'public', 'images', 'ubisoft-vintage.svg');

function hasFile(p) { try { return fs.existsSync(p); } catch { return false; } }

// ── Helper: add dark background + subtle bg image to every slide ──────────────
function addBase(slide) {
  slide.background = { color: BG };
  if (hasFile(BG_IMG)) {
    slide.addImage({ path: BG_IMG, x: 0, y: 0, w: '100%', h: '100%', transparency: 87 });
  }
}

// ── Helper: add logo to bottom-left of every slide ────────────────────────────
function addLogo(slide) {
  if (hasFile(LOGO_PATH)) {
    slide.addImage({ path: LOGO_PATH, x: 0.4, y: 6.9, w: 1.5, h: 0.35, transparency: 25 });
  }
}

// ── Helper: section label pill ────────────────────────────────────────────────
function eyebrow(slide, text, y = 0.7) {
  slide.addText(text.toUpperCase(), {
    x: 0.6, y, w: 8, h: 0.28,
    fontSize: 9, bold: true, color: ACCENT, charSpacing: 2.5,
  });
  slide.addShape(pptx.ShapeType.rect, {
    x: 0.6, y: y + 0.32, w: 0.45, h: 0.04,
    fill: { color: ACCENT }, line: { color: ACCENT },
  });
}

// ── Helper: section title ─────────────────────────────────────────────────────
function title(slide, line1, line2, highlightWord, y = 1.05) {
  slide.addText(line1, {
    x: 0.6, y, w: 11.8, h: 0.7,
    fontSize: 32, bold: true, color: TEXT_PRI, fontFace: 'Calibri',
  });
  if (line2) {
    const parts = highlightWord
      ? [
          { text: line2.replace(highlightWord, '').trim() + ' ', options: { color: TEXT_PRI } },
          { text: highlightWord, options: { color: ACCENT } },
        ]
      : [{ text: line2, options: { color: TEXT_PRI } }];
    slide.addText(parts, {
      x: 0.6, y: y + 0.68, w: 11.8, h: 0.7,
      fontSize: 32, bold: true, fontFace: 'Calibri',
    });
  }
}

// ── Helper: body text ─────────────────────────────────────────────────────────
function body(slide, text, x, y, w, h, opts = {}) {
  slide.addText(text, {
    x, y, w, h,
    fontSize: opts.size || 11, color: opts.color || TEXT_SEC,
    fontFace: 'Calibri', wrap: true, valign: 'top',
    bold: opts.bold || false,
    ...opts,
  });
}

// ── Helper: strip alpha suffix so pptxgenjs gets valid 6-digit hex ────────────
function hex6(c) { return (c || BORDER).slice(0, 6); }

// ── Helper: card ──────────────────────────────────────────────────────────────
function card(slide, x, y, w, h, opts = {}) {
  slide.addShape(pptx.ShapeType.roundRect, {
    x, y, w, h,
    fill: { color: hex6(opts.fill) || CARD_BG },
    line: { color: hex6(opts.border) || BORDER, width: 0.5 },
    rectRadius: 0.1,
  });
}

// ── Helper: stat card ─────────────────────────────────────────────────────────
function statCard(slide, x, y, num, label, color) {
  card(slide, x, y, 2.8, 1.5);
  slide.addText(num, {
    x: x + 0.1, y: y + 0.18, w: 2.6, h: 0.7,
    fontSize: 34, bold: true, color, fontFace: 'Calibri', align: 'center',
  });
  slide.addText(label, {
    x: x + 0.1, y: y + 0.9, w: 2.6, h: 0.5,
    fontSize: 10, color: TEXT_SEC, fontFace: 'Calibri', align: 'center', wrap: true,
  });
}

// ── Helper: feature row ───────────────────────────────────────────────────────
function featureRow(slide, x, y, icon, label, desc, color) {
  slide.addShape(pptx.ShapeType.roundRect, {
    x, y, w: 5.4, h: 0.72,
    fill: { color: CARD_BG }, line: { color: BORDER, width: 0.5 }, rectRadius: 0.08,
  });
  slide.addText(icon, { x: x + 0.12, y: y + 0.14, w: 0.5, h: 0.44, fontSize: 18 });
  slide.addText(label, { x: x + 0.7, y: y + 0.08, w: 4.5, h: 0.28, fontSize: 11, bold: true, color: TEXT_PRI, fontFace: 'Calibri' });
  slide.addText(desc,  { x: x + 0.7, y: y + 0.36, w: 4.5, h: 0.28, fontSize: 9.5, color: TEXT_SEC, fontFace: 'Calibri', wrap: true });
}


// ══════════════════════════════════════════════════════════════════════════════
// SLIDE 1 — TITLE
// ══════════════════════════════════════════════════════════════════════════════
{
  const s = pptx.addSlide();
  addBase(s);

  // Radial glow accent
  s.addShape(pptx.ShapeType.ellipse, {
    x: 8, y: -1, w: 7, h: 7,
    fill: { type: 'solid', color: ACCENT, transparency: 92 },
    line: { color: ACCENT, transparency: 100 },
  });

  if (hasFile(LOGO_PATH)) {
    s.addImage({ path: LOGO_PATH, x: 0.7, y: 0.6, w: 2.8, h: 0.65 });
  }

  s.addText('REPORTER', {
    x: 0.7, y: 1.55, w: 8, h: 1.0,
    fontSize: 54, bold: true, color: TEXT_PRI, fontFace: 'Calibri',
    charSpacing: 8,
  });
  s.addText('PLATFORM', {
    x: 0.7, y: 2.5, w: 8, h: 0.55,
    fontSize: 18, bold: true, color: CYAN, fontFace: 'Calibri',
    charSpacing: 5,
  });

  // Underline accent
  s.addShape(pptx.ShapeType.rect, {
    x: 0.7, y: 3.12, w: 5.5, h: 0.04,
    fill: { color: ACCENT }, line: { color: ACCENT },
  });

  s.addText('Strategic Vision · QC Technology Group · Ubisoft', {
    x: 0.7, y: 3.28, w: 10, h: 0.35,
    fontSize: 13, color: TEXT_SEC, fontFace: 'Calibri',
  });

  s.addText('CONFIDENTIAL · INTERNAL', {
    x: 0.7, y: 4.0, w: 4, h: 0.3,
    fontSize: 9, bold: true, color: RED, charSpacing: 2.5,
  });

  s.addText('The Defect Management & Crash Handling Platform\nfor the Gaming Industry', {
    x: 0.7, y: 4.6, w: 9, h: 0.8,
    fontSize: 15, color: TEXT_SEC, fontFace: 'Calibri', wrap: true,
  });

  addLogo(s);
}


// ══════════════════════════════════════════════════════════════════════════════
// SLIDE 2 — EXECUTIVE SUMMARY
// ══════════════════════════════════════════════════════════════════════════════
{
  const s = pptx.addSlide();
  addBase(s);
  eyebrow(s, 'Executive Summary');
  title(s, 'The Opportunity', 'In One Slide');

  const points = [
    { icon: '📌', t: '18 years', d: 'Bloomberg Reporter has been in production. It is trusted, embedded, and irreplaceable — running on every production machine at Ubisoft.' },
    { icon: '👥', t: '5,000+ users', d: 'Active users across all Ubisoft studios. No other QC tool has this reach. This is our moat.' },
    { icon: '⚠️', t: 'The problem', d: 'The legacy architecture makes innovation slow and painful. It is hard to maintain, hard to scale, and increasingly hard to justify to studios.' },
    { icon: '🏆', t: 'The insight', d: 'We don\'t need to replace Reporter — we need to evolve it into a platform. The platform always wins. Steam, Airbnb, Uber don\'t "make things." They own the layer everything else connects to.' },
    { icon: '🚀', t: 'The proposal', d: 'Reporter Web: an open, web-first platform for defect management, crash handling, and AI-assisted QC. Modular. Integration-ready. Scalable outside Ubisoft.' },
    { icon: '💰', t: 'The ceiling', d: 'Internal adoption → cross-studio standard → external licensing to other gaming studios. The IP, infrastructure, and user base already exist.' },
  ];

  points.forEach((p, i) => {
    const col = i % 2;
    const row = Math.floor(i / 2);
    const x = 0.6 + col * 6.2;
    const y = 2.05 + row * 1.45;
    card(s, x, y, 5.8, 1.3);
    s.addText(p.icon, { x: x + 0.15, y: y + 0.28, w: 0.55, h: 0.55, fontSize: 22 });
    s.addText(p.t, { x: x + 0.75, y: y + 0.1, w: 4.8, h: 0.35, fontSize: 12, bold: true, color: TEXT_PRI, fontFace: 'Calibri' });
    s.addText(p.d, { x: x + 0.75, y: y + 0.45, w: 4.8, h: 0.72, fontSize: 9.5, color: TEXT_SEC, fontFace: 'Calibri', wrap: true });
  });

  addLogo(s);
}


// ══════════════════════════════════════════════════════════════════════════════
// SLIDE 3 — THE LEGACY CHALLENGE
// ══════════════════════════════════════════════════════════════════════════════
{
  const s = pptx.addSlide();
  addBase(s);
  eyebrow(s, 'Bloomberg Reporter · 2006 → Today');
  title(s, '18 Years of Trust.', 'Time to Evolve.', 'Evolve.');

  // Stats row
  statCard(s, 0.6,  2.0, '18',     'Years in\nProduction', ACCENT);
  statCard(s, 3.6,  2.0, '5,000+', 'Active Users\nAll Studios', GREEN);
  statCard(s, 6.6,  2.0, '100%',   'Coverage on Every\nProduction Machine', PURPLE);
  statCard(s, 9.6,  2.0, 'Legacy', 'Architecture:\nNow a Bottleneck', AMBER);

  // Challenge block
  card(s, 0.6, 3.7, 5.8, 2.6, { border: AMBER + '55', fill: '1A1500' });
  s.addText('⚠  The Challenge', { x: 0.8, y: 3.85, w: 5.4, h: 0.35, fontSize: 12, bold: true, color: AMBER, fontFace: 'Calibri' });
  s.addText(
    'Bloomberg Reporter was designed for a different era. Its architecture makes every new integration expensive, every new feature slow to ship, and every bug fix a risk to a system thousands of people depend on daily. The tool that made us essential is now the thing holding us back.',
    { x: 0.8, y: 4.25, w: 5.4, h: 1.85, fontSize: 10.5, color: TEXT_SEC, fontFace: 'Calibri', wrap: true }
  );

  // Opportunity block
  card(s, 6.8, 3.7, 5.8, 2.6, { border: ACCENT + '55', fill: '0D1520' });
  s.addText('🚀  The Opportunity', { x: 7.0, y: 3.85, w: 5.4, h: 0.35, fontSize: 12, bold: true, color: ACCENT, fontFace: 'Calibri' });
  s.addText(
    'We have something no other team at Ubisoft has: 5,000 users, every production machine, 18 years of institutional trust. We don\'t need to start from scratch — we need to open the architecture and let the ecosystem grow around us. That\'s how you build a platform.',
    { x: 7.0, y: 4.25, w: 5.4, h: 1.85, fontSize: 10.5, color: TEXT_SEC, fontFace: 'Calibri', wrap: true }
  );

  addLogo(s);
}


// ══════════════════════════════════════════════════════════════════════════════
// SLIDE 4 — WHY PLATFORM WINS
// ══════════════════════════════════════════════════════════════════════════════
{
  const s = pptx.addSlide();
  addBase(s);
  eyebrow(s, 'Strategic Thinking');
  title(s, 'The Platform Always Wins', '');

  const analogies = [
    { icon: '🎮', name: 'Steam', sub: 'Doesn\'t make games', desc: 'Owns the distribution layer. Every game makes Steam more valuable, not less.' },
    { icon: '🏠', name: 'Airbnb', sub: 'Doesn\'t own rooms', desc: 'Owns the marketplace. Every host makes Airbnb more powerful, not the other way around.' },
    { icon: '🚗', name: 'Uber', sub: 'Doesn\'t own cars', desc: 'Owns the connection layer. Every driver and rider makes the network stronger.' },
  ];

  analogies.forEach((a, i) => {
    const x = 0.6 + i * 4.2;
    card(s, x, 1.8, 3.9, 3.0, { border: ACCENT + '44', fill: '0D1520' });
    s.addText(a.icon, { x: x + 0.15, y: 1.95, w: 0.8, h: 0.8, fontSize: 36 });
    s.addText(a.name, { x: x + 1.0, y: 2.0, w: 2.7, h: 0.42, fontSize: 18, bold: true, color: TEXT_PRI, fontFace: 'Calibri' });
    s.addText(a.sub,  { x: x + 1.0, y: 2.42, w: 2.7, h: 0.3, fontSize: 10, color: AMBER, fontFace: 'Calibri', bold: true });
    s.addText(a.desc, { x: x + 0.15, y: 2.88, w: 3.6, h: 1.6, fontSize: 10.5, color: TEXT_SEC, fontFace: 'Calibri', wrap: true });
  });

  // Arrow
  s.addText('→', { x: 4.3, y: 2.9, w: 0.5, h: 0.7, fontSize: 28, color: ACCENT, align: 'center' });
  s.addText('→', { x: 8.5, y: 2.9, w: 0.5, h: 0.7, fontSize: 28, color: ACCENT, align: 'center' });

  card(s, 0.6, 5.1, 12.4, 1.35, { border: ACCENT + '33', fill: '0D1520' });
  s.addText('Bloomberg Reporter Web', { x: 0.9, y: 5.22, w: 4, h: 0.38, fontSize: 13, bold: true, color: ACCENT, fontFace: 'Calibri' });
  s.addText(
    'We already have the infrastructure, the users, and the trust. Reporter Web is the integration layer that QA Companion, Eagle Eye, QC Assistant, DevConsoles, Jira, GitLab — and eventually external studios — all connect to. The platform doesn\'t "win" by being the best tool. It wins by being the layer everything else needs.',
    { x: 4.8, y: 5.22, w: 7.9, h: 1.1, fontSize: 10.5, color: TEXT_SEC, fontFace: 'Calibri', wrap: true }
  );

  addLogo(s);
}


// ══════════════════════════════════════════════════════════════════════════════
// SLIDE 5 — REPORTER WEB PLATFORM PILLARS
// ══════════════════════════════════════════════════════════════════════════════
{
  const s = pptx.addSlide();
  addBase(s);
  eyebrow(s, 'Reporter Web — Platform Architecture');
  title(s, 'Six Pillars of the Platform', '');

  const pillars = [
    { icon: '🐛', color: ACCENT,  title: 'Defect Management',  desc: 'Every bug, crash, and QC report in one structured, searchable place. Single source of truth across all studios.' },
    { icon: '💥', color: RED,     title: 'Crash Handling',     desc: 'Real-time crash ingestion, automatic triage, and routing. From dump to ticket in seconds.' },
    { icon: '🤖', color: GREEN,   title: 'AI Integration',     desc: 'QA Companion, QC Assistant, Eagle Eye plug in as modules — not separate tools.' },
    { icon: '🌐', color: PURPLE,  title: 'Portal & Web UI',    desc: 'Bloomberg Portal: web-first access for producers, directors, and stakeholders. No install needed.' },
    { icon: '📡', color: CYAN,    title: 'Integration Layer',  desc: 'Jira, GitLab, DevConsole, Perforce, Slack. The hub everything routes through.' },
    { icon: '📈', color: AMBER,   title: 'Scale Beyond',       desc: 'Internal → cross-studio standard → external licensing. The IP and infrastructure are already there.' },
  ];

  pillars.forEach((p, i) => {
    const col = i % 3;
    const row = Math.floor(i / 3);
    const x = 0.5 + col * 4.35;
    const y = 1.8 + row * 2.3;
    card(s, x, y, 4.1, 2.1);
    s.addShape(pptx.ShapeType.rect, { x, y, w: 4.1, h: 0.06, fill: { color: p.color }, line: { color: p.color } });
    s.addText(p.icon,  { x: x + 0.15, y: y + 0.22, w: 0.6, h: 0.6, fontSize: 24 });
    s.addText(p.title, { x: x + 0.8, y: y + 0.22, w: 3.1, h: 0.4, fontSize: 12, bold: true, color: TEXT_PRI, fontFace: 'Calibri' });
    s.addText(p.desc,  { x: x + 0.15, y: y + 0.78, w: 3.8, h: 1.1, fontSize: 9.5, color: TEXT_SEC, fontFace: 'Calibri', wrap: true });
  });

  addLogo(s);
}


// ══════════════════════════════════════════════════════════════════════════════
// SLIDE 6 — BLOOMBERG PORTAL (live product)
// ══════════════════════════════════════════════════════════════════════════════
{
  const s = pptx.addSlide();
  addBase(s);
  eyebrow(s, 'Live Product · In Production Today');
  title(s, 'Bloomberg Portal:', 'Web Access, Everywhere', 'Everywhere');

  const features = [
    { icon: '🟢', t: 'Live in Production',    d: 'tgdp.bloom.ubisoft.org — accessible to all Ubisoft users. No install required.' },
    { icon: '🔍', t: 'Filtered Occurrence View', d: 'Deep-link to any defect. See all occurrences, stack traces, platforms in one view.' },
    { icon: '📊', t: 'Cross-Studio Visibility', d: 'Directors and PMs monitor defect trends across all titles from a single URL.' },
    { icon: '🔗', t: 'Integration Ready',      d: 'Portal\'s data API is the foundation for QA Companion AI analysis and Eagle Eye detection.' },
    { icon: '🌐', t: 'Platform-First Design',  d: 'Portal becomes the primary UI as Reporter Web matures. Legacy client progressively decoupled.' },
    { icon: '🚀', t: 'Next: Reporter Web',     d: 'All new features ship web-first. Portal evolves into the main interface for all users.' },
  ];

  features.forEach((f, i) => {
    const col = i % 2;
    const row = Math.floor(i / 2);
    featureRow(s, 0.6 + col * 5.7, 2.0 + row * 1.0, f.icon, f.t, f.d);
  });

  // URL badge
  card(s, 0.6, 5.25, 12.4, 0.9, { border: GREEN + '44', fill: '0D1A0D' });
  s.addText('🌐  Live URL:', { x: 0.85, y: 5.4, w: 1.5, h: 0.35, fontSize: 11, bold: true, color: GREEN, fontFace: 'Calibri' });
  s.addText('https://tgdp.bloom.ubisoft.org', { x: 2.3, y: 5.4, w: 7, h: 0.35, fontSize: 11, color: ACCENT, fontFace: 'Courier New', bold: true });
  s.addText('Requires Ubisoft internal network', { x: 9.3, y: 5.48, w: 3.5, h: 0.28, fontSize: 9, color: TEXT_MUT, fontFace: 'Calibri' });

  addLogo(s);
}


// ══════════════════════════════════════════════════════════════════════════════
// SLIDE 7 — INTEGRATION ECOSYSTEM
// ══════════════════════════════════════════════════════════════════════════════
{
  const s = pptx.addSlide();
  addBase(s);
  eyebrow(s, 'Platform Ecosystem');
  title(s, 'Integration Layer:', 'Everything Plugs In', 'Everything');

  const items = [
    { status: 'LIVE',   color: GREEN,  icon: '✓', title: 'Jira 153',          desc: 'Bidirectional defect sync' },
    { status: 'LIVE',   color: GREEN,  icon: '✓', title: 'Bloomberg Portal',  desc: 'Web occurrence view' },
    { status: 'LIVE',   color: GREEN,  icon: '✓', title: 'GitLab CI/CD',      desc: 'MR & release tracking' },
    { status: 'NEXT',   color: ACCENT, icon: '→', title: 'QA Companion',      desc: 'AI triage & deduplication' },
    { status: 'NEXT',   color: ACCENT, icon: '→', title: 'Eagle Eye',         desc: 'Anomaly & spike detection' },
    { status: 'NEXT',   color: ACCENT, icon: '→', title: 'QC Assistant',      desc: 'Workflow AI & automation' },
    { status: 'NEXT',   color: ACCENT, icon: '→', title: 'Xray / Test Mgmt',  desc: 'Test ↔ defect link' },
    { status: 'FUTURE', color: TEXT_SEC,icon:'◦', title: 'DevConsole Crashes','desc': 'PlayStation / Xbox / Nintendo' },
    { status: 'FUTURE', color: TEXT_SEC,icon:'◦', title: 'CODEX Builds',      desc: 'Defect → build traceability' },
    { status: 'FUTURE', color: TEXT_SEC,icon:'◦', title: 'Slack / Teams',     desc: 'Smart alerting & digests' },
    { status: 'FUTURE', color: TEXT_SEC,icon:'◦', title: 'External Studios',  desc: 'Third-party API access' },
    { status: 'FUTURE', color: TEXT_SEC,icon:'◦', title: 'Perforce',          desc: 'Changelist ↔ bug traceability' },
  ];

  items.forEach((item, i) => {
    const col = i % 4;
    const row = Math.floor(i / 4);
    const x = 0.6 + col * 3.2;
    const y = 2.0 + row * 1.5;
    card(s, x, y, 3.0, 1.3, { border: item.color + '44' });
    s.addText(item.icon + ' ' + item.status, {
      x: x + 0.12, y: y + 0.1, w: 2.76, h: 0.28,
      fontSize: 8, bold: true, color: item.color, charSpacing: 1.5, fontFace: 'Calibri',
    });
    s.addText(item.title, { x: x + 0.12, y: y + 0.42, w: 2.76, h: 0.32, fontSize: 11, bold: true, color: TEXT_PRI, fontFace: 'Calibri' });
    s.addText(item.desc,  { x: x + 0.12, y: y + 0.74, w: 2.76, h: 0.44, fontSize: 9.5, color: TEXT_SEC, fontFace: 'Calibri', wrap: true });
  });

  addLogo(s);
}


// ══════════════════════════════════════════════════════════════════════════════
// SLIDE 8 — MEET BLOOM (AI Agent)
// ══════════════════════════════════════════════════════════════════════════════
{
  const s = pptx.addSlide();
  addBase(s);

  // Large radial glow behind avatar area
  s.addShape(pptx.ShapeType.ellipse, {
    x: 6.8, y: 0.5, w: 6.5, h: 6.5,
    fill: { type: 'solid', color: ACCENT, transparency: 93 },
    line: { color: ACCENT, transparency: 100 },
  });
  s.addShape(pptx.ShapeType.ellipse, {
    x: 7.4, y: 1.0, w: 5.5, h: 5.5,
    fill: { type: 'solid', color: PURPLE, transparency: 95 },
    line: { color: PURPLE, transparency: 100 },
  });

  eyebrow(s, 'Meet Bloom · Bloomberg AI Agent');
  title(s, 'Bloom:', 'Your Bloomberg Reporter AI', 'AI');

  // Avatar circle
  s.addShape(pptx.ShapeType.ellipse, {
    x: 8.8, y: 1.5, w: 3.5, h: 3.5,
    fill: { color: '0F1A2E' },
    line: { color: ACCENT, width: 2.5 },
  });
  // Inner glow ring
  s.addShape(pptx.ShapeType.ellipse, {
    x: 8.95, y: 1.65, w: 3.2, h: 3.2,
    fill: { type: 'solid', color: ACCENT, transparency: 90 },
    line: { color: ACCENT, transparency: 100 },
  });
  // Bloomberg B logo inside avatar
  if (hasFile(path.join(__dirname, 'public', 'Bloomberg_Logo_Black.png'))) {
    s.addImage({
      path: path.join(__dirname, 'public', 'Bloomberg_Logo_Black.png'),
      x: 9.55, y: 2.25, w: 2.0, h: 2.0,
    });
  }
  // "BLOOM" label under avatar
  s.addText('BLOOM', {
    x: 8.8, y: 5.15, w: 3.5, h: 0.45,
    fontSize: 14, bold: true, color: ACCENT, fontFace: 'Calibri',
    align: 'center', charSpacing: 5,
  });
  s.addText('Bloomberg AI · Online', {
    x: 8.8, y: 5.58, w: 3.5, h: 0.3,
    fontSize: 9, color: GREEN, fontFace: 'Calibri', align: 'center',
  });

  // Chat bubble
  s.addShape(pptx.ShapeType.roundRect, {
    x: 0.6, y: 2.0, w: 7.8, h: 1.55,
    fill: { color: '0D1520' },
    line: { color: ACCENT, width: 0.8 },
    rectRadius: 0.14,
  });
  // bubble tail (triangle pointing right)
  s.addShape(pptx.ShapeType.rtTriangle, {
    x: 8.35, y: 2.45, w: 0.3, h: 0.3,
    fill: { color: ACCENT },
    line: { color: ACCENT },
    rotate: 180,
  });
  s.addText([
    { text: 'Hi! I\'m Bloom ', options: { color: TEXT_PRI, bold: true } },
    { text: '— your Bloomberg Reporter AI assistant.\nI can help you navigate defect reports, crashes, Jira issues, and the full platform. Ask me anything!', options: { color: TEXT_SEC } },
  ], {
    x: 0.85, y: 2.15, w: 7.3, h: 1.25,
    fontSize: 11.5, fontFace: 'Calibri', wrap: true,
  });

  // Feature pills / what Bloom can do
  const blooms = [
    { icon: '🐛', label: 'Defect triage',        desc: 'AI-classifies incoming reports by severity, type, and likely owner' },
    { icon: '💥', label: 'Crash analysis',        desc: 'Reads dump files, suggests root cause, flags duplicates automatically' },
    { icon: '🔗', label: 'Jira sync',             desc: 'Creates and updates Jira tickets from natural language — no forms' },
    { icon: '📊', label: 'Portal navigation',     desc: 'Answers "how many P1s this sprint?" by querying Portal live' },
    { icon: '🤖', label: 'QA Companion bridge',   desc: 'Orchestrates QA Companion AI modules from a single chat interface' },
    { icon: '💬', label: 'Natural language',       desc: 'Just ask. No commands, no menus. Bloom figures out the rest.' },
  ];

  blooms.forEach((b, i) => {
    const col = i % 3;
    const row = Math.floor(i / 3);
    const x = 0.6 + col * 2.65;
    const y = 3.85 + row * 1.38;
    s.addShape(pptx.ShapeType.roundRect, {
      x, y, w: 2.5, h: 1.25,
      fill: { color: CARD_BG }, line: { color: hex6(ACCENT), width: 0.5 }, rectRadius: 0.1,
    });
    s.addText(b.icon,  { x: x + 0.1,  y: y + 0.1,  w: 0.55, h: 0.55, fontSize: 22 });
    s.addText(b.label, { x: x + 0.68, y: y + 0.1,  w: 1.72, h: 0.35, fontSize: 10.5, bold: true, color: TEXT_PRI,  fontFace: 'Calibri' });
    s.addText(b.desc,  { x: x + 0.1,  y: y + 0.58, w: 2.3,  h: 0.58, fontSize: 8.5,  color: TEXT_SEC, fontFace: 'Calibri', wrap: true });
  });

  addLogo(s);
}


// ══════════════════════════════════════════════════════════════════════════════
// SLIDE 9 — SUBSCRIPTION TIERS  (was 8)
// ══════════════════════════════════════════════════════════════════════════════
{
  const s = pptx.addSlide();
  addBase(s);
  eyebrow(s, 'Platform Tiers · Subscription Model');
  title(s, 'Three Tiers.', 'One Platform.', 'Platform.');

  const tiers = [
    {
      name: 'Vintage',
      color: AMBER,
      tag: 'CURRENT BASELINE',
      tagColor: AMBER,
      desc: 'Classic defect reporting. Reliability without complexity. The foundation every team starts from.',
      products: ['Bloomberg Reporter ✓', 'Bloomberg Portal (read-only) ✓', 'QA Companion  ✗', 'QC Assistant  ✗', 'Eagle Eye  ✗'],
      productColors: [GREEN, GREEN, TEXT_MUT, TEXT_MUT, TEXT_MUT],
    },
    {
      name: 'Basic',
      color: ACCENT,
      tag: 'RECOMMENDED · M5 TARGET',
      tagColor: ACCENT,
      desc: 'Full Reporter + AI-assisted triage. The right choice for active QA teams and leads.',
      products: ['Bloomberg Reporter ✓', 'Bloomberg Portal (full) ✓', 'QA Companion (AI) ✓', 'QC Assistant  ✗', 'Eagle Eye  ✗'],
      productColors: [GREEN, GREEN, GREEN, TEXT_MUT, TEXT_MUT],
    },
    {
      name: 'Premium',
      color: PURPLE,
      tag: 'LONG-TERM VISION',
      tagColor: PURPLE,
      desc: 'The full platform — every AI module, every integration. For teams leading the QC transformation.',
      products: ['Bloomberg Reporter ✓', 'Bloomberg Portal (full) ✓', 'QA Companion (AI) ✓', 'QC Assistant (AI) ✓', 'Eagle Eye (AI) ✓'],
      productColors: [GREEN, GREEN, GREEN, GREEN, GREEN],
    },
  ];

  tiers.forEach((tier, i) => {
    const x = 0.55 + i * 4.35;
    // Card
    s.addShape(pptx.ShapeType.roundRect, {
      x, y: 1.8, w: 4.1, h: 5.5,
      fill: { color: CARD_BG }, line: { color: tier.color, width: i === 1 ? 1.2 : 0.5 }, rectRadius: 0.12,
    });
    // Top bar
    s.addShape(pptx.ShapeType.rect, { x, y: 1.8, w: 4.1, h: 0.07, fill: { color: tier.color }, line: { color: tier.color } });
    // Tag
    s.addText(tier.tag, { x: x + 0.15, y: 1.93, w: 3.8, h: 0.25, fontSize: 7.5, bold: true, color: tier.tagColor, charSpacing: 1, fontFace: 'Calibri' });
    // Name
    s.addText(tier.name, { x: x + 0.15, y: 2.22, w: 3.8, h: 0.52, fontSize: 16, bold: true, color: TEXT_PRI, fontFace: 'Calibri' });
    // Desc
    s.addText(tier.desc, { x: x + 0.15, y: 2.78, w: 3.8, h: 0.88, fontSize: 9.5, color: TEXT_SEC, fontFace: 'Calibri', wrap: true });
    // Divider
    s.addShape(pptx.ShapeType.rect, { x: x + 0.15, y: 3.7, w: 3.8, h: 0.02, fill: { color: BORDER }, line: { color: BORDER } });
    // Products
    tier.products.forEach((p, pi) => {
      s.addText(p, { x: x + 0.2, y: 3.82 + pi * 0.55, w: 3.7, h: 0.44, fontSize: 10, color: tier.productColors[pi], fontFace: 'Calibri' });
    });
  });

  addLogo(s);
}


// ══════════════════════════════════════════════════════════════════════════════
// SLIDE 10 — VISION ROADMAP / NEXT STEPS  (was 9)
// ══════════════════════════════════════════════════════════════════════════════
{
  const s = pptx.addSlide();
  addBase(s);
  eyebrow(s, 'Proposed Next Steps');
  title(s, 'From Legacy to Platform:', 'The Path Forward', 'Path Forward');

  const steps = [
    { phase: 'NOW · M4', color: GREEN,  title: 'QA Companion Integration',    desc: 'Phase 1 + Phase 2 complete. AI triage live. Reporter ↔ QA Companion pipeline established.' },
    { phase: 'NOW · M4', color: GREEN,  title: 'Reporter Web Architecture',   desc: 'Begin decoupling legacy client. Bloomberg Portal as primary web interface. API-first design.' },
    { phase: 'M5',       color: ACCENT, title: 'Basic Tier Live',             desc: 'Full QA Companion integration. Reporter Web replaces 80% of legacy client use cases.' },
    { phase: 'M5',       color: ACCENT, title: 'Integration Layer Open',      desc: 'Xray, Eagle Eye, QC Assistant connect through standardised Reporter API.' },
    { phase: 'M6+',      color: PURPLE, title: 'Premium Tier',                desc: 'Full AI suite live. QC Assistant + Eagle Eye fully integrated. External studio pilots.' },
    { phase: 'FUTURE',   color: AMBER,  title: 'Scale Beyond Ubisoft',        desc: 'License to external gaming studios. The IP, data model, and infrastructure are the product.' },
  ];

  steps.forEach((step, i) => {
    const col = i % 2;
    const row = Math.floor(i / 2);
    const x = 0.6 + col * 6.5;
    const y = 2.0 + row * 1.5;
    card(s, x, y, 6.1, 1.3, { border: step.color + '44' });
    s.addText(step.phase, { x: x + 0.15, y: y + 0.1, w: 2, h: 0.28, fontSize: 8, bold: true, color: step.color, charSpacing: 1.5, fontFace: 'Calibri' });
    s.addText(step.title, { x: x + 0.15, y: y + 0.42, w: 5.8, h: 0.35, fontSize: 12, bold: true, color: TEXT_PRI, fontFace: 'Calibri' });
    s.addText(step.desc,  { x: x + 0.15, y: y + 0.77, w: 5.8, h: 0.45, fontSize: 9.5, color: TEXT_SEC, fontFace: 'Calibri', wrap: true });
  });

  addLogo(s);
}


// ══════════════════════════════════════════════════════════════════════════════
// SLIDE 11 — CLOSING / CALL TO ACTION  (was 10)
// ══════════════════════════════════════════════════════════════════════════════
{
  const s = pptx.addSlide();
  addBase(s);

  s.addShape(pptx.ShapeType.ellipse, {
    x: 7.5, y: -1, w: 8, h: 8,
    fill: { type: 'solid', color: ACCENT, transparency: 93 },
    line: { color: ACCENT, transparency: 100 },
  });

  if (hasFile(LOGO_PATH)) {
    s.addImage({ path: LOGO_PATH, x: 0.7, y: 0.6, w: 2.5, h: 0.58 });
  }

  s.addText('The question isn\'t whether\nBloomberg Reporter should evolve.', {
    x: 0.7, y: 1.4, w: 11, h: 1.5,
    fontSize: 30, bold: true, color: TEXT_PRI, fontFace: 'Calibri', wrap: true,
  });
  s.addText('The question is: do we build a better tool,\nor do we build the platform the industry runs on?', {
    x: 0.7, y: 2.9, w: 11, h: 1.4,
    fontSize: 22, color: TEXT_SEC, fontFace: 'Calibri', wrap: true,
  });

  s.addShape(pptx.ShapeType.rect, { x: 0.7, y: 4.35, w: 5.5, h: 0.04, fill: { color: ACCENT }, line: { color: ACCENT } });

  s.addText([
    { text: 'Reporter Web', options: { color: ACCENT, bold: true } },
    { text: ' is the answer. Not a rebuild — a transformation.\nFrom a tool to the platform the entire QC Technology group grows on.', options: { color: TEXT_SEC } },
  ], {
    x: 0.7, y: 4.5, w: 11, h: 1.0,
    fontSize: 13, fontFace: 'Calibri', wrap: true,
  });

  s.addText('Cristian Negoescu · Team Lead, Bloomberg Reporter · QC Technology Group · Ubisoft', {
    x: 0.7, y: 6.3, w: 11, h: 0.3,
    fontSize: 9.5, color: TEXT_MUT, fontFace: 'Calibri',
  });

  addLogo(s);
}


// ── Write file ────────────────────────────────────────────────────────────────
const OUT = path.join(__dirname, 'Bloomberg_Reporter_Platform.pptx');
pptx.writeFile({ fileName: OUT })
  .then(() => console.log(`\n✓ Saved: ${OUT}\n`))
  .catch(e => console.error('Error:', e));
