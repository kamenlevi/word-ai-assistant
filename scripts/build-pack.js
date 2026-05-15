#!/usr/bin/env node
/* build-pack.js — MAINTAINER-ONLY script.
 *
 * Pulls images from Pexels / Unsplash / Pixabay APIs to populate the local
 * image asset pack. Run with:
 *   PEXELS_KEY=xxx UNSPLASH_KEY=yyy PIXABAY_KEY=zzz node scripts/build-pack.js
 *
 * Output: assets/pack/manifest.json (rewritten) + assets/pack/*.jpg files.
 *
 * End users just click "Download asset pack" in Settings; this script is for
 * the project maintainer to grow the manifest over time.
 */

const fs = require('fs');
const path = require('path');
const https = require('https');
const crypto = require('crypto');

const PEXELS_KEY   = process.env.PEXELS_KEY;
const UNSPLASH_KEY = process.env.UNSPLASH_KEY;
const PIXABAY_KEY  = process.env.PIXABAY_KEY;

const PACK_DIR = path.join(__dirname, '..', 'assets', 'pack');
const MANIFEST = path.join(PACK_DIR, 'manifest.json');
fs.mkdirSync(PACK_DIR, { recursive: true });

const CATEGORIES = [
  // ── 14 categories reused from PPT ─────────────────────────────────────
  { name: 'people',                queries: ['team meeting', 'business people', 'diverse workplace', 'professional portrait'] },
  { name: 'technology',            queries: ['laptop code', 'data center', 'circuit board', 'modern computer'] },
  { name: 'finance',               queries: ['stock chart', 'business growth', 'money calculator', 'finance dashboard'] },
  { name: 'nature',                queries: ['forest landscape', 'mountain sunrise', 'ocean horizon', 'green hills'] },
  { name: 'charts-diagrams',       queries: ['analytics dashboard', 'data visualization', 'business graphs'] },
  { name: 'abstract-backgrounds',  queries: ['blue gradient', 'geometric pattern', 'minimal abstract', 'modern texture'] },
  { name: 'transport',             queries: ['highway car', 'airplane sky', 'modern train', 'shipping logistics'] },
  { name: 'food',                  queries: ['coffee cup', 'healthy meal', 'restaurant table', 'fresh produce'] },
  { name: 'education',             queries: ['library books', 'classroom learning', 'graduation', 'study notes'] },
  { name: 'healthcare',            queries: ['medical stethoscope', 'hospital corridor', 'doctor consultation', 'wellness'] },
  { name: 'cities',                queries: ['skyline night', 'urban architecture', 'street modern', 'metropolitan'] },
  { name: 'office-workplace',      queries: ['minimal desk', 'open plan office', 'home office', 'modern workspace'] },
  { name: 'travel',                queries: ['world map', 'airport terminal', 'destination beach', 'adventure landscape'] },
  { name: 'sports',                queries: ['running athletic', 'team huddle', 'fitness motion', 'stadium crowd'] },

  // ── 4 NEW Word-specific categories ────────────────────────────────────
  { name: 'document-icons',        queries: ['pdf icon', 'contract signed', 'certificate scroll', 'briefcase legal', 'gavel court'] },
  { name: 'watermarks',            queries: ['draft stamp', 'confidential stamp', 'approved stamp', 'copy stamp', 'sample stamp'] },
  { name: 'letterhead-graphics',   queries: ['letterhead banner', 'corporate header', 'monogram simple', 'two-tone bar'] },
  { name: 'page-borders',          queries: ['laurel wreath border', 'classic frame border', 'art deco border', 'academic frame'] },
];

const PER_QUERY = 5;
const MAX_PER_CAT = 50;

function get(url, headers) {
  return new Promise((resolve, reject) => {
    https.get(url, { headers: headers || {} }, res => {
      let data = '';
      res.on('data', c => data += c);
      res.on('end', () => {
        if (res.statusCode >= 400) return reject(new Error(`HTTP ${res.statusCode}: ${data.slice(0, 200)}`));
        resolve(data);
      });
      res.on('error', reject);
    }).on('error', reject);
  });
}

function download(url, dest) {
  return new Promise((resolve, reject) => {
    https.get(url, res => {
      if (res.statusCode >= 300 && res.statusCode < 400 && res.headers.location) {
        return download(res.headers.location, dest).then(resolve, reject);
      }
      if (res.statusCode >= 400) return reject(new Error('HTTP ' + res.statusCode));
      const f = fs.createWriteStream(dest);
      res.pipe(f);
      f.on('finish', () => f.close(() => resolve()));
      f.on('error', reject);
    }).on('error', reject);
  });
}

async function pexelsSearch(query, perPage) {
  if (!PEXELS_KEY) return [];
  const url = `https://api.pexels.com/v1/search?query=${encodeURIComponent(query)}&per_page=${perPage}&orientation=landscape`;
  const data = JSON.parse(await get(url, { Authorization: PEXELS_KEY }));
  return (data.photos || []).map(p => ({
    source: p.src.large || p.src.original,
    w: p.width, h: p.height,
    attribution: `Photo by ${p.photographer} (Pexels)`,
    license: 'Pexels License',
  }));
}

async function unsplashSearch(query, perPage) {
  if (!UNSPLASH_KEY) return [];
  const url = `https://api.unsplash.com/search/photos?query=${encodeURIComponent(query)}&per_page=${perPage}&orientation=landscape`;
  const data = JSON.parse(await get(url, { Authorization: `Client-ID ${UNSPLASH_KEY}` }));
  return (data.results || []).map(p => ({
    source: p.urls.regular,
    w: p.width, h: p.height,
    attribution: `Photo by ${p.user.name} (Unsplash)`,
    license: 'Unsplash License',
  }));
}

async function pixabaySearch(query, perPage) {
  if (!PIXABAY_KEY) return [];
  const url = `https://pixabay.com/api/?key=${PIXABAY_KEY}&q=${encodeURIComponent(query)}&image_type=photo&orientation=horizontal&per_page=${perPage}`;
  const data = JSON.parse(await get(url));
  return (data.hits || []).map(p => ({
    source: p.largeImageURL,
    w: p.imageWidth, h: p.imageHeight,
    attribution: `Image by ${p.user} (Pixabay)`,
    license: 'Pixabay License',
  }));
}

async function main() {
  if (!PEXELS_KEY && !UNSPLASH_KEY && !PIXABAY_KEY) {
    console.error('At least one of PEXELS_KEY, UNSPLASH_KEY, PIXABAY_KEY env vars must be set.');
    process.exit(1);
  }

  const existing = (() => { try { return JSON.parse(fs.readFileSync(MANIFEST, 'utf8')); } catch { return { images: [] }; } })();
  const existingIds = new Set(existing.images.map(i => i.id));
  const newImages = [...existing.images];

  for (const cat of CATEGORIES) {
    let counter = newImages.filter(i => i.category === cat.name).length;
    for (const q of cat.queries) {
      if (counter >= MAX_PER_CAT) break;
      const candidates = [
        ...await pexelsSearch(q, PER_QUERY).catch(() => []),
        ...await unsplashSearch(q, PER_QUERY).catch(() => []),
        ...await pixabaySearch(q, PER_QUERY).catch(() => []),
      ];
      for (const c of candidates) {
        if (counter >= MAX_PER_CAT) break;
        const id = `${cat.name}-${String(counter+1).padStart(3, '0')}`;
        if (existingIds.has(id)) { counter++; continue; }
        const tags = q.split(/\s+/).filter(Boolean);
        const dest = path.join(PACK_DIR, id + '.jpg');
        try {
          if (!fs.existsSync(dest)) await download(c.source, dest);
          const buf = fs.readFileSync(dest);
          const sha256 = crypto.createHash('sha256').update(buf).digest('hex');
          newImages.push({
            id, category: cat.name, tags,
            style: 'photo',
            w: c.w, h: c.h, ext: '.jpg',
            sha256,
            source: c.source,
            license: c.license,
            attribution: c.attribution,
          });
          counter++;
          console.log('+', id);
        } catch (e) {
          console.warn('skip', q, '—', e.message);
        }
      }
    }
  }

  const out = {
    version: 'v1.0.0',
    tarballUrl: null,
    license: 'See per-image license / attribution.',
    sources: ['pexels.com', 'unsplash.com', 'pixabay.com'],
    categories: CATEGORIES.map(c => c.name),
    images: newImages,
  };
  fs.writeFileSync(MANIFEST, JSON.stringify(out, null, 2));
  console.log(`\nDone. ${newImages.length} total images across ${CATEGORIES.length} categories.`);
  console.log('Manifest written to ' + MANIFEST);
}

main().catch(e => { console.error(e); process.exit(1); });
