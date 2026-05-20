#!/usr/bin/env node
/**
 * Export species CSV + tagged Excel for one batch (local, no dev server).
 *
 * Usage:
 *   node scripts/export-batch-local.js --list P_E2
 *   node scripts/export-batch-local.js --batchId <ExcelData _id from dashboard URL>
 *   node scripts/export-batch-local.js --uploadGroupId <uuid>
 *
 * Optional:
 *   --out ./exports          (default: ./exports in project root)
 *
 * Requires MONGODB_URI in .env.local
 */

const fs = require('fs');
const path = require('path');
const mongoose = require('mongoose');
const XLSX = require('xlsx');

require('dotenv').config({ path: path.join(__dirname, '..', '.env.local') });

const FULLWIDTH_DOT = '\uff0e';

const TAG_COLUMNS = [
  'Blurry',
  'Low-light',
  'Body part',
  'Blends in',
  'Unidentifiable to taxonomic level by human ground-truth',
  'Other',
  'Similar species that does not occur in the area',
];

function imageBasename(p) {
  return String(p).replace(/^.*\//, '').trim();
}

function imagePathsForRow(row) {
  if (row.imagePaths?.length) return row.imagePaths;
  if (row.filenames) {
    return row.filenames
      .split(',')
      .map((x) => x.trim())
      .filter(Boolean);
  }
  return [];
}

function collectBasenamesFromDocs(docs) {
  const out = new Set();
  for (const doc of docs) {
    for (const row of doc.data || []) {
      for (const p of imagePathsForRow(row)) {
        const bn = imageBasename(p);
        if (bn) out.add(bn);
      }
    }
  }
  return out;
}

function inferYearFromBasename(basename) {
  const m = basename.match(/_(\d{2})\d{4}_/i);
  return m ? m[1] : null;
}

function inferDominantYearFromBasenames(basenames) {
  const counts = {};
  for (const bn of basenames) {
    const y = inferYearFromBasename(bn);
    if (y) counts[y] = (counts[y] || 0) + 1;
  }
  let best = null;
  let max = 0;
  for (const [y, c] of Object.entries(counts)) {
    if (c > max) {
      max = c;
      best = y;
    }
  }
  return best;
}

function filterBasenamesByYear(basenames, year) {
  const re = new RegExp(`_${year}\\d{4}_`, 'i');
  const out = new Set();
  for (const b of basenames) {
    if (re.test(b)) out.add(b);
  }
  return out;
}

function resolveExportImageScope(anchor, sessionDocs) {
  const anchorBasenames = collectBasenamesFromDocs([anchor]);
  const dominantYear = inferDominantYearFromBasenames(anchorBasenames);
  const allInSession = collectBasenamesFromDocs(sessionDocs);
  if (!dominantYear) return allInSession;
  const yearFiltered = filterBasenamesByYear(allInSession, dominantYear);
  return yearFiltered.size > 0 ? yearFiltered : allInSession;
}

function inferDeploymentIdFromImageBasename(basename) {
  const base = basename.replace(/^.*\//, '').replace(/\.[^.]+$/i, '');
  const m = base.match(/^(P_[A-Z]_\d+_\d{6})/i);
  if (m) return m[1];
  const m2 = base.match(/^(PA\d+_\d{6})/i);
  if (m2) return m2[1];
  return base;
}

function escapeCsvCell(value) {
  const s = String(value ?? '');
  if (/[",\n\r]/.test(s)) return `"${s.replace(/"/g, '""')}"`;
  return s;
}

function cleanTaxValue(value) {
  const s = String(value ?? '').trim();
  return s.toLowerCase() === 'nan' ? '' : s;
}

function taxonomicLevelFromSheetName(sheetName) {
  const s = sheetName.toLowerCase();
  if (s.includes('common')) return 'common_name';
  if (s.includes('species')) return 'species';
  if (s.includes('genus')) return 'genus';
  if (s.includes('family')) return 'family';
  if (s.includes('order')) return 'order';
  if (s.includes('class')) return 'class';
  return null;
}

function lookupGlobalSpecies(speciesByPath, bn) {
  if (!bn) return '';
  const direct = (key) => {
    if (!key) return '';
    const esc = key.replace(/\./g, FULLWIDTH_DOT);
    return speciesByPath[key] || speciesByPath[esc] || '';
  };
  let v = direct(bn);
  if (v) return String(v).trim();
  for (const key of Object.keys(speciesByPath)) {
    const val = speciesByPath[key];
    if (val == null || String(val).trim() === '') continue;
    const keyBase = key.replaceAll(FULLWIDTH_DOT, '.').replace(/^.*\//, '');
    if (keyBase === bn) return String(val).trim();
  }
  return '';
}

function mergeGlobalSpecies(sessionDocs, exportScope) {
  const merged = {};
  for (const doc of sessionDocs) {
    const g = doc.globalImageSpecies || {};
    for (const [k, v] of Object.entries(g)) {
      if (v != null && String(v).trim() !== '') merged[k] = String(v).trim();
    }
  }
  const out = {};
  for (const [key, val] of Object.entries(merged)) {
    const bn = key.replaceAll(FULLWIDTH_DOT, '.').replace(/^.*\//, '');
    if (exportScope.has(bn)) out[key] = val;
  }
  return out;
}

function buildSessionTagIndex(sessionDocs) {
  const index = new Map();
  for (const sheet of sessionDocs) {
    const bySheet = sheet.sheetSpecificImageTags || {};
    for (const map of Object.values(bySheet)) {
      if (!map || typeof map !== 'object') continue;
      for (const [rawKey, value] of Object.entries(map)) {
        if (!Array.isArray(value) || value.length === 0) continue;
        const bn = rawKey.replaceAll(FULLWIDTH_DOT, '.').replace(/^.*\//, '');
        if (bn) index.set(bn, value);
      }
    }
  }
  return index;
}

async function loadSessionSheets(anchor, ExcelData) {
  if (anchor.uploadGroupId) {
    return ExcelData.find({ uploadGroupId: anchor.uploadGroupId }).sort({
      sheetName: 1,
    });
  }
  const uploadedAt = anchor.uploadedAt;
  if (uploadedAt) {
    const windowMs = 5 * 60 * 1000;
    const start = new Date(uploadedAt.getTime() - windowMs);
    const end = new Date(uploadedAt.getTime() + windowMs);
    const nearby = await ExcelData.find({
      filename: anchor.filename,
      uploadedAt: { $gte: start, $lte: end },
    }).sort({ sheetName: 1 });
    if (nearby.length) return nearby;
  }
  return ExcelData.find({ filename: anchor.filename }).sort({ sheetName: 1 });
}

function generateCsv(sessionDocs, speciesByPath, exportScope) {
  const headers = [
    'deployment_id',
    'filename',
    'class',
    'order',
    'family',
    'genus',
    'species',
    'common_name',
  ];
  const byImage = new Map();

  for (const sheetDoc of sessionDocs) {
    const level = taxonomicLevelFromSheetName(sheetDoc.sheetName || '');
    if (!level) continue;
    for (const row of sheetDoc.data || []) {
      const human = cleanTaxValue(row.human);
      for (const rawPath of imagePathsForRow(row)) {
        const bn = imageBasename(rawPath);
        if (!bn || !exportScope.has(bn)) continue;
        if (!byImage.has(bn)) {
          byImage.set(bn, {
            deployment_id: inferDeploymentIdFromImageBasename(bn),
            filename: bn,
            class: '',
            order: '',
            family: '',
            genus: '',
            species: '',
            common_name: '',
          });
        }
        byImage.get(bn)[level] = human;
      }
    }
  }

  for (const rec of byImage.values()) {
    const sp = lookupGlobalSpecies(speciesByPath, rec.filename);
    if (sp) rec.species = sp;
  }

  const lines = [headers.join(',')];
  for (const rec of byImage.values()) {
    lines.push(
      [
        rec.deployment_id,
        rec.filename,
        rec.class,
        rec.order,
        rec.family,
        rec.genus,
        rec.species,
        rec.common_name,
      ]
        .map(escapeCsvCell)
        .join(',')
    );
  }
  return { csv: lines.join('\n'), rowCount: byImage.size };
}

function generateTaggedExcel(sessionDocs, tagIndex, exportScope) {
  const workbook = XLSX.utils.book_new();
  const uniqueSheets = sessionDocs
    .filter(
      (s, i, self) => i === self.findIndex((x) => x.sheetName === s.sheetName)
    )
    .sort((a, b) => (a.sheetName || '').localeCompare(b.sheetName || ''));

  let totalTaggedImages = 0;

  for (const sheet of uniqueSheets) {
    const groupedData = new Map();

    for (const row of sheet.data || []) {
      const key = `${row.human}_vs_${row.ai}`;
      if (!groupedData.has(key)) {
        groupedData.set(key, {
          human: row.human,
          ai: row.ai,
          taggedImages: Object.fromEntries(TAG_COLUMNS.map((t) => [t, []])),
        });
      }
      const group = groupedData.get(key);

      for (const imagePath of imagePathsForRow(row)) {
        const bn = imageBasename(imagePath);
        if (!exportScope.has(bn)) continue;
        const tags = tagIndex.get(bn);
        if (!tags?.length) continue;
        for (const tag of tags) {
          if (group.taggedImages[tag] && !group.taggedImages[tag].includes(imagePath)) {
            group.taggedImages[tag].push(imagePath);
            totalTaggedImages++;
          }
        }
      }
    }

    const worksheetData = [['Human', 'AI', ...TAG_COLUMNS, 'Notable images']];
    for (const group of groupedData.values()) {
      worksheetData.push([
        group.human,
        group.ai,
        ...TAG_COLUMNS.map((t) => group.taggedImages[t].join(', ')),
        '',
      ]);
    }

    const ws = XLSX.utils.aoa_to_sheet(worksheetData);
    ws['!cols'] = [
      { wch: 20 },
      { wch: 20 },
      ...TAG_COLUMNS.map(() => ({ wch: 35 })),
      { wch: 35 },
    ];
    XLSX.utils.book_append_sheet(workbook, ws, sheet.sheetName);
  }

  return { buffer: XLSX.write(workbook, { type: 'buffer', bookType: 'xlsx' }), totalTaggedImages };
}

function parseArgs(argv) {
  const args = { list: null, batchId: null, uploadGroupId: null, out: path.join(__dirname, '..', 'exports') };
  for (let i = 2; i < argv.length; i++) {
    const a = argv[i];
    if (a === '--list' && argv[i + 1]) {
      args.list = argv[++i];
    } else if (a.startsWith('--list=')) {
      args.list = a.split('=')[1];
    } else if (a.startsWith('--batchId=')) {
      args.batchId = a.split('=')[1];
    } else if (a === '--batchId' && argv[i + 1]) {
      args.batchId = argv[++i];
    } else if (a.startsWith('--uploadGroupId=')) {
      args.uploadGroupId = a.split('=')[1];
    } else if (a === '--uploadGroupId' && argv[i + 1]) {
      args.uploadGroupId = argv[++i];
    } else if (a.startsWith('--out=')) {
      args.out = a.split('=')[1];
    } else if (a === '--out' && argv[i + 1]) {
      args.out = argv[++i];
    }
  }
  return args;
}

async function listBatches(ExcelData, projectToken) {
  const rx = new RegExp(projectToken.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), 'i');
  const docs = await ExcelData.find({
    filename: rx,
    sheetName: 'Class',
  })
    .select('_id filename uploadGroupId uploadedAt data sheetSpecificImageTags')
    .sort({ uploadedAt: -1 })
    .lean();

  console.log(`\nBatches matching "${projectToken}" (Class sheets, newest first):\n`);
  console.log(
    'batchId'.padEnd(26),
    'uploaded'.padEnd(12),
    'images'.padEnd(8),
    'year'.padEnd(6),
    'tags(21)'.padEnd(10),
    'uploadGroupId'
  );
  console.log('-'.repeat(90));

  for (const d of docs) {
    let imgs = 0;
    for (const row of d.data || []) {
      for (const p of imagePathsForRow(row)) {
        imgs++;
      }
    }
    const dominant = inferDominantYearFromBasenames(collectBasenamesFromDocs([d]));
    let tag21 = 0;
    for (const map of Object.values(d.sheetSpecificImageTags || {})) {
      for (const [k, tags] of Object.entries(map || {})) {
        if (!tags?.length) continue;
        if (inferYearFromBasename(k.replaceAll(FULLWIDTH_DOT, '.')) === '21') tag21++;
      }
    }
    console.log(
      String(d._id).padEnd(26),
      (d.uploadedAt ? d.uploadedAt.toISOString().slice(0, 10) : '?').padEnd(12),
      String(imgs).padEnd(8),
      String(dominant || '?').padEnd(6),
      String(tag21).padEnd(10),
      (d.uploadGroupId || '(none)').slice(0, 36)
    );
  }
  console.log('\nCopy a batchId from the dashboard URL or above, then run:');
  console.log('  node scripts/export-batch-local.js --batchId <id>\n');
}

async function main() {
  const args = parseArgs(process.argv);

  if (!process.env.MONGODB_URI) {
    console.error('Missing MONGODB_URI in .env.local');
    process.exit(1);
  }

  await mongoose.connect(process.env.MONGODB_URI);
  const ExcelData =
    mongoose.models.ExcelData ||
    mongoose.model(
      'ExcelData',
      new mongoose.Schema({}, { strict: false }),
      'exceldatas'
    );

  if (args.list) {
    await listBatches(ExcelData, args.list);
    await mongoose.disconnect();
    return;
  }

  let anchor = null;
  if (args.batchId) {
    anchor = await ExcelData.findById(args.batchId);
  } else if (args.uploadGroupId) {
    anchor = await ExcelData.findOne({
      uploadGroupId: args.uploadGroupId,
      sheetName: 'Class',
    });
  } else {
    console.error('Provide --batchId <id> or --uploadGroupId <uuid> or --list P_E2');
    process.exit(1);
  }

  if (!anchor) {
    console.error('Batch not found.');
    process.exit(1);
  }

  const sessionDocs = await loadSessionSheets(anchor, ExcelData);
  const exportScope = resolveExportImageScope(anchor, sessionDocs);
  const dominantYear = inferDominantYearFromBasenames(collectBasenamesFromDocs([anchor]));
  const speciesByPath = mergeGlobalSpecies(sessionDocs, exportScope);
  const tagIndex = buildSessionTagIndex(sessionDocs);

  let tagsInScope = 0;
  for (const bn of exportScope) {
    if (tagIndex.has(bn)) tagsInScope++;
  }

  const { csv, rowCount } = generateCsv(sessionDocs, speciesByPath, exportScope);
  const { buffer, totalTaggedImages } = generateTaggedExcel(
    sessionDocs,
    tagIndex,
    exportScope
  );

  const safeName = anchor.filename.replace(/[^a-zA-Z0-9._-]+/g, '_').slice(0, 80);
  const baseName = anchor.filename.replace(/\.xlsx$/i, '');
  fs.mkdirSync(args.out, { recursive: true });

  const csvPath = path.join(
    args.out,
    `updated_species_${safeName}_from_sheet.csv`
  );
  const xlsxPath = path.join(
    args.out,
    `tagged_data_${baseName.replace(/[^a-zA-Z0-9._-]+/g, '_')}.xlsx`
  );

  fs.writeFileSync(csvPath, csv, 'utf8');
  fs.writeFileSync(xlsxPath, buffer);

  console.log('\nExport complete\n');
  console.log('  Batch:      ', anchor.filename);
  console.log('  Sheet:      ', anchor.sheetName);
  console.log('  Session:    ', anchor.uploadGroupId || '(by upload time)');
  console.log('  Image year: ', dominantYear || '(mixed)');
  console.log('  Images:     ', exportScope.size);
  console.log('  CSV rows:   ', rowCount);
  console.log('  Tagged imgs:', tagsInScope, '(unique images with tags in scope)');
  console.log('  Tag cells:  ', totalTaggedImages, '(rows in tag columns)');
  console.log('\n  CSV:  ', csvPath);
  console.log('  Excel:', xlsxPath);
  console.log('');

  if (tagsInScope === 0) {
    console.warn(
      'Warning: no image tags on filenames for this year/session. Excel tag columns will be empty.'
    );
    console.warn(
      'Try --list P_E2 and pick a batchId with tags(21) > 0, or another uploadGroupId.\n'
    );
  }

  await mongoose.disconnect();
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});
