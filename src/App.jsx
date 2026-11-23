import React, { useState, useEffect, useMemo, useRef, useCallback } from 'react';
import {
  Building2,
  MapPin,
  Droplet,
  Plus,
  Trash2,
  Edit3,
  History,
  Settings,
  Sun,
  Moon,
  CheckCircle2,
  Check,
  X,
  Activity,
  FileText,
  ChevronLeft,
  ChevronDown,
  ChevronUp,
  AlertTriangle,
  Gauge,
  ShieldCheck,
  Clock,
  Wrench,
  LayoutDashboard,
  Table,
  UserCheck,
  Search,
  RotateCcw,
  ArrowUpRight,
  Layers,
  Menu,
  Calendar,
  SlidersHorizontal,
  Download,
  Upload,
  FileSpreadsheet,
  Bell,
  QrCode,
  Printer,
  Camera,
  Image as ImageIcon,
  Filter,
  ArrowUpDown,
  ZoomIn
} from 'lucide-react';

/**
 * Water Dispensers & Tanks Tracker - V4.0.0 PRO (Field Operations & Offline Suite)
 * -----------------------------------------------------------------------------------------
 * Full-featured SaaS application shell in RTL Hebrew:
 * - Persistent SaaS Sidebar Navigation + Sticky Glassmorphic Top Bar
 * - Smart Action Center (מרכז התראות ודברים דחופים) for expired/expiring tags, <20% filters, & faults
 * - Quick Technician & Escort Name History chips inside Filter Renewal modal
 * - Zero-dependency Offline QR Code Label Generator (ISO/IEC 18004) & Deep-Linking per Dispenser
 * - Official A4 Printable / PDF Inspection Report (multi-page print-safe layout)
 * - Advanced Dashboard Filtering (Faulty / Expired / Area) & Sorting (Lowest Filter % / Nearest Expiry)
 * - Mobile Camera Capture & Offline Compressed Photo Documentation per Dispenser
 * - Native Excel (.xlsx) Export, UTF-8 CSV & Sparse SpreadsheetML Fallback + Deduplicated Import
 */

const STORAGE_KEY = 'aquaops_water_tracker_v3_state';
const SHEETJS_CDN_URL = 'https://cdn.sheetjs.com/xlsx-0.20.3/package/dist/xlsx.full.min.js';

// Helper: Extract local YYYY-MM-DD without UTC timezone shift
const getLocalIsoDate = (dateObj = new Date()) => {
  const year = dateObj.getFullYear();
  const month = String(dateObj.getMonth() + 1).padStart(2, '0');
  const day = String(dateObj.getDate()).padStart(2, '0');
  return `${year}-${month}-${day}`;
};

// Helper: Safely parse YYYY-MM-DD in local time and reject calendar overflow (e.g. Feb 31)
const parseLocalIsoDate = (isoDate) => {
  if (!isoDate || typeof isoDate !== 'string') return null;
  const match = isoDate.trim().match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
  if (!match) return null;
  const year = Number(match[1]);
  const month = Number(match[2]);
  const day = Number(match[3]);
  if (month < 1 || month > 12 || day < 1 || day > 31) return null;
  const parsed = new Date(year, month - 1, day);
  if (
    isNaN(parsed.getTime()) ||
    parsed.getFullYear() !== year ||
    parsed.getMonth() !== month - 1 ||
    parsed.getDate() !== day
  ) {
    return null;
  }
  return parsed;
};

// Helper: Calculate days from today to a YYYY-MM-DD date string
const getDaysUntilDate = (isoDate) => {
  const target = parseLocalIsoDate(isoDate);
  if (!target) return null;
  const now = new Date();
  const todayMidnight = new Date(now.getFullYear(), now.getMonth(), now.getDate());
  return Math.round((target.getTime() - todayMidnight.getTime()) / (1000 * 60 * 60 * 24));
};

// Helper: Format ISO YYYY-MM-DD to DD/MM/YYYY
const formatDateIL = (isoDate) => {
  if (!isoDate) return '';
  const parts = String(isoDate).split('-');
  if (parts.length !== 3) return String(isoDate);
  return `${parts[2]}/${parts[1]}/${parts[0]}`;
};

// Helper: Safely normalize percentage values (treats empty/null cells as fallback instead of 0)
const toPercent = (val, fallback = 100) => {
  if (val === null || val === undefined || String(val).trim() === '') return fallback;
  const cleaned = String(val).replace('%', '').trim();
  const parsed = Number(cleaned);
  return isNaN(parsed) ? fallback : Math.max(0, Math.min(100, Math.round(parsed)));
};

// Helper: Normalize various date representations (Excel serial, DD/MM/YYYY, YYYY-MM-DD, Date) to YYYY-MM-DD
const normalizeImportedDate = (rawVal) => {
  if (rawVal === null || rawVal === undefined) return '';
  if (rawVal instanceof Date && !isNaN(rawVal.getTime())) {
    return getLocalIsoDate(rawVal);
  }
  if (typeof rawVal === 'number' && rawVal > 20000 && rawVal < 100000) {
    const utcDays = Math.floor(rawVal - 25569);
    const utcMs = utcDays * 86400 * 1000;
    const d = new Date(utcMs);
    const year = d.getUTCFullYear();
    const month = String(d.getUTCMonth() + 1).padStart(2, '0');
    const day = String(d.getUTCDate()).padStart(2, '0');
    const candidate = `${year}-${month}-${day}`;
    return parseLocalIsoDate(candidate) ? candidate : '';
  }
  const str = String(rawVal).trim();
  if (!str || str === '-' || str === 'טרם בוצע' || str === '__/__/____') return '';

  const isoMatch = str.match(/^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})/);
  if (isoMatch) {
    const candidate = `${isoMatch[1]}-${isoMatch[2].padStart(2, '0')}-${isoMatch[3].padStart(2, '0')}`;
    return parseLocalIsoDate(candidate) ? candidate : '';
  }

  const ilMatch = str.match(/^(\d{1,2})[-/.](\d{1,2})[-/.](\d{4})/);
  if (ilMatch) {
    const candidate = `${ilMatch[3]}-${ilMatch[2].padStart(2, '0')}-${ilMatch[1].padStart(2, '0')}`;
    return parseLocalIsoDate(candidate) ? candidate : '';
  }

  return '';
};

// Helper: Dynamically load SheetJS (XLSX) library on demand with clean error recovery
const loadXLSXLibrary = () => {
  return new Promise((resolve, reject) => {
    if (typeof window !== 'undefined' && window.XLSX) {
      resolve(window.XLSX);
      return;
    }
    const existingScript = document.querySelector(`script[src="${SHEETJS_CDN_URL}"]`);
    if (existingScript) {
      if (existingScript.dataset.status === 'loaded' && window.XLSX) {
        resolve(window.XLSX);
        return;
      }
      if (existingScript.dataset.status === 'error') {
        existingScript.remove();
      } else {
        existingScript.addEventListener(
          'load',
          () => {
            if (window.XLSX) resolve(window.XLSX);
            else reject(new Error('XLSX global unavailable'));
          },
          { once: true }
        );
        existingScript.addEventListener('error', () => reject(new Error('XLSX load failed')), {
          once: true
        });
        return;
      }
    }
    const script = document.createElement('script');
    script.src = SHEETJS_CDN_URL;
    script.async = true;
    script.onload = () => {
      script.dataset.status = 'loaded';
      if (window.XLSX) resolve(window.XLSX);
      else {
        script.remove();
        reject(new Error('XLSX global unavailable'));
      }
    };
    script.onerror = () => {
      script.dataset.status = 'error';
      script.remove();
      reject(new Error('XLSX CDN unreachable'));
    };
    document.head.appendChild(script);
  });
};

// Helper: Escape XML special characters for offline SpreadsheetML (.xls) fallback
const escapeXml = (val) =>
  String(val ?? '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');

// Helper: Load persisted state safely from localStorage
const loadSavedState = () => {
  try {
    const raw = window.localStorage.getItem(STORAGE_KEY);
    if (!raw) return null;
    return JSON.parse(raw);
  } catch {
    return null;
  }
};

// Helper: Compress & resize mobile camera / gallery image via HTML5 Canvas for lightweight offline storage
const compressImageFile = (file, maxDimension = 800, quality = 0.66) => {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (event) => {
      const result = event.target?.result;
      if (typeof result !== 'string') {
        reject(new Error('Invalid image data'));
        return;
      }
      const img = new Image();
      img.onload = () => {
        let { width, height } = img;
        if (width > maxDimension || height > maxDimension) {
          if (width > height) {
            height = Math.round((height * maxDimension) / width);
            width = maxDimension;
          } else {
            width = Math.round((width * maxDimension) / height);
            height = maxDimension;
          }
        }
        const canvas = document.createElement('canvas');
        canvas.width = width;
        canvas.height = height;
        const ctx = canvas.getContext('2d');
        if (!ctx) {
          reject(new Error('Canvas context unavailable'));
          return;
        }
        ctx.fillStyle = '#FFFFFF';
        ctx.fillRect(0, 0, width, height);
        ctx.drawImage(img, 0, 0, width, height);
        const compressedDataUrl = canvas.toDataURL('image/jpeg', quality);
        resolve(compressedDataUrl);
      };
      img.onerror = () => reject(new Error('Image decode failed'));
      img.src = result;
    };
    reader.onerror = () => reject(new Error('File read failed'));
    reader.readAsDataURL(file);
  });
};

// ============================================================================
// ZERO-DEPENDENCY OFFLINE QR CODE GENERATOR (ISO/IEC 18004 Byte Mode, V1-V10)
// ============================================================================
const QR_VERSION_SPECS = [
  null,
  { size: 21, totalData: 19, blocks: [[19, 7]], align: [] }, // V1
  { size: 25, totalData: 34, blocks: [[34, 10]], align: [6, 18] }, // V2
  { size: 29, totalData: 55, blocks: [[55, 15]], align: [6, 22] }, // V3
  { size: 33, totalData: 80, blocks: [[80, 20]], align: [6, 26] }, // V4
  { size: 37, totalData: 108, blocks: [[108, 26]], align: [6, 30] }, // V5
  { size: 41, totalData: 136, blocks: [[68, 18], [68, 18]], align: [6, 34] }, // V6
  { size: 45, totalData: 156, blocks: [[78, 20], [78, 20]], align: [6, 22, 38], verBits: 0x07c94 }, // V7
  { size: 49, totalData: 194, blocks: [[97, 24], [97, 24]], align: [6, 24, 42], verBits: 0x085bc }, // V8
  { size: 53, totalData: 232, blocks: [[116, 30], [116, 30]], align: [6, 26, 46], verBits: 0x09a99 }, // V9
  {
    size: 57,
    totalData: 274,
    blocks: [
      [68, 18],
      [68, 18],
      [69, 18],
      [69, 18]
    ],
    align: [6, 28, 50],
    verBits: 0x0a4d3
  } // V10
];

const QR_FORMAT_BITS_L = [0x77c4, 0x72f3, 0x7daa, 0x789d, 0x662f, 0x6318, 0x6c41, 0x6976];

const gfMultiply = (x, y) => {
  let z = 0;
  for (let i = 7; i >= 0; i--) {
    z = (z << 1) ^ ((z >> 7) * 0x11d);
    z ^= ((y >> i) & 1) * x;
  }
  return z;
};

const computeReedSolomonEC = (dataBytes, ecCount) => {
  let gen = [1];
  let root = 1;
  for (let i = 0; i < ecCount; i++) {
    const nextGen = new Array(gen.length + 1).fill(0);
    for (let j = 0; j < gen.length; j++) {
      nextGen[j] ^= gen[j];
      nextGen[j + 1] ^= gfMultiply(gen[j], root);
    }
    gen = nextGen;
    root = gfMultiply(root, 2);
  }

  const remainder = new Array(ecCount).fill(0);
  for (let i = 0; i < dataBytes.length; i++) {
    const factor = dataBytes[i] ^ remainder.shift();
    remainder.push(0);
    for (let j = 0; j < ecCount; j++) {
      remainder[j] ^= gfMultiply(gen[j + 1], factor);
    }
  }
  return remainder;
};

const generateQrMatrix = (text) => {
  if (!text || typeof text !== 'string') return null;
  const bytes = new TextEncoder().encode(text);
  let version = 1;
  while (version <= 10) {
    const countBits = version <= 9 ? 8 : 16;
    const capacityBits = QR_VERSION_SPECS[version].totalData * 8;
    if (4 + countBits + bytes.length * 8 <= capacityBits) break;
    version++;
  }
  if (version > 10) return null;

  const spec = QR_VERSION_SPECS[version];
  const N = spec.size;
  const countBits = version <= 9 ? 8 : 16;

  // Build data bit stream (Byte mode 0100)
  const bits = [];
  const pushBits = (val, len) => {
    for (let i = len - 1; i >= 0; i--) {
      bits.push((val >> i) & 1);
    }
  };
  pushBits(0b0100, 4);
  pushBits(bytes.length, countBits);
  bytes.forEach((b) => pushBits(b, 8));

  const maxDataBits = spec.totalData * 8;
  const terminatorLen = Math.min(4, maxDataBits - bits.length);
  pushBits(0, terminatorLen);
  while (bits.length % 8 !== 0) bits.push(0);

  const dataCodewords = [];
  for (let i = 0; i < bits.length; i += 8) {
    let byte = 0;
    for (let j = 0; j < 8; j++) byte = (byte << 1) | bits[i + j];
    dataCodewords.push(byte);
  }
  let padToggle = true;
  while (dataCodewords.length < spec.totalData) {
    dataCodewords.push(padToggle ? 0xec : 0x11);
    padToggle = !padToggle;
  }

  // Split into RS blocks & compute EC codewords
  const dataBlocks = [];
  const ecBlocks = [];
  let offset = 0;
  for (const [blockDataLen, blockEcLen] of spec.blocks) {
    const blockData = dataCodewords.slice(offset, offset + blockDataLen);
    offset += blockDataLen;
    dataBlocks.push(blockData);
    ecBlocks.push(computeReedSolomonEC(blockData, blockEcLen));
  }

  // Interleave data & EC codewords
  const finalCodewords = [];
  const maxBlockData = Math.max(...dataBlocks.map((b) => b.length));
  for (let i = 0; i < maxBlockData; i++) {
    for (const blk of dataBlocks) {
      if (i < blk.length) finalCodewords.push(blk[i]);
    }
  }
  const maxBlockEc = Math.max(...ecBlocks.map((b) => b.length));
  for (let i = 0; i < maxBlockEc; i++) {
    for (const blk of ecBlocks) {
      if (i < blk.length) finalCodewords.push(blk[i]);
    }
  }

  const allBits = [];
  for (const cw of finalCodewords) {
    for (let i = 7; i >= 0; i--) allBits.push((cw >> i) & 1);
  }

  // Initialize module grid
  const modules = Array.from({ length: N }, () => new Array(N).fill(false));
  const isFunction = Array.from({ length: N }, () => new Array(N).fill(false));
  const setFunc = (r, c, dark) => {
    if (r >= 0 && r < N && c >= 0 && c < N) {
      modules[r][c] = dark;
      isFunction[r][c] = true;
    }
  };

  // Finder patterns + separators
  const drawFinder = (cr, cc) => {
    for (let dy = -4; dy <= 4; dy++) {
      for (let dx = -4; dx <= 4; dx++) {
        const dist = Math.max(Math.abs(dx), Math.abs(dy));
        setFunc(cr + dy, cc + dx, dist !== 2 && dist !== 4);
      }
    }
  };
  drawFinder(3, 3);
  drawFinder(3, N - 4);
  drawFinder(N - 4, 3);

  // Alignment patterns
  for (const cy of spec.align) {
    for (const cx of spec.align) {
      if (
        (cy === 6 && cx === 6) ||
        (cy === 6 && cx === N - 7) ||
        (cy === N - 7 && cx === 6)
      ) {
        continue;
      }
      for (let dy = -2; dy <= 2; dy++) {
        for (let dx = -2; dx <= 2; dx++) {
          const dist = Math.max(Math.abs(dx), Math.abs(dy));
          setFunc(cy + dy, cx + dx, dist !== 1);
        }
      }
    }
  }

  // Timing patterns
  for (let i = 8; i <= N - 9; i++) {
    setFunc(6, i, i % 2 === 0);
    setFunc(i, 6, i % 2 === 0);
  }

  // Dark module & reserve format info
  setFunc(N - 8, 8, true);
  for (let i = 0; i <= 8; i++) {
    isFunction[8][i] = true;
    isFunction[i][8] = true;
  }
  for (let i = 0; i < 8; i++) isFunction[8][N - 8 + i] = true;
  for (let i = 0; i < 7; i++) isFunction[N - 7 + i][8] = true;

  // Version info bits (V7+)
  if (spec.verBits) {
    for (let i = 0; i < 18; i++) {
      const bit = Boolean((spec.verBits >> i) & 1);
      const r = Math.floor(i / 3);
      const c = N - 11 + (i % 3);
      setFunc(r, c, bit);
      setFunc(c, r, bit);
    }
  }

  // Place data bits in zig-zag pattern
  let bitIdx = 0;
  let upward = true;
  for (let right = N - 1; right >= 1; right -= 2) {
    if (right === 6) right = 5;
    for (let vert = 0; vert < N; vert++) {
      const r = upward ? N - 1 - vert : vert;
      for (let j = 0; j < 2; j++) {
        const c = right - j;
        if (!isFunction[r][c]) {
          modules[r][c] = bitIdx < allBits.length ? Boolean(allBits[bitIdx]) : false;
          bitIdx++;
        }
      }
    }
    upward = !upward;
  }

  // Evaluate mask patterns 0..7 & place format info (ISO/IEC 18004 MSB->LSB order)
  const applyMaskAndFormat = (grid, maskId) => {
    for (let r = 0; r < N; r++) {
      for (let c = 0; c < N; c++) {
        if (isFunction[r][c]) continue;
        let invert = false;
        switch (maskId) {
          case 0:
            invert = (r + c) % 2 === 0;
            break;
          case 1:
            invert = r % 2 === 0;
            break;
          case 2:
            invert = c % 3 === 0;
            break;
          case 3:
            invert = (r + c) % 3 === 0;
            break;
          case 4:
            invert = (Math.floor(r / 2) + Math.floor(c / 3)) % 2 === 0;
            break;
          case 5:
            invert = ((r * c) % 2) + ((r * c) % 3) === 0;
            break;
          case 6:
            invert = (((r * c) % 2) + ((r * c) % 3)) % 2 === 0;
            break;
          case 7:
            invert = (((r + c) % 2) + ((r * c) % 3)) % 2 === 0;
            break;
          default:
            break;
        }
        if (invert) grid[r][c] = !grid[r][c];
      }
    }

    const fmt = QR_FORMAT_BITS_L[maskId];
    // Top-left copy: bit 14 (MSB) at (8,0) .. bit 0 (LSB) at (0,8)
    for (let i = 0; i < 6; i++) grid[8][i] = Boolean((fmt >> (14 - i)) & 1);
    grid[8][7] = Boolean((fmt >> 8) & 1);
    grid[8][8] = Boolean((fmt >> 7) & 1);
    grid[7][8] = Boolean((fmt >> 6) & 1);
    for (let i = 0; i < 6; i++) grid[5 - i][8] = Boolean((fmt >> (5 - i)) & 1);

    // Bottom-left & top-right split copy
    for (let i = 0; i < 7; i++) grid[N - 1 - i][8] = Boolean((fmt >> (14 - i)) & 1);
    for (let i = 0; i < 8; i++) grid[8][N - 8 + i] = Boolean((fmt >> (7 - i)) & 1);
  };

  let bestScore = Infinity;
  let bestGrid = null;

  for (let m = 0; m < 8; m++) {
    const candidate = modules.map((row) => [...row]);
    applyMaskAndFormat(candidate, m);

    // Penalty scoring (N1 adjacent runs + N2 2x2 blocks + N4 dark balance)
    let score = 0;
    for (let r = 0; r < N; r++) {
      let runLen = 1;
      for (let c = 1; c < N; c++) {
        if (candidate[r][c] === candidate[r][c - 1]) {
          runLen++;
          if (runLen === 5) score += 3;
          else if (runLen > 5) score += 1;
        } else {
          runLen = 1;
        }
      }
    }
    for (let c = 0; c < N; c++) {
      let runLen = 1;
      for (let r = 1; r < N; r++) {
        if (candidate[r][c] === candidate[r - 1][c]) {
          runLen++;
          if (runLen === 5) score += 3;
          else if (runLen > 5) score += 1;
        } else {
          runLen = 1;
        }
      }
    }
    let darkCount = 0;
    for (let r = 0; r < N; r++) {
      for (let c = 0; c < N; c++) {
        if (candidate[r][c]) darkCount++;
        if (
          r < N - 1 &&
          c < N - 1 &&
          candidate[r][c] === candidate[r][c + 1] &&
          candidate[r][c] === candidate[r + 1][c] &&
          candidate[r][c] === candidate[r + 1][c + 1]
        ) {
          score += 3;
        }
      }
    }
    const darkRatio = (darkCount * 100) / (N * N);
    score += Math.floor(Math.abs(darkRatio - 50) / 5) * 10;

    if (score < bestScore) {
      bestScore = score;
      bestGrid = candidate;
    }
  }

  return bestGrid;
};

// Reusable SVG QR Code Renderer Component
const QrCodeSvg = ({ value, size = 210 }) => {
  const matrix = useMemo(() => generateQrMatrix(value), [value]);
  if (!matrix) {
    return (
      <div className="p-4 rounded-xl border border-amber-500/30 bg-amber-500/10 text-amber-700 dark:text-amber-300 text-xs font-semibold">
        כתובת המתקן ארוכה מדי עבור קוד QR (מקסימום 271 בתים בקידוד UTF-8).
      </div>
    );
  }
  const n = matrix.length;
  const quietZone = 4;
  const total = n + quietZone * 2;

  return (
    <svg
      width={size}
      height={size}
      viewBox={`0 0 ${total} ${total}`}
      shapeRendering="crispEdges"
      className="rounded-xl bg-white p-2 border border-slate-200 shadow-sm mx-auto"
    >
      <rect x="0" y="0" width={total} height={total} fill="#FFFFFF" />
      {matrix.map((row, r) =>
        row.map((isDark, c) =>
          isDark ? (
            <rect
              key={`${r}-${c}`}
              x={c + quietZone}
              y={r + quietZone}
              width="1"
              height="1"
              fill="#0F172A"
            />
          ) : null
        )
      )}
    </svg>
  );
};

export default function App() {
  const APP_VERSION = 'V4.0.0 PRO';

  // Load initial state once from localStorage (defaults to clean empty state)
  const savedState = useMemo(() => loadSavedState(), []);

  // Theme, Mobile Drawer & Active View State
  const [darkMode, setDarkMode] = useState(() => Boolean(savedState?.darkMode));
  const [mobileMenuOpen, setMobileMenuOpen] = useState(false);
  const [activeView, setActiveView] = useState('tracker'); // 'tracker' | 'dashboard' | 'dispenser_history' | 'history' | 'settings'

  // Core Hierarchical Data State
  const [areas, setAreas] = useState(() => (Array.isArray(savedState?.areas) ? savedState.areas : []));
  const [lastGlobalUpdateDate, setLastGlobalUpdateDate] = useState(
    () => savedState?.lastGlobalUpdateDate || ''
  );
  const [actionHistory, setActionHistory] = useState(() =>
    Array.isArray(savedState?.actionHistory) ? savedState.actionHistory : []
  );

  // Persistent Quick-Pick Name History for Technicians & Escorts
  const [savedTechnicians, setSavedTechnicians] = useState(() =>
    Array.isArray(savedState?.savedTechnicians) ? savedState.savedTechnicians : []
  );
  const [savedEscorts, setSavedEscorts] = useState(() =>
    Array.isArray(savedState?.savedEscorts) ? savedState.savedEscorts : []
  );

  // Active Selection States
  const [selectedAreaId, setSelectedAreaId] = useState(() => savedState?.selectedAreaId || null);
  const [selectedLocationId, setSelectedLocationId] = useState(
    () => savedState?.selectedLocationId || null
  );

  // Area Inline Add & Rename States
  const [isAddingArea, setIsAddingArea] = useState(false);
  const [newAreaName, setNewAreaName] = useState('');
  const [editingAreaId, setEditingAreaId] = useState(null);
  const [editingAreaNameValue, setEditingAreaNameValue] = useState('');

  // Location Inline Add & Rename States
  const [isAddingLocation, setIsAddingLocation] = useState(false);
  const [newLocName, setNewLocName] = useState('');
  const [newItemName, setNewItemName] = useState('');
  const [editingLocationId, setEditingLocationId] = useState(null);
  const [editingLocationNameValue, setEditingLocationNameValue] = useState('');

  // Filter Renewal Modal States (Requires Technician, Escort & Renewal Date)
  const [showRenewalModal, setShowRenewalModal] = useState(false);
  const [technicianName, setTechnicianName] = useState('');
  const [escortName, setEscortName] = useState('');
  const [renewalDate, setRenewalDate] = useState(() => getLocalIsoDate());
  const [resetFiltersOnRenewal, setResetFiltersOnRenewal] = useState(true);
  const [renewalError, setRenewalError] = useState('');

  // Action Center Collapse State
  const [actionCenterExpanded, setActionCenterExpanded] = useState(true);

  // QR Label Modal State
  const [showQrModal, setShowQrModal] = useState(false);

  // Official A4 / PDF Inspection Report Modal State
  const [showReportModal, setShowReportModal] = useState(false);
  const [reportInspectorName, setReportInspectorName] = useState('');

  // Photo Capture & Lightbox States
  const [photoCategory, setPhotoCategory] = useState('תג שמיש מתקן');
  const [isCompressingPhoto, setIsCompressingPhoto] = useState(false);
  const [lightboxPhoto, setLightboxPhoto] = useState(null);
  const cameraInputRef = useRef(null);
  const galleryInputRef = useRef(null);

  // Custom Confirmation Modal State
  const [confirmModal, setConfirmModal] = useState({
    isOpen: false,
    type: null, // 'area' | 'location' | 'reset_all'
    id: null,
    title: '',
    message: ''
  });

  // Dashboard Filter & Sort States
  const [trackingPeriod, setTrackingPeriod] = useState('6m'); // '6m' | '1y'
  const [dashboardSearch, setDashboardSearch] = useState('');
  const [dashboardStatusFilter, setDashboardStatusFilter] = useState('all'); // 'all' | 'faulty' | 'expired'
  const [dashboardAreaFilter, setDashboardAreaFilter] = useState('all');
  const [dashboardSortBy, setDashboardSortBy] = useState('default'); // 'default' | 'lowest_filter' | 'nearest_expiry'

  // Excel Backup Import/Export States & Refs
  const [isBackupBusy, setIsBackupBusy] = useState(false);
  const fileInputRef = useRef(null);

  // Floating Toast State & Timer Ref
  const [toast, setToast] = useState(null);
  const toastTimerRef = useRef(null);

  // Guard refs
  const lastLogMetaRef = useRef({ key: '', timestamp: 0 });
  const qrHandledRef = useRef(false);
  const storageWarnedRef = useRef(false);

  // Display non-intrusive SaaS toast notification without overlapping timers
  const notify = useCallback((message, duration = 3500) => {
    setToast(message);
    if (toastTimerRef.current) clearTimeout(toastTimerRef.current);
    toastTimerRef.current = setTimeout(() => {
      setToast(null);
      toastTimerRef.current = null;
    }, duration);
  }, []);

  // Sync `.dark` class on <html>
  useEffect(() => {
    const root = document.documentElement;
    if (darkMode) {
      root.classList.add('dark');
      root.style.colorScheme = 'dark';
    } else {
      root.classList.remove('dark');
      root.style.colorScheme = 'light';
    }
  }, [darkMode]);

  // Persist application state automatically to localStorage with storage quota warning
  useEffect(() => {
    try {
      const payload = {
        darkMode,
        areas,
        lastGlobalUpdateDate,
        actionHistory,
        savedTechnicians,
        savedEscorts,
        selectedAreaId,
        selectedLocationId
      };
      window.localStorage.setItem(STORAGE_KEY, JSON.stringify(payload));
      storageWarnedRef.current = false;
    } catch {
      if (!storageWarnedRef.current) {
        storageWarnedRef.current = true;
        notify(
          'שגיאה בשמירה המקומית: נפח האחסון בדפדפן מלא (ייתכן מעודף תמונות). בצע ייצוא Excel ומחק תמונות ישנות.',
          6000
        );
      }
    }
  }, [
    darkMode,
    areas,
    lastGlobalUpdateDate,
    actionHistory,
    savedTechnicians,
    savedEscorts,
    selectedAreaId,
    selectedLocationId,
    notify
  ]);

  // Clean up toast timer on unmount
  useEffect(() => {
    return () => {
      if (toastTimerRef.current) clearTimeout(toastTimerRef.current);
    };
  }, []);

  // QR Code Deep-Linking Handler: Runs once on mount if query params exist
  useEffect(() => {
    if (qrHandledRef.current) return;
    try {
      const params = new URLSearchParams(window.location.search);
      const qrAreaParam = params.get('area');
      const qrLocParam = params.get('loc');
      if (!qrAreaParam && !qrLocParam) {
        qrHandledRef.current = true;
        return;
      }

      qrHandledRef.current = true;

      if (areas.length === 0) {
        notify('לא נמצאו מתקנים שמורים במכשיר זה לפתיחת קוד ה-QR', 4500);
      } else {
        let matchedArea = areas.find(
          (a) =>
            String(a.id) === qrAreaParam ||
            a.name.trim().toLowerCase() === String(qrAreaParam || '').trim().toLowerCase()
        );
        let matchedLoc = matchedArea?.locations.find(
          (l) =>
            String(l.id) === qrLocParam ||
            l.name.trim().toLowerCase() === String(qrLocParam || '').trim().toLowerCase()
        );

        // Fallback: search by loc across all areas if area wasn't matched
        if (!matchedLoc && qrLocParam) {
          for (const a of areas) {
            const found = a.locations.find((l) => String(l.id) === qrLocParam);
            if (found) {
              matchedArea = a;
              matchedLoc = found;
              break;
            }
          }
        }

        if (matchedArea) {
          setSelectedAreaId(matchedArea.id);
          if (matchedLoc) {
            setSelectedLocationId(matchedLoc.id);
          }
          setActiveView('tracker');
          notify(
            `נפתח מתקן מסריקת QR: ${matchedLoc ? matchedLoc.dispenserName : matchedArea.name}`,
            4200
          );
        } else {
          notify('המתקן בקוד ה-QR לא נמצא במכשיר זה', 4200);
        }
      }

      // Clean query string safely across http/https/file protocols
      const url = new URL(window.location.href);
      url.searchParams.delete('area');
      url.searchParams.delete('loc');
      window.history.replaceState({}, document.title, url.toString());
    } catch {
      qrHandledRef.current = true;
    }
  }, [areas, notify]);

  // Keep active selection valid when items are added or deleted
  useEffect(() => {
    if (areas.length === 0) {
      if (selectedAreaId !== null) setSelectedAreaId(null);
      if (selectedLocationId !== null) setSelectedLocationId(null);
      return;
    }
    const activeArea = areas.find((a) => a.id === selectedAreaId);
    if (!activeArea) {
      setSelectedAreaId(areas[0].id);
      setSelectedLocationId(areas[0].locations[0]?.id || null);
      return;
    }
    if (
      selectedLocationId &&
      !activeArea.locations.some((loc) => loc.id === selectedLocationId)
    ) {
      setSelectedLocationId(activeArea.locations[0]?.id || null);
    }
  }, [areas, selectedAreaId, selectedLocationId]);

  // Reset dashboard area filter if the selected area was deleted
  useEffect(() => {
    if (dashboardAreaFilter !== 'all' && !areas.some((a) => a.id === dashboardAreaFilter)) {
      setDashboardAreaFilter('all');
    }
  }, [areas, dashboardAreaFilter]);

  // Generate current local ISO date and formatted Hebrew timestamp
  const getNowStrings = () => {
    const now = new Date();
    const todayStr = getLocalIsoDate(now);
    const timeStr = now.toLocaleTimeString('he-IL', {
      hour: '2-digit',
      minute: '2-digit',
      second: '2-digit'
    });
    return {
      todayStr,
      timeStr,
      fullDateTime: `${formatDateIL(todayStr)} ${timeStr}`
    };
  };

  // Helper to open Filter Renewal Modal with fresh default today's date
  const openFilterRenewalModal = () => {
    setRenewalDate(getLocalIsoDate());
    setRenewalError('');
    setShowRenewalModal(true);
  };

  // Calculate 6-month or 1-year next filter renewal date & status badge in local time
  const getRenewalCycleInfo = (lastRefreshDate, period) => {
    if (!lastRefreshDate) {
      return {
        nextDueDate: 'טרם נקבע',
        statusLabel: 'טרם בוצע רענון',
        badgeClass: darkMode
          ? 'bg-slate-800 text-slate-400 border-slate-700'
          : 'bg-slate-100 text-slate-600 border-slate-200'
      };
    }

    const baseDate = parseLocalIsoDate(lastRefreshDate);
    if (!baseDate) {
      return {
        nextDueDate: '-',
        statusLabel: 'תאריך לא תקין',
        badgeClass: 'bg-slate-500/10 text-slate-500 border-slate-500/20'
      };
    }

    const months = period === '6m' ? 6 : 12;
    const targetMonthIndex = baseDate.getMonth() + months;
    const dueDate = new Date(baseDate.getFullYear(), targetMonthIndex, baseDate.getDate());

    const expectedMonth = ((targetMonthIndex % 12) + 12) % 12;
    if (dueDate.getMonth() !== expectedMonth) {
      dueDate.setDate(0);
    }

    const dueIso = getLocalIsoDate(dueDate);
    const now = new Date();
    const todayMidnight = new Date(now.getFullYear(), now.getMonth(), now.getDate());
    const diffDays = Math.round(
      (dueDate.getTime() - todayMidnight.getTime()) / (1000 * 60 * 60 * 24)
    );

    if (diffDays < 0) {
      return {
        nextDueDate: formatDateIL(dueIso),
        statusLabel: `פג תוקף (${Math.abs(diffDays)} ימים)`,
        badgeClass: 'bg-rose-500/10 text-rose-600 dark:text-rose-400 border-rose-500/25'
      };
    }
    if (diffDays <= 30) {
      return {
        nextDueDate: formatDateIL(dueIso),
        statusLabel: `מתקרב (בעוד ${diffDays} ימים)`,
        badgeClass: 'bg-amber-500/10 text-amber-600 dark:text-amber-400 border-amber-500/25'
      };
    }
    return {
      nextDueDate: formatDateIL(dueIso),
      statusLabel: 'בתוקף',
      badgeClass: 'bg-emerald-500/10 text-emerald-600 dark:text-emerald-400 border-emerald-500/25'
    };
  };

  // Hebrew field labels for audit trail
  const fieldLabels = {
    dispenserName: 'שם פריט / מתקן',
    filterMaga: 'פילטר מגה',
    filterAlef1: 'פילטר אלף 1 (%)',
    filterAlef2: 'פילטר אלף 2 (%)',
    tagExpiry: 'תוקף תג שמיש',
    clockType: 'סוג שעון',
    systemType: 'סוג מערכת',
    tankSensitivity: 'רגישות מיכל',
    cartridgeCarriage: 'עגלה למיכל',
    tankTagExpiry: 'תוקף תג שמיש למיכל',
    notes: 'הערות'
  };

  // Derived active Area & Location
  const currentArea = useMemo(
    () => areas.find((a) => a.id === selectedAreaId) || null,
    [areas, selectedAreaId]
  );

  const currentLocation = useMemo(
    () => currentArea?.locations.find((l) => l.id === selectedLocationId) || null,
    [currentArea, selectedLocationId]
  );

  // Flattened inventory across all areas
  const allDispensers = useMemo(() => {
    return areas.flatMap((area) =>
      (area.locations || []).map((loc) => ({
        ...loc,
        areaId: area.id,
        areaName: area.name
      }))
    );
  }, [areas]);

  // Combined Unique History of Technicians & Escorts
  const technicianHistoryList = useMemo(() => {
    const set = new Set(savedTechnicians.map((n) => n.trim()).filter(Boolean));
    allDispensers.forEach((d) => {
      (d.filterRefreshHistory || []).forEach((rec) => {
        if (rec.technician?.trim()) set.add(rec.technician.trim());
      });
    });
    return Array.from(set);
  }, [savedTechnicians, allDispensers]);

  const escortHistoryList = useMemo(() => {
    const set = new Set(savedEscorts.map((n) => n.trim()).filter(Boolean));
    allDispensers.forEach((d) => {
      (d.filterRefreshHistory || []).forEach((rec) => {
        if (rec.escort?.trim()) set.add(rec.escort.trim());
      });
    });
    return Array.from(set);
  }, [savedEscorts, allDispensers]);

  // =========================================================================
  // SMART URGENT ACTION CENTER (מרכז התראות ודברים דחופים)
  // =========================================================================
  const urgentAlerts = useMemo(() => {
    const alerts = [];
    const seenKeys = new Set();

    const pushUniqueAlert = (alertObj) => {
      const dedupeKey = `${alertObj.locationId}|${alertObj.message}`;
      if (seenKeys.has(dedupeKey)) return;
      seenKeys.add(dedupeKey);
      alerts.push(alertObj);
    };

    allDispensers.forEach((item) => {
      const diffDisp = item.tagExpiry ? getDaysUntilDate(item.tagExpiry) : null;
      const diffTank = item.tankTagExpiry ? getDaysUntilDate(item.tankTagExpiry) : null;
      const dispNeedsAlert = diffDisp !== null && diffDisp <= 30;
      const tankNeedsAlert = diffTank !== null && diffTank <= 30;
      const pct1 = toPercent(item.filterAlef1, 100);
      const pct2 = toPercent(item.filterAlef2, 100);

      // 1 & 2. Dispenser & Tank Tag Expiry (combined when both have the exact same expiry date)
      if (dispNeedsAlert && tankNeedsAlert && item.tagExpiry === item.tankTagExpiry) {
        if (diffDisp < 0) {
          pushUniqueAlert({
            id: `${item.id}-both-tags-exp`,
            areaId: item.areaId,
            areaName: item.areaName,
            locationId: item.id,
            locationName: item.name,
            dispenserName: item.dispenserName,
            category: 'tag',
            severity: 'critical',
            message: `פג תוקף תג שמיש למתקן ולמיכל (${formatDateIL(item.tagExpiry)} • לפני ${Math.abs(diffDisp)} ימים)`
          });
        } else {
          pushUniqueAlert({
            id: `${item.id}-both-tags-soon`,
            areaId: item.areaId,
            areaName: item.areaName,
            locationId: item.id,
            locationName: item.name,
            dispenserName: item.dispenserName,
            category: 'tag',
            severity: 'warning',
            message:
              diffDisp === 0
                ? `תוקף תג שמיש למתקן ולמיכל פג היום (${formatDateIL(item.tagExpiry)})`
                : `תוקף תג שמיש למתקן ולמיכל עומד לפוג בעוד ${diffDisp} ימים (${formatDateIL(item.tagExpiry)})`
          });
        }
      } else {
        if (dispNeedsAlert) {
          if (diffDisp < 0) {
            pushUniqueAlert({
              id: `${item.id}-disp-tag-exp`,
              areaId: item.areaId,
              areaName: item.areaName,
              locationId: item.id,
              locationName: item.name,
              dispenserName: item.dispenserName,
              category: 'tag',
              severity: 'critical',
              message: `[מתקן] פג תוקף תג שמיש (${formatDateIL(item.tagExpiry)} • לפני ${Math.abs(diffDisp)} ימים)`
            });
          } else {
            pushUniqueAlert({
              id: `${item.id}-disp-tag-soon`,
              areaId: item.areaId,
              areaName: item.areaName,
              locationId: item.id,
              locationName: item.name,
              dispenserName: item.dispenserName,
              category: 'tag',
              severity: 'warning',
              message:
                diffDisp === 0
                  ? `[מתקן] תוקף תג שמיש פג היום (${formatDateIL(item.tagExpiry)})`
                  : `[מתקן] תוקף תג שמיש עומד לפוג בעוד ${diffDisp} ימים (${formatDateIL(item.tagExpiry)})`
            });
          }
        }

        if (tankNeedsAlert) {
          if (diffTank < 0) {
            pushUniqueAlert({
              id: `${item.id}-tank-tag-exp`,
              areaId: item.areaId,
              areaName: item.areaName,
              locationId: item.id,
              locationName: item.name,
              dispenserName: item.dispenserName,
              category: 'tag',
              severity: 'critical',
              message: `[מיכל] פג תוקף תג שמיש (${formatDateIL(item.tankTagExpiry)} • לפני ${Math.abs(diffTank)} ימים)`
            });
          } else {
            pushUniqueAlert({
              id: `${item.id}-tank-tag-soon`,
              areaId: item.areaId,
              areaName: item.areaName,
              locationId: item.id,
              locationName: item.name,
              dispenserName: item.dispenserName,
              category: 'tag',
              severity: 'warning',
              message:
                diffTank === 0
                  ? `[מיכל] תוקף תג שמיש פג היום (${formatDateIL(item.tankTagExpiry)})`
                  : `[מיכל] תוקף תג שמיש עומד לפוג בעוד ${diffTank} ימים (${formatDateIL(item.tankTagExpiry)})`
            });
          }
        }
      }

      // 3. Filter Alef 1 or Filter Alef 2 below 20%
      if (pct1 < 20) {
        pushUniqueAlert({
          id: `${item.id}-alef1-low`,
          areaId: item.areaId,
          areaName: item.areaName,
          locationId: item.id,
          locationName: item.name,
          dispenserName: item.dispenserName,
          category: 'filter_low',
          severity: pct1 <= 10 ? 'critical' : 'warning',
          message: `פילטר אלף 1 ירד ל-${pct1}% (מתחת ל-20%)`
        });
      }
      if (pct2 < 20) {
        pushUniqueAlert({
          id: `${item.id}-alef2-low`,
          areaId: item.areaId,
          areaName: item.areaName,
          locationId: item.id,
          locationName: item.name,
          dispenserName: item.dispenserName,
          category: 'filter_low',
          severity: pct2 <= 10 ? 'critical' : 'warning',
          message: `פילטר אלף 2 ירד ל-${pct2}% (מתחת ל-20%)`
        });
      }

      // 4. Filter Maga or Tank Sensitivity marked as 'לא תקין'
      if (item.filterMaga === 'לא תקין') {
        pushUniqueAlert({
          id: `${item.id}-maga-fault`,
          areaId: item.areaId,
          areaName: item.areaName,
          locationId: item.id,
          locationName: item.name,
          dispenserName: item.dispenserName,
          category: 'fault',
          severity: 'critical',
          message: 'פילטר מגה מסומן כ"לא תקין"'
        });
      }
      if (item.tankSensitivity === 'לא תקין') {
        pushUniqueAlert({
          id: `${item.id}-tank-fault`,
          areaId: item.areaId,
          areaName: item.areaName,
          locationId: item.id,
          locationName: item.name,
          dispenserName: item.dispenserName,
          category: 'fault',
          severity: 'critical',
          message: 'רגישות מיכל מסומנת כ"לא תקין"'
        });
      }
    });

    // Sort critical alerts first, then warnings
    return alerts.sort((a, b) => {
      if (a.severity === b.severity) return 0;
      return a.severity === 'critical' ? -1 : 1;
    });
  }, [allDispensers]);

  // Helper: Check if a dispenser has any tag expired or expiring within 30 days
  const hasExpiredOrExpiringTag = (item) => {
    const d1 = getDaysUntilDate(item.tagExpiry);
    const d2 = getDaysUntilDate(item.tankTagExpiry);
    return (d1 !== null && d1 <= 30) || (d2 !== null && d2 <= 30);
  };

  // Helper: Overall location health indicator
  const isLocationHealthy = (loc) =>
    (loc.filterMaga || 'תקין') === 'תקין' &&
    (loc.tankSensitivity || 'תקין') === 'תקין' &&
    toPercent(loc.filterAlef1, 100) >= 20 &&
    toPercent(loc.filterAlef2, 100) >= 20;

  // Advanced Filtered & Sorted Dispensers for Dashboard
  const filteredDispensers = useMemo(() => {
    let list = [...allDispensers];

    if (dashboardAreaFilter !== 'all') {
      list = list.filter((item) => item.areaId === dashboardAreaFilter);
    }

    if (dashboardStatusFilter === 'faulty') {
      list = list.filter((item) => !isLocationHealthy(item));
    } else if (dashboardStatusFilter === 'expired') {
      list = list.filter((item) => hasExpiredOrExpiringTag(item));
    }

    if (dashboardSearch.trim()) {
      const q = dashboardSearch.trim().toLowerCase();
      list = list.filter(
        (item) =>
          item.areaName.toLowerCase().includes(q) ||
          item.name.toLowerCase().includes(q) ||
          (item.dispenserName || '').toLowerCase().includes(q) ||
          (item.systemType || '').toLowerCase().includes(q) ||
          (item.notes || '').toLowerCase().includes(q)
      );
    }

    if (dashboardSortBy === 'lowest_filter') {
      list.sort((a, b) => {
        const minA = Math.min(toPercent(a.filterAlef1, 100), toPercent(a.filterAlef2, 100));
        const minB = Math.min(toPercent(b.filterAlef1, 100), toPercent(b.filterAlef2, 100));
        return minA - minB;
      });
    } else if (dashboardSortBy === 'nearest_expiry') {
      list.sort((a, b) => {
        const getNearest = (it) => {
          const d1 = getDaysUntilDate(it.tagExpiry);
          const d2 = getDaysUntilDate(it.tankTagExpiry);
          const vals = [d1, d2].filter((v) => v !== null);
          return vals.length > 0 ? Math.min(...vals) : 999999;
        };
        return getNearest(a) - getNearest(b);
      });
    }

    return list;
  }, [
    allDispensers,
    dashboardAreaFilter,
    dashboardStatusFilter,
    dashboardSearch,
    dashboardSortBy
  ]);

  // Top-left live Table Update Date
  const headerTableUpdateDate = currentLocation?.lastTableUpdateDate || lastGlobalUpdateDate;
  const todayLocalIso = getLocalIsoDate();

  // QR Deep Link URL for current dispenser (safe across http/https/file)
  const currentDispenserQrUrl = useMemo(() => {
    if (!currentArea || !currentLocation || typeof window === 'undefined') return '';
    try {
      const url = new URL(window.location.href);
      url.search = '';
      url.hash = '';
      url.searchParams.set('area', currentArea.id);
      url.searchParams.set('loc', currentLocation.id);
      return url.toString();
    } catch {
      const baseUrl = `${window.location.origin || ''}${window.location.pathname || ''}`;
      return `${baseUrl}?area=${encodeURIComponent(currentArea.id)}&loc=${encodeURIComponent(
        currentLocation.id
      )}`;
    }
  }, [currentArea, currentLocation]);

  // =========================================================================
  // MOBILE CAMERA & GALLERY PHOTO DOCUMENTATION HANDLERS
  // =========================================================================
  const handlePhotoFileSelected = async (e) => {
    const file = e.target.files?.[0];
    if (!file || !selectedAreaId || !selectedLocationId || !currentLocation) return;

    setIsCompressingPhoto(true);
    try {
      const compressedDataUrl = await compressImageFile(file, 800, 0.66);
      const { todayStr, fullDateTime } = getNowStrings();
      const newPhoto = {
        id: `photo-${Date.now()}-${Math.random().toString(36).slice(2, 6)}`,
        dataUrl: compressedDataUrl,
        category: photoCategory,
        date: fullDateTime,
        isoDate: todayStr
      };

      setAreas((prev) =>
        prev.map((area) => {
          if (area.id !== selectedAreaId) return area;
          return {
            ...area,
            locations: area.locations.map((loc) =>
              loc.id === selectedLocationId
                ? {
                    ...loc,
                    photos: [newPhoto, ...(loc.photos || [])],
                    lastTableUpdateDate: todayStr
                  }
                : loc
            )
          };
        })
      );

      setLastGlobalUpdateDate(todayStr);
      setActionHistory((prev) => [
        {
          id: Date.now() + Math.random(),
          date: fullDateTime,
          user: 'מפעיל מערכת',
          description: `אזור "${currentArea?.name}" > מקום "${currentLocation.name}" > מתקן "${currentLocation.dispenserName}": נוספה תמונת תיעוד (${photoCategory})`
        },
        ...prev
      ]);
      notify(`התמונה נשמרה וכווצה בהצלחה (${photoCategory})`);
    } catch {
      notify('שגיאה בעיבוד התמונה. נסה לצלם או לבחור תמונה אחרת.');
    } finally {
      setIsCompressingPhoto(false);
      if (cameraInputRef.current) cameraInputRef.current.value = '';
      if (galleryInputRef.current) galleryInputRef.current.value = '';
    }
  };

  const handleDeletePhoto = (photoId) => {
    if (!selectedAreaId || !selectedLocationId || !currentLocation) return;
    const { todayStr, fullDateTime } = getNowStrings();

    setAreas((prev) =>
      prev.map((area) => {
        if (area.id !== selectedAreaId) return area;
        return {
          ...area,
          locations: area.locations.map((loc) =>
            loc.id === selectedLocationId
              ? {
                  ...loc,
                  photos: (loc.photos || []).filter((p) => p.id !== photoId),
                  lastTableUpdateDate: todayStr
                }
              : loc
          )
        };
      })
    );

    if (lightboxPhoto?.id === photoId) setLightboxPhoto(null);
    setActionHistory((prev) => [
      {
        id: Date.now() + Math.random(),
        date: fullDateTime,
        user: 'מפעיל מערכת',
        description: `נמחקה תמונת תיעוד ממתקן "${currentLocation.dispenserName}"`
      },
      ...prev
    ]);
    notify('התמונה נמחקה');
  };

  // Download QR Code + Dispenser Label as high-res PNG image
  const handleDownloadQrPng = () => {
    if (!currentLocation || !currentArea || !currentDispenserQrUrl) return;
    const matrix = generateQrMatrix(currentDispenserQrUrl);
    if (!matrix) {
      notify('לא ניתן ליצור תווית PNG: כתובת הקישור ארוכה מדי לקוד QR');
      return;
    }

    const n = matrix.length;
    const scale = 10;
    const quiet = 4;
    const qrPixelSize = (n + quiet * 2) * scale;
    const footerHeight = 96;
    const canvas = document.createElement('canvas');
    canvas.width = qrPixelSize;
    canvas.height = qrPixelSize + footerHeight;

    const ctx = canvas.getContext('2d');
    if (!ctx) return;

    ctx.fillStyle = '#FFFFFF';
    ctx.fillRect(0, 0, canvas.width, canvas.height);

    ctx.fillStyle = '#0F172A';
    for (let r = 0; r < n; r++) {
      for (let c = 0; c < n; c++) {
        if (matrix[r][c]) {
          ctx.fillRect((c + quiet) * scale, (r + quiet) * scale, scale, scale);
        }
      }
    }

    ctx.strokeStyle = '#E2E8F0';
    ctx.lineWidth = 2;
    ctx.beginPath();
    ctx.moveTo(20, qrPixelSize + 6);
    ctx.lineTo(canvas.width - 20, qrPixelSize + 6);
    ctx.stroke();

    const maxLabelWidth = canvas.width - 24;
    if ('direction' in ctx) ctx.direction = 'rtl';
    ctx.textAlign = 'center';

    ctx.fillStyle = '#0F172A';
    ctx.font = 'bold 18px sans-serif';
    ctx.fillText(currentLocation.dispenserName || '', canvas.width / 2, qrPixelSize + 34, maxLabelWidth);

    ctx.fillStyle = '#475569';
    ctx.font = '14px sans-serif';
    ctx.fillText(
      `אזור: ${currentArea.name} | מקום: ${currentLocation.name}`,
      canvas.width / 2,
      qrPixelSize + 58,
      maxLabelWidth
    );

    ctx.fillStyle = '#4F46E5';
    ctx.font = 'bold 11px sans-serif';
    ctx.fillText(
      'AquaOps PRO • סרוק לפתיחת כרטיס מתקן',
      canvas.width / 2,
      qrPixelSize + 80,
      maxLabelWidth
    );

    const safeFilePart = (str) => String(str || '').replace(/[\\/:*?"<>|]/g, '_');
    const url = canvas.toDataURL('image/png');
    const a = document.createElement('a');
    a.href = url;
    a.download = `QR_${safeFilePart(currentArea.name)}_${safeFilePart(
      currentLocation.dispenserName
    )}.png`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    notify('תווית ה-QR הורדה כתמונה (PNG)!');
  };

  // =========================================================================
  // EXCEL BACKUP EXPORT & NON-DESTRUCTIVE DEDUPLICATED IMPORT
  // =========================================================================
  const buildBackupDatasets = () => {
    const dispensersRows = [];
    const renewalsRows = [];

    areas.forEach((area) => {
      if (!area.locations || area.locations.length === 0) {
        dispensersRows.push({
          מזהה_אזור: area.id,
          שם_אזור: area.name,
          מזהה_מקום: '',
          שם_מקום: '',
          שם_מתקן: '',
          פילטר_מגה: '',
          פילטר_אלף_1: '',
          פילטר_אלף_2: '',
          תוקף_תג_מתקן: '',
          סוג_שעון: '',
          סוג_מערכת: '',
          תאריך_עדכון_טבלה: '',
          תאריך_רענון_אחרון: '',
          רגישות_מיכל: '',
          עגלה_למיכל: '',
          תוקף_תג_מיכל: '',
          מספר_תמונות: 0,
          הערות: ''
        });
        return;
      }

      area.locations.forEach((loc) => {
        dispensersRows.push({
          מזהה_אזור: area.id,
          שם_אזור: area.name,
          מזהה_מקום: loc.id,
          שם_מקום: loc.name,
          שם_מתקן: loc.dispenserName || '',
          פילטר_מגה: loc.filterMaga || 'תקין',
          פילטר_אלף_1: toPercent(loc.filterAlef1, 100),
          פילטר_אלף_2: toPercent(loc.filterAlef2, 100),
          תוקף_תג_מתקן: loc.tagExpiry || '',
          סוג_שעון: loc.clockType || 'דיגיטלי',
          סוג_מערכת: loc.systemType || 'טורי',
          תאריך_עדכון_טבלה: loc.lastTableUpdateDate || '',
          תאריך_רענון_אחרון: loc.lastFilterRefreshDate || '',
          רגישות_מיכל: loc.tankSensitivity || 'תקין',
          עגלה_למיכל: loc.cartridgeCarriage || 'יש עגלה',
          תוקף_תג_מיכל: loc.tankTagExpiry || '',
          מספר_תמונות: (loc.photos || []).length,
          הערות: loc.notes || ''
        });

        (loc.filterRefreshHistory || []).forEach((rec) => {
          renewalsRows.push({
            מזהה_רענון: String(rec.id),
            מזהה_אזור: area.id,
            שם_אזור: area.name,
            מזהה_מקום: loc.id,
            שם_מקום: loc.name,
            שם_מתקן: loc.dispenserName || '',
            תאריך_מלא: rec.date || '',
            תאריך_ISO: rec.isoDate || '',
            טכנאי: rec.technician || '',
            מלווה: rec.escort || '',
            פעולה: rec.action || ''
          });
        });
      });
    });

    const historyRows = actionHistory.map((log) => ({
      מזהה_פעולה: String(log.id),
      תאריך: log.date || '',
      מבצע: log.user || '',
      תיאור: log.description || ''
    }));

    return { dispensersRows, renewalsRows, historyRows };
  };

  const handleExportExcelBackup = async () => {
    if (areas.length === 0 && actionHistory.length === 0) {
      notify('אין נתונים במערכת לייצוא כרגע');
      return;
    }

    setIsBackupBusy(true);
    const { todayStr, fullDateTime } = getNowStrings();
    const { dispensersRows, renewalsRows, historyRows } = buildBackupDatasets();

    const dispenserHeaders = [
      'מזהה_אזור',
      'שם_אזור',
      'מזהה_מקום',
      'שם_מקום',
      'שם_מתקן',
      'פילטר_מגה',
      'פילטר_אלף_1',
      'פילטר_אלף_2',
      'תוקף_תג_מתקן',
      'סוג_שעון',
      'סוג_מערכת',
      'תאריך_עדכון_טבלה',
      'תאריך_רענון_אחרון',
      'רגישות_מיכל',
      'עגלה_למיכל',
      'תוקף_תג_מיכל',
      'מספר_תמונות',
      'הערות'
    ];
    const renewalHeaders = [
      'מזהה_רענון',
      'מזהה_אזור',
      'שם_אזור',
      'מזהה_מקום',
      'שם_מקום',
      'שם_מתקן',
      'תאריך_מלא',
      'תאריך_ISO',
      'טכנאי',
      'מלווה',
      'פעולה'
    ];
    const historyHeaders = ['מזהה_פעולה', 'תאריך', 'מבצע', 'תיאור'];

    try {
      const XLSX = await loadXLSXLibrary();
      const wb = XLSX.utils.book_new();
      wb.Workbook = { Views: [{ RTL: true }] };

      const wsDispensers = XLSX.utils.json_to_sheet(dispensersRows, {
        header: dispenserHeaders
      });
      wsDispensers['!cols'] = dispenserHeaders.map((h) => ({
        wch: h === 'הערות' ? 32 : h.includes('שם') ? 22 : 18
      }));
      XLSX.utils.book_append_sheet(wb, wsDispensers, 'מתקנים_ומיכלים');

      const wsRenewals = XLSX.utils.json_to_sheet(renewalsRows, {
        header: renewalHeaders
      });
      wsRenewals['!cols'] = renewalHeaders.map((h) => ({
        wch: h === 'פעולה' ? 40 : 18
      }));
      XLSX.utils.book_append_sheet(wb, wsRenewals, 'היסטוריית_רענונים');

      const wsHistory = XLSX.utils.json_to_sheet(historyRows, {
        header: historyHeaders
      });
      wsHistory['!cols'] = [{ wch: 18 }, { wch: 22 }, { wch: 22 }, { wch: 55 }];
      XLSX.utils.book_append_sheet(wb, wsHistory, 'יומן_פעולות');

      XLSX.writeFile(wb, `AquaOps_Backup_${todayStr}.xlsx`);
      notify('קובץ גיבוי Excel (.xlsx) יוצא בהצלחה!');
    } catch {
      const buildXmlSheet = (sheetName, headers, rows) => {
        const headerCells = headers
          .map((h) => `<Cell><Data ss:Type="String">${escapeXml(h)}</Data></Cell>`)
          .join('');
        const bodyRows = rows
          .map((row) => {
            const cells = headers
              .map((h) => {
                const val = row[h] ?? '';
                const type = typeof val === 'number' && !isNaN(val) ? 'Number' : 'String';
                return `<Cell><Data ss:Type="${type}">${escapeXml(val)}</Data></Cell>`;
              })
              .join('');
            return `<Row>${cells}</Row>`;
          })
          .join('');
        return `<Worksheet ss:Name="${escapeXml(sheetName)}" ss:RightToLeft="1"><Table><Row>${headerCells}</Row>${bodyRows}</Table></Worksheet>`;
      };

      const xmlContent =
        `<?xml version="1.0" encoding="UTF-8"?>` +
        `<?mso-application progid="Excel.Sheet"?>` +
        `<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet" xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet">` +
        buildXmlSheet('מתקנים_ומיכלים', dispenserHeaders, dispensersRows) +
        buildXmlSheet('היסטוריית_רענונים', renewalHeaders, renewalsRows) +
        buildXmlSheet('יומן_פעולות', historyHeaders, historyRows) +
        `</Workbook>`;

      const blob = new Blob([xmlContent], {
        type: 'application/vnd.ms-excel;charset=utf-8;'
      });
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = `AquaOps_Backup_${todayStr}.xls`;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      URL.revokeObjectURL(url);
      notify('קובץ גיבוי Excel (.xls אופליין) יוצא בהצלחה!');
    } finally {
      setActionHistory((prev) => [
        {
          id: Date.now() + Math.random(),
          date: fullDateTime,
          user: 'מפעיל מערכת',
          description: `יוצא קובץ גיבוי Excel (${areas.length} אזורים, ${allDispensers.length} מתקנים)`
        },
        ...prev
      ]);
      setIsBackupBusy(false);
    }
  };

  // SpreadsheetML parser with full ss:Index support for sparse/empty XML cells
  const parseSpreadsheetMLFallback = (xmlText) => {
    const parser = new DOMParser();
    const doc = parser.parseFromString(xmlText, 'application/xml');
    const worksheets = Array.from(doc.getElementsByTagName('Worksheet'));
    const result = {};

    worksheets.forEach((ws) => {
      const name = ws.getAttribute('ss:Name') || ws.getAttribute('Name') || '';
      const rows = Array.from(ws.getElementsByTagName('Row'));
      if (rows.length === 0) {
        result[name] = [];
        return;
      }

      const getRowValues = (rowEl) => {
        const vals = [];
        let colIdx = 0;
        Array.from(rowEl.getElementsByTagName('Cell')).forEach((cell) => {
          const ssIndexRaw =
            cell.getAttribute('ss:Index') ||
            cell.getAttributeNS('urn:schemas-microsoft-com:office:spreadsheet', 'Index') ||
            cell.getAttribute('Index');
          if (ssIndexRaw) {
            const targetIdx = Number(ssIndexRaw) - 1;
            if (!isNaN(targetIdx)) {
              while (colIdx < targetIdx) {
                vals.push('');
                colIdx += 1;
              }
            }
          }
          const dataEl = cell.getElementsByTagName('Data')[0];
          vals.push(dataEl ? dataEl.textContent || '' : '');
          colIdx += 1;
        });
        return vals;
      };

      const headers = getRowValues(rows[0]);
      result[name] = rows.slice(1).map((r) => {
        const vals = getRowValues(r);
        const obj = {};
        headers.forEach((h, idx) => {
          if (h) obj[h] = vals[idx] ?? '';
        });
        return obj;
      });
    });

    return result;
  };

  const mergeImportedDatasets = ({ dispensersRows, renewalsRows, historyRows }) => {
    let addedAreasCount = 0;
    let addedLocationsCount = 0;
    let addedRenewalsCount = 0;
    let skippedDuplicatesCount = 0;
    let latestImportedDate = lastGlobalUpdateDate;

    const mergedAreas = areas.map((area) => ({
      ...area,
      locations: (area.locations || []).map((loc) => ({
        ...loc,
        photos: [...(loc.photos || [])],
        filterRefreshHistory: [...(loc.filterRefreshHistory || [])]
      }))
    }));

    const usedAreaIds = new Set(mergedAreas.map((a) => String(a.id)));
    const usedLocationIds = new Set(
      mergedAreas.flatMap((a) => a.locations.map((l) => String(l.id)))
    );

    const findMatchingArea = (impAreaId, impAreaName) => {
      const cleanName = String(impAreaName || '').trim().toLowerCase();
      if (impAreaId) {
        const byId = mergedAreas.find((a) => String(a.id) === String(impAreaId));
        if (byId) return byId;
      }
      if (cleanName) {
        return mergedAreas.find((a) => a.name.trim().toLowerCase() === cleanName) || null;
      }
      return null;
    };

    const findMatchingLocation = (areaObj, impLocId, impLocName) => {
      const cleanLocName = String(impLocName || '').trim().toLowerCase();
      if (impLocId) {
        const byId = areaObj.locations.find((l) => String(l.id) === String(impLocId));
        if (byId) return byId;
      }
      if (cleanLocName) {
        return areaObj.locations.find((l) => l.name.trim().toLowerCase() === cleanLocName) || null;
      }
      return null;
    };

    dispensersRows.forEach((row, index) => {
      const rawAreaName = String(row['שם_אזור'] ?? row.areaName ?? '').trim();
      const rawAreaId = String(row['מזהה_אזור'] ?? row.areaId ?? '').trim();
      if (!rawAreaName) return;

      let targetArea = findMatchingArea(rawAreaId, rawAreaName);
      if (!targetArea) {
        const newAreaId =
          rawAreaId && !usedAreaIds.has(rawAreaId)
            ? rawAreaId
            : `area-imp-${Date.now()}-${index}-${Math.random().toString(36).slice(2, 6)}`;
        usedAreaIds.add(newAreaId);
        targetArea = {
          id: newAreaId,
          name: rawAreaName,
          locations: []
        };
        mergedAreas.push(targetArea);
        addedAreasCount += 1;
      }

      const rawLocName = String(row['שם_מקום'] ?? row.name ?? '').trim();
      const rawLocId = String(row['מזהה_מקום'] ?? row.id ?? '').trim();
      if (!rawLocName) return;

      const existingLoc = findMatchingLocation(targetArea, rawLocId, rawLocName);
      if (existingLoc) {
        skippedDuplicatesCount += 1;
        return;
      }

      const newLocId =
        rawLocId && !usedLocationIds.has(rawLocId)
          ? rawLocId
          : `loc-imp-${Date.now()}-${index}-${Math.random().toString(36).slice(2, 6)}`;
      usedLocationIds.add(newLocId);

      const importedUpdateDate = normalizeImportedDate(
        row['תאריך_עדכון_טבלה'] ?? row.lastTableUpdateDate
      );
      if (importedUpdateDate && (!latestImportedDate || importedUpdateDate > latestImportedDate)) {
        latestImportedDate = importedUpdateDate;
      }

      targetArea.locations.push({
        id: newLocId,
        name: rawLocName,
        dispenserName:
          String(row['שם_מתקן'] ?? row.dispenserName ?? rawLocName).trim() || rawLocName,
        filterMaga:
          String(row['פילטר_מגה'] ?? row.filterMaga ?? '').trim() === 'לא תקין'
            ? 'לא תקין'
            : 'תקין',
        filterAlef1: toPercent(row['פילטר_אלף_1'] ?? row.filterAlef1, 100),
        filterAlef2: toPercent(row['פילטר_אלף_2'] ?? row.filterAlef2, 100),
        tagExpiry: normalizeImportedDate(row['תוקף_תג_מתקן'] ?? row.tagExpiry),
        clockType: ['דיגיטלי', 'אנלוגי', 'מכני'].includes(
          String(row['סוג_שעון'] ?? row.clockType ?? '').trim()
        )
          ? String(row['סוג_שעון'] ?? row.clockType).trim()
          : 'דיגיטלי',
        systemType:
          String(row['סוג_מערכת'] ?? row.systemType ?? '').trim() === 'מקבילי'
            ? 'מקבילי'
            : 'טורי',
        lastTableUpdateDate: importedUpdateDate,
        lastFilterRefreshDate: normalizeImportedDate(
          row['תאריך_רענון_אחרון'] ?? row.lastFilterRefreshDate
        ),
        tankSensitivity:
          String(row['רגישות_מיכל'] ?? row.tankSensitivity ?? '').trim() === 'לא תקין'
            ? 'לא תקין'
            : 'תקין',
        cartridgeCarriage:
          String(row['עגלה_למיכל'] ?? row.cartridgeCarriage ?? '').trim() === 'אין עגלה'
            ? 'אין עגלה'
            : 'יש עגלה',
        tankTagExpiry: normalizeImportedDate(row['תוקף_תג_מיכל'] ?? row.tankTagExpiry),
        notes: String(row['הערות'] ?? row.notes ?? '').trim(),
        photos: Array.isArray(row.photos) ? row.photos : [],
        filterRefreshHistory: []
      });
      addedLocationsCount += 1;
    });

    renewalsRows.forEach((row, index) => {
      const rawAreaName = String(row['שם_אזור'] ?? row.areaName ?? '').trim();
      const rawAreaId = String(row['מזהה_אזור'] ?? row.areaId ?? '').trim();
      const rawLocName = String(row['שם_מקום'] ?? row.locationName ?? '').trim();
      const rawLocId = String(row['מזהה_מקום'] ?? row.locationId ?? '').trim();

      const targetArea = findMatchingArea(rawAreaId, rawAreaName);
      if (!targetArea) return;
      const targetLoc = findMatchingLocation(targetArea, rawLocId, rawLocName);
      if (!targetLoc) return;

      const recId = String(row['מזהה_רענון'] ?? row.id ?? '').trim();
      const isoDate = normalizeImportedDate(row['תאריך_ISO'] ?? row.isoDate ?? row['תאריך_מלא']);
      const dateDisplay =
        String(row['תאריך_מלא'] ?? row.date ?? '').trim() || (isoDate ? formatDateIL(isoDate) : '');
      const technician = String(row['טכנאי'] ?? row.technician ?? '').trim();
      const escort = String(row['מלווה'] ?? row.escort ?? '').trim();
      const action =
        String(row['פעולה'] ?? row.action ?? '').trim() ||
        `בוצע חידוש פילטרים למתקן "${targetLoc.dispenserName}"`;

      if (!technician && !escort && !dateDisplay) return;

      const signature = `${isoDate}|${technician.toLowerCase()}|${escort.toLowerCase()}|${action}`;
      const isDuplicate = targetLoc.filterRefreshHistory.some((existing) => {
        if (recId && String(existing.id) === recId) return true;
        const existingSig = `${existing.isoDate || ''}|${(existing.technician || '')
          .trim()
          .toLowerCase()}|${(existing.escort || '').trim().toLowerCase()}|${(
          existing.action || ''
        ).trim()}`;
        return existingSig === signature;
      });

      if (isDuplicate) {
        skippedDuplicatesCount += 1;
        return;
      }

      targetLoc.filterRefreshHistory.push({
        id: recId || `ren-imp-${Date.now()}-${index}-${Math.random().toString(36).slice(2, 6)}`,
        date: dateDisplay,
        isoDate,
        technician,
        escort,
        action
      });
      addedRenewalsCount += 1;

      if (isoDate && (!targetLoc.lastFilterRefreshDate || isoDate > targetLoc.lastFilterRefreshDate)) {
        targetLoc.lastFilterRefreshDate = isoDate;
      }
    });

    mergedAreas.forEach((area) => {
      area.locations.forEach((loc) => {
        if (loc.filterRefreshHistory.length > 1) {
          loc.filterRefreshHistory.sort((a, b) =>
            (b.isoDate || '').localeCompare(a.isoDate || '')
          );
        }
      });
    });

    const mergedHistory = [...actionHistory];
    const existingLogSignatures = new Set(
      mergedHistory.map((log) => `${log.date}|${String(log.description || '').trim()}`)
    );
    const existingLogIds = new Set(mergedHistory.map((log) => String(log.id)));
    let addedLogsCount = 0;

    historyRows.forEach((row, index) => {
      const logId = String(row['מזהה_פעולה'] ?? row.id ?? '').trim();
      const logDate = String(row['תאריך'] ?? row.date ?? '').trim();
      const logUser = String(row['מבצע'] ?? row.user ?? 'מפעיל מערכת').trim();
      const logDesc = String(row['תיאור'] ?? row.description ?? '').trim();
      if (!logDesc) return;

      const sig = `${logDate}|${logDesc}`;
      if ((logId && existingLogIds.has(logId)) || existingLogSignatures.has(sig)) {
        return;
      }
      existingLogSignatures.add(sig);
      if (logId) existingLogIds.add(logId);

      mergedHistory.push({
        id: logId || `log-imp-${Date.now()}-${index}-${Math.random().toString(36).slice(2, 6)}`,
        date: logDate,
        user: logUser,
        description: logDesc
      });
      addedLogsCount += 1;
    });

    if (
      addedAreasCount === 0 &&
      addedLocationsCount === 0 &&
      addedRenewalsCount === 0 &&
      addedLogsCount === 0 &&
      skippedDuplicatesCount === 0
    ) {
      notify('לא נמצאו שורות תקינות לייבוא בקובץ (ודא שקיימת עמודת שם_אזור).', 4500);
      return;
    }

    const { fullDateTime } = getNowStrings();
    const summaryLog = {
      id: Date.now() + Math.random(),
      date: fullDateTime,
      user: 'מפעיל מערכת',
      description: `ייבוא גיבוי Excel (מיזוג ללא דריסה): נוספו ${addedAreasCount} אזורים, ${addedLocationsCount} מתקנים ו-${addedRenewalsCount} רענונים (דולגו ${skippedDuplicatesCount} כפילויות)`
    };

    setAreas(mergedAreas);
    if (latestImportedDate) setLastGlobalUpdateDate(latestImportedDate);
    setActionHistory([summaryLog, ...mergedHistory]);

    if (!selectedAreaId && mergedAreas.length > 0) {
      setSelectedAreaId(mergedAreas[0].id);
      setSelectedLocationId(mergedAreas[0].locations[0]?.id || null);
    }

    notify(
      `הייבוא הושלם: נוספו ${addedAreasCount} אזורים, ${addedLocationsCount} מתקנים ו-${addedRenewalsCount} רענונים (דולגו ${skippedDuplicatesCount} כפילויות)`,
      4500
    );
  };

  const handleImportExcelFile = async (e) => {
    const file = e.target.files?.[0];
    if (!file) return;

    setIsBackupBusy(true);
    try {
      const lowerName = file.name.toLowerCase();

      if (lowerName.endsWith('.json')) {
        const text = await file.text();
        const parsed = JSON.parse(text);
        const dispensersRows = [];
        const renewalsRows = [];
        (parsed.areas || []).forEach((a) => {
          if (!a.locations?.length) {
            dispensersRows.push({ areaId: a.id, areaName: a.name });
          } else {
            a.locations.forEach((l) => {
              dispensersRows.push({ ...l, areaId: a.id, areaName: a.name });
              (l.filterRefreshHistory || []).forEach((r) => {
                renewalsRows.push({
                  ...r,
                  areaId: a.id,
                  areaName: a.name,
                  locationId: l.id,
                  locationName: l.name
                });
              });
            });
          }
        });
        mergeImportedDatasets({
          dispensersRows,
          renewalsRows,
          historyRows: Array.isArray(parsed.actionHistory) ? parsed.actionHistory : []
        });
        return;
      }

      const headSnippet = await file.slice(0, 250).text();
      if (headSnippet.includes('schemas-microsoft-com:office:spreadsheet')) {
        const fullXml = await file.text();
        const sheets = parseSpreadsheetMLFallback(fullXml);
        mergeImportedDatasets({
          dispensersRows: sheets['מתקנים_ומיכלים'] || [],
          renewalsRows: sheets['היסטוריית_רענונים'] || [],
          historyRows: sheets['יומן_פעולות'] || []
        });
        return;
      }

      const XLSX = await loadXLSXLibrary();
      let wb;
      if (lowerName.endsWith('.csv')) {
        const csvText = (await file.text()).replace(/^\uFEFF/, '');
        wb = XLSX.read(csvText, { type: 'string', cellDates: true });
      } else {
        const buffer = await file.arrayBuffer();
        wb = XLSX.read(buffer, { type: 'array', cellDates: true });
      }

      const firstSheetName = wb.SheetNames.includes('מתקנים_ומיכלים')
        ? 'מתקנים_ומיכלים'
        : wb.SheetNames[0];

      const dispensersRows = firstSheetName
        ? XLSX.utils.sheet_to_json(wb.Sheets[firstSheetName], { defval: '' })
        : [];
      const renewalsRows = wb.SheetNames.includes('היסטוריית_רענונים')
        ? XLSX.utils.sheet_to_json(wb.Sheets['היסטוריית_רענונים'], { defval: '' })
        : [];
      const historyRows = wb.SheetNames.includes('יומן_פעולות')
        ? XLSX.utils.sheet_to_json(wb.Sheets['יומן_פעולות'], { defval: '' })
        : [];

      mergeImportedDatasets({ dispensersRows, renewalsRows, historyRows });
    } catch {
      notify('שגיאה בקריאת הקובץ. ודא שמדובר בקובץ גיבוי תקין (.xlsx / .xls / .csv / .json).', 4000);
    } finally {
      setIsBackupBusy(false);
      if (fileInputRef.current) {
        fileInputRef.current.value = '';
      }
    }
  };

  /**
   * Live update for any dispenser / tank / notes field.
   */
  const handleUpdateField = (field, rawValue) => {
    if (!selectedAreaId || !selectedLocationId || !currentLocation) return;

    let value = rawValue;
    if (field === 'filterAlef1' || field === 'filterAlef2') {
      value = toPercent(rawValue, 0);
    }

    const { todayStr, fullDateTime } = getNowStrings();
    const labelHe = fieldLabels[field] || field;

    setAreas((prev) =>
      prev.map((area) => {
        if (area.id !== selectedAreaId) return area;
        return {
          ...area,
          locations: area.locations.map((loc) =>
            loc.id === selectedLocationId
              ? { ...loc, [field]: value, lastTableUpdateDate: todayStr }
              : loc
          )
        };
      })
    );

    setLastGlobalUpdateDate(todayStr);

    const effectiveDispenserName =
      field === 'dispenserName' ? value || '(ללא שם)' : currentLocation.dispenserName;

    const formattedVal =
      field === 'filterAlef1' || field === 'filterAlef2' ? `${value}%` : value || '(ריק)';

    const logDescription = `אזור "${currentArea?.name}" > מקום "${currentLocation.name}" > מתקן "${effectiveDispenserName}": עודכן '${labelHe}' ל-${formattedVal}`;
    const logKey = `${selectedLocationId}:${field}`;
    const nowMs = Date.now();
    const isRapidRepeat =
      lastLogMetaRef.current.key === logKey && nowMs - lastLogMetaRef.current.timestamp < 2000;

    lastLogMetaRef.current = { key: logKey, timestamp: nowMs };

    setActionHistory((prev) => {
      if (isRapidRepeat && prev.length > 0) {
        const updatedFirst = {
          ...prev[0],
          date: fullDateTime,
          description: logDescription
        };
        return [updatedFirst, ...prev.slice(1)];
      }
      return [
        {
          id: nowMs + Math.random(),
          date: fullDateTime,
          user: 'מפעיל מערכת',
          description: logDescription
        },
        ...prev
      ];
    });
  };

  /**
   * Confirm Filter Renewal + Save Technician & Escort names to Quick-Pick History
   */
  const handleConfirmFilterRenewal = () => {
    if (!currentLocation || !currentArea) return;

    if (!technicianName.trim() || !escortName.trim()) {
      setRenewalError('חובה להזין שם טכנאי ושם מלווה כדי לתעד את רענון הפילטרים.');
      return;
    }

    if (!renewalDate || !parseLocalIsoDate(renewalDate)) {
      setRenewalError('חובה לבחור תאריך ביצוע רענון תקין.');
      return;
    }

    const { todayStr, timeStr, fullDateTime } = getNowStrings();
    const cleanTech = technicianName.trim();
    const cleanEscort = escortName.trim();
    const formattedRenewalDate = formatDateIL(renewalDate);
    const recordTimestamp =
      renewalDate === todayStr ? `${formattedRenewalDate} ${timeStr}` : formattedRenewalDate;

    setSavedTechnicians((prev) => [cleanTech, ...prev.filter((n) => n.trim() !== cleanTech)]);
    setSavedEscorts((prev) => [cleanEscort, ...prev.filter((n) => n.trim() !== cleanEscort)]);

    const renewalRecord = {
      id: Date.now() + Math.random(),
      date: recordTimestamp,
      isoDate: renewalDate,
      technician: cleanTech,
      escort: cleanEscort,
      action: resetFiltersOnRenewal
        ? `בוצע חידוש פילטרים מלא (איפוס ל-100%) למתקן "${currentLocation.dispenserName}"`
        : `בוצע חידוש פילטרים למתקן "${currentLocation.dispenserName}"`
    };

    setAreas((prev) =>
      prev.map((area) => {
        if (area.id !== selectedAreaId) return area;
        return {
          ...area,
          locations: area.locations.map((loc) => {
            if (loc.id !== selectedLocationId) return loc;

            const updatedHistory = [renewalRecord, ...(loc.filterRefreshHistory || [])].sort(
              (a, b) => (b.isoDate || '').localeCompare(a.isoDate || '')
            );

            const newestRefreshIso = updatedHistory[0]?.isoDate || renewalDate;

            return {
              ...loc,
              ...(resetFiltersOnRenewal
                ? {
                    filterMaga: 'תקין',
                    filterAlef1: 100,
                    filterAlef2: 100
                  }
                : {}),
              lastFilterRefreshDate: newestRefreshIso,
              lastTableUpdateDate: todayStr,
              filterRefreshHistory: updatedHistory
            };
          })
        };
      })
    );

    setLastGlobalUpdateDate(todayStr);

    setActionHistory((prev) => [
      {
        id: Date.now() + Math.random(),
        date: fullDateTime,
        user: `${cleanTech} (מלווה: ${cleanEscort})`,
        description: `רענון פילטרים למתקן "${currentLocation.dispenserName}" (אזור ${currentArea.name} / ${currentLocation.name}) בתאריך ${formattedRenewalDate} | טכנאי: ${cleanTech} | מלווה: ${cleanEscort}`
      },
      ...prev
    ]);

    setShowRenewalModal(false);
    setTechnicianName('');
    setEscortName('');
    setRenewalDate(getLocalIsoDate());
    setRenewalError('');
    notify(`רענון הפילטרים (${formattedRenewalDate}) תועד בהצלחה!`);
  };

  // --- AREA CRUD ---
  const handleAddArea = () => {
    const trimmed = newAreaName.trim();
    if (!trimmed) {
      notify('יש להזין שם או מספר אזור');
      return;
    }

    const newId = `area-${Date.now()}-${Math.random().toString(36).slice(2, 6)}`;
    const { fullDateTime } = getNowStrings();

    setAreas((prev) => [...prev, { id: newId, name: trimmed, locations: [] }]);
    setSelectedAreaId(newId);
    setSelectedLocationId(null);
    setNewAreaName('');
    setIsAddingArea(false);

    setActionHistory((prev) => [
      {
        id: Date.now() + Math.random(),
        date: fullDateTime,
        user: 'מפעיל מערכת',
        description: `נוסף אזור חדש: "${trimmed}"`
      },
      ...prev
    ]);

    notify(`אזור "${trimmed}" נוסף בהצלחה`);
  };

  const handleSaveAreaRename = (areaId) => {
    const trimmed = editingAreaNameValue.trim();
    if (!trimmed) {
      notify('שם האזור לא יכול להיות ריק');
      return;
    }

    const oldArea = areas.find((a) => a.id === areaId);
    const { fullDateTime } = getNowStrings();

    setAreas((prev) => prev.map((a) => (a.id === areaId ? { ...a, name: trimmed } : a)));
    setEditingAreaId(null);

    if (oldArea && oldArea.name !== trimmed) {
      setActionHistory((prev) => [
        {
          id: Date.now() + Math.random(),
          date: fullDateTime,
          user: 'מפעיל מערכת',
          description: `שונה שם אזור מ-"${oldArea.name}" ל-"${trimmed}"`
        },
        ...prev
      ]);
    }
    notify('שם האזור עודכן');
  };

  const requestDeleteArea = (areaId, e) => {
    e.stopPropagation();
    const area = areas.find((a) => a.id === areaId);
    if (!area) return;

    setConfirmModal({
      isOpen: true,
      type: 'area',
      id: areaId,
      title: `מחיקת אזור "${area.name}"`,
      message: `האם למחוק לצמיתות את האזור "${area.name}" וכל ${area.locations.length} המקומות והמתקנים המשויכים אליו?`
    });
  };

  // --- LOCATION CRUD ---
  const handleAddLocation = () => {
    const trimmedLoc = newLocName.trim();
    const trimmedItem = newItemName.trim();
    if (!selectedAreaId) return;
    if (!trimmedLoc || !trimmedItem) {
      notify('יש למלא גם שם מקום וגם שם מתקן מים');
      return;
    }

    const newLocId = `loc-${Date.now()}-${Math.random().toString(36).slice(2, 6)}`;
    const { todayStr, fullDateTime } = getNowStrings();

    setAreas((prev) =>
      prev.map((area) => {
        if (area.id !== selectedAreaId) return area;
        return {
          ...area,
          locations: [
            ...area.locations,
            {
              id: newLocId,
              name: trimmedLoc,
              dispenserName: trimmedItem,
              filterMaga: 'תקין',
              filterAlef1: 100,
              filterAlef2: 100,
              tagExpiry: todayStr,
              clockType: 'דיגיטלי',
              systemType: 'טורי',
              lastTableUpdateDate: todayStr,
              lastFilterRefreshDate: '',
              tankSensitivity: 'תקין',
              cartridgeCarriage: 'יש עגלה',
              tankTagExpiry: todayStr,
              notes: '',
              photos: [],
              filterRefreshHistory: []
            }
          ]
        };
      })
    );

    setSelectedLocationId(newLocId);
    setLastGlobalUpdateDate(todayStr);
    setNewLocName('');
    setNewItemName('');
    setIsAddingLocation(false);

    setActionHistory((prev) => [
      {
        id: Date.now() + Math.random(),
        date: fullDateTime,
        user: 'מפעיל מערכת',
        description: `באזור "${currentArea?.name}": נוסף מקום "${trimmedLoc}" עם מתקן "${trimmedItem}"`
      },
      ...prev
    ]);

    notify(`מקום "${trimmedLoc}" ומתקן "${trimmedItem}" נוספו`);
  };

  const handleSaveLocationRename = (locId) => {
    const trimmed = editingLocationNameValue.trim();
    if (!trimmed) {
      notify('שם המקום לא יכול להיות ריק');
      return;
    }

    const oldLoc = currentArea?.locations.find((l) => l.id === locId);
    const { fullDateTime } = getNowStrings();

    setAreas((prev) =>
      prev.map((area) => {
        if (area.id !== selectedAreaId) return area;
        return {
          ...area,
          locations: area.locations.map((l) => (l.id === locId ? { ...l, name: trimmed } : l))
        };
      })
    );
    setEditingLocationId(null);

    if (oldLoc && oldLoc.name !== trimmed) {
      setActionHistory((prev) => [
        {
          id: Date.now() + Math.random(),
          date: fullDateTime,
          user: 'מפעיל מערכת',
          description: `באזור "${currentArea?.name}": שונה שם מקום מ-"${oldLoc.name}" ל-"${trimmed}"`
        },
        ...prev
      ]);
    }
    notify('שם המקום עודכן');
  };

  const requestDeleteLocation = (locId, e) => {
    e.stopPropagation();
    const loc = currentArea?.locations.find((l) => l.id === locId);
    if (!loc) return;

    setConfirmModal({
      isOpen: true,
      type: 'location',
      id: locId,
      title: `מחיקת מקום "${loc.name}"`,
      message: `האם למחוק לצמיתות את המקום "${loc.name}" ואת מתקן המים "${loc.dispenserName}" המשויך אליו?`
    });
  };

  // Execute Custom Confirmation Modal
  const handleExecuteConfirm = () => {
    const { fullDateTime } = getNowStrings();

    if (confirmModal.type === 'area') {
      const target = areas.find((a) => a.id === confirmModal.id);
      const remaining = areas.filter((a) => a.id !== confirmModal.id);
      setAreas(remaining);

      if (selectedAreaId === confirmModal.id) {
        setSelectedAreaId(remaining[0]?.id || null);
        setSelectedLocationId(remaining[0]?.locations[0]?.id || null);
      }

      setActionHistory((prev) => [
        {
          id: Date.now() + Math.random(),
          date: fullDateTime,
          user: 'מפעיל מערכת',
          description: `נמחק אזור "${target?.name || ''}"`
        },
        ...prev
      ]);
      notify('האזור נמחק');
    } else if (confirmModal.type === 'location') {
      const target = currentArea?.locations.find((l) => l.id === confirmModal.id);
      const remainingLocs = currentArea?.locations.filter((l) => l.id !== confirmModal.id) || [];

      setAreas((prev) =>
        prev.map((area) =>
          area.id === selectedAreaId ? { ...area, locations: remainingLocs } : area
        )
      );

      if (selectedLocationId === confirmModal.id) {
        setSelectedLocationId(remainingLocs[0]?.id || null);
      }

      setActionHistory((prev) => [
        {
          id: Date.now() + Math.random(),
          date: fullDateTime,
          user: 'מפעיל מערכת',
          description: `באזור "${currentArea?.name}": נמחק מקום "${target?.name || ''}"`
        },
        ...prev
      ]);
      notify('המקום נמחק');
    } else if (confirmModal.type === 'reset_all') {
      setAreas([]);
      setSelectedAreaId(null);
      setSelectedLocationId(null);
      setLastGlobalUpdateDate('');
      setActionHistory([]);
      setSavedTechnicians([]);
      setSavedEscorts([]);
      try {
        window.localStorage.removeItem(STORAGE_KEY);
      } catch {
        // Ignore storage errors
      }
      notify('כל נתוני המערכת אופסו');
    }

    setConfirmModal({ isOpen: false, type: null, id: null, title: '', message: '' });
  };

  // Modern SaaS Design System Surface & Border Helpers
  const saasCard = darkMode
    ? 'bg-[#111827] border-slate-800 shadow-sm'
    : 'bg-white border-slate-200/90 shadow-[0_1px_3px_rgba(15,23,42,0.04)]';

  const saasPanelInset = darkMode
    ? 'bg-slate-900/75 border-slate-800'
    : 'bg-slate-50/80 border-slate-200/80';

  const saasInput = darkMode
    ? 'bg-slate-950 border-slate-700/90 text-slate-100 placeholder-slate-500 focus:border-indigo-500 focus:ring-2 focus:ring-indigo-500/20'
    : 'bg-white border-slate-300/90 text-slate-900 placeholder-slate-400 focus:border-indigo-600 focus:ring-2 focus:ring-indigo-600/15';

  // Helper: Continuous percentage-based color shade & gradient scale
  const getFilterShade = (pct) => {
    const clamped = toPercent(pct, 100);
    const hue = Math.round((clamped / 100) * 145);
    const startHue = Math.max(0, hue - 22);
    const lightness = darkMode ? 54 : 43;

    return {
      pct: clamped,
      color: `hsl(${hue}, 82%, ${lightness}%)`,
      gradient: `linear-gradient(to left, hsl(${startHue}, 86%, 50%), hsl(${hue}, 82%, ${lightness}%))`,
      tint: `hsla(${hue}, 85%, 50%, ${darkMode ? 0.13 : 0.08})`,
      border: `hsla(${hue}, 80%, 48%, ${darkMode ? 0.35 : 0.28})`,
      levelLabel: clamped <= 20 ? 'נמוך' : clamped <= 60 ? 'בינוני' : 'גבוה'
    };
  };

  const navItems = [
    { id: 'tracker', label: 'מעקב מתקנים', icon: Layers, badge: allDispensers.length },
    { id: 'dashboard', label: 'Dashboard ומעקב תקופתי', icon: LayoutDashboard },
    {
      id: 'dispenser_history',
      label: 'רענונים למתקן הנבחר',
      icon: Wrench,
      badge: currentLocation?.filterRefreshHistory?.length || 0
    },
    { id: 'history', label: 'יומן פעולות מערכת', icon: History, badge: actionHistory.length },
    { id: 'settings', label: 'הגדרות וגיבוי Excel', icon: Settings }
  ];

  const viewTitles = {
    tracker: 'ניהול ומעקב מתקני מים',
    dashboard: 'Dashboard ומעקב רענונים תקופתי',
    dispenser_history: 'היסטוריית רענון פילטרים למתקן',
    history: 'יומן פעולות ושינויים במערכת',
    settings: 'הגדרות מערכת וגיבוי Excel'
  };

  // Reusable Action Center Component (rendered in Tracker & Dashboard)
  const renderActionCenterBanner = () => {
    if (urgentAlerts.length === 0) {
      return (
        <section
          className={`rounded-2xl border px-4 py-3 flex items-center justify-between gap-3 ${
            darkMode
              ? 'bg-emerald-950/25 border-emerald-500/25 text-emerald-300'
              : 'bg-emerald-50/70 border-emerald-200/80 text-emerald-800'
          }`}
        >
          <div className="flex items-center gap-2.5 text-xs font-bold">
            <span className="w-7 h-7 rounded-lg bg-emerald-500/15 text-emerald-600 dark:text-emerald-400 flex items-center justify-center shrink-0">
              <CheckCircle2 size={16} />
            </span>
            <span>מרכז התראות: כל המתקנים, הפילטרים ותגי השמיש תקינים ובתוקף!</span>
          </div>
          <span className="text-[11px] font-bold px-2.5 py-0.5 rounded-full bg-emerald-500/15">
            0 חריגים
          </span>
        </section>
      );
    }

    const criticalCount = urgentAlerts.filter((a) => a.severity === 'critical').length;
    const warningCount = urgentAlerts.length - criticalCount;

    return (
      <section
        className={`rounded-2xl border overflow-hidden transition-all ${
          darkMode
            ? 'bg-[#14121D] border-rose-500/35 shadow-sm'
            : 'bg-white border-rose-200 shadow-sm'
        }`}
      >
        <div
          onClick={() => setActionCenterExpanded(!actionCenterExpanded)}
          className={`px-4 py-3 flex flex-wrap items-center justify-between gap-3 cursor-pointer ${
            darkMode ? 'bg-rose-500/10' : 'bg-rose-50/75'
          }`}
        >
          <div className="flex items-center gap-2.5">
            <span className="w-8 h-8 rounded-xl bg-rose-600 text-white flex items-center justify-center shadow-sm shrink-0">
              <Bell size={16} />
            </span>
            <div>
              <div className="flex items-center gap-2">
                <h2 className="text-xs sm:text-sm font-extrabold text-rose-700 dark:text-rose-300">
                  מרכז התראות ודברים דחופים (Action Center)
                </h2>
                <span className="px-2 py-0.5 rounded-full text-[11px] font-black bg-rose-600 text-white tabular-nums">
                  {urgentAlerts.length}
                </span>
              </div>
              <p className="text-[11px] opacity-75">
                {criticalCount > 0 && `${criticalCount} קריטיים (לא תקין / פג תוקף)`}
                {criticalCount > 0 && warningCount > 0 && ' • '}
                {warningCount > 0 && `${warningCount} אזהרות (מתחת ל-20% / פג תוך 30 יום)`}
              </p>
            </div>
          </div>

          <div className="flex items-center gap-2">
            <span className="text-[11px] font-bold opacity-70 hidden sm:inline">
              {actionCenterExpanded ? 'הסתר פירוט' : 'הצג פירוט'}
            </span>
            <span className="p-1.5 rounded-lg hover:bg-black/5 dark:hover:bg-white/5">
              {actionCenterExpanded ? <ChevronUp size={16} /> : <ChevronDown size={16} />}
            </span>
          </div>
        </div>

        {actionCenterExpanded && (
          <div className="p-3.5 max-h-56 overflow-y-auto divide-y dark:divide-slate-800/80 space-y-1.5">
            {urgentAlerts.map((alert) => (
              <div
                key={alert.id}
                onClick={() => {
                  setSelectedAreaId(alert.areaId);
                  setSelectedLocationId(alert.locationId);
                  setActiveView('tracker');
                }}
                className={`pt-2 first:pt-0 flex flex-col sm:flex-row sm:items-center justify-between gap-2 cursor-pointer group rounded-lg px-2 py-1.5 transition ${
                  darkMode ? 'hover:bg-slate-800/60' : 'hover:bg-slate-50'
                }`}
              >
                <div className="flex items-start sm:items-center gap-2.5 min-w-0">
                  <span
                    className={`mt-0.5 sm:mt-0 w-2.5 h-2.5 rounded-full shrink-0 ${
                      alert.severity === 'critical' ? 'bg-rose-600' : 'bg-amber-500'
                    }`}
                  />
                  <div className="min-w-0">
                    <div className="text-xs font-bold truncate">
                      אזור {alert.areaName} &gt; {alert.locationName} &gt;{' '}
                      <span className="text-indigo-600 dark:text-indigo-400">
                        {alert.dispenserName}
                      </span>
                    </div>
                    <div
                      className={`text-[11px] font-semibold ${
                        alert.severity === 'critical'
                          ? 'text-rose-600 dark:text-rose-400'
                          : 'text-amber-600 dark:text-amber-400'
                      }`}
                    >
                      {alert.message}
                    </div>
                  </div>
                </div>

                <span className="text-[11px] font-bold text-indigo-600 dark:text-indigo-400 shrink-0 inline-flex items-center gap-1 group-hover:underline">
                  טפל עכשיו <ArrowUpRight size={13} />
                </span>
              </div>
            ))}
          </div>
        )}
      </section>
    );
  };

  return (
    <div
      dir="rtl"
      className={`print-flow min-h-screen flex font-sans antialiased transition-colors duration-200 ${
        darkMode ? 'bg-[#090D16] text-slate-100' : 'bg-[#F8FAFC] text-slate-900'
      }`}
    >
      {/* Multi-page A4 Print Stylesheet (active only while Report Modal is open) */}
      {showReportModal && (
        <style>{`
          @media print {
            @page {
              size: A4 portrait;
              margin: 12mm;
            }
            .no-print {
              display: none !important;
            }
            .print-flow {
              display: block !important;
              min-height: auto !important;
              height: auto !important;
              background: #ffffff !important;
              color: #0f172a !important;
            }
            .print-reset {
              position: static !important;
              inset: auto !important;
              width: 100% !important;
              max-width: none !important;
              max-height: none !important;
              height: auto !important;
              overflow: visible !important;
              background: #ffffff !important;
              color: #0f172a !important;
              padding: 0 !important;
              margin: 0 !important;
              box-shadow: none !important;
              border: none !important;
              border-radius: 0 !important;
            }
            #aquaops-printable-report tr {
              break-inside: avoid;
              page-break-inside: avoid;
            }
            #aquaops-printable-report thead {
              display: table-header-group;
            }
          }
        `}</style>
      )}

      {/* Hidden File Inputs for Excel Backup & Mobile Camera / Gallery */}
      <input
        ref={fileInputRef}
        type="file"
        accept=".xlsx,.xls,.csv,.json"
        onChange={handleImportExcelFile}
        className="hidden"
      />
      <input
        ref={cameraInputRef}
        type="file"
        accept="image/*"
        capture="environment"
        onChange={handlePhotoFileSelected}
        className="hidden"
      />
      <input
        ref={galleryInputRef}
        type="file"
        accept="image/*"
        onChange={handlePhotoFileSelected}
        className="hidden"
      />

      {/* ===================================================================== */}
      {/* PERSISTENT MODERN SAAS SIDEBAR (Desktop Right Rail)                   */}
      {/* ===================================================================== */}
      <aside
        className={`no-print hidden lg:flex lg:w-64 xl:w-72 shrink-0 flex-col justify-between border-l sticky top-0 h-screen z-30 ${
          darkMode ? 'bg-[#0D1320] border-slate-800/90' : 'bg-white border-slate-200/90'
        }`}
      >
        <div>
          {/* SaaS Product Brand Header */}
          <div
            onClick={() => setActiveView('tracker')}
            className={`px-5 py-5 border-b flex items-center justify-between cursor-pointer ${
              darkMode ? 'border-slate-800/90' : 'border-slate-100'
            }`}
          >
            <div className="flex items-center gap-3">
              <div className="w-9 h-9 rounded-xl bg-indigo-600 text-white flex items-center justify-center shadow-sm shadow-indigo-600/25">
                <Droplet size={19} />
              </div>
              <div>
                <div className="flex items-center gap-2">
                  <span className="font-bold text-sm tracking-tight">מעקב מתקני מים</span>
                </div>
                <span className="text-[11px] font-medium opacity-55 block">
                  AquaOps Field Suite
                </span>
              </div>
            </div>
            <span
              className={`px-2 py-0.5 text-[10px] font-bold rounded-md border ${
                darkMode
                  ? 'bg-indigo-500/10 text-indigo-400 border-indigo-500/25'
                  : 'bg-indigo-50 text-indigo-700 border-indigo-200/80'
              }`}
            >
              {APP_VERSION}
            </span>
          </div>

          {/* Primary SaaS Sidebar Navigation */}
          <div className="p-3.5 space-y-1">
            <div className="px-3 py-1.5 text-[11px] font-semibold tracking-wider uppercase opacity-45">
              ניווט ראשי
            </div>
            {navItems.map((item) => {
              const Icon = item.icon;
              const isActive = activeView === item.id;
              return (
                <button
                  key={item.id}
                  onClick={() => setActiveView(item.id)}
                  className={`w-full flex items-center justify-between px-3.5 py-2.5 rounded-xl text-xs font-semibold transition-all cursor-pointer ${
                    isActive
                      ? darkMode
                        ? 'bg-indigo-600/15 text-indigo-300 border border-indigo-500/30'
                        : 'bg-indigo-50/90 text-indigo-700 border border-indigo-200/70 shadow-[0_1px_2px_rgba(79,70,229,0.06)]'
                      : darkMode
                      ? 'text-slate-400 hover:text-slate-100 hover:bg-slate-800/50 border border-transparent'
                      : 'text-slate-600 hover:text-slate-900 hover:bg-slate-100/70 border border-transparent'
                  }`}
                >
                  <span className="flex items-center gap-2.5">
                    <Icon
                      size={16}
                      className={
                        isActive ? 'text-indigo-600 dark:text-indigo-400' : 'opacity-70'
                      }
                    />
                    <span>{item.label}</span>
                  </span>
                  {typeof item.badge === 'number' && (
                    <span
                      className={`px-2 py-0.5 text-[10px] font-bold rounded-full tabular-nums ${
                        isActive
                          ? 'bg-indigo-600 text-white'
                          : darkMode
                          ? 'bg-slate-800 text-slate-400'
                          : 'bg-slate-100 text-slate-600'
                      }`}
                    >
                      {item.badge}
                    </span>
                  )}
                </button>
              );
            })}

            {/* Official PDF Report Quick Sidebar Trigger */}
            <button
              onClick={() => setShowReportModal(true)}
              className={`w-full flex items-center justify-between px-3.5 py-2.5 rounded-xl text-xs font-semibold transition-all cursor-pointer mt-2 ${
                darkMode
                  ? 'text-indigo-300 bg-indigo-500/10 hover:bg-indigo-500/20 border border-indigo-500/25'
                  : 'text-indigo-700 bg-indigo-50/60 hover:bg-indigo-100/80 border border-indigo-200/70'
              }`}
            >
              <span className="flex items-center gap-2.5">
                <Printer size={16} className="text-indigo-600 dark:text-indigo-400" />
                <span>הפק דוח ביקורת PDF</span>
              </span>
            </button>
          </div>
        </div>

        {/* Sidebar Bottom Workspace Card, Excel Quick Actions & Theme Switcher */}
        <div
          className={`p-4 border-t space-y-2.5 ${
            darkMode ? 'border-slate-800/90' : 'border-slate-100'
          }`}
        >
          <div className={`p-3 rounded-xl border text-xs space-y-1.5 ${saasPanelInset}`}>
            <div className="flex items-center justify-between font-semibold">
              <span className="opacity-70">סיכום מהיר</span>
              <span className="text-[10px] px-1.5 py-0.5 rounded bg-emerald-500/10 text-emerald-600 dark:text-emerald-400 font-bold">
                OFFLINE READY
              </span>
            </div>
            <div className="flex items-center justify-between text-[11px] opacity-80">
              <span>אזורים: {areas.length}</span>
              <span>מתקנים: {allDispensers.length}</span>
            </div>
          </div>

          <div className="grid grid-cols-2 gap-2">
            <button
              onClick={handleExportExcelBackup}
              disabled={isBackupBusy}
              className="px-2.5 py-2 rounded-xl bg-emerald-600 hover:bg-emerald-700 disabled:opacity-50 text-white text-[11px] font-bold flex items-center justify-center gap-1.5 shadow-sm transition cursor-pointer"
              title="ייצוא גיבוי מלא לקובץ Excel"
            >
              <Download size={13} />
              <span>ייצוא Excel</span>
            </button>
            <button
              onClick={() => fileInputRef.current?.click()}
              disabled={isBackupBusy}
              className={`px-2.5 py-2 rounded-xl border text-[11px] font-bold flex items-center justify-center gap-1.5 transition cursor-pointer ${
                darkMode
                  ? 'bg-slate-900 border-slate-700 text-slate-200 hover:bg-slate-800'
                  : 'bg-white border-slate-200 text-slate-700 hover:bg-slate-50'
              }`}
              title="ייבוא קובץ Excel ומיזוג ללא כפילויות"
            >
              <Upload size={13} />
              <span>ייבוא Excel</span>
            </button>
          </div>

          <button
            onClick={() => setDarkMode(!darkMode)}
            className={`w-full flex items-center justify-between px-3.5 py-2.5 rounded-xl border text-xs font-semibold transition cursor-pointer ${
              darkMode
                ? 'bg-slate-900 border-slate-800 text-slate-200 hover:bg-slate-800'
                : 'bg-white border-slate-200 text-slate-700 hover:bg-slate-50'
            }`}
          >
            <span className="flex items-center gap-2">
              {darkMode ? (
                <Sun size={15} className="text-amber-400" />
              ) : (
                <Moon size={15} className="text-slate-500" />
              )}
              <span>{darkMode ? 'מצב תצוגה בהיר' : 'מצב תצוגה כהה'}</span>
            </span>
          </button>
        </div>
      </aside>

      {/* MOBILE NAVIGATION DRAWER */}
      {mobileMenuOpen && (
        <div className="no-print fixed inset-0 z-50 lg:hidden flex">
          <div
            className="fixed inset-0 bg-slate-950/50 backdrop-blur-sm"
            onClick={() => setMobileMenuOpen(false)}
          />
          <aside
            className={`relative w-72 max-w-[85vw] h-full flex flex-col justify-between z-10 border-l ${
              darkMode ? 'bg-[#0D1320] border-slate-800' : 'bg-white border-slate-200'
            }`}
          >
            <div>
              <div className="px-5 py-4 border-b dark:border-slate-800 flex items-center justify-between">
                <div className="flex items-center gap-2.5">
                  <div className="w-8 h-8 rounded-lg bg-indigo-600 text-white flex items-center justify-center">
                    <Droplet size={17} />
                  </div>
                  <span className="font-bold text-sm">מעקב מתקני מים</span>
                  <span className="text-[10px] font-bold px-1.5 py-0.5 rounded bg-indigo-500/10 text-indigo-500">
                    {APP_VERSION}
                  </span>
                </div>
                <button
                  onClick={() => setMobileMenuOpen(false)}
                  className="p-1.5 rounded-lg hover:bg-slate-100 dark:hover:bg-slate-800 cursor-pointer"
                >
                  <X size={18} />
                </button>
              </div>
              <div className="p-3 space-y-1">
                {navItems.map((item) => {
                  const Icon = item.icon;
                  const isActive = activeView === item.id;
                  return (
                    <button
                      key={item.id}
                      onClick={() => {
                        setActiveView(item.id);
                        setMobileMenuOpen(false);
                      }}
                      className={`w-full flex items-center justify-between px-3.5 py-2.5 rounded-xl text-xs font-semibold transition cursor-pointer ${
                        isActive
                          ? 'bg-indigo-600 text-white'
                          : 'opacity-75 hover:opacity-100 hover:bg-slate-100 dark:hover:bg-slate-800'
                      }`}
                    >
                      <span className="flex items-center gap-2.5">
                        <Icon size={16} />
                        <span>{item.label}</span>
                      </span>
                      {typeof item.badge === 'number' && (
                        <span className="text-[10px] font-bold px-2 py-0.5 rounded-full bg-black/10 dark:bg-white/10">
                          {item.badge}
                        </span>
                      )}
                    </button>
                  );
                })}

                <button
                  onClick={() => {
                    setMobileMenuOpen(false);
                    setShowReportModal(true);
                  }}
                  className="w-full flex items-center gap-2.5 px-3.5 py-2.5 rounded-xl text-xs font-bold text-indigo-600 dark:text-indigo-400 bg-indigo-500/10 mt-2 cursor-pointer"
                >
                  <Printer size={16} />
                  <span>הפק דוח ביקורת PDF</span>
                </button>
              </div>
            </div>
            <div className="p-4 border-t dark:border-slate-800 space-y-2">
              <div className="grid grid-cols-2 gap-2">
                <button
                  onClick={() => {
                    setMobileMenuOpen(false);
                    handleExportExcelBackup();
                  }}
                  className="py-2 px-2.5 rounded-xl bg-emerald-600 text-white text-xs font-bold flex items-center justify-center gap-1.5 cursor-pointer"
                >
                  <Download size={14} /> ייצוא Excel
                </button>
                <button
                  onClick={() => {
                    setMobileMenuOpen(false);
                    fileInputRef.current?.click();
                  }}
                  className="py-2 px-2.5 rounded-xl border dark:border-slate-700 text-xs font-bold flex items-center justify-center gap-1.5 cursor-pointer"
                >
                  <Upload size={14} /> ייבוא Excel
                </button>
              </div>
              <button
                onClick={() => setDarkMode(!darkMode)}
                className="w-full py-2.5 px-3 rounded-xl border dark:border-slate-700 text-xs font-semibold flex items-center justify-center gap-2 cursor-pointer"
              >
                {darkMode ? <Sun size={15} className="text-amber-400" /> : <Moon size={15} />}
                <span>{darkMode ? 'עבור למצב בהיר' : 'עבור למצב כהה'}</span>
              </button>
            </div>
          </aside>
        </div>
      )}

      {/* ===================================================================== */}
      {/* MAIN SAAS WORKSPACE CONTENT SHELL                                     */}
      {/* ===================================================================== */}
      <div className="print-flow flex-1 flex flex-col min-w-0">
        {/* Sticky SaaS Top Header Bar */}
        <header
          className={`no-print sticky top-0 z-20 border-b backdrop-blur-md px-4 sm:px-6 lg:px-8 py-3.5 flex flex-wrap items-center justify-between gap-3 ${
            darkMode
              ? 'bg-[#090D16]/85 border-slate-800/90'
              : 'bg-white/85 border-slate-200/90'
          }`}
        >
          {/* Right Side: Mobile Menu Trigger + Page Title & Breadcrumb */}
          <div className="flex items-center gap-3 min-w-0">
            <button
              onClick={() => setMobileMenuOpen(true)}
              className={`lg:hidden p-2 rounded-xl border cursor-pointer ${
                darkMode
                  ? 'bg-slate-900 border-slate-800 text-slate-200'
                  : 'bg-white border-slate-200 text-slate-700'
              }`}
              aria-label="פתח תפריט"
            >
              <Menu size={18} />
            </button>

            <div className="min-w-0">
              <div className="flex items-center gap-2">
                <h1 className="text-base sm:text-lg font-bold tracking-tight truncate">
                  {viewTitles[activeView]}
                </h1>
                <span className="lg:hidden px-2 py-0.5 text-[10px] font-bold rounded-md bg-indigo-500/10 text-indigo-600 dark:text-indigo-400">
                  {APP_VERSION}
                </span>
              </div>

              {/* SaaS Breadcrumb */}
              <div className="flex items-center gap-1.5 text-[11px] font-medium opacity-60 truncate">
                <span>אזור</span>
                <span>/</span>
                <span
                  className={currentArea ? 'text-indigo-600 dark:text-indigo-400 font-semibold' : ''}
                >
                  {currentArea?.name || 'לא נבחר'}
                </span>
                <span>/</span>
                <span
                  className={
                    currentLocation ? 'text-indigo-600 dark:text-indigo-400 font-semibold' : ''
                  }
                >
                  {currentLocation?.name || 'לא נבחר'}
                </span>
              </div>
            </div>
          </div>

          {/* Top-Left Side: Urgent Alert Counter + Official PDF Report + Excel Backup + Live Date */}
          <div className="flex flex-wrap items-center gap-2">
            {/* Live Action Center Alert Counter Button */}
            <button
              onClick={() => {
                if (activeView !== 'tracker' && activeView !== 'dashboard') {
                  setActiveView('tracker');
                  setActionCenterExpanded(true);
                } else {
                  setActionCenterExpanded((prev) => !prev);
                }
              }}
              className={`px-3 py-2 rounded-xl border text-xs font-bold flex items-center gap-1.5 transition cursor-pointer ${
                urgentAlerts.length > 0
                  ? 'bg-rose-600 hover:bg-rose-700 text-white border-rose-600 shadow-sm'
                  : darkMode
                  ? 'bg-emerald-950/40 border-emerald-500/30 text-emerald-400'
                  : 'bg-emerald-50 border-emerald-200 text-emerald-700'
              }`}
              title="מרכז התראות ודברים דחופים"
            >
              <Bell size={14} />
              <span className="hidden sm:inline">התראות:</span>
              <span className="tabular-nums font-black">{urgentAlerts.length}</span>
            </button>

            {/* Official Printable / PDF Report Button */}
            <button
              onClick={() => setShowReportModal(true)}
              className={`px-3 py-2 rounded-xl border text-xs font-bold flex items-center gap-1.5 transition cursor-pointer ${
                darkMode
                  ? 'bg-[#111827] border-slate-700 text-indigo-300 hover:bg-slate-800'
                  : 'bg-indigo-50/70 border-indigo-200 text-indigo-700 hover:bg-indigo-100/80'
              }`}
              title="הפק דוח ביקורת רשמי להדפסה או PDF"
            >
              <Printer size={14} />
              <span className="hidden md:inline">דוח PDF</span>
            </button>

            <div className="hidden xl:flex items-center gap-2">
              <button
                onClick={handleExportExcelBackup}
                disabled={isBackupBusy}
                className="px-3 py-2 rounded-xl bg-emerald-600 hover:bg-emerald-700 disabled:opacity-50 text-white text-xs font-bold flex items-center gap-1.5 shadow-sm transition cursor-pointer"
                title="ייצוא קובץ גיבוי Excel (.xlsx)"
              >
                <Download size={14} />
                <span>ייצוא Excel</span>
              </button>

              <button
                onClick={() => fileInputRef.current?.click()}
                disabled={isBackupBusy}
                className={`px-3 py-2 rounded-xl border text-xs font-bold flex items-center gap-1.5 transition cursor-pointer ${
                  darkMode
                    ? 'bg-[#111827] border-slate-700 text-slate-200 hover:bg-slate-800'
                    : 'bg-white border-slate-200 text-slate-700 hover:bg-slate-50'
                }`}
                title="ייבוא קובץ גיבוי Excel (מיזוג ללא דריסה וללא כפילויות)"
              >
                <Upload size={14} className="text-indigo-600 dark:text-indigo-400" />
                <span>ייבוא Excel</span>
              </button>
            </div>

            <div
              className={`px-3 py-1.5 rounded-xl border flex items-center gap-2.5 ${
                darkMode ? 'bg-[#111827] border-slate-800' : 'bg-slate-50 border-slate-200/90'
              }`}
            >
              <div className="text-right">
                <span className="block text-[10px] font-semibold opacity-60">
                  עדכון טבלה:
                </span>
                <span className="font-mono font-bold text-xs sm:text-sm tabular-nums text-indigo-600 dark:text-indigo-400">
                  {headerTableUpdateDate ? formatDateIL(headerTableUpdateDate) : '__/__/____'}
                </span>
              </div>
              <div
                className={`w-7 h-7 rounded-lg flex items-center justify-center ${
                  headerTableUpdateDate === todayLocalIso
                    ? 'bg-emerald-500/15 text-emerald-600 dark:text-emerald-400'
                    : 'bg-slate-500/10 opacity-45'
                }`}
              >
                <Activity size={15} />
              </div>
            </div>
          </div>
        </header>

        {/* SAAS TOAST NOTIFICATION */}
        {toast && (
          <div className="no-print fixed bottom-5 left-5 z-50 bg-slate-900 dark:bg-slate-800 text-white px-4 py-3 rounded-xl shadow-xl border border-slate-700/80 flex items-center gap-2.5 max-w-md">
            <CheckCircle2 size={17} className="text-emerald-400 shrink-0" />
            <span className="text-xs sm:text-sm font-semibold">{toast}</span>
          </div>
        )}

        {/* =================================================================== */}
        {/* FILTER RENEWAL MODAL (With Quick Technician & Escort Name History)  */}
        {/* =================================================================== */}
        {showRenewalModal && (
          <div className="no-print fixed inset-0 z-50 bg-slate-950/60 backdrop-blur-sm flex items-center justify-center p-4 overflow-y-auto">
            <div
              className={`w-full max-w-md p-6 rounded-2xl shadow-2xl border my-auto ${
                darkMode
                  ? 'bg-[#111827] border-slate-800 text-slate-100'
                  : 'bg-white border-slate-200 text-slate-900'
              }`}
            >
              <div className="flex items-center justify-between pb-3.5 mb-4 border-b dark:border-slate-800">
                <div className="flex items-center gap-2.5">
                  <div className="w-8 h-8 rounded-lg bg-emerald-500/10 text-emerald-600 dark:text-emerald-400 flex items-center justify-center">
                    <Wrench size={17} />
                  </div>
                  <h3 className="text-base font-bold">תיעוד רענון פילטרים</h3>
                </div>
                <button
                  onClick={() => {
                    setShowRenewalModal(false);
                    setRenewalError('');
                  }}
                  className="p-1.5 rounded-lg hover:bg-slate-100 dark:hover:bg-slate-800 cursor-pointer"
                >
                  <X size={17} />
                </button>
              </div>

              <div className="mb-4 p-3 rounded-xl bg-emerald-500/10 border border-emerald-500/20 text-xs font-medium text-emerald-700 dark:text-emerald-300">
                מתקן: <strong className="font-bold">{currentLocation?.dispenserName}</strong> (אזור{' '}
                {currentArea?.name} / {currentLocation?.name})
              </div>

              {renewalError && (
                <div className="mb-4 p-3 rounded-xl bg-rose-500/10 border border-rose-500/30 text-rose-600 dark:text-rose-400 text-xs font-semibold flex items-center gap-2">
                  <AlertTriangle size={15} className="shrink-0" />
                  <span>{renewalError}</span>
                </div>
              )}

              <div className="space-y-4 mb-6">
                <div>
                  <label className="block text-xs font-semibold mb-1.5 flex items-center gap-1.5">
                    <Calendar size={14} className="text-emerald-600 dark:text-emerald-400" /> תאריך ביצוע הרענון:
                  </label>
                  <input
                    type="date"
                    value={renewalDate}
                    onChange={(e) => {
                      setRenewalDate(e.target.value);
                      if (renewalError) setRenewalError('');
                    }}
                    className={`w-full px-3.5 py-2.5 rounded-xl border text-sm font-semibold outline-none transition ${saasInput}`}
                  />
                </div>

                {/* Technician Input + Quick History Chips */}
                <div>
                  <label className="block text-xs font-semibold mb-1.5 flex items-center gap-1.5">
                    <UserCheck size={14} className="text-emerald-600 dark:text-emerald-400" /> שם טכנאי:
                  </label>
                  <input
                    type="text"
                    placeholder="הזן שם טכנאי או בחר מהיסטוריה..."
                    value={technicianName}
                    onChange={(e) => {
                      setTechnicianName(e.target.value);
                      if (renewalError) setRenewalError('');
                    }}
                    className={`w-full px-3.5 py-2.5 rounded-xl border text-sm font-medium outline-none transition ${saasInput}`}
                    autoFocus
                  />
                  {technicianHistoryList.length > 0 && (
                    <div className="mt-2">
                      <span className="block text-[10px] font-semibold opacity-60 mb-1">
                        בחירה מהירה של טכנאי:
                      </span>
                      <div className="flex flex-wrap gap-1.5">
                        {technicianHistoryList.map((name) => {
                          const isSelected = technicianName.trim() === name;
                          return (
                            <span
                              key={name}
                              className={`inline-flex items-center gap-1 px-2.5 py-1 rounded-lg border text-[11px] font-semibold transition ${
                                isSelected
                                  ? 'bg-emerald-600 text-white border-emerald-600'
                                  : darkMode
                                  ? 'bg-slate-900 border-slate-700 text-slate-300 hover:border-emerald-500/50'
                                  : 'bg-slate-50 border-slate-200 text-slate-700 hover:border-emerald-400'
                              }`}
                            >
                              <button
                                type="button"
                                onClick={() => {
                                  setTechnicianName(name);
                                  if (renewalError) setRenewalError('');
                                }}
                                className="cursor-pointer"
                              >
                                {name}
                              </button>
                              {savedTechnicians.includes(name) && (
                                <button
                                  type="button"
                                  onClick={() =>
                                    setSavedTechnicians((prev) => prev.filter((n) => n !== name))
                                  }
                                  className="opacity-60 hover:opacity-100 mr-0.5 cursor-pointer"
                                  title="הסר מהרשימה"
                                >
                                  ×
                                </button>
                              )}
                            </span>
                          );
                        })}
                      </div>
                    </div>
                  )}
                </div>

                {/* Escort Input + Quick History Chips */}
                <div>
                  <label className="block text-xs font-semibold mb-1.5 flex items-center gap-1.5">
                    <ShieldCheck size={14} className="text-indigo-600 dark:text-indigo-400" /> שם מלווה:
                  </label>
                  <input
                    type="text"
                    placeholder="הזן שם מלווה או בחר מהיסטוריה..."
                    value={escortName}
                    onChange={(e) => {
                      setEscortName(e.target.value);
                      if (renewalError) setRenewalError('');
                    }}
                    onKeyDown={(e) => e.key === 'Enter' && handleConfirmFilterRenewal()}
                    className={`w-full px-3.5 py-2.5 rounded-xl border text-sm font-medium outline-none transition ${saasInput}`}
                  />
                  {escortHistoryList.length > 0 && (
                    <div className="mt-2">
                      <span className="block text-[10px] font-semibold opacity-60 mb-1">
                        בחירה מהירה של מלווה:
                      </span>
                      <div className="flex flex-wrap gap-1.5">
                        {escortHistoryList.map((name) => {
                          const isSelected = escortName.trim() === name;
                          return (
                            <span
                              key={name}
                              className={`inline-flex items-center gap-1 px-2.5 py-1 rounded-lg border text-[11px] font-semibold transition ${
                                isSelected
                                  ? 'bg-indigo-600 text-white border-indigo-600'
                                  : darkMode
                                  ? 'bg-slate-900 border-slate-700 text-slate-300 hover:border-indigo-500/50'
                                  : 'bg-slate-50 border-slate-200 text-slate-700 hover:border-indigo-400'
                              }`}
                            >
                              <button
                                type="button"
                                onClick={() => {
                                  setEscortName(name);
                                  if (renewalError) setRenewalError('');
                                }}
                                className="cursor-pointer"
                              >
                                {name}
                              </button>
                              {savedEscorts.includes(name) && (
                                <button
                                  type="button"
                                  onClick={() =>
                                    setSavedEscorts((prev) => prev.filter((n) => n !== name))
                                  }
                                  className="opacity-60 hover:opacity-100 mr-0.5 cursor-pointer"
                                  title="הסר מהרשימה"
                                >
                                  ×
                                </button>
                              )}
                            </span>
                          );
                        })}
                      </div>
                    </div>
                  )}
                </div>

                <label
                  className={`flex items-center gap-2.5 p-3 rounded-xl border text-xs font-semibold cursor-pointer select-none ${saasPanelInset}`}
                >
                  <input
                    type="checkbox"
                    checked={resetFiltersOnRenewal}
                    onChange={(e) => setResetFiltersOnRenewal(e.target.checked)}
                    className="w-4 h-4 rounded accent-emerald-600 cursor-pointer"
                  />
                  <span>אפס אחוזי פילטר אלף 1 ו-2 ל-100% וסמן פילטר מגה כתקין</span>
                </label>
              </div>

              <div className="flex gap-2.5">
                <button
                  onClick={handleConfirmFilterRenewal}
                  className="flex-1 bg-emerald-600 hover:bg-emerald-700 text-white py-2.5 rounded-xl text-xs font-bold shadow-sm transition cursor-pointer"
                >
                  אשר ותעד רענון
                </button>
                <button
                  onClick={() => {
                    setShowRenewalModal(false);
                    setRenewalError('');
                  }}
                  className="px-4 bg-slate-100 dark:bg-slate-800 hover:bg-slate-200 dark:hover:bg-slate-700 py-2.5 rounded-xl text-xs font-semibold transition cursor-pointer"
                >
                  ביטול
                </button>
              </div>
            </div>
          </div>
        )}

        {/* =================================================================== */}
        {/* QR CODE LABEL MODAL PER DISPENSER (OFFLINE READY)                   */}
        {/* =================================================================== */}
        {showQrModal && currentLocation && currentArea && (
          <div className="no-print fixed inset-0 z-50 bg-slate-950/60 backdrop-blur-sm flex items-center justify-center p-4">
            <div
              className={`w-full max-w-sm p-6 rounded-2xl shadow-2xl border text-center ${
                darkMode
                  ? 'bg-[#111827] border-slate-800 text-slate-100'
                  : 'bg-white border-slate-200 text-slate-900'
              }`}
            >
              <div className="flex items-center justify-between pb-3 mb-4 border-b dark:border-slate-800">
                <div className="flex items-center gap-2">
                  <div className="w-8 h-8 rounded-lg bg-indigo-500/10 text-indigo-600 dark:text-indigo-400 flex items-center justify-center">
                    <QrCode size={17} />
                  </div>
                  <h3 className="text-sm font-bold">תווית QR למתקן</h3>
                </div>
                <button
                  onClick={() => setShowQrModal(false)}
                  className="p-1.5 rounded-lg hover:bg-slate-100 dark:hover:bg-slate-800 cursor-pointer"
                >
                  <X size={17} />
                </button>
              </div>

              <div className="mb-4">
                <QrCodeSvg value={currentDispenserQrUrl} size={210} />
              </div>

              <div className={`p-3 rounded-xl border mb-5 space-y-1 ${saasPanelInset}`}>
                <div className="text-sm font-extrabold text-indigo-600 dark:text-indigo-400">
                  {currentLocation.dispenserName}
                </div>
                <div className="text-xs font-semibold opacity-75">
                  אזור: {currentArea.name} • מקום: {currentLocation.name}
                </div>
                <p className="text-[11px] opacity-60 pt-1">
                  הדפס או הדבק תווית זו על המתקן. סריקה במצלמת הפלאפון תפתח ישירות את כרטיס המתקן!
                </p>
              </div>

              <div className="grid grid-cols-2 gap-2.5">
                <button
                  onClick={handleDownloadQrPng}
                  className="py-2.5 px-3 rounded-xl bg-indigo-600 hover:bg-indigo-700 text-white text-xs font-bold flex items-center justify-center gap-1.5 shadow-sm transition cursor-pointer"
                >
                  <Download size={14} /> הורד תווית PNG
                </button>
                <button
                  onClick={() => setShowQrModal(false)}
                  className="py-2.5 px-3 rounded-xl bg-slate-100 dark:bg-slate-800 hover:bg-slate-200 dark:hover:bg-slate-700 text-xs font-semibold transition cursor-pointer"
                >
                  סגור
                </button>
              </div>
            </div>
          </div>
        )}

        {/* =================================================================== */}
        {/* PHOTO LIGHTBOX PREVIEW MODAL                                        */}
        {/* =================================================================== */}
        {lightboxPhoto && (
          <div
            className="no-print fixed inset-0 z-50 bg-slate-950/85 backdrop-blur-sm flex items-center justify-center p-4"
            onClick={() => setLightboxPhoto(null)}
          >
            <div
              className="max-w-2xl w-full bg-slate-900 border border-slate-700 rounded-2xl overflow-hidden shadow-2xl"
              onClick={(e) => e.stopPropagation()}
            >
              <div className="px-4 py-3 border-b border-slate-800 flex items-center justify-between text-white">
                <div className="text-xs font-bold">
                  {lightboxPhoto.category} •{' '}
                  <span className="font-mono opacity-75">{lightboxPhoto.date}</span>
                </div>
                <div className="flex items-center gap-2">
                  <button
                    onClick={() => handleDeletePhoto(lightboxPhoto.id)}
                    className="px-2.5 py-1 rounded-lg bg-rose-600 hover:bg-rose-700 text-white text-xs font-bold flex items-center gap-1 cursor-pointer"
                  >
                    <Trash2 size={13} /> מחק תמונה
                  </button>
                  <button
                    onClick={() => setLightboxPhoto(null)}
                    className="p-1.5 rounded-lg hover:bg-slate-800 cursor-pointer"
                  >
                    <X size={18} />
                  </button>
                </div>
              </div>
              <div className="p-2 bg-black flex items-center justify-center max-h-[75vh]">
                <img
                  src={lightboxPhoto.dataUrl}
                  alt={lightboxPhoto.category}
                  className="max-h-[70vh] w-auto object-contain rounded-lg"
                />
              </div>
            </div>
          </div>
        )}

        {/* =================================================================== */}
        {/* OFFICIAL A4 PRINTABLE / PDF INSPECTION REPORT MODAL                 */}
        {/* =================================================================== */}
        {showReportModal && (
          <div className="print-reset fixed inset-0 z-50 bg-slate-950/70 backdrop-blur-sm flex items-center justify-center p-2 sm:p-4 overflow-y-auto">
            <div className="print-reset w-full max-w-5xl bg-white text-slate-900 rounded-2xl shadow-2xl border border-slate-200 overflow-hidden my-auto">
              {/* Top Modal Toolbar (Hidden when printing) */}
              <div className="no-print px-5 py-3.5 bg-slate-900 text-white flex flex-wrap items-center justify-between gap-3">
                <div className="flex items-center gap-2.5">
                  <Printer size={18} className="text-indigo-400" />
                  <span className="text-sm font-bold">תצוגה מקדימה: דוח ביקורת רשמי להדפסה / PDF</span>
                </div>

                <div className="flex flex-wrap items-center gap-2">
                  <input
                    type="text"
                    placeholder="שם עורך הביקורת (אופציונלי)..."
                    value={reportInspectorName}
                    onChange={(e) => setReportInspectorName(e.target.value)}
                    className="px-3 py-1.5 rounded-lg bg-slate-800 border border-slate-700 text-xs text-white placeholder-slate-400 outline-none"
                  />
                  <button
                    onClick={() => window.print()}
                    className="px-4 py-1.5 rounded-lg bg-indigo-600 hover:bg-indigo-700 text-white text-xs font-bold flex items-center gap-1.5 shadow-sm cursor-pointer"
                  >
                    <Printer size={14} /> הדפס / שמור כ-PDF
                  </button>
                  <button
                    onClick={() => setShowReportModal(false)}
                    className="p-1.5 rounded-lg hover:bg-slate-800 text-slate-300 cursor-pointer"
                  >
                    <X size={18} />
                  </button>
                </div>
              </div>

              {/* Printable A4 Sheet Content */}
              <div
                id="aquaops-printable-report"
                dir="rtl"
                className="print-reset p-6 sm:p-10 space-y-6 max-h-[82vh] overflow-y-auto bg-white text-slate-900"
              >
                {/* Official Report Header */}
                <div className="border-b-2 border-slate-900 pb-4 flex flex-col sm:flex-row sm:items-center justify-between gap-3">
                  <div>
                    <span className="text-xs font-bold text-indigo-700 tracking-wider uppercase block">
                      AQUAOPS PRO • מערכת ניהול ובקרת איכות מים
                    </span>
                    <h2 className="text-xl sm:text-2xl font-black text-slate-900 mt-0.5">
                      דוח ביקורת ותקינות מתקני מים ומיכלים
                    </h2>
                  </div>
                  <div className="text-xs space-y-1 text-right sm:text-left">
                    <div>
                      <strong>תאריך הפקת דוח:</strong>{' '}
                      <span className="font-mono">{getNowStrings().fullDateTime}</span>
                    </div>
                    {reportInspectorName.trim() && (
                      <div>
                        <strong>עורך הביקורת:</strong> {reportInspectorName.trim()}
                      </div>
                    )}
                  </div>
                </div>

                {/* Summary Boxes */}
                <div className="grid grid-cols-2 sm:grid-cols-4 gap-3">
                  <div className="p-3 rounded-xl border border-slate-300 bg-slate-50">
                    <span className="text-[11px] font-bold text-slate-500 block">סה״כ אזורים</span>
                    <span className="text-xl font-black text-slate-900">{areas.length}</span>
                  </div>
                  <div className="p-3 rounded-xl border border-slate-300 bg-slate-50">
                    <span className="text-[11px] font-bold text-slate-500 block">סה״כ מתקנים ומיכלים</span>
                    <span className="text-xl font-black text-slate-900">{allDispensers.length}</span>
                  </div>
                  <div className="p-3 rounded-xl border border-emerald-300 bg-emerald-50">
                    <span className="text-[11px] font-bold text-emerald-700 block">מערכות תקינות</span>
                    <span className="text-xl font-black text-emerald-800">
                      {allDispensers.filter((d) => isLocationHealthy(d)).length}
                    </span>
                  </div>
                  <div className="p-3 rounded-xl border border-rose-300 bg-rose-50">
                    <span className="text-[11px] font-bold text-rose-700 block">חריגים / התראות</span>
                    <span className="text-xl font-black text-rose-800">{urgentAlerts.length}</span>
                  </div>
                </div>

                {/* Section 1: Urgent Exceptions */}
                <div className="space-y-2">
                  <h3 className="text-sm font-black text-slate-900 border-b border-slate-300 pb-1.5">
                    1. סיכום חריגים והתראות הדורשות טיפול ({urgentAlerts.length})
                  </h3>
                  {urgentAlerts.length === 0 ? (
                    <p className="text-xs text-emerald-700 font-semibold py-2">
                      לא נמצאו חריגים בביקורת זו — כל המתקנים, הפילטרים ותגי השמיש תקינים ובתוקף.
                    </p>
                  ) : (
                    <table className="w-full text-right border-collapse text-xs border border-slate-300">
                      <thead>
                        <tr className="bg-slate-100 border-b border-slate-300 font-bold">
                          <th className="py-2 px-3 border-l border-slate-300">אזור</th>
                          <th className="py-2 px-3 border-l border-slate-300">מיקום</th>
                          <th className="py-2 px-3 border-l border-slate-300">מתקן</th>
                          <th className="py-2 px-3">פירוט החריג / ההתראה</th>
                        </tr>
                      </thead>
                      <tbody className="divide-y divide-slate-200">
                        {urgentAlerts.map((al) => (
                          <tr key={al.id}>
                            <td className="py-2 px-3 border-l border-slate-300 font-bold">
                              {al.areaName}
                            </td>
                            <td className="py-2 px-3 border-l border-slate-300">
                              {al.locationName}
                            </td>
                            <td className="py-2 px-3 border-l border-slate-300 font-bold">
                              {al.dispenserName}
                            </td>
                            <td className="py-2 px-3 text-rose-700 font-semibold">{al.message}</td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  )}
                </div>

                {/* Section 2: Complete Inventory Table */}
                <div className="space-y-2">
                  <h3 className="text-sm font-black text-slate-900 border-b border-slate-300 pb-1.5">
                    2. טבלת סטטוס מתקנים ומיכלים מפורטת ({allDispensers.length})
                  </h3>
                  {allDispensers.length === 0 ? (
                    <p className="text-xs text-slate-500 py-2">אין מתקנים רשומים במערכת.</p>
                  ) : (
                    <table className="w-full text-right border-collapse text-[11px] border border-slate-300">
                      <thead>
                        <tr className="bg-slate-100 border-b border-slate-300 font-bold">
                          <th className="py-2 px-2 border-l border-slate-300">אזור / מיקום</th>
                          <th className="py-2 px-2 border-l border-slate-300">שם מתקן</th>
                          <th className="py-2 px-2 border-l border-slate-300">מגה</th>
                          <th className="py-2 px-2 border-l border-slate-300">אלף 1 / 2</th>
                          <th className="py-2 px-2 border-l border-slate-300">תוקף תג מתקן</th>
                          <th className="py-2 px-2 border-l border-slate-300">רגישות מיכל</th>
                          <th className="py-2 px-2 border-l border-slate-300">תוקף תג מיכל</th>
                          <th className="py-2 px-2 border-l border-slate-300">רענון אחרון</th>
                          <th className="py-2 px-2">הערות</th>
                        </tr>
                      </thead>
                      <tbody className="divide-y divide-slate-200">
                        {allDispensers.map((d) => (
                          <tr key={d.id}>
                            <td className="py-2 px-2 border-l border-slate-300">
                              <strong>{d.areaName}</strong> / {d.name}
                            </td>
                            <td className="py-2 px-2 border-l border-slate-300 font-bold">
                              {d.dispenserName}
                            </td>
                            <td className="py-2 px-2 border-l border-slate-300 font-semibold">
                              {d.filterMaga}
                            </td>
                            <td className="py-2 px-2 border-l border-slate-300 font-mono">
                              {toPercent(d.filterAlef1, 100)}% / {toPercent(d.filterAlef2, 100)}%
                            </td>
                            <td className="py-2 px-2 border-l border-slate-300 font-mono">
                              {d.tagExpiry ? formatDateIL(d.tagExpiry) : '-'}
                            </td>
                            <td className="py-2 px-2 border-l border-slate-300">
                              {d.tankSensitivity} ({d.cartridgeCarriage})
                            </td>
                            <td className="py-2 px-2 border-l border-slate-300 font-mono">
                              {d.tankTagExpiry ? formatDateIL(d.tankTagExpiry) : '-'}
                            </td>
                            <td className="py-2 px-2 border-l border-slate-300 font-mono">
                              {d.lastFilterRefreshDate ? formatDateIL(d.lastFilterRefreshDate) : '-'}
                            </td>
                            <td className="py-2 px-2">{d.notes || '-'}</td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  )}
                </div>

                {/* Official Signature Block */}
                <div className="pt-8 mt-6 border-t border-slate-300 flex items-center justify-between text-xs text-slate-600">
                  <div>
                    <strong>חתימת מבצע הביקורת:</strong> ___________________________
                  </div>
                  <div>
                    <strong>אישור אחראי תחום:</strong> ___________________________
                  </div>
                </div>
              </div>
            </div>
          </div>
        )}

        {/* CUSTOM CONFIRMATION MODAL (Delete / Reset) */}
        {confirmModal.isOpen && (
          <div className="no-print fixed inset-0 z-50 bg-slate-950/60 backdrop-blur-sm flex items-center justify-center p-4">
            <div
              className={`w-full max-w-md p-6 rounded-2xl shadow-2xl border ${
                darkMode
                  ? 'bg-[#111827] border-slate-800 text-slate-100'
                  : 'bg-white border-slate-200 text-slate-900'
              }`}
            >
              <div className="flex items-center gap-3 mb-3 text-rose-600 dark:text-rose-400">
                <div className="p-2 rounded-xl bg-rose-500/10">
                  <AlertTriangle size={20} />
                </div>
                <h3 className="text-base font-bold">{confirmModal.title}</h3>
              </div>
              <p className="text-xs sm:text-sm opacity-80 leading-relaxed mb-6">
                {confirmModal.message}
              </p>
              <div className="flex gap-2.5">
                <button
                  onClick={handleExecuteConfirm}
                  className="flex-1 bg-rose-600 hover:bg-rose-700 text-white py-2.5 rounded-xl text-xs font-bold shadow-sm transition cursor-pointer"
                >
                  אישור מחיקה
                </button>
                <button
                  onClick={() =>
                    setConfirmModal({ isOpen: false, type: null, id: null, title: '', message: '' })
                  }
                  className="px-4 bg-slate-100 dark:bg-slate-800 hover:bg-slate-200 dark:hover:bg-slate-700 py-2.5 rounded-xl text-xs font-semibold transition cursor-pointer"
                >
                  ביטול
                </button>
              </div>
            </div>
          </div>
        )}

        {/* =================================================================== */}
        {/* WORKSPACE BODY                                                      */}
        {/* =================================================================== */}
        <main className="no-print flex-1 max-w-[1520px] w-full mx-auto px-4 sm:px-6 lg:px-8 py-6">
          {/* ================================================================= */}
          {/* VIEW 1: TRACKER (Action Center + Areas Bar + Master-Detail View)  */}
          {/* ================================================================= */}
          {activeView === 'tracker' && (
            <div className="space-y-5">
              {/* SMART URGENT ACTION CENTER */}
              {allDispensers.length > 0 && renderActionCenterBanner()}

              {/* STEP 1: AREAS HORIZONTAL SAAS BAR */}
              <section className={`rounded-2xl border p-4 ${saasCard}`}>
                <div className="flex flex-wrap items-center justify-between gap-3 mb-3">
                  <div className="flex items-center gap-2">
                    <span className="w-6 h-6 rounded-lg bg-indigo-600/10 text-indigo-600 dark:text-indigo-400 flex items-center justify-center">
                      <Building2 size={14} />
                    </span>
                    <h2 className="text-xs sm:text-sm font-bold">אזורים במערכת</h2>
                    <span
                      className={`text-[11px] font-semibold px-2 py-0.5 rounded-full tabular-nums ${
                        darkMode ? 'bg-slate-800 text-slate-300' : 'bg-slate-100 text-slate-600'
                      }`}
                    >
                      {areas.length}
                    </span>
                  </div>

                  {!isAddingArea && (
                    <button
                      onClick={() => setIsAddingArea(true)}
                      className="px-3 py-1.5 rounded-xl bg-indigo-600 hover:bg-indigo-700 text-white text-xs font-semibold flex items-center gap-1.5 shadow-sm transition cursor-pointer"
                    >
                      <Plus size={14} /> הוסף אזור
                    </button>
                  )}
                </div>

                {/* Areas Pills Strip */}
                <div className="flex items-center gap-2 overflow-x-auto pb-1">
                  {areas.length === 0 && !isAddingArea && (
                    <div
                      className={`w-full py-3.5 px-4 rounded-xl border border-dashed text-center text-xs font-medium opacity-65 ${
                        darkMode ? 'border-slate-800' : 'border-slate-300'
                      }`}
                    >
                      אין אזורים מוגדרים עדיין. לחץ על &quot;+ הוסף אזור&quot; או ייבא קובץ גיבוי Excel כדי להתחיל.
                    </div>
                  )}

                  {areas.map((area) => {
                    const isSelected = selectedAreaId === area.id;
                    return (
                      <div
                        key={area.id}
                        onClick={() => {
                          setSelectedAreaId(area.id);
                          setSelectedLocationId(area.locations[0]?.id || null);
                        }}
                        className={`shrink-0 flex items-center gap-2.5 px-3.5 py-2 rounded-xl border cursor-pointer transition-all ${
                          isSelected
                            ? 'bg-indigo-600 text-white border-indigo-600 shadow-sm'
                            : darkMode
                            ? 'bg-slate-900 border-slate-800 text-slate-300 hover:border-slate-700'
                            : 'bg-slate-50 border-slate-200/90 text-slate-700 hover:bg-slate-100/80'
                        }`}
                      >
                        {editingAreaId === area.id ? (
                          <div
                            className="flex items-center gap-1.5"
                            onClick={(e) => e.stopPropagation()}
                          >
                            <input
                              value={editingAreaNameValue}
                              onChange={(e) => setEditingAreaNameValue(e.target.value)}
                              className="w-28 px-2.5 py-1 rounded-lg bg-white text-slate-900 text-xs font-semibold outline-none"
                              autoFocus
                              onKeyDown={(e) => {
                                if (e.key === 'Enter') handleSaveAreaRename(area.id);
                                if (e.key === 'Escape') setEditingAreaId(null);
                              }}
                            />
                            <button
                              onClick={() => handleSaveAreaRename(area.id)}
                              className="p-1.5 bg-indigo-800 text-white rounded-lg cursor-pointer"
                              title="שמור"
                            >
                              <Check size={12} />
                            </button>
                            <button
                              onClick={() => setEditingAreaId(null)}
                              className="p-1.5 bg-slate-600 text-white rounded-lg cursor-pointer"
                              title="ביטול"
                            >
                              <X size={12} />
                            </button>
                          </div>
                        ) : (
                          <>
                            <span className="font-bold text-xs sm:text-sm">{area.name}</span>
                            <span
                              className={`text-[11px] px-2 py-0.5 rounded-md font-semibold tabular-nums ${
                                isSelected
                                  ? 'bg-white/20 text-white'
                                  : darkMode
                                  ? 'bg-slate-800 text-slate-400'
                                  : 'bg-white text-slate-500 border border-slate-200/70'
                              }`}
                            >
                              {area.locations.length} מקומות
                            </span>

                            <div
                              className="flex items-center gap-0.5 mr-0.5"
                              onClick={(e) => e.stopPropagation()}
                            >
                              <button
                                onClick={() => {
                                  setEditingAreaId(area.id);
                                  setEditingAreaNameValue(area.name);
                                }}
                                className={`p-1 rounded-lg transition cursor-pointer ${
                                  isSelected
                                    ? 'hover:bg-white/20 text-white'
                                    : 'hover:bg-slate-200 dark:hover:bg-slate-800 text-slate-400 hover:text-slate-600'
                                }`}
                                title="שנה שם אזור"
                              >
                                <Edit3 size={13} />
                              </button>
                              <button
                                onClick={(e) => requestDeleteArea(area.id, e)}
                                className={`p-1 rounded-lg transition cursor-pointer ${
                                  isSelected
                                    ? 'hover:bg-rose-500/40 text-white'
                                    : 'hover:bg-rose-500/15 text-rose-500'
                                }`}
                                title="מחק אזור"
                              >
                                <Trash2 size={13} />
                              </button>
                            </div>
                          </>
                        )}
                      </div>
                    );
                  })}

                  {/* Inline Add Area Input */}
                  {isAddingArea && (
                    <div
                      className={`shrink-0 flex items-center gap-1.5 p-1 rounded-xl border border-indigo-500 ${
                        darkMode ? 'bg-slate-900' : 'bg-white'
                      }`}
                    >
                      <input
                        type="text"
                        placeholder="שם או מספר אזור..."
                        value={newAreaName}
                        onChange={(e) => setNewAreaName(e.target.value)}
                        onKeyDown={(e) => {
                          if (e.key === 'Enter') handleAddArea();
                          if (e.key === 'Escape') {
                            setIsAddingArea(false);
                            setNewAreaName('');
                          }
                        }}
                        className={`w-36 px-2.5 py-1.5 rounded-lg text-xs font-medium outline-none ${
                          darkMode ? 'bg-slate-800 text-white' : 'bg-slate-50 text-slate-900'
                        }`}
                        autoFocus
                      />
                      <button
                        onClick={handleAddArea}
                        className="px-3 py-1.5 bg-indigo-600 hover:bg-indigo-700 text-white rounded-lg text-xs font-semibold cursor-pointer"
                      >
                        הוסף
                      </button>
                      <button
                        onClick={() => {
                          setIsAddingArea(false);
                          setNewAreaName('');
                        }}
                        className="p-1.5 rounded-lg hover:bg-slate-200 dark:hover:bg-slate-800 cursor-pointer"
                      >
                        <X size={14} />
                      </button>
                    </div>
                  )}
                </div>
              </section>

              {/* STEP 2 & 3: SAAS MASTER-DETAIL WORKSPACE */}
              <div className="grid grid-cols-1 lg:grid-cols-12 gap-5 items-start">
                {/* MASTER PANE (4 cols): LOCATIONS LIST IN SELECTED AREA */}
                <section className={`lg:col-span-4 rounded-2xl border p-4 sm:p-5 ${saasCard}`}>
                  <div className="flex items-center justify-between pb-3.5 mb-3.5 border-b dark:border-slate-800">
                    <div className="flex items-center gap-2">
                      <span className="w-7 h-7 rounded-lg bg-sky-500/10 text-sky-600 dark:text-sky-400 flex items-center justify-center">
                        <MapPin size={15} />
                      </span>
                      <div>
                        <h2 className="text-sm font-bold">
                          מקומות {currentArea ? `באזור ${currentArea.name}` : ''}
                        </h2>
                        <span className="text-[11px] opacity-55">
                          בחר מקום לצפייה ועריכת המתקן והמיכל
                        </span>
                      </div>
                    </div>
                  </div>

                  {!currentArea ? (
                    <div className="py-12 text-center text-xs font-medium opacity-55">
                      בחר אזור למעלה כדי להציג ולנהל מקומות.
                    </div>
                  ) : (
                    <div className="space-y-2.5">
                      {currentArea.locations.length === 0 && !isAddingLocation && (
                        <div
                          className={`py-8 px-4 rounded-xl border border-dashed text-center text-xs font-medium opacity-60 ${
                            darkMode ? 'border-slate-800' : 'border-slate-300'
                          }`}
                        >
                          אין מקומות רשומים באזור &quot;{currentArea.name}&quot;.
                        </div>
                      )}

                      {currentArea.locations.map((loc) => {
                        const isSelected = selectedLocationId === loc.id;
                        const healthy = isLocationHealthy(loc);
                        const tagAlert = hasExpiredOrExpiringTag(loc);
                        return (
                          <div
                            key={loc.id}
                            onClick={() => setSelectedLocationId(loc.id)}
                            className={`p-3.5 rounded-xl border cursor-pointer transition-all flex items-center justify-between gap-2.5 ${
                              isSelected
                                ? darkMode
                                  ? 'bg-indigo-600/15 border-indigo-500/50 text-white'
                                  : 'bg-indigo-50/80 border-indigo-300 text-indigo-950'
                                : darkMode
                                ? 'bg-slate-900/70 border-slate-800 hover:border-slate-700'
                                : 'bg-white border-slate-200/80 hover:border-slate-300 hover:bg-slate-50/60'
                            }`}
                          >
                            {editingLocationId === loc.id ? (
                              <div
                                className="flex items-center gap-1.5 w-full"
                                onClick={(e) => e.stopPropagation()}
                              >
                                <input
                                  value={editingLocationNameValue}
                                  onChange={(e) => setEditingLocationNameValue(e.target.value)}
                                  className={`flex-1 min-w-0 px-2.5 py-1.5 rounded-lg border text-xs font-semibold outline-none ${saasInput}`}
                                  autoFocus
                                  onKeyDown={(e) => {
                                    if (e.key === 'Enter') handleSaveLocationRename(loc.id);
                                    if (e.key === 'Escape') setEditingLocationId(null);
                                  }}
                                />
                                <button
                                  onClick={() => handleSaveLocationRename(loc.id)}
                                  className="p-2 bg-indigo-600 text-white rounded-lg cursor-pointer"
                                  title="שמור"
                                >
                                  <Check size={13} />
                                </button>
                                <button
                                  onClick={() => setEditingLocationId(null)}
                                  className="p-2 bg-slate-500 text-white rounded-lg cursor-pointer"
                                  title="ביטול"
                                >
                                  <X size={13} />
                                </button>
                              </div>
                            ) : (
                              <>
                                <div className="flex items-center gap-3 min-w-0">
                                  <span
                                    className={`w-2.5 h-2.5 rounded-full shrink-0 ${
                                      !healthy
                                        ? 'bg-rose-500'
                                        : tagAlert
                                        ? 'bg-amber-500'
                                        : 'bg-emerald-500'
                                    }`}
                                    title={
                                      !healthy
                                        ? 'תקלה / פילטר נמוך'
                                        : tagAlert
                                        ? 'תוקף תג קרוב או פג'
                                        : 'מערכת תקינה'
                                    }
                                  />
                                  <div className="truncate">
                                    <div className="font-bold text-xs sm:text-sm truncate">
                                      {loc.name}
                                    </div>
                                    <div className="text-[11px] opacity-65 truncate">
                                      מתקן: {loc.dispenserName}
                                    </div>
                                  </div>
                                </div>

                                <div
                                  className="flex items-center gap-1 shrink-0"
                                  onClick={(e) => e.stopPropagation()}
                                >
                                  <button
                                    onClick={() => {
                                      setEditingLocationId(loc.id);
                                      setEditingLocationNameValue(loc.name);
                                    }}
                                    className="p-1.5 rounded-lg hover:bg-slate-200/80 dark:hover:bg-slate-800 text-slate-400 hover:text-slate-700 dark:hover:text-slate-200 transition cursor-pointer"
                                    title="שנה שם מקום"
                                  >
                                    <Edit3 size={14} />
                                  </button>
                                  <button
                                    onClick={(e) => requestDeleteLocation(loc.id, e)}
                                    className="p-1.5 rounded-lg hover:bg-rose-500/15 text-rose-500 transition cursor-pointer"
                                    title="מחק מקום"
                                  >
                                    <Trash2 size={14} />
                                  </button>
                                </div>
                              </>
                            )}
                          </div>
                        );
                      })}

                      {/* Inline Add Location + Dispenser Form */}
                      {isAddingLocation ? (
                        <div
                          className={`p-3.5 rounded-xl border border-indigo-500/50 space-y-2.5 ${saasPanelInset}`}
                        >
                          <input
                            type="text"
                            placeholder="שם המקום (לדוגמה: מקום - 3)..."
                            value={newLocName}
                            onChange={(e) => setNewLocName(e.target.value)}
                            className={`w-full px-3 py-2 rounded-lg border text-xs font-medium outline-none ${saasInput}`}
                            autoFocus
                          />
                          <input
                            type="text"
                            placeholder="שם פריט / מתקן מים..."
                            value={newItemName}
                            onChange={(e) => setNewItemName(e.target.value)}
                            onKeyDown={(e) => {
                              if (e.key === 'Enter') handleAddLocation();
                              if (e.key === 'Escape') {
                                setIsAddingLocation(false);
                                setNewLocName('');
                                setNewItemName('');
                              }
                            }}
                            className={`w-full px-3 py-2 rounded-lg border text-xs font-medium outline-none ${saasInput}`}
                          />
                          <div className="flex gap-2">
                            <button
                              onClick={handleAddLocation}
                              className="flex-1 bg-indigo-600 hover:bg-indigo-700 text-white py-2 rounded-lg text-xs font-semibold cursor-pointer"
                            >
                              שמור מקום
                            </button>
                            <button
                              onClick={() => {
                                setIsAddingLocation(false);
                                setNewLocName('');
                                setNewItemName('');
                              }}
                              className="px-3 bg-slate-200 dark:bg-slate-800 py-2 rounded-lg text-xs font-semibold cursor-pointer"
                            >
                              ביטול
                            </button>
                          </div>
                        </div>
                      ) : (
                        <button
                          onClick={() => setIsAddingLocation(true)}
                          className={`w-full py-2.5 rounded-xl border border-dashed text-xs font-semibold flex items-center justify-center gap-1.5 transition cursor-pointer ${
                            darkMode
                              ? 'border-slate-700 text-indigo-400 hover:bg-slate-900'
                              : 'border-slate-300 text-indigo-600 hover:bg-indigo-50/50'
                          }`}
                        >
                          <Plus size={15} /> הוסף מקום ומתקן
                        </button>
                      )}
                    </div>
                  )}
                </section>

                {/* DETAIL INSPECTOR PANE (8 cols): DISPENSER + WATER TANK + PHOTOS + NOTES */}
                <div className="lg:col-span-8">
                  {!currentLocation ? (
                    <div
                      className={`rounded-2xl border p-10 flex flex-col items-center justify-center text-center min-h-[420px] ${saasCard}`}
                    >
                      <div className="w-14 h-14 rounded-2xl bg-indigo-500/10 text-indigo-600 dark:text-indigo-400 flex items-center justify-center mb-4">
                        <Droplet size={28} />
                      </div>
                      <h3 className="text-base font-bold mb-1.5">אין מתקן מים מוצג כעת</h3>
                      <p className="text-xs opacity-65 max-w-md leading-relaxed">
                        {areas.length === 0
                          ? 'התחל ביצירת אזור חדש למעלה, או ייבא קובץ גיבוי Excel כדי לטעון נתונים קיימים.'
                          : 'בחר מקום מרשימת המקומות מימין או הוסף מקום חדש כדי לנהל את פרטי המתקן, המיכל וההערות.'}
                      </p>
                    </div>
                  ) : (
                    <div className="space-y-5">
                      {/* SaaS Inspector Action Header Card */}
                      <div
                        className={`rounded-2xl border p-4 sm:p-5 flex flex-col xl:flex-row xl:items-center justify-between gap-4 ${saasCard}`}
                      >
                        <div className="flex-1 min-w-0">
                          <label className="block text-[11px] font-semibold opacity-60 mb-1">
                            שם פריט / מתקן מים:
                          </label>
                          <input
                            type="text"
                            value={currentLocation.dispenserName || ''}
                            onChange={(e) => handleUpdateField('dispenserName', e.target.value)}
                            className={`w-full max-w-md px-3.5 py-2 rounded-xl border font-bold text-sm outline-none transition ${saasInput}`}
                          />
                        </div>

                        <div className="flex flex-wrap items-center gap-2">
                          <div
                            className={`px-3 py-2 rounded-xl border text-xs flex items-center gap-1.5 ${saasPanelInset}`}
                          >
                            <Clock size={14} className="text-emerald-600 dark:text-emerald-400" />
                            <span className="opacity-70">רענון אחרון:</span>
                            <strong className="font-mono tabular-nums">
                              {currentLocation.lastFilterRefreshDate
                                ? formatDateIL(currentLocation.lastFilterRefreshDate)
                                : 'טרם בוצע'}
                            </strong>
                          </div>

                          <button
                            onClick={openFilterRenewalModal}
                            className="px-3.5 py-2 rounded-xl bg-emerald-600 hover:bg-emerald-700 text-white text-xs font-bold flex items-center gap-1.5 shadow-sm transition cursor-pointer"
                          >
                            <Wrench size={14} /> רענון פילטרים
                          </button>

                          <button
                            onClick={() => setShowQrModal(true)}
                            className={`px-3.5 py-2 rounded-xl border text-xs font-bold flex items-center gap-1.5 transition cursor-pointer ${
                              darkMode
                                ? 'bg-indigo-500/15 border-indigo-500/30 text-indigo-300 hover:bg-indigo-500/25'
                                : 'bg-indigo-50 border-indigo-200 text-indigo-700 hover:bg-indigo-100'
                            }`}
                            title="הצג והורד קוד QR לסריקה מהירה מהשטח"
                          >
                            <QrCode size={14} /> תווית QR
                          </button>

                          <button
                            onClick={() => setActiveView('dispenser_history')}
                            className={`px-3 py-2 rounded-xl border text-xs font-semibold flex items-center gap-1.5 transition cursor-pointer ${
                              darkMode
                                ? 'bg-slate-900 border-slate-700 hover:bg-slate-800 text-slate-200'
                                : 'bg-white border-slate-200 hover:bg-slate-50 text-slate-700'
                            }`}
                          >
                            <History size={14} />
                            רענונים ({currentLocation.filterRefreshHistory?.length || 0})
                          </button>
                        </div>
                      </div>

                      {/* Two-Column SaaS Specifications Grid: Water Dispenser + Water Tank */}
                      <div className="grid grid-cols-1 xl:grid-cols-12 gap-5">
                        {/* WATER DISPENSER SPEC SHEET (7 cols) */}
                        <section
                          className={`xl:col-span-7 rounded-2xl border p-5 space-y-5 ${saasCard}`}
                        >
                          <div className="flex items-center justify-between pb-3 border-b dark:border-slate-800">
                            <div className="flex items-center gap-2">
                              <span className="w-7 h-7 rounded-lg bg-indigo-500/10 text-indigo-600 dark:text-indigo-400 flex items-center justify-center">
                                <Gauge size={16} />
                              </span>
                              <h3 className="text-sm font-bold">פרטי מתקן מים</h3>
                            </div>
                            <span className="text-[11px] font-medium opacity-55">
                              פילטרים, שעון ותוקף תג
                            </span>
                          </div>

                          {/* Percentage Filters: Filter Alef 1 & Filter Alef 2 */}
                          <div className="grid grid-cols-1 sm:grid-cols-2 gap-3.5">
                            {[
                              {
                                key: 'filterAlef1',
                                title: 'פילטר אלף 1 (%)',
                                val: toPercent(currentLocation.filterAlef1, 100)
                              },
                              {
                                key: 'filterAlef2',
                                title: 'פילטר אלף 2 (%)',
                                val: toPercent(currentLocation.filterAlef2, 100)
                              }
                            ].map((flt) => {
                              const shade = getFilterShade(flt.val);
                              return (
                                <div
                                  key={flt.key}
                                  style={{
                                    borderColor: shade.border,
                                    backgroundColor: shade.tint
                                  }}
                                  className="p-3.5 rounded-xl border space-y-2.5 transition-colors duration-200"
                                >
                                  <div className="flex items-center justify-between gap-2">
                                    <div>
                                      <span className="text-xs font-semibold block">{flt.title}</span>
                                      <span
                                        style={{ color: shade.color }}
                                        className="text-[10px] font-bold"
                                      >
                                        רמת תקינות: {shade.levelLabel}
                                      </span>
                                    </div>
                                    <div className="flex items-center gap-1">
                                      <input
                                        type="number"
                                        min="0"
                                        max="100"
                                        value={flt.val}
                                        onChange={(e) => handleUpdateField(flt.key, e.target.value)}
                                        style={{ color: shade.color, borderColor: shade.border }}
                                        className={`w-14 px-1.5 py-0.5 text-center font-mono font-bold text-xs rounded-md border outline-none tabular-nums ${saasInput}`}
                                      />
                                      <span
                                        style={{ color: shade.color }}
                                        className="text-xs font-bold"
                                      >
                                        %
                                      </span>
                                    </div>
                                  </div>

                                  {/* Shaded Percentage Scale Bar */}
                                  <div className="space-y-1">
                                    <div className="relative w-full h-2.5 rounded-full bg-slate-200/90 dark:bg-slate-800 overflow-hidden">
                                      <div
                                        className="h-full rounded-full transition-all duration-200"
                                        style={{
                                          width: `${shade.pct}%`,
                                          background: shade.gradient
                                        }}
                                      />
                                    </div>
                                    <div className="flex justify-between text-[9px] font-mono opacity-50 px-0.5">
                                      <span>0%</span>
                                      <span>50%</span>
                                      <span>100%</span>
                                    </div>
                                  </div>

                                  <input
                                    type="range"
                                    min="0"
                                    max="100"
                                    value={flt.val}
                                    onChange={(e) => handleUpdateField(flt.key, e.target.value)}
                                    style={{ accentColor: shade.color }}
                                    className="w-full cursor-pointer"
                                  />
                                </div>
                              );
                            })}
                          </div>

                          {/* Segmented Controls: Filter Maga & System Type */}
                          <div className="grid grid-cols-1 sm:grid-cols-2 gap-3.5">
                            <div className="space-y-1.5">
                              <label className="block text-xs font-semibold opacity-75">
                                פילטר מגה:
                              </label>
                              <div
                                className={`grid grid-cols-2 gap-1 p-1 rounded-xl border ${saasPanelInset}`}
                              >
                                {['תקין', 'לא תקין'].map((status) => (
                                  <button
                                    key={status}
                                    onClick={() => handleUpdateField('filterMaga', status)}
                                    className={`py-1.5 rounded-lg text-xs font-semibold transition cursor-pointer ${
                                      (currentLocation.filterMaga || 'תקין') === status
                                        ? status === 'תקין'
                                          ? 'bg-emerald-600 text-white shadow-sm'
                                          : 'bg-rose-600 text-white shadow-sm'
                                        : 'opacity-65 hover:opacity-100'
                                    }`}
                                  >
                                    {status}
                                  </button>
                                ))}
                              </div>
                            </div>

                            <div className="space-y-1.5">
                              <label className="block text-xs font-semibold opacity-75">
                                סוג מערכת (טורי / מקבילי):
                              </label>
                              <div
                                className={`grid grid-cols-2 gap-1 p-1 rounded-xl border ${saasPanelInset}`}
                              >
                                {['טורי', 'מקבילי'].map((sysType) => (
                                  <button
                                    key={sysType}
                                    onClick={() => handleUpdateField('systemType', sysType)}
                                    className={`py-1.5 rounded-lg text-xs font-semibold transition cursor-pointer ${
                                      (currentLocation.systemType || 'טורי') === sysType
                                        ? 'bg-indigo-600 text-white shadow-sm'
                                        : 'opacity-65 hover:opacity-100'
                                    }`}
                                  >
                                    {sysType}
                                  </button>
                                ))}
                              </div>
                            </div>
                          </div>

                          {/* Date & Select Inputs: Tag Expiry & Clock Type */}
                          <div className="grid grid-cols-1 sm:grid-cols-2 gap-3.5">
                            <div className="space-y-1.5">
                              <label className="block text-xs font-semibold opacity-75">
                                תוקף תג שמיש:
                              </label>
                              <input
                                type="date"
                                value={currentLocation.tagExpiry || ''}
                                onChange={(e) => handleUpdateField('tagExpiry', e.target.value)}
                                className={`w-full px-3 py-2 rounded-xl border text-xs font-semibold outline-none transition ${saasInput}`}
                              />
                            </div>

                            <div className="space-y-1.5">
                              <label className="block text-xs font-semibold opacity-75">
                                סוג שעון:
                              </label>
                              <select
                                value={currentLocation.clockType || 'דיגיטלי'}
                                onChange={(e) => handleUpdateField('clockType', e.target.value)}
                                className={`w-full px-3 py-2 rounded-xl border text-xs font-semibold outline-none transition ${saasInput}`}
                              >
                                <option value="דיגיטלי">דיגיטלי</option>
                                <option value="אנלוגי">אנלוגי</option>
                                <option value="מכני">מכני</option>
                              </select>
                            </div>
                          </div>
                        </section>

                        {/* WATER TANK + NOTES COLUMN (5 cols) */}
                        <div className="xl:col-span-5 flex flex-col gap-5">
                          {/* WATER TANK SPEC CARD */}
                          <section className={`rounded-2xl border p-5 space-y-4 ${saasCard}`}>
                            <div className="flex items-center justify-between pb-3 border-b dark:border-slate-800">
                              <div className="flex items-center gap-2">
                                <span className="w-7 h-7 rounded-lg bg-sky-500/10 text-sky-600 dark:text-sky-400 flex items-center justify-center">
                                  <Droplet size={16} />
                                </span>
                                <h3 className="text-sm font-bold">מיכל מים</h3>
                              </div>
                              <span className="text-[11px] font-medium opacity-55">
                                רגישות, עגלה ותוקף
                              </span>
                            </div>

                            {/* Tank Sensitivity */}
                            <div className="space-y-1.5">
                              <label className="block text-xs font-semibold opacity-75">
                                רגישות מיכל:
                              </label>
                              <div
                                className={`grid grid-cols-2 gap-1 p-1 rounded-xl border ${saasPanelInset}`}
                              >
                                {['תקין', 'לא תקין'].map((status) => (
                                  <button
                                    key={status}
                                    onClick={() => handleUpdateField('tankSensitivity', status)}
                                    className={`py-1.5 rounded-lg text-xs font-semibold transition cursor-pointer ${
                                      (currentLocation.tankSensitivity || 'תקין') === status
                                        ? status === 'תקין'
                                          ? 'bg-emerald-600 text-white shadow-sm'
                                          : 'bg-rose-600 text-white shadow-sm'
                                        : 'opacity-65 hover:opacity-100'
                                    }`}
                                  >
                                    {status}
                                  </button>
                                ))}
                              </div>
                            </div>

                            {/* Cartridge Carriage */}
                            <div className="space-y-1.5">
                              <label className="block text-xs font-semibold opacity-75">
                                עגלה למיכל:
                              </label>
                              <div
                                className={`grid grid-cols-2 gap-1 p-1 rounded-xl border ${saasPanelInset}`}
                              >
                                {['יש עגלה', 'אין עגלה'].map((status) => (
                                  <button
                                    key={status}
                                    onClick={() => handleUpdateField('cartridgeCarriage', status)}
                                    className={`py-1.5 rounded-lg text-xs font-semibold transition cursor-pointer ${
                                      (currentLocation.cartridgeCarriage || 'יש עגלה') === status
                                        ? 'bg-indigo-600 text-white shadow-sm'
                                        : 'opacity-65 hover:opacity-100'
                                    }`}
                                  >
                                    {status}
                                  </button>
                                ))}
                              </div>
                            </div>

                            {/* Tank Tag Expiry */}
                            <div className="space-y-1.5">
                              <label className="block text-xs font-semibold opacity-75">
                                תוקף תג שמיש למיכל:
                              </label>
                              <input
                                type="date"
                                value={currentLocation.tankTagExpiry || ''}
                                onChange={(e) => handleUpdateField('tankTagExpiry', e.target.value)}
                                className={`w-full px-3 py-2 rounded-xl border text-xs font-semibold outline-none transition ${saasInput}`}
                              />
                            </div>
                          </section>

                          {/* NOTES CARD */}
                          <section
                            className={`rounded-2xl border p-5 flex-1 flex flex-col min-h-[160px] ${saasCard}`}
                          >
                            <div className="flex items-center gap-2 mb-2.5">
                              <span className="w-6 h-6 rounded-lg bg-amber-500/10 text-amber-600 dark:text-amber-400 flex items-center justify-center">
                                <FileText size={14} />
                              </span>
                              <h3 className="text-sm font-bold">הערות</h3>
                            </div>
                            <textarea
                              value={currentLocation.notes || ''}
                              onChange={(e) => handleUpdateField('notes', e.target.value)}
                              placeholder="הוסף הערות תפעול, תקלות או דגשים עבור המתקן והמיכל..."
                              className={`flex-1 w-full p-3 rounded-xl border resize-none text-xs font-medium outline-none transition ${saasInput}`}
                            />
                          </section>
                        </div>
                      </div>

                      {/* ================================================================= */}
                      {/* MOBILE CAMERA & COMPRESSED PHOTO DOCUMENTATION CARD               */}
                      {/* ================================================================= */}
                      <section className={`rounded-2xl border p-5 space-y-4 ${saasCard}`}>
                        <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3 pb-3 border-b dark:border-slate-800">
                          <div className="flex items-center gap-2">
                            <span className="w-7 h-7 rounded-lg bg-purple-500/10 text-purple-600 dark:text-purple-400 flex items-center justify-center">
                              <Camera size={16} />
                            </span>
                            <div>
                              <h3 className="text-sm font-bold">
                                תיעוד ויזואלי (צילום תג שמיש / תקלה)
                              </h3>
                              <span className="text-[11px] opacity-60">
                                דחיסה אוטומטית לשמירה קלת משקל באופליין ({(currentLocation.photos || []).length} תמונות)
                              </span>
                            </div>
                          </div>

                          <div className="flex flex-wrap items-center gap-2">
                            <select
                              value={photoCategory}
                              onChange={(e) => setPhotoCategory(e.target.value)}
                              className={`px-3 py-2 rounded-xl border text-xs font-semibold outline-none ${saasInput}`}
                            >
                              <option value="תג שמיש מתקן">תג שמיש מתקן</option>
                              <option value="תג שמיש מיכל">תג שמיש מיכל</option>
                              <option value="תיעוד תקלה">תיעוד תקלה</option>
                              <option value="כללי">כללי</option>
                            </select>

                            <button
                              onClick={() => cameraInputRef.current?.click()}
                              disabled={isCompressingPhoto}
                              className="px-3.5 py-2 rounded-xl bg-indigo-600 hover:bg-indigo-700 disabled:opacity-50 text-white text-xs font-bold flex items-center gap-1.5 shadow-sm transition cursor-pointer"
                            >
                              <Camera size={14} />
                              <span>{isCompressingPhoto ? 'מעבד...' : 'צלם תמונה'}</span>
                            </button>

                            <button
                              onClick={() => galleryInputRef.current?.click()}
                              disabled={isCompressingPhoto}
                              className={`px-3 py-2 rounded-xl border text-xs font-semibold flex items-center gap-1.5 transition cursor-pointer ${
                                darkMode
                                  ? 'bg-slate-900 border-slate-700 hover:bg-slate-800 text-slate-200'
                                  : 'bg-white border-slate-200 hover:bg-slate-50 text-slate-700'
                              }`}
                            >
                              <ImageIcon size={14} />
                              <span>גלריה</span>
                            </button>
                          </div>
                        </div>

                        {(currentLocation.photos || []).length === 0 ? (
                          <div
                            className={`py-6 px-4 rounded-xl border border-dashed text-center text-xs font-medium opacity-60 ${
                              darkMode ? 'border-slate-800' : 'border-slate-300'
                            }`}
                          >
                            אין תמונות שמורות למתקן זה. לחץ על &quot;צלם תמונה&quot; כדי לצלם תג שמיש או תקלה ישירות מהפלאפון.
                          </div>
                        ) : (
                          <div className="grid grid-cols-2 sm:grid-cols-3 md:grid-cols-4 gap-3">
                            {(currentLocation.photos || []).map((photo) => (
                              <div
                                key={photo.id}
                                className={`group relative rounded-xl border overflow-hidden ${saasPanelInset}`}
                              >
                                <div
                                  onClick={() => setLightboxPhoto(photo)}
                                  className="aspect-square w-full bg-slate-950 overflow-hidden cursor-pointer relative"
                                >
                                  <img
                                    src={photo.dataUrl}
                                    alt={photo.category}
                                    className="w-full h-full object-cover group-hover:scale-105 transition-transform duration-200"
                                  />
                                  <div className="absolute inset-0 bg-black/0 group-hover:bg-black/30 transition flex items-center justify-center">
                                    <ZoomIn
                                      size={20}
                                      className="text-white opacity-0 group-hover:opacity-100 transition"
                                    />
                                  </div>
                                </div>
                                <div className="p-2 flex items-center justify-between gap-1">
                                  <div className="min-w-0">
                                    <span className="block text-[11px] font-bold truncate">
                                      {photo.category}
                                    </span>
                                    <span className="block text-[10px] font-mono opacity-60 truncate">
                                      {photo.date}
                                    </span>
                                  </div>
                                  <button
                                    onClick={() => handleDeletePhoto(photo.id)}
                                    className="p-1.5 rounded-lg hover:bg-rose-500/15 text-rose-500 shrink-0 cursor-pointer"
                                    title="מחק תמונה"
                                  >
                                    <Trash2 size={13} />
                                  </button>
                                </div>
                              </div>
                            ))}
                          </div>
                        )}
                      </section>
                    </div>
                  )}
                </div>
              </div>
            </div>
          )}

          {/* ================================================================= */}
          {/* VIEW 2: SAAS DASHBOARD (Action Center + KPIs + Advanced Filters)  */}
          {/* ================================================================= */}
          {activeView === 'dashboard' && (
            <div className="space-y-6">
              {/* SMART URGENT ACTION CENTER */}
              {allDispensers.length > 0 && renderActionCenterBanner()}

              {/* Top KPI Metrics Row */}
              <div className="grid grid-cols-1 sm:grid-cols-2 xl:grid-cols-4 gap-4">
                {[
                  {
                    label: 'סה״כ אזורים',
                    value: areas.length,
                    icon: Building2,
                    accent: 'text-indigo-600 dark:text-indigo-400 bg-indigo-500/10'
                  },
                  {
                    label: 'סה״כ מקומות ומתקנים',
                    value: allDispensers.length,
                    icon: Droplet,
                    accent: 'text-sky-600 dark:text-sky-400 bg-sky-500/10'
                  },
                  {
                    label: 'רענוני פילטרים שבוצעו',
                    value: allDispensers.reduce(
                      (acc, d) => acc + (d.filterRefreshHistory?.length || 0),
                      0
                    ),
                    icon: Wrench,
                    accent: 'text-emerald-600 dark:text-emerald-400 bg-emerald-500/10'
                  },
                  {
                    label: 'התראות וחריגים דחופים',
                    value: urgentAlerts.length,
                    icon: Bell,
                    accent:
                      urgentAlerts.length > 0
                        ? 'text-rose-600 dark:text-rose-400 bg-rose-500/10'
                        : 'text-emerald-600 dark:text-emerald-400 bg-emerald-500/10'
                  }
                ].map((kpi, idx) => {
                  const Icon = kpi.icon;
                  return (
                    <div
                      key={idx}
                      className={`rounded-2xl border p-5 flex items-center justify-between ${saasCard}`}
                    >
                      <div>
                        <span className="block text-xs font-semibold opacity-60">{kpi.label}</span>
                        <span className="text-2xl sm:text-3xl font-bold tabular-nums mt-1 block">
                          {kpi.value}
                        </span>
                      </div>
                      <div
                        className={`w-11 h-11 rounded-xl flex items-center justify-center ${kpi.accent}`}
                      >
                        <Icon size={21} />
                      </div>
                    </div>
                  );
                })}
              </div>

              {/* 6-Month / 1-Year Periodic Filter Renewal Tracking Table */}
              <section className={`rounded-2xl border p-5 sm:p-6 ${saasCard}`}>
                <div className="flex flex-col md:flex-row md:items-center justify-between gap-4 mb-5">
                  <div>
                    <h2 className="text-base font-bold flex items-center gap-2">
                      <Calendar size={18} className="text-emerald-600 dark:text-emerald-400" />
                      טבלת רענונים (מעקב של חצי שנה או שנה) לפי כל מתקן
                    </h2>
                    <p className="text-xs opacity-60 mt-0.5">
                      מעקב מועדי חידוש פילטרים וסטטוס תוקף לכל מתקן מים במערכת
                    </p>
                  </div>

                  <div className={`flex p-1 rounded-xl border self-start ${saasPanelInset}`}>
                    <button
                      onClick={() => setTrackingPeriod('6m')}
                      className={`px-3.5 py-1.5 rounded-lg text-xs font-semibold transition cursor-pointer ${
                        trackingPeriod === '6m'
                          ? 'bg-indigo-600 text-white shadow-sm'
                          : 'opacity-70 hover:opacity-100'
                      }`}
                    >
                      חצי שנה (6 חודשים)
                    </button>
                    <button
                      onClick={() => setTrackingPeriod('1y')}
                      className={`px-3.5 py-1.5 rounded-lg text-xs font-semibold transition cursor-pointer ${
                        trackingPeriod === '1y'
                          ? 'bg-indigo-600 text-white shadow-sm'
                          : 'opacity-70 hover:opacity-100'
                      }`}
                    >
                      שנה (12 חודשים)
                    </button>
                  </div>
                </div>

                {allDispensers.length === 0 ? (
                  <div className="py-12 text-center text-xs font-medium opacity-55">
                    אין מתקני מים רשומים במערכת להצגה בטבלת המעקב.
                  </div>
                ) : (
                  <div className="overflow-x-auto">
                    <table className="w-full text-right border-collapse min-w-[780px]">
                      <thead>
                        <tr
                          className={`border-b text-[11px] font-bold uppercase tracking-wider ${
                            darkMode
                              ? 'border-slate-800 bg-slate-900/60 text-slate-400'
                              : 'border-slate-200 bg-slate-50 text-slate-500'
                          }`}
                        >
                          <th className="py-3 px-4 rounded-r-lg">אזור</th>
                          <th className="py-3 px-4">מיקום</th>
                          <th className="py-3 px-4">שם מתקן</th>
                          <th className="py-3 px-4">תאריך רענון אחרון</th>
                          <th className="py-3 px-4">טכנאי / מלווה אחרון</th>
                          <th className="py-3 px-4">
                            מועד רענון הבא ({trackingPeriod === '6m' ? 'חצי שנה' : 'שנה'})
                          </th>
                          <th className="py-3 px-4">סה״כ רענונים</th>
                          <th className="py-3 px-4 rounded-l-lg">סטטוס</th>
                        </tr>
                      </thead>
                      <tbody className="divide-y dark:divide-slate-800 text-xs font-medium">
                        {allDispensers.map((item) => {
                          const cycle = getRenewalCycleInfo(
                            item.lastFilterRefreshDate,
                            trackingPeriod
                          );
                          const latest = item.filterRefreshHistory?.[0];
                          return (
                            <tr
                              key={item.id}
                              onClick={() => {
                                setSelectedAreaId(item.areaId);
                                setSelectedLocationId(item.id);
                                setActiveView('tracker');
                              }}
                              className="cursor-pointer hover:bg-slate-50/80 dark:hover:bg-slate-900/50 transition"
                            >
                              <td className="py-3.5 px-4 font-bold">{item.areaName}</td>
                              <td className="py-3.5 px-4">{item.name}</td>
                              <td className="py-3.5 px-4 font-bold text-indigo-600 dark:text-indigo-400">
                                {item.dispenserName}
                              </td>
                              <td className="py-3.5 px-4 font-mono tabular-nums">
                                {item.lastFilterRefreshDate
                                  ? formatDateIL(item.lastFilterRefreshDate)
                                  : 'טרם בוצע'}
                              </td>
                              <td className="py-3.5 px-4">
                                {latest ? `${latest.technician} / ${latest.escort}` : '-'}
                              </td>
                              <td className="py-3.5 px-4 font-mono font-bold tabular-nums">
                                {cycle.nextDueDate}
                              </td>
                              <td className="py-3.5 px-4 font-mono tabular-nums">
                                {item.filterRefreshHistory?.length || 0}
                              </td>
                              <td className="py-3.5 px-4">
                                <span
                                  className={`px-2.5 py-1 rounded-full border text-[11px] font-semibold inline-block ${cycle.badgeClass}`}
                                >
                                  {cycle.statusLabel}
                                </span>
                              </td>
                            </tr>
                          );
                        })}
                      </tbody>
                    </table>
                  </div>
                )}
              </section>

              {/* Full Comprehensive Dispensers Breakdown Table with Advanced Filtering & Sorting */}
              <section className={`rounded-2xl border p-5 sm:p-6 ${saasCard}`}>
                <div className="flex flex-col lg:flex-row lg:items-center justify-between gap-4 mb-4">
                  <div>
                    <h2 className="text-base font-bold flex items-center gap-2">
                      <Table size={18} className="text-indigo-600 dark:text-indigo-400" />
                      פירוט מלא של כל המתקנים במערכת ({filteredDispensers.length})
                    </h2>
                    <p className="text-xs opacity-60 mt-0.5">
                      כולל סינון חריגים, מיון לפי אחוז פילטר נמוך וסקלת גוונים השוואתית — לחץ על שורה לעריכה
                    </p>
                  </div>

                  <div className="flex flex-wrap items-center gap-2.5">
                    <button
                      onClick={() => setShowReportModal(true)}
                      className="px-3.5 py-2 rounded-xl bg-indigo-600 hover:bg-indigo-700 text-white text-xs font-bold flex items-center gap-1.5 shadow-sm transition cursor-pointer"
                    >
                      <Printer size={14} /> הפק דוח PDF / הדפסה
                    </button>

                    <div className="relative w-full sm:w-60">
                      <Search
                        size={15}
                        className="absolute right-3.5 top-1/2 -translate-y-1/2 opacity-45"
                      />
                      <input
                        type="text"
                        placeholder="חיפוש אזור, מקום או מתקן..."
                        value={dashboardSearch}
                        onChange={(e) => setDashboardSearch(e.target.value)}
                        className={`w-full pr-9 pl-3.5 py-2 rounded-xl border text-xs font-medium outline-none transition ${saasInput}`}
                      />
                    </div>
                  </div>
                </div>

                {/* ADVANCED FILTER & SORT TOOLBAR */}
                <div
                  className={`p-3 rounded-xl border mb-5 flex flex-wrap items-center justify-between gap-3 ${saasPanelInset}`}
                >
                  {/* Status Quick Filter Pills */}
                  <div className="flex flex-wrap items-center gap-1.5">
                    <span className="text-[11px] font-bold opacity-65 flex items-center gap-1 ml-1">
                      <Filter size={13} /> סינון מהיר:
                    </span>
                    {[
                      { id: 'all', label: `הכל (${allDispensers.length})` },
                      {
                        id: 'faulty',
                        label: `רק לא תקינים / נמוכים (${
                          allDispensers.filter((d) => !isLocationHealthy(d)).length
                        })`
                      },
                      {
                        id: 'expired',
                        label: `פגי תוקף / מתקרבים (${
                          allDispensers.filter((d) => hasExpiredOrExpiringTag(d)).length
                        })`
                      }
                    ].map((tab) => (
                      <button
                        key={tab.id}
                        onClick={() => setDashboardStatusFilter(tab.id)}
                        className={`px-3 py-1.5 rounded-lg text-xs font-bold transition cursor-pointer ${
                          dashboardStatusFilter === tab.id
                            ? tab.id === 'faulty'
                              ? 'bg-rose-600 text-white shadow-sm'
                              : tab.id === 'expired'
                              ? 'bg-amber-600 text-white shadow-sm'
                              : 'bg-indigo-600 text-white shadow-sm'
                            : darkMode
                            ? 'bg-slate-900 text-slate-300 hover:bg-slate-800 border border-slate-700'
                            : 'bg-white text-slate-700 hover:bg-slate-100 border border-slate-200'
                        }`}
                      >
                        {tab.label}
                      </button>
                    ))}
                  </div>

                  {/* Area Filter + Sort Dropdown + Reset */}
                  <div className="flex flex-wrap items-center gap-2">
                    <select
                      value={dashboardAreaFilter}
                      onChange={(e) => setDashboardAreaFilter(e.target.value)}
                      className={`px-3 py-1.5 rounded-lg border text-xs font-semibold outline-none ${saasInput}`}
                      title="סינון לפי אזור"
                    >
                      <option value="all">כל האזורים</option>
                      {areas.map((a) => (
                        <option key={a.id} value={a.id}>
                          אזור: {a.name}
                        </option>
                      ))}
                    </select>

                    <div className="flex items-center gap-1">
                      <ArrowUpDown size={13} className="opacity-60" />
                      <select
                        value={dashboardSortBy}
                        onChange={(e) => setDashboardSortBy(e.target.value)}
                        className={`px-3 py-1.5 rounded-lg border text-xs font-semibold outline-none ${saasInput}`}
                        title="מיון רשימה"
                      >
                        <option value="default">מיון: סדר רגיל</option>
                        <option value="lowest_filter">מיון: אחוז פילטר הנמוך ביותר</option>
                        <option value="nearest_expiry">מיון: תוקף תג הקרוב ביותר</option>
                      </select>
                    </div>

                    {(dashboardStatusFilter !== 'all' ||
                      dashboardAreaFilter !== 'all' ||
                      dashboardSortBy !== 'default' ||
                      dashboardSearch.trim() !== '') && (
                      <button
                        onClick={() => {
                          setDashboardStatusFilter('all');
                          setDashboardAreaFilter('all');
                          setDashboardSortBy('default');
                          setDashboardSearch('');
                        }}
                        className="px-2.5 py-1.5 rounded-lg bg-rose-500/10 text-rose-600 dark:text-rose-400 text-xs font-bold hover:bg-rose-500/20 cursor-pointer"
                      >
                        נקה סינון
                      </button>
                    )}
                  </div>
                </div>

                {filteredDispensers.length === 0 ? (
                  <div className="py-12 text-center text-xs font-medium opacity-55">
                    {allDispensers.length === 0
                      ? 'אין נתונים להצגה. הוסף אזורים, מקומות ומתקנים במסך המעקב.'
                      : 'לא נמצאו מתקנים התואמים לסינון הנבחר.'}
                  </div>
                ) : (
                  <div className="overflow-x-auto">
                    <table className="w-full text-right border-collapse min-w-[1100px]">
                      <thead>
                        <tr
                          className={`border-b text-[11px] font-bold uppercase tracking-wider ${
                            darkMode
                              ? 'border-slate-800 bg-slate-900/60 text-slate-400'
                              : 'border-slate-200 bg-slate-50 text-slate-500'
                          }`}
                        >
                          <th className="py-3 px-4 rounded-r-lg">אזור</th>
                          <th className="py-3 px-4">מיקום</th>
                          <th className="py-3 px-4">שם מתקן</th>
                          <th className="py-3 px-4">פילטר מגה</th>
                          <th className="py-3 px-4">אלף 1</th>
                          <th className="py-3 px-4">אלף 2</th>
                          <th className="py-3 px-4 min-w-[195px]">
                            השוואת פילטרים (אלף 1 מול אלף 2)
                          </th>
                          <th className="py-3 px-4">שעון / מערכת</th>
                          <th className="py-3 px-4">תוקף תג</th>
                          <th className="py-3 px-4">מיכל מים</th>
                          <th className="py-3 px-4">הערות</th>
                          <th className="py-3 px-4 rounded-l-lg">עדכון טבלה</th>
                        </tr>
                      </thead>
                      <tbody className="divide-y dark:divide-slate-800 text-xs font-medium">
                        {filteredDispensers.map((item) => {
                          const shade1 = getFilterShade(item.filterAlef1);
                          const shade2 = getFilterShade(item.filterAlef2);
                          const diff = Math.abs(shade1.pct - shade2.pct);
                          const dispTagDays = getDaysUntilDate(item.tagExpiry);

                          return (
                            <tr
                              key={item.id}
                              onClick={() => {
                                setSelectedAreaId(item.areaId);
                                setSelectedLocationId(item.id);
                                setActiveView('tracker');
                              }}
                              className="cursor-pointer hover:bg-indigo-50/50 dark:hover:bg-slate-900/60 transition"
                            >
                              <td className="py-3.5 px-4 font-bold">{item.areaName}</td>
                              <td className="py-3.5 px-4">{item.name}</td>
                              <td className="py-3.5 px-4 font-bold text-indigo-600 dark:text-indigo-400">
                                <span className="inline-flex items-center gap-1">
                                  {item.dispenserName}
                                  <ArrowUpRight size={13} className="opacity-50" />
                                </span>
                              </td>
                              <td className="py-3.5 px-4">
                                <span
                                  className={`px-2 py-0.5 rounded-full text-[11px] font-semibold ${
                                    (item.filterMaga || 'תקין') === 'תקין'
                                      ? 'bg-emerald-500/10 text-emerald-600 dark:text-emerald-400'
                                      : 'bg-rose-500/10 text-rose-600 dark:text-rose-400'
                                  }`}
                                >
                                  {item.filterMaga || 'תקין'}
                                </span>
                              </td>
                              <td className="py-3.5 px-4 font-mono tabular-nums">
                                <span
                                  style={{
                                    color: shade1.color,
                                    backgroundColor: shade1.tint,
                                    borderColor: shade1.border
                                  }}
                                  className="px-2 py-0.5 rounded-md border font-bold text-[11px]"
                                >
                                  {shade1.pct}%
                                </span>
                              </td>
                              <td className="py-3.5 px-4 font-mono tabular-nums">
                                <span
                                  style={{
                                    color: shade2.color,
                                    backgroundColor: shade2.tint,
                                    borderColor: shade2.border
                                  }}
                                  className="px-2 py-0.5 rounded-md border font-bold text-[11px]"
                                >
                                  {shade2.pct}%
                                </span>
                              </td>
                              {/* Comparative Shaded Scale for Filter Alef 1 vs Filter Alef 2 */}
                              <td className="py-3.5 px-4">
                                <div className={`p-2 rounded-xl border space-y-1.5 ${saasPanelInset}`}>
                                  <div className="flex items-center gap-2">
                                    <span className="text-[10px] font-bold opacity-70 w-6 shrink-0">
                                      א׳1
                                    </span>
                                    <div className="relative flex-1 h-2 rounded-full bg-slate-200 dark:bg-slate-800 overflow-hidden">
                                      <div
                                        className="h-full rounded-full transition-all duration-300"
                                        style={{
                                          width: `${shade1.pct}%`,
                                          background: shade1.gradient
                                        }}
                                      />
                                    </div>
                                    <span
                                      style={{ color: shade1.color }}
                                      className="font-mono text-[10px] font-bold w-8 text-left tabular-nums shrink-0"
                                    >
                                      {shade1.pct}%
                                    </span>
                                  </div>

                                  <div className="flex items-center gap-2">
                                    <span className="text-[10px] font-bold opacity-70 w-6 shrink-0">
                                      א׳2
                                    </span>
                                    <div className="relative flex-1 h-2 rounded-full bg-slate-200 dark:bg-slate-800 overflow-hidden">
                                      <div
                                        className="h-full rounded-full transition-all duration-300"
                                        style={{
                                          width: `${shade2.pct}%`,
                                          background: shade2.gradient
                                        }}
                                      />
                                    </div>
                                    <span
                                      style={{ color: shade2.color }}
                                      className="font-mono text-[10px] font-bold w-8 text-left tabular-nums shrink-0"
                                    >
                                      {shade2.pct}%
                                    </span>
                                  </div>

                                  <div className="flex items-center justify-between text-[10px] pt-0.5 border-t border-slate-200/70 dark:border-slate-800 opacity-75">
                                    <span>הפרש:</span>
                                    <span className="font-mono font-bold tabular-nums">
                                      {diff === 0 ? 'מאוזן (0%)' : `${diff}%`}
                                    </span>
                                  </div>
                                </div>
                              </td>
                              <td className="py-3.5 px-4">
                                {item.clockType} /{' '}
                                <span className="font-bold text-indigo-600 dark:text-indigo-400">
                                  {item.systemType}
                                </span>
                              </td>
                              <td className="py-3.5 px-4 font-mono tabular-nums">
                                <span
                                  className={
                                    dispTagDays !== null && dispTagDays < 0
                                      ? 'text-rose-600 dark:text-rose-400 font-bold'
                                      : dispTagDays !== null && dispTagDays <= 30
                                      ? 'text-amber-600 dark:text-amber-400 font-bold'
                                      : ''
                                  }
                                >
                                  {item.tagExpiry ? formatDateIL(item.tagExpiry) : '-'}
                                </span>
                              </td>
                              <td className="py-3.5 px-4">
                                <div className="space-y-0.5">
                                  <span
                                    className={`block font-semibold ${
                                      (item.tankSensitivity || 'תקין') === 'תקין'
                                        ? 'text-emerald-600 dark:text-emerald-400'
                                        : 'text-rose-600 dark:text-rose-400'
                                    }`}
                                  >
                                    {item.tankSensitivity || 'תקין'}
                                  </span>
                                  <span className="block text-[11px] opacity-65">
                                    {item.cartridgeCarriage} | תוקף:{' '}
                                    {item.tankTagExpiry ? formatDateIL(item.tankTagExpiry) : '-'}
                                  </span>
                                </div>
                              </td>
                              <td className="py-3.5 px-4 max-w-[180px] truncate opacity-75">
                                {item.notes || '-'}
                              </td>
                              <td className="py-3.5 px-4 font-mono tabular-nums text-indigo-600 dark:text-indigo-400 font-semibold">
                                {item.lastTableUpdateDate
                                  ? formatDateIL(item.lastTableUpdateDate)
                                  : '-'}
                              </td>
                            </tr>
                          );
                        })}
                      </tbody>
                    </table>
                  </div>
                )}
              </section>
            </div>
          )}

          {/* ================================================================= */}
          {/* VIEW 3: INDIVIDUAL DISPENSER FILTER REFRESH HISTORY               */}
          {/* ================================================================= */}
          {activeView === 'dispenser_history' && (
            <div className="max-w-4xl mx-auto">
              <div className={`p-6 rounded-2xl border ${saasCard}`}>
                <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4 pb-4 mb-5 border-b dark:border-slate-800">
                  <div>
                    <h2 className="text-base sm:text-lg font-bold flex items-center gap-2">
                      <Wrench size={19} className="text-emerald-600 dark:text-emerald-400" />
                      היסטוריית רענון פילטרים למתקן
                    </h2>
                    {currentLocation && (
                      <p className="text-xs opacity-65 mt-0.5">
                        מתקן: <strong>{currentLocation.dispenserName}</strong> | אזור:{' '}
                        {currentArea?.name} | מקום: {currentLocation.name}
                      </p>
                    )}
                  </div>

                  <div className="flex flex-wrap items-center gap-2">
                    {currentLocation && (
                      <button
                        onClick={openFilterRenewalModal}
                        className="px-3.5 py-2 rounded-xl bg-emerald-600 hover:bg-emerald-700 text-white text-xs font-bold flex items-center gap-1.5 shadow-sm transition cursor-pointer"
                      >
                        <Wrench size={14} /> תעד רענון חדש
                      </button>
                    )}
                    <button
                      onClick={() => setActiveView('tracker')}
                      className={`px-3.5 py-2 rounded-xl border text-xs font-semibold flex items-center gap-1.5 cursor-pointer ${
                        darkMode
                          ? 'bg-slate-900 border-slate-700 hover:bg-slate-800'
                          : 'bg-white border-slate-200 hover:bg-slate-50'
                      }`}
                    >
                      <ChevronLeft size={15} /> חזרה למתקן
                    </button>
                  </div>
                </div>

                {allDispensers.length > 1 && (
                  <div className="mb-5">
                    <label className="block text-xs font-semibold opacity-70 mb-1.5">
                      בחר מתקן לצפייה בהיסטוריית הרענונים שלו:
                    </label>
                    <select
                      value={selectedLocationId || ''}
                      onChange={(e) => {
                        const target = allDispensers.find((d) => d.id === e.target.value);
                        if (target) {
                          setSelectedAreaId(target.areaId);
                          setSelectedLocationId(target.id);
                        }
                      }}
                      className={`w-full sm:max-w-md px-3.5 py-2 rounded-xl border text-xs font-semibold outline-none transition ${saasInput}`}
                    >
                      {allDispensers.map((d) => (
                        <option key={d.id} value={d.id}>
                          אזור {d.areaName} &gt; {d.name} &gt; {d.dispenserName}
                        </option>
                      ))}
                    </select>
                  </div>
                )}

                {!currentLocation ? (
                  <div className="py-14 text-center space-y-2">
                    <AlertTriangle size={32} className="mx-auto text-amber-500 opacity-80" />
                    <p className="text-xs sm:text-sm font-semibold opacity-70">
                      לא נבחר מתקן מים כרגע. בחר אזור ומקום במסך המעקב כדי לצפות בהיסטוריה.
                    </p>
                  </div>
                ) : (currentLocation.filterRefreshHistory || []).length === 0 ? (
                  <div className="py-14 text-center text-xs font-medium opacity-55">
                    טרם בוצעו רענוני פילטרים עבור המתקן &quot;{currentLocation.dispenserName}&quot;.
                  </div>
                ) : (
                  <div className="space-y-3">
                    {currentLocation.filterRefreshHistory.map((rec) => (
                      <div
                        key={rec.id}
                        className={`p-4 rounded-xl border flex flex-col sm:flex-row sm:items-center justify-between gap-3 ${saasPanelInset}`}
                      >
                        <div className="space-y-1.5">
                          <p className="text-xs sm:text-sm font-bold text-emerald-700 dark:text-emerald-300">
                            {rec.action}
                          </p>
                          <div className="flex flex-wrap items-center gap-3 text-xs font-medium opacity-80">
                            <span className="inline-flex items-center gap-1">
                              <UserCheck
                                size={13}
                                className="text-emerald-600 dark:text-emerald-400"
                              />
                              טכנאי: <strong>{rec.technician}</strong>
                            </span>
                            <span className="inline-flex items-center gap-1">
                              <ShieldCheck
                                size={13}
                                className="text-indigo-600 dark:text-indigo-400"
                              />
                              מלווה: <strong>{rec.escort}</strong>
                            </span>
                          </div>
                        </div>

                        <span className="font-mono text-xs font-semibold px-3 py-1 rounded-lg bg-white dark:bg-slate-950 border dark:border-slate-800 shrink-0 text-emerald-600 dark:text-emerald-400 tabular-nums">
                          {rec.date}
                        </span>
                      </div>
                    ))}
                  </div>
                )}
              </div>
            </div>
          )}

          {/* ================================================================= */}
          {/* VIEW 4: GLOBAL ACTION LOG                                         */}
          {/* ================================================================= */}
          {activeView === 'history' && (
            <div className="max-w-4xl mx-auto">
              <div className={`p-6 rounded-2xl border ${saasCard}`}>
                <div className="flex items-center justify-between pb-4 mb-5 border-b dark:border-slate-800">
                  <h2 className="text-base sm:text-lg font-bold flex items-center gap-2">
                    <History size={19} className="text-indigo-600 dark:text-indigo-400" />
                    יומן פעולות ושינויים במערכת
                  </h2>
                  <button
                    onClick={() => setActiveView('tracker')}
                    className={`px-3.5 py-2 rounded-xl border text-xs font-semibold flex items-center gap-1.5 cursor-pointer ${
                      darkMode
                        ? 'bg-slate-900 border-slate-700 hover:bg-slate-800'
                        : 'bg-white border-slate-200 hover:bg-slate-50'
                    }`}
                  >
                    <ChevronLeft size={15} /> חזרה למעקב
                  </button>
                </div>

                {actionHistory.length === 0 ? (
                  <div className="py-14 text-center text-xs font-medium opacity-55">
                    אין פעולות מתועדות ביומן עדיין.
                  </div>
                ) : (
                  <div className="space-y-2.5">
                    {actionHistory.map((log) => (
                      <div
                        key={log.id}
                        className={`p-3.5 rounded-xl border flex flex-col sm:flex-row sm:items-center justify-between gap-3 ${saasPanelInset}`}
                      >
                        <div>
                          <p className="text-xs sm:text-sm font-semibold">{log.description}</p>
                          <p className="text-[11px] opacity-60 mt-0.5">מבצע: {log.user}</p>
                        </div>
                        <span className="font-mono text-xs font-semibold px-3 py-1 rounded-lg bg-white dark:bg-slate-950 border dark:border-slate-800 shrink-0 text-indigo-600 dark:text-indigo-400 tabular-nums">
                          {log.date}
                        </span>
                      </div>
                    ))}
                  </div>
                )}
              </div>
            </div>
          )}

          {/* ================================================================= */}
          {/* VIEW 5: SAAS SETTINGS & EXCEL BACKUP MANAGEMENT                   */}
          {/* ================================================================= */}
          {activeView === 'settings' && (
            <div className="max-w-2xl mx-auto space-y-5">
              {/* Excel Backup & Smart Merge Card */}
              <div className={`p-6 rounded-2xl border space-y-4 ${saasCard}`}>
                <div className="flex items-center justify-between pb-4 border-b dark:border-slate-800">
                  <h2 className="text-base sm:text-lg font-bold flex items-center gap-2">
                    <FileSpreadsheet size={20} className="text-emerald-600 dark:text-emerald-400" />
                    גיבוי, דוחות וייבוא נתונים
                  </h2>
                  <span className="text-[11px] font-bold px-2.5 py-0.5 rounded-md bg-emerald-500/10 text-emerald-600 dark:text-emerald-400">
                    XLSX / PDF
                  </span>
                </div>

                <div
                  className={`p-4 rounded-xl border text-xs leading-relaxed space-y-1.5 ${saasPanelInset}`}
                >
                  <p className="font-bold text-slate-800 dark:text-slate-100">
                    מנגנון מיזוג חכם ללא דריסה וללא כפילויות:
                  </p>
                  <ul className="list-disc list-inside opacity-75 space-y-1">
                    <li>
                      ייצוא יוצר קובץ Excel מסודר עם גיליונות נפרדים למתקנים, היסטוריית רענונים ויומן פעולות.
                    </li>
                    <li>
                      ייבוא קובץ גיבוי <strong>אינו דורס</strong> אזורים או מתקנים קיימים במערכת.
                    </li>
                    <li>
                      ניתן להפיק בכל רגע <strong>דוח ביקורת רשמי להדפסה / PDF</strong> של כל המערכת.
                    </li>
                  </ul>
                </div>

                <div className="grid grid-cols-1 sm:grid-cols-3 gap-3 pt-1">
                  <button
                    onClick={handleExportExcelBackup}
                    disabled={isBackupBusy}
                    className="py-3 px-4 rounded-xl bg-emerald-600 hover:bg-emerald-700 disabled:opacity-50 text-white text-xs font-bold flex items-center justify-center gap-2 shadow-sm transition cursor-pointer"
                  >
                    <Download size={16} />
                    <span>ייצוא Excel</span>
                  </button>

                  <button
                    onClick={() => fileInputRef.current?.click()}
                    disabled={isBackupBusy}
                    className={`py-3 px-4 rounded-xl border text-xs font-bold flex items-center justify-center gap-2 transition cursor-pointer ${
                      darkMode
                        ? 'bg-slate-900 border-slate-700 text-slate-100 hover:bg-slate-800'
                        : 'bg-white border-slate-300 text-slate-800 hover:bg-slate-50'
                    }`}
                  >
                    <Upload size={16} className="text-indigo-600 dark:text-indigo-400" />
                    <span>ייבוא Excel</span>
                  </button>

                  <button
                    onClick={() => setShowReportModal(true)}
                    className="py-3 px-4 rounded-xl bg-indigo-600 hover:bg-indigo-700 text-white text-xs font-bold flex items-center justify-center gap-2 shadow-sm transition cursor-pointer"
                  >
                    <Printer size={16} />
                    <span>הפק דוח PDF</span>
                  </button>
                </div>
              </div>

              {/* Display & System Reset Card */}
              <div className={`p-6 rounded-2xl border space-y-4 ${saasCard}`}>
                <div className="flex items-center justify-between pb-4 border-b dark:border-slate-800">
                  <h2 className="text-base sm:text-lg font-bold flex items-center gap-2">
                    <SlidersHorizontal size={19} className="text-indigo-600 dark:text-indigo-400" />
                    הגדרות תצוגה ומערכת
                  </h2>
                  <span className="text-xs font-bold px-2.5 py-0.5 rounded-md bg-indigo-500/10 text-indigo-600 dark:text-indigo-400">
                    {APP_VERSION}
                  </span>
                </div>

                <div
                  className={`p-4 rounded-xl border flex items-center justify-between gap-3 ${saasPanelInset}`}
                >
                  <div>
                    <span className="block font-bold text-sm">מצב תצוגה (בהיר / כהה)</span>
                    <span className="text-xs opacity-60">התאמת בהירות הממשק לסביבת העבודה</span>
                  </div>
                  <button
                    onClick={() => setDarkMode(!darkMode)}
                    className={`px-4 py-2 rounded-xl text-xs font-semibold flex items-center gap-2 cursor-pointer ${
                      darkMode ? 'bg-amber-400 text-slate-950' : 'bg-slate-900 text-white'
                    }`}
                  >
                    {darkMode ? <Sun size={14} /> : <Moon size={14} />}
                    {darkMode ? 'מצב בהיר' : 'מצב כהה'}
                  </button>
                </div>

                <div className="p-4 rounded-xl border border-rose-500/20 bg-rose-500/5 flex items-center justify-between gap-3">
                  <div>
                    <span className="block font-bold text-sm text-rose-600 dark:text-rose-400">
                      איפוס נתוני מערכת
                    </span>
                    <span className="text-xs opacity-65">
                      מחיקת כל האזורים, המקומות, המתקנים וההיסטוריות
                    </span>
                  </div>
                  <button
                    onClick={() =>
                      setConfirmModal({
                        isOpen: true,
                        type: 'reset_all',
                        id: null,
                        title: 'איפוס כל נתוני המערכת',
                        message:
                          'האם לאפס ולמחוק לצמיתות את כל האזורים, המקומות, המתקנים וההיסטוריות במערכת?'
                      })
                    }
                    className="px-4 py-2 rounded-xl bg-rose-600 hover:bg-rose-700 text-white text-xs font-semibold flex items-center gap-1.5 cursor-pointer"
                  >
                    <RotateCcw size={14} /> איפוס מלא
                  </button>
                </div>
              </div>
            </div>
          )}
        </main>
      </div>
    </div>
  );
}